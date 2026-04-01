import { ItemView, WorkspaceLeaf, Notice, Menu, FuzzySuggestModal, App, TFile, setIcon, requestUrl } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist/types/src/display/api';
import type PdfCanvasAiPlugin from '../main';
import { HIGHLIGHT_COLORS, COLOR_HEX } from '../types/annotations';
import type { HighlightColor, PageRect, Highlight } from '../types/annotations';
import { DEFAULT_COLOR_LABELS } from '../settings';

/** POS abbreviation → readable label for the embedded WordNet dictionary */
const POS_LABELS: Record<string, string> = {
  n: 'noun',
  v: 'verb',
  adj: 'adjective',
  adv: 'adverb',
};

interface PdfMetadata {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  creationDate?: string;
}

/** Shape returned by the free Dictionary API (dictionaryapi.dev). */
interface DictApiEntry {
  word?: string;
  phonetic?: string;
  meanings?: {
    partOfSpeech?: string;
    definitions?: { definition?: string; example?: string }[];
  }[];
}

/** Shape of the locally-cached WordNet dictionary JSON. */
type LocalDict = Record<string, [string, string][]>;

export const PDF_VIEWER_VIEW_TYPE = 'pdf-tools-viewer';

const MIN_SCALE = 0.5;
const MAX_SCALE = 4.0;
const FALLBACK_SCALE = 1.5;

export class PdfViewerView extends ItemView {
  private plugin: PdfCanvasAiPlugin;

  // PDF state
  private currentFile: TFile | null = null;
  private pdfDoc: PDFDocumentProxy | null = null;
  private currentScale = FALLBACK_SCALE;
  private renderedPages = new Set<number>();
  private pageObserver: IntersectionObserver | null = null;
  private loadGeneration = 0;
  private renderTasks = new Map<number, { cancel(): void }>();

  // Text selection state
  private selectedText = '';
  private selectedRects: PageRect[] = [];
  private selectedPageNum = 0;

  // DOM refs — library panel
  private libraryEl!: HTMLElement;
  private fileListEl!: HTMLElement;
  private annotationsEl!: HTMLElement;
  private librarySearchEl!: HTMLInputElement;

  // DOM refs — viewer panel
  private pagesEl!: HTMLElement;
  private viewportEl!: HTMLElement;
  private filenameLabelEl!: HTMLElement;
  private pageInfoEl!: HTMLElement;
  private selectionMenuEl!: HTMLElement;
  private searchBarEl!: HTMLElement;
  private searchInputEl!: HTMLInputElement;
  private searchResultsEl!: HTMLElement;
  private outlineEl!: HTMLElement;
  private outlineSectionEl!: HTMLElement;
  private dictionaryResultsEl!: HTMLElement;
  private dictionarySectionEl!: HTMLElement;

  // Search state
  private searchOpen = false;
  private searchMatches: { page: number; index: number }[] = [];
  private searchCurrentIdx = -1;

  // Scroll-based page tracking
  private pageTrackingObserver: IntersectionObserver | null = null;
  private visiblePages = new Set<number>();

  // PDF metadata
  private currentMetadata: PdfMetadata | null = null;
  private metadataEl!: HTMLElement;

  // Color filter state
  private activeColorFilter: HighlightColor | null = null;
  private colorFilterEl!: HTMLElement;

  // Cross-PDF annotation search
  private annoSearchInputEl!: HTMLInputElement;
  private annoSearchActive = false;

  // Embedded dictionary cache (loaded once from dictionary.json)
  private localDict: LocalDict | null = null;
  private localDictLoading = false;

  // Event handler refs for cleanup
  private mousedownHandler: ((e: MouseEvent) => void) | null = null;

  constructor(leaf: WorkspaceLeaf, plugin: PdfCanvasAiPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.currentScale = plugin.settings.defaultZoom ?? FALLBACK_SCALE;
  }

  getViewType(): string {
    return PDF_VIEWER_VIEW_TYPE;
  }

  getDisplayText(): string {
    return this.currentFile?.name ?? 'PDF Library';
  }

  getIcon(): string {
    return 'file-text';
  }

  async onOpen(): Promise<void> {
    await Promise.resolve();
    this.buildUI();
    this.refreshFileList();
  }

  // State persistence — allows Obsidian to open PDFs via file explorer, links, etc.
  getState(): Record<string, unknown> {
    const state = super.getState();
    if (this.currentFile) {
      state.file = this.currentFile.path;
    }
    return state;
  }

  async setState(state: unknown, result: { history: boolean }): Promise<void> {
    const s = state as Record<string, unknown>;
    if (s?.file && typeof s.file === 'string') {
      const file = this.app.vault.getAbstractFileByPath(s.file);
      if (file instanceof TFile && file.extension === 'pdf') {
        await this.loadFile(file);
      }
    }
    await super.setState(state, result);
  }

  async onClose(): Promise<void> {
    if (this.mousedownHandler) {
      document.removeEventListener('mousedown', this.mousedownHandler);
      this.mousedownHandler = null;
    }
    for (const task of this.renderTasks.values()) {
      task.cancel();
    }
    this.renderTasks.clear();
    this.pageObserver?.disconnect();
    await this.pdfDoc?.destroy();
  }

  // ─── Public API ────────────────────────────────────────────────────────────

  getCurrentFile(): TFile | null {
    return this.currentFile;
  }

  async loadFile(file: TFile): Promise<void> {
    if (this.currentFile?.path === file.path) return;

    this.hideSelectionMenu();
    this.currentFile = file;
    this.currentMetadata = null;
    const gen = ++this.loadGeneration;

    // Clear cross-PDF annotation search when loading a new file
    this.annoSearchActive = false;
    if (this.annoSearchInputEl) {
      this.annoSearchInputEl.value = '';
      const clearBtn = this.annoSearchInputEl.parentElement?.querySelector('.pcai-anno-search-clear') as HTMLElement | null;
      if (clearBtn) clearBtn.addClass('pcai-hidden');
    }

    this.renderedPages.clear();
    void this.pdfDoc?.destroy();
    this.pdfDoc = null;
    this.pageObserver?.disconnect();
    this.pagesEl.empty();
    this.filenameLabelEl.setText(file.name);
    this.metadataEl.empty();
    this.metadataEl.addClass('pcai-hidden');
    (this.leaf as unknown as { updateHeader?(): void }).updateHeader?.();

    // Highlight the active file in the library list
    this.highlightActiveFile();
    this.refreshAnnotations();

    const loadingEl = this.pagesEl.createDiv({ cls: 'pcai-pdf-loading', text: 'Loading PDF\u2026' });

    try {
      const buffer = await this.app.vault.readBinary(file);
      if (gen !== this.loadGeneration) return;

      this.pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
      if (gen !== this.loadGeneration) {
        void this.pdfDoc.destroy();
        this.pdfDoc = null;
        return;
      }
    } catch (err) {
      if (gen !== this.loadGeneration) return;
      loadingEl.setText(`Error loading PDF: ${err instanceof Error ? err.message : String(err)}`);
      console.error('PDF Tools \u2014 loadFile error:', err);
      return;
    }

    loadingEl.remove();
    this.pageInfoEl.setText(`0 / ${this.pdfDoc.numPages}`);

    // Extract PDF metadata
    try {
      const metaResult = await this.pdfDoc.getMetadata();
      if (gen !== this.loadGeneration) return;
      const info = metaResult?.info as Record<string, unknown> | undefined;
      if (info) {
        this.currentMetadata = {
          title: (info.Title as string) || undefined,
          author: (info.Author as string) || undefined,
          subject: (info.Subject as string) || undefined,
          keywords: (info.Keywords as string) || undefined,
          creationDate: (info.CreationDate as string) || undefined,
        };
        this.updateMetadataDisplay();
      }
    } catch (err) {
      console.warn('PDF Tools: metadata extraction failed:', err);
    }

    await this.createPagePlaceholders();
    if (gen !== this.loadGeneration) return;
    this.loadHighlightsForCurrentFile();
    void this.loadOutline();

    // Resume reading position if enabled
    if (this.plugin.settings.resumeLastPage) {
      const progress = this.plugin.annotationStore.getReadingProgress(file.path);
      if (progress && progress.lastPage > 1) {
        // Small delay to let placeholders render before scrolling
        setTimeout(() => this.scrollToPage(progress.lastPage), 100);
      }
    }
  }

  // ─── UI construction ───────────────────────────────────────────────────────

  private buildUI(): void {
    const root = this.containerEl.children[1] as HTMLElement;
    root.empty();
    root.addClass('pcai-pdf-root');

    // Two-panel layout: library | viewer
    const splitContainer = root.createDiv('pcai-split');

    this.buildLibraryPanel(splitContainer);
    this.buildViewerPanel(splitContainer);
    this.buildSelectionMenu(root);
    this.setupSelectionListener();
  }

  // ─── Library panel (left sidebar) ──────────────────────────────────────────

  private buildLibraryPanel(parent: HTMLElement): void {
    this.libraryEl = parent.createDiv('pcai-library');

    // ── File list section ──
    const { content: filesContent } = this.createCollapsibleSection(
      this.libraryEl, 'PDFs', 'pcai-files-section',
    );

    this.librarySearchEl = filesContent.createEl('input', {
      cls: 'pcai-library-search',
      attr: { type: 'text', placeholder: 'Filter\u2026' },
    });
    this.librarySearchEl.addEventListener('input', () => this.refreshFileList());

    this.fileListEl = filesContent.createDiv('pcai-file-list');

    // ── Outline / TOC section ──
    const { section: outlineSection, content: outlineContent } = this.createCollapsibleSection(
      this.libraryEl, 'Outline', 'pcai-outline-section', true,
    );
    this.outlineSectionEl = outlineSection;
    this.outlineEl = outlineContent;

    // ── Dictionary section ──
    const { section: dictSection, content: dictContent } = this.createCollapsibleSection(
      this.libraryEl, 'Dictionary', 'pcai-dictionary-section', true,
    );
    this.dictionarySectionEl = dictSection;

    const dictInput = dictContent.createEl('input', {
      cls: 'pcai-dict-search',
      attr: { type: 'text', placeholder: 'Look up a word\u2026' },
    });
    dictInput.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        const word = dictInput.value.trim();
        if (word) void this.lookupWord(word);
      }
    });
    this.dictionaryResultsEl = dictContent.createDiv('pcai-dict-results');

    // ── Annotations section ──
    const { content: annoContent } = this.createCollapsibleSection(
      this.libraryEl, 'Annotations', 'pcai-annotations-section',
    );

    // Export + Summarize buttons row
    const annoActionsRow = annoContent.createDiv({ cls: 'pcai-anno-header-actions' });

    const exportBtn = annoActionsRow.createEl('button', {
      cls: 'pcai-icon-btn pcai-anno-header-btn clickable-icon',
      attr: { 'aria-label': 'Export annotations to Markdown' },
    });
    setIcon(exportBtn, 'download');
    exportBtn.addEventListener('click', () => void this.exportAnnotations());

    if (this.plugin.settings.enableAi) {
      const summarizeBtn = annoActionsRow.createEl('button', {
        cls: 'pcai-icon-btn pcai-anno-header-btn clickable-icon',
        attr: { 'aria-label': 'Summarize annotations with AI' },
      });
      setIcon(summarizeBtn, 'sparkles');
      summarizeBtn.addEventListener('click', () => void this.summarizeAnnotations());
    }

    // Cross-PDF annotation search
    const annoSearchRow = annoContent.createDiv({ cls: 'pcai-anno-search-row' });
    this.annoSearchInputEl = annoSearchRow.createEl('input', {
      cls: 'pcai-anno-search-input',
      attr: { type: 'text', placeholder: 'Search all annotations\u2026' },
    });
    this.annoSearchInputEl.addEventListener('input', () => this.onAnnoSearchInput());

    const annoSearchClearBtn = annoSearchRow.createEl('button', {
      cls: 'pcai-icon-btn pcai-anno-search-clear clickable-icon',
      attr: { 'aria-label': 'Clear search' },
    });
    setIcon(annoSearchClearBtn, 'x');
    annoSearchClearBtn.addClass('pcai-hidden');
    annoSearchClearBtn.addEventListener('click', () => {
      this.annoSearchInputEl.value = '';
      annoSearchClearBtn.addClass('pcai-hidden');
      this.annoSearchActive = false;
      this.refreshAnnotations();
    });

    // Color filter row
    this.colorFilterEl = annoContent.createDiv('pcai-color-filter-row');
    this.buildColorFilterRow();

    this.annotationsEl = annoContent.createDiv();
    this.annotationsEl.addClass('pcai-annotation-list');
    this.annotationsEl.createDiv({
      cls: 'pcai-anno-empty',
      text: 'Open a PDF to see annotations',
    });
  }

  private createCollapsibleSection(
    parent: HTMLElement,
    title: string,
    extraClass?: string,
    startCollapsed = false,
  ): { section: HTMLElement; content: HTMLElement } {
    const section = parent.createDiv({ cls: `pcai-library-section ${extraClass ?? ''}` });
    if (startCollapsed) section.addClass('pcai-collapsed');

    const header = section.createDiv({ cls: 'pcai-library-section-header pcai-collapsible' });
    const chevron = header.createSpan({ cls: 'pcai-collapse-chevron' });
    setIcon(chevron, 'chevron-right');
    header.createSpan({ text: title });

    header.addEventListener('click', () => {
      section.toggleClass('pcai-collapsed', !section.hasClass('pcai-collapsed'));
    });

    const content = section.createDiv({ cls: 'pcai-section-content' });
    return { section, content };
  }

  private buildColorFilterRow(): void {
    this.colorFilterEl.empty();
    const labels = this.plugin.settings.colorLabels ?? DEFAULT_COLOR_LABELS;

    // "All" button
    const allBtn = this.colorFilterEl.createEl('button', {
      cls: `pcai-color-filter-btn ${this.activeColorFilter === null ? 'pcai-color-filter-active' : ''}`,
      text: 'All',
    });
    allBtn.addEventListener('click', () => {
      this.activeColorFilter = null;
      this.buildColorFilterRow();
      this.refreshAnnotations();
    });

    for (const color of HIGHLIGHT_COLORS) {
      const btn = this.colorFilterEl.createEl('button', {
        cls: `pcai-color-filter-btn ${this.activeColorFilter === color ? 'pcai-color-filter-active' : ''}`,
        attr: { title: labels[color] ?? color },
      });
      const dot = btn.createSpan({ cls: 'pcai-color-filter-dot' });
      dot.setCssStyles({ backgroundColor: COLOR_HEX[color] });
      btn.addEventListener('click', () => {
        this.activeColorFilter = this.activeColorFilter === color ? null : color;
        this.buildColorFilterRow();
        this.refreshAnnotations();
      });
    }
  }

  private refreshFileList(): void {
    this.fileListEl.empty();
    const filter = this.librarySearchEl?.value?.toLowerCase() ?? '';

    const pdfFiles = this.app.vault
      .getFiles()
      .filter((f) => f.extension === 'pdf')
      .filter((f) => !filter || f.path.toLowerCase().includes(filter))
      .sort((a, b) => a.path.localeCompare(b.path));

    if (pdfFiles.length === 0) {
      this.fileListEl.createDiv({
        cls: 'pcai-file-empty',
        text: filter ? 'No matching PDFs' : 'No PDFs in vault',
      });
      return;
    }

    for (const file of pdfFiles) {
      const item = this.fileListEl.createDiv({
        cls: 'pcai-file-item-wrapper',
        attr: { 'data-path': file.path },
      });

      const row = item.createDiv({
        cls: 'pcai-file-item',
      });

      // Show annotation count as a subtle badge
      const annotations = this.plugin.annotationStore.getForFile(file.path);
      const hasAnnotations = annotations.length > 0;

      row.createSpan({ cls: 'pcai-file-name', text: file.basename });
      if (hasAnnotations) {
        row.createSpan({
          cls: 'pcai-file-badge',
          text: String(annotations.length),
        });
      }

      if (this.currentFile?.path === file.path) {
        item.addClass('pcai-file-active');
      }

      row.addEventListener('click', () => {
        void this.loadFile(file).catch((e: unknown) =>
          console.error('PDF Tools \u2014 loadFile error:', e),
        );
      });

      // Reading progress bar
      const progress = this.plugin.annotationStore.getReadingProgress(file.path);
      if (progress && progress.totalPages > 0) {
        const pct = Math.min(100, Math.round((progress.lastPage / progress.totalPages) * 100));
        const progressBar = item.createDiv({ cls: 'pcai-file-progress' });
        const fill = progressBar.createDiv({ cls: 'pcai-file-progress-fill' });
        fill.setCssStyles({ width: `${pct}%` });
        progressBar.title = `${pct}% read (page ${progress.lastPage}/${progress.totalPages})`;
      }

      // Tooltip with full path
      item.title = file.path;
    }
  }

  private highlightActiveFile(): void {
    this.fileListEl.querySelectorAll('.pcai-file-item-wrapper').forEach((el) => {
      el.removeClass('pcai-file-active');
      if (el.getAttribute('data-path') === this.currentFile?.path) {
        el.addClass('pcai-file-active');
      }
    });
  }

  // ─── Annotations panel ─────────────────────────────────────────────────────

  refreshAnnotations(): void {
    // If cross-PDF search is active, don't overwrite search results
    if (this.annoSearchActive) return;

    this.annotationsEl.empty();

    if (!this.currentFile) {
      this.annotationsEl.createDiv({
        cls: 'pcai-anno-empty',
        text: 'Open a PDF to see annotations',
      });
      return;
    }

    let highlights = this.plugin.annotationStore.getForFile(this.currentFile.path);

    if (highlights.length === 0) {
      this.annotationsEl.createDiv({
        cls: 'pcai-anno-empty',
        text: 'No annotations yet. Select text to highlight.',
      });
      return;
    }

    // Apply color filter
    if (this.activeColorFilter) {
      highlights = highlights.filter((h) => h.color === this.activeColorFilter);
      if (highlights.length === 0) {
        const label = this.plugin.settings.colorLabels?.[this.activeColorFilter] ?? this.activeColorFilter;
        this.annotationsEl.createDiv({
          cls: 'pcai-anno-empty',
          text: `No "${label}" highlights yet.`,
        });
        return;
      }
    }

    // Group by page
    const byPage = new Map<number, Highlight[]>();
    for (const h of highlights) {
      const arr = byPage.get(h.pageNumber) ?? [];
      arr.push(h);
      byPage.set(h.pageNumber, arr);
    }

    const sortedPages = [...byPage.keys()].sort((a, b) => a - b);

    for (const page of sortedPages) {
      const pageGroup = this.annotationsEl.createDiv('pcai-anno-page-group');
      pageGroup.createDiv({
        cls: 'pcai-anno-page-label',
        text: `Page ${page}`,
      });

      const pageHighlights = byPage.get(page)!;
      for (const h of pageHighlights) {
        const card = pageGroup.createDiv('pcai-anno-card');
        card.addEventListener('click', (e) => {
          // Don't navigate if clicking an action button
          if ((e.target as HTMLElement).closest('.pcai-anno-actions')) return;
          this.scrollToPage(h.pageNumber);
        });

        // Color dot + label + text
        const row = card.createDiv('pcai-anno-row');
        const dot = row.createSpan('pcai-anno-dot');
        dot.setCssProps({ '--dot-color': COLOR_HEX[h.color] });
        const colorLabel = this.plugin.settings.colorLabels?.[h.color] ?? DEFAULT_COLOR_LABELS[h.color];
        dot.title = colorLabel;

        const textEl = row.createSpan({
          cls: 'pcai-anno-text',
          text: h.text.length > 120 ? h.text.slice(0, 120) + '\u2026' : h.text,
        });
        textEl.title = h.text;

        // Note indicator
        if (h.notePath) {
          const noteLink = card.createDiv({ cls: 'pcai-anno-note pcai-anno-note-link' });
          const basename = h.notePath.split('/').pop()?.replace(/\.md$/, '') ?? h.notePath;
          noteLink.setText(basename);
          noteLink.title = 'Open note';
          noteLink.addEventListener('click', (e) => {
            e.stopPropagation();
            void this.editAnnotationNote(h);
          });
        } else if (h.note) {
          card.createDiv({ cls: 'pcai-anno-note', text: h.note });
        }

        // Actions row
        const actions = card.createDiv('pcai-anno-actions');

        const gotoBtn = actions.createEl('button', {
          cls: 'pcai-anno-action',
          text: 'Go to',
        });
        gotoBtn.addEventListener('click', () => this.scrollToPage(h.pageNumber));

        if (this.plugin.settings.enableAi) {
          const askBtn = actions.createEl('button', {
            cls: 'pcai-anno-action',
            text: 'Ask AI',
          });
          askBtn.addEventListener('click', () => void this.askAboutHighlight(h));
        }

        const hasNote = Boolean(h.notePath || h.note);
        const noteBtn = actions.createEl('button', {
          cls: 'pcai-anno-action',
          text: hasNote ? 'Open note' : 'Add note',
        });
        noteBtn.addEventListener('click', () => void this.editAnnotationNote(h));

        const delBtn = actions.createEl('button', {
          cls: 'pcai-anno-action pcai-anno-action--danger',
          text: 'Delete',
        });
        delBtn.addEventListener('click', () => {
          this.plugin.annotationStore.remove(h.id);
          this.refreshAnnotations();
          this.applyHighlightsToPage(h.pageNumber);
          this.refreshFileList();
        });
      }
    }
  }

  private async editAnnotationNote(h: Highlight): Promise<void> {
    // If a note file already exists, open it
    if (h.notePath) {
      const existing = this.app.vault.getAbstractFileByPath(h.notePath);
      if (existing instanceof TFile) {
        await this.app.workspace.getLeaf('tab').openFile(existing);
        return;
      }
      // File was deleted externally — clear the stale path and fall through
      this.plugin.annotationStore.updateNotePath(h.id, '');
    }

    // Determine output path: same directory as the PDF
    const pdfDir = this.currentFile?.parent?.path ?? '';
    const pdfBasename = this.currentFile?.basename ?? 'PDF';
    const shortId = h.id.slice(0, 8);
    const noteFilename = `${pdfBasename} - Note - ${shortId}`;
    const notePath = pdfDir ? `${pdfDir}/${noteFilename}.md` : `${noteFilename}.md`;

    const labels = this.plugin.settings.colorLabels ?? DEFAULT_COLOR_LABELS;
    const colorLabel = labels[h.color] ?? h.color;

    // If there's an existing inline note (old format), include it as initial content
    const existingNote = h.note?.trim() ?? '';

    const lines: string[] = [
      '---',
      `annotation-target: "[[${this.currentFile?.path ?? h.pdfPath}]]"`,
      `annotation-id: ${h.id}`,
      `page: ${h.pageNumber}`,
      `highlight-color: ${h.color}`,
      `highlight-label: ${colorLabel}`,
      `created: ${new Date(h.createdAt).toISOString().split('T')[0]}`,
      '---',
      '',
      `> [!quote] Highlight from page ${h.pageNumber}`,
      `> ${h.text}`,
      '',
      '## Note',
      '',
    ];

    if (existingNote) {
      lines.push(existingNote);
    }
    lines.push('');

    try {
      const file = await this.app.vault.create(notePath, lines.join('\n'));
      this.plugin.annotationStore.updateNotePath(h.id, notePath);
      // Clear old inline note if it was migrated
      if (existingNote) {
        this.plugin.annotationStore.updateNote(h.id, '');
      }
      this.refreshAnnotations();
      await this.app.workspace.getLeaf('tab').openFile(file);
    } catch (err) {
      console.error('PDF Tools: note creation error:', err);
      new Notice(`Failed to create note: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  private scrollToPage(pageNum: number): void {
    const targetEl = this.pagesEl.querySelector(`[data-page="${pageNum}"]`);
    if (targetEl) {
      targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
      if (this.pdfDoc) {
        this.pageInfoEl.setText(`${pageNum} / ${this.pdfDoc.numPages}`);
      }
    }
  }

  // ─── Viewer panel (right side) ─────────────────────────────────────────────

  private buildViewerPanel(parent: HTMLElement): void {
    const viewerPanel = parent.createDiv('pcai-viewer-panel');

    this.buildToolbar(viewerPanel);
    this.buildSearchBar(viewerPanel);
    this.buildViewport(viewerPanel);
  }

  private buildToolbar(root: HTMLElement): void {
    const bar = root.createDiv('pcai-pdf-toolbar');

    const openBtn = bar.createEl('button', {
      cls: 'pcai-icon-btn pcai-open-btn clickable-icon',
      attr: { 'aria-label': 'Open PDF from vault' },
    });
    setIcon(openBtn, 'folder-open');
    openBtn.addEventListener('click', () => this.openFilePicker());

    // Filename + metadata container
    const nameBlock = bar.createDiv({ cls: 'pcai-pdf-name-block' });
    this.filenameLabelEl = nameBlock.createSpan({ cls: 'pcai-pdf-filename', text: 'Select a PDF' });
    this.metadataEl = nameBlock.createDiv({ cls: 'pcai-pdf-metadata' });
    this.metadataEl.addClass('pcai-hidden');

    const navGroup = bar.createDiv('pcai-pdf-nav');

    const prevBtn = navGroup.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Previous page' } });
    setIcon(prevBtn, 'chevron-left');
    prevBtn.addEventListener('click', () => this.navigatePage(-1));

    this.pageInfoEl = navGroup.createSpan({ cls: 'pcai-pdf-page-info pcai-page-info-clickable', text: '\u2014' });
    this.pageInfoEl.setAttribute('title', 'Click to jump to page');
    this.pageInfoEl.addEventListener('click', () => this.showPageJumpInput());

    const nextBtn = navGroup.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Next page' } });
    setIcon(nextBtn, 'chevron-right');
    nextBtn.addEventListener('click', () => this.navigatePage(1));

    const searchBtn = navGroup.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Search (Ctrl+F)' } });
    setIcon(searchBtn, 'search');
    searchBtn.addEventListener('click', () => this.toggleSearch());

    const zoomGroup = bar.createDiv('pcai-pdf-zoom');

    const zoomOutBtn = zoomGroup.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Zoom out' } });
    setIcon(zoomOutBtn, 'minus');
    zoomOutBtn.addEventListener('click', () => this.adjustZoom(-0.25));

    const zoomInBtn = zoomGroup.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Zoom in' } });
    setIcon(zoomInBtn, 'plus');
    zoomInBtn.addEventListener('click', () => this.adjustZoom(0.25));

    const fitBtn = zoomGroup.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Fit to width' } });
    setIcon(fitBtn, 'maximize-2');
    fitBtn.addEventListener('click', () => this.fitToWidth());

    const outlineBtn = bar.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Table of contents' } });
    setIcon(outlineBtn, 'list');
    outlineBtn.addEventListener('click', () => this.toggleOutline());

    const canvasBtn = bar.createEl('button', { cls: 'pcai-icon-btn clickable-icon', attr: { 'aria-label': 'Add to canvas' } });
    setIcon(canvasBtn, 'layout-dashboard');
    canvasBtn.addEventListener('click', () => this.addCurrentPdfToCanvas());

    if (this.plugin.settings.enableAi) {
      const aiBtn = bar.createEl('button', { cls: 'pcai-ask-btn mod-cta', text: 'Ask AI' });
      aiBtn.addEventListener('click', () => void this.askAboutCurrentPdf());
    }
  }

  private buildViewport(root: HTMLElement): void {
    this.viewportEl = root.createDiv('pcai-pdf-viewport');
    this.pagesEl = this.viewportEl.createDiv('pcai-pdf-pages');

    // Track current page on scroll
    this.viewportEl.addEventListener('scroll', () => this.updateCurrentPageOnScroll());

    // Keyboard shortcuts
    this.viewportEl.setAttribute('tabindex', '0');
    this.viewportEl.addEventListener('keydown', (e: KeyboardEvent) => this.handleKeyboard(e));
  }

  private buildSelectionMenu(root: HTMLElement): void {
    this.selectionMenuEl = root.createDiv('pcai-sel-menu');
    this.selectionMenuEl.addClass('pcai-hidden');

    const colors = this.selectionMenuEl.createDiv('pcai-sel-colors');
    const defaultColor = this.plugin.settings.defaultHighlightColor ?? 'yellow';
    for (const color of HIGHLIGHT_COLORS) {
      const dot = colors.createEl('button', {
        cls: `pcai-sel-dot ${color === defaultColor ? 'pcai-sel-dot--default' : ''}`,
        attr: { title: `Highlight ${color}${color === defaultColor ? ' (default)' : ''}` },
      });
      dot.setCssProps({ '--dot-color': COLOR_HEX[color] });
      dot.addEventListener('mousedown', (e) => {
        e.preventDefault();
        this.createHighlight(color);
      });
    }

    this.selectionMenuEl.createDiv('pcai-sel-divider');

    const actions = this.selectionMenuEl.createDiv('pcai-sel-actions');

    if (this.plugin.settings.enableAi) {
      const askBtn = actions.createEl('button', { cls: 'pcai-sel-btn pcai-sel-btn--accent', text: 'Ask AI' });
      askBtn.addEventListener('mousedown', (e) => {
        e.preventDefault();
        void this.askAboutSelection();
      });
    }

    const copyBtn = actions.createEl('button', { cls: 'pcai-sel-btn', text: 'Copy' });
    copyBtn.addEventListener('mousedown', (e) => {
      e.preventDefault();
      void navigator.clipboard.writeText(this.selectedText).catch((err: unknown) => {
        new Notice('Copy failed.');
        console.error('PDF Tools \u2014 clipboard error:', err);
      });
      this.hideSelectionMenu();
      window.getSelection()?.removeAllRanges();
    });

    const defineBtn = actions.createEl('button', { cls: 'pcai-sel-btn', text: 'Define' });
    defineBtn.addEventListener('mousedown', (e) => {
      e.preventDefault();
      const text = this.selectedText.trim();
      if (text) void this.lookupWord(text);
      this.hideSelectionMenu();
      window.getSelection()?.removeAllRanges();
    });

    this.mousedownHandler = (e: MouseEvent) => {
      if (!this.selectionMenuEl.contains(e.target as Node)) {
        this.hideSelectionMenu();
      }
    };
    document.addEventListener('mousedown', this.mousedownHandler);
  }

  // ─── PDF rendering ─────────────────────────────────────────────────────────

  private async createPagePlaceholders(): Promise<void> {
    if (!this.pdfDoc) return;

    this.pageObserver?.disconnect();

    this.pageObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const num = parseInt(entry.target.getAttribute('data-page') ?? '0', 10);
          if (num > 0 && !this.renderedPages.has(num)) {
            this.renderedPages.add(num);
            void this.renderPage(num).catch((e: unknown) => {
              console.error('PDF Tools \u2014 renderPage error:', e);
            });
          }
        }
      },
      { root: this.pagesEl.parentElement, rootMargin: '300px', threshold: 0 },
    );

    for (let i = 1; i <= this.pdfDoc.numPages; i++) {
      const page = await this.pdfDoc.getPage(i);
      const vp = page.getViewport({ scale: this.currentScale });

      const wrapper = this.pagesEl.createDiv({
        cls: 'pcai-page-wrapper pcai-page-placeholder',
        attr: { 'data-page': String(i) },
      });
      wrapper.setCssStyles({ width: `${vp.width}px`, height: `${vp.height}px` });

      this.pageObserver.observe(wrapper);
    }
  }

  private async renderPage(pageNum: number): Promise<void> {
    if (!this.pdfDoc) return;

    this.renderTasks.get(pageNum)?.cancel();

    const wrapper = this.pagesEl.querySelector(`[data-page="${pageNum}"]`);
    if (!wrapper) return;
    wrapper.removeClass('pcai-page-placeholder');

    const page = await this.pdfDoc.getPage(pageNum);
    const vp = page.getViewport({ scale: this.currentScale });

    const dpr = window.devicePixelRatio || 1;
    const canvas = wrapper.createEl('canvas', { cls: 'pcai-page-canvas' });
    canvas.width = Math.round(vp.width * dpr);
    canvas.height = Math.round(vp.height * dpr);
    canvas.setCssStyles({ width: `${vp.width}px`, height: `${vp.height}px` });
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    if (dpr !== 1) ctx.scale(dpr, dpr);

    const renderTask = page.render({ canvasContext: ctx, viewport: vp });
    this.renderTasks.set(pageNum, renderTask);
    try {
      await renderTask.promise;
    } catch (err) {
      if ((err as { name?: string }).name === 'RenderingCancelledException') return;
      throw err;
    } finally {
      this.renderTasks.delete(pageNum);
    }

    const textLayerEl = wrapper.createDiv({ cls: 'textLayer' });
    textLayerEl.setCssStyles({ width: `${vp.width}px`, height: `${vp.height}px` });
    textLayerEl.setCssProps({ '--scale-factor': String(this.currentScale) });

    const textContent = await page.getTextContent();
    const textRenderTask = pdfjsLib.renderTextLayer({
      textContentSource: textContent,
      container: textLayerEl,
      viewport: vp,
      textDivs: [],
    });
    await textRenderTask.promise;

    wrapper.createDiv({ cls: 'pcai-annotation-layer' });

    if (pageNum === 1) {
      this.pageInfoEl.setText(`1 / ${this.pdfDoc.numPages}`);
    }

    this.applyHighlightsToPage(pageNum);
  }

  private async rerenderAllPages(): Promise<void> {
    if (!this.pdfDoc) return;

    for (const task of this.renderTasks.values()) {
      task.cancel();
    }
    this.renderTasks.clear();
    this.renderedPages.clear();
    this.pageObserver?.disconnect();

    // Clear all page elements and rebuild from scratch
    this.pagesEl.empty();
    await this.createPagePlaceholders();
    this.loadHighlightsForCurrentFile();
  }

  // ─── Navigation and zoom ───────────────────────────────────────────────────

  private navigatePage(delta: number): void {
    if (!this.pdfDoc) return;
    const wrappers = Array.from(this.pagesEl.querySelectorAll<HTMLElement>('[data-page]'));
    const viewportEl = this.pagesEl.parentElement!;
    const viewportCenter = viewportEl.scrollTop + viewportEl.clientHeight / 2;

    let currentPage = 1;
    for (const wrapper of wrappers) {
      const top = wrapper.offsetTop;
      if (top <= viewportCenter) currentPage = parseInt(wrapper.getAttribute('data-page')!, 10);
    }

    const target = Math.max(1, Math.min(this.pdfDoc.numPages, currentPage + delta));
    const targetEl = this.pagesEl.querySelector(`[data-page="${target}"]`);
    if (targetEl) targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    this.pageInfoEl.setText(`${target} / ${this.pdfDoc.numPages}`);
  }

  private adjustZoom(delta: number): void {
    const newScale = Math.max(MIN_SCALE, Math.min(MAX_SCALE, this.currentScale + delta));
    if (newScale === this.currentScale) return;
    this.currentScale = newScale;
    void this.rerenderAllPages().catch((e: unknown) => console.error('PDF Tools \u2014 zoom error:', e));
  }

  private fitToWidth(): void {
    if (!this.pdfDoc) return;
    const viewportWidth = (this.pagesEl.parentElement?.clientWidth ?? 600) - 32;
    void this.pdfDoc.getPage(1).then((page) => {
      const naturalWidth = page.getViewport({ scale: 1 }).width;
      this.currentScale = viewportWidth / naturalWidth;
      void this.rerenderAllPages().catch((e: unknown) => console.error('PDF Tools \u2014 fitToWidth error:', e));
    }).catch((e: unknown) => console.error('PDF Tools \u2014 fitToWidth error:', e));
  }

  // ─── Current page tracking on scroll ────────────────────────────────────────

  private updateCurrentPageOnScroll(): void {
    if (!this.pdfDoc) return;
    const wrappers = Array.from(this.pagesEl.querySelectorAll<HTMLElement>('[data-page]'));
    const scrollTop = this.viewportEl.scrollTop;
    const viewportCenter = scrollTop + this.viewportEl.clientHeight / 2;

    let currentPage = 1;
    for (const wrapper of wrappers) {
      if (wrapper.offsetTop <= viewportCenter) {
        currentPage = parseInt(wrapper.getAttribute('data-page') ?? '1', 10);
      }
    }
    this.pageInfoEl.setText(`${currentPage} / ${this.pdfDoc.numPages}`);

    // Update reading progress
    if (this.currentFile) {
      this.plugin.annotationStore.updateReadingProgress(
        this.currentFile.path,
        currentPage,
        this.pdfDoc.numPages,
      );
    }
  }

  // ─── Page jump ────────────────────────────────────────────────────────────

  private showPageJumpInput(): void {
    if (!this.pdfDoc) return;

    const current = this.pageInfoEl.getText().split('/')[0]?.trim() ?? '1';
    const input = document.createElement('input');
    input.type = 'number';
    input.value = current;
    input.min = '1';
    input.max = String(this.pdfDoc.numPages);
    input.className = 'pcai-page-jump-input';

    this.pageInfoEl.empty();
    this.pageInfoEl.appendChild(input);
    input.focus();
    input.select();

    const commit = () => {
      const val = parseInt(input.value, 10);
      if (!isNaN(val) && val >= 1 && val <= this.pdfDoc!.numPages) {
        this.scrollToPage(val);
      } else {
        this.updateCurrentPageOnScroll();
      }
    };

    input.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        commit();
      } else if (e.key === 'Escape') {
        this.updateCurrentPageOnScroll();
      }
    });
    input.addEventListener('blur', () => commit());
  }

  // ─── Keyboard shortcuts ───────────────────────────────────────────────────

  private handleKeyboard(e: KeyboardEvent): void {
    // Don't capture when typing in an input
    if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;

    if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
      e.preventDefault();
      this.toggleSearch();
      return;
    }

    switch (e.key) {
      case '+':
      case '=':
        e.preventDefault();
        this.adjustZoom(0.25);
        break;
      case '-':
        e.preventDefault();
        this.adjustZoom(-0.25);
        break;
      case 'ArrowLeft':
        if (!e.shiftKey) this.navigatePage(-1);
        break;
      case 'ArrowRight':
        if (!e.shiftKey) this.navigatePage(1);
        break;
      case 'Home':
        e.preventDefault();
        this.scrollToPage(1);
        break;
      case 'End':
        if (this.pdfDoc) {
          e.preventDefault();
          this.scrollToPage(this.pdfDoc.numPages);
        }
        break;
    }
  }

  // ─── Search ───────────────────────────────────────────────────────────────

  private buildSearchBar(root: HTMLElement): void {
    this.searchBarEl = root.createDiv('pcai-search-bar');
    this.searchBarEl.addClass('pcai-hidden');

    this.searchInputEl = this.searchBarEl.createEl('input', {
      cls: 'pcai-search-input',
      attr: { type: 'text', placeholder: 'Search in PDF\u2026' },
    });

    const prevBtn = this.searchBarEl.createEl('button', { cls: 'pcai-icon-btn', text: '\u25B2' });
    prevBtn.addEventListener('click', () => this.navigateSearch(-1));

    const nextBtn = this.searchBarEl.createEl('button', { cls: 'pcai-icon-btn', text: '\u25BC' });
    nextBtn.addEventListener('click', () => this.navigateSearch(1));

    this.searchResultsEl = this.searchBarEl.createSpan({ cls: 'pcai-search-results-count' });

    const closeBtn = this.searchBarEl.createEl('button', { cls: 'pcai-icon-btn', text: '\u00D7' });
    closeBtn.addEventListener('click', () => this.toggleSearch());

    this.searchInputEl.addEventListener('input', () => void this.executeSearch());

    this.searchInputEl.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.navigateSearch(e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        this.toggleSearch();
      }
    });
  }

  private toggleSearch(): void {
    this.searchOpen = !this.searchOpen;
    this.searchBarEl.toggleClass('pcai-hidden', !this.searchOpen);
    if (this.searchOpen) {
      this.searchInputEl.focus();
      this.searchInputEl.select();
    } else {
      this.clearSearchHighlights();
      this.searchMatches = [];
      this.searchCurrentIdx = -1;
      this.searchResultsEl.setText('');
    }
  }

  private async executeSearch(): Promise<void> {
    const query = this.searchInputEl.value.trim().toLowerCase();
    this.clearSearchHighlights();
    this.searchMatches = [];
    this.searchCurrentIdx = -1;

    if (!query || !this.pdfDoc) {
      this.searchResultsEl.setText('');
      return;
    }

    for (let i = 1; i <= this.pdfDoc.numPages; i++) {
      try {
        const page = await this.pdfDoc.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items
          .map((item) => ('str' in item ? item.str : ''))
          .join(' ')
          .toLowerCase();

        let startIdx = 0;
        while (true) {
          const idx = pageText.indexOf(query, startIdx);
          if (idx === -1) break;
          this.searchMatches.push({ page: i, index: idx });
          startIdx = idx + 1;
        }
      } catch {
        // Skip unreadable pages
      }
    }

    if (this.searchMatches.length > 0) {
      this.searchCurrentIdx = 0;
      this.searchResultsEl.setText(`1 / ${this.searchMatches.length}`);
      this.highlightSearchMatch(this.searchMatches[0]);
    } else {
      this.searchResultsEl.setText('No results');
    }
  }

  private navigateSearch(direction: number): void {
    if (this.searchMatches.length === 0) return;
    this.searchCurrentIdx = (this.searchCurrentIdx + direction + this.searchMatches.length) % this.searchMatches.length;
    this.searchResultsEl.setText(`${this.searchCurrentIdx + 1} / ${this.searchMatches.length}`);
    this.highlightSearchMatch(this.searchMatches[this.searchCurrentIdx]);
  }

  private highlightSearchMatch(match: { page: number; index: number }): void {
    // Scroll to the page
    this.scrollToPage(match.page);

    // Highlight matching text spans in the text layer
    this.clearSearchHighlights();
    const wrapper = this.pagesEl.querySelector(`[data-page="${match.page}"]`);
    if (!wrapper) return;
    const textLayer = wrapper.querySelector('.textLayer');
    if (!textLayer) return;

    const spans = textLayer.querySelectorAll('span');
    const query = this.searchInputEl.value.trim().toLowerCase();
    for (const span of Array.from(spans)) {
      if (span.textContent?.toLowerCase().includes(query)) {
        span.addClass('pcai-search-highlight');
      }
    }
  }

  private clearSearchHighlights(): void {
    this.pagesEl.querySelectorAll('.pcai-search-highlight').forEach((el) => {
      el.removeClass('pcai-search-highlight');
    });
  }

  // ─── PDF Outline / TOC ────────────────────────────────────────────────────

  private async loadOutline(): Promise<void> {
    if (!this.pdfDoc) return;
    this.outlineEl.empty();

    try {
      const outline = await this.pdfDoc.getOutline();
      if (!outline || outline.length === 0) {
        this.outlineEl.createDiv({ cls: 'pcai-outline-empty', text: 'No bookmarks in this PDF' });
        return;
      }
      this.renderOutlineItems(outline, this.outlineEl, 0);
    } catch {
      this.outlineEl.createDiv({ cls: 'pcai-outline-empty', text: 'Could not load outline' });
    }
  }

  private renderOutlineItems(items: { title?: string; items?: unknown[]; dest?: unknown }[], parent: HTMLElement, depth: number): void {
    for (const item of items) {
      const row = parent.createDiv({
        cls: 'pcai-outline-item',
      });
      row.setCssStyles({ paddingLeft: `${12 + depth * 16}px` });
      row.createSpan({ text: item.title ?? '(untitled)' });

      row.addEventListener('click', () => {
        void this.navigateToOutlineItem(item);
      });

      if (item.items && item.items.length > 0) {
        this.renderOutlineItems(item.items as { title?: string; items?: unknown[]; dest?: unknown }[], parent, depth + 1);
      }
    }
  }

  private async navigateToOutlineItem(item: { dest?: unknown }): Promise<void> {
    if (!this.pdfDoc) return;
    try {
      const dest = typeof item.dest === 'string'
        ? await this.pdfDoc.getDestination(item.dest)
        : item.dest;
      if (!dest || !Array.isArray(dest)) return;

      const pageIdx = await this.pdfDoc.getPageIndex(dest[0] as { num: number; gen: number });
      this.scrollToPage(pageIdx + 1);
    } catch {
      // Some destinations can't be resolved
    }
  }

  private toggleOutline(): void {
    const isCollapsed = this.outlineSectionEl.hasClass('pcai-collapsed');
    this.outlineSectionEl.toggleClass('pcai-collapsed', !isCollapsed);
    if (isCollapsed) {
      this.outlineSectionEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }

  // ─── Text selection ────────────────────────────────────────────────────────

  private setupSelectionListener(): void {
    this.containerEl.addEventListener('mouseup', (e: MouseEvent) => {
      setTimeout(() => this.checkSelection(e), 10);
    });
  }

  private checkSelection(_e: MouseEvent): void {
    const selection = window.getSelection();
    if (!selection || selection.isCollapsed) {
      this.hideSelectionMenu();
      return;
    }

    const text = selection.toString().trim();
    if (!text) {
      this.hideSelectionMenu();
      return;
    }

    const range = selection.getRangeAt(0);
    const startNode = range.startContainer;
    const pageWrapper = (startNode.nodeType === Node.TEXT_NODE
      ? startNode.parentElement
      : startNode as Element
    )?.closest('.pcai-page-wrapper') as HTMLElement | null;

    if (!pageWrapper) {
      this.hideSelectionMenu();
      return;
    }

    const pageNum = parseInt(pageWrapper.getAttribute('data-page')!, 10);
    const pageRect = pageWrapper.getBoundingClientRect();
    const clientRects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0);

    if (clientRects.length === 0) {
      this.hideSelectionMenu();
      return;
    }

    this.selectedRects = clientRects.map((r) => ({
      x: (r.left - pageRect.left) / pageRect.width,
      y: (r.top - pageRect.top) / pageRect.height,
      width: r.width / pageRect.width,
      height: r.height / pageRect.height,
    }));
    this.selectedText = text;
    this.selectedPageNum = pageNum;

    const lastRect = clientRects[clientRects.length - 1];
    const containerRect = this.containerEl.getBoundingClientRect();
    const menuLeft = Math.min(lastRect.left - containerRect.left, containerRect.width - 280);
    const menuTop = lastRect.bottom - containerRect.top + 6;

    this.selectionMenuEl.setCssStyles({ left: `${Math.max(0, menuLeft)}px`, top: `${menuTop}px` });
    this.selectionMenuEl.removeClass('pcai-hidden');
  }

  private hideSelectionMenu(): void {
    this.selectionMenuEl.addClass('pcai-hidden');
  }

  // ─── Highlights ────────────────────────────────────────────────────────────

  private createHighlight(color: HighlightColor): void {
    if (!this.currentFile || !this.selectedText || this.selectedRects.length === 0) return;

    const h = this.plugin.annotationStore.add(
      this.currentFile.path,
      this.selectedPageNum,
      this.selectedText,
      color,
      this.selectedRects,
    );

    this.renderHighlightOnPage(this.selectedPageNum, h);

    // Flash newly created highlights
    const wrapper = this.pagesEl.querySelector(`[data-page="${this.selectedPageNum}"]`);
    if (wrapper) {
      wrapper.querySelectorAll(`[data-highlight-id="${h.id}"]`).forEach((el) => {
        el.addClass('pcai-highlight-new');
      });
    }

    this.hideSelectionMenu();
    window.getSelection()?.removeAllRanges();

    // Refresh the annotations sidebar and file list (badge count)
    this.refreshAnnotations();
    this.refreshFileList();
  }

  private loadHighlightsForCurrentFile(): void {
    for (const pageNum of this.renderedPages) {
      this.applyHighlightsToPage(pageNum);
    }
  }

  private applyHighlightsToPage(pageNum: number): void {
    if (!this.currentFile) return;
    const wrapper = this.pagesEl.querySelector(`[data-page="${pageNum}"]`);
    if (!wrapper) return;
    const layer = wrapper.querySelector('.pcai-annotation-layer');
    if (!layer) return;

    layer.empty();
    this.plugin.annotationStore
      .getForPage(this.currentFile.path, pageNum)
      .forEach((h) => this.renderHighlightOnPage(pageNum, h));
  }

  private renderHighlightOnPage(pageNum: number, h: Highlight): void {
    const wrapper = this.pagesEl.querySelector(`[data-page="${pageNum}"]`);
    if (!wrapper) return;

    const layer = wrapper.querySelector('.pcai-annotation-layer');
    if (!layer) return;

    for (const rect of h.rects) {
      const div = layer.createDiv({ cls: `pcai-highlight pcai-hl-${h.color}` });
      div.setCssStyles({
        left: `${rect.x * 100}%`,
        top: `${rect.y * 100}%`,
        width: `${rect.width * 100}%`,
        height: `${rect.height * 100}%`,
        backgroundColor: COLOR_HEX[h.color],
      });
      div.title = h.text;
      div.dataset.highlightId = h.id;

      div.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.showHighlightContextMenu(h, e);
      });
    }
  }

  private showHighlightContextMenu(h: Highlight, e: MouseEvent): void {
    const menu = new Menu();
    if (this.plugin.settings.enableAi) {
      menu.addItem((item) =>
        item
          .setTitle('Ask AI about this highlight')
          .setIcon('bot')
          .onClick(() => void this.askAboutHighlight(h)),
      );
    }
    menu.addItem((item) =>
      item
        .setTitle('Add note')
        .setIcon('pencil')
        .onClick(() => this.editAnnotationNote(h)),
    );
    menu.addItem((item) =>
      item
        .setTitle('Delete highlight')
        .setIcon('trash')
        .onClick(() => {
          this.plugin.annotationStore.remove(h.id);
          this.applyHighlightsToPage(h.pageNumber);
          this.refreshAnnotations();
          this.refreshFileList();
        }),
    );
    menu.showAtMouseEvent(e);
  }

  // ─── AI integration ────────────────────────────────────────────────────────

  private async askAboutCurrentPdf(): Promise<void> {
    await this.plugin.activateAiSidebar();
    const view = this.plugin.getAiSidebarView();
    if (view) {
      if (this.currentFile) view.setCurrentPdf(this.currentFile);
      view.setContextScope('pdf');
    }
    new Notice('Type your question in the sidebar.');
  }

  private async askAboutSelection(): Promise<void> {
    const text = this.selectedText;
    this.hideSelectionMenu();
    window.getSelection()?.removeAllRanges();

    await this.plugin.activateAiSidebar();
    const view = this.plugin.getAiSidebarView();
    if (view) {
      if (this.currentFile) view.setCurrentPdf(this.currentFile);
      view.prefillQuestion(`Regarding this passage:\n\n> ${text}\n\n`);
      view.setContextScope('pdf');
    }
  }

  private async askAboutHighlight(h: Highlight): Promise<void> {
    await this.plugin.activateAiSidebar();
    const view = this.plugin.getAiSidebarView();
    if (view) {
      if (this.currentFile) view.setCurrentPdf(this.currentFile);
      view.prefillQuestion(`Regarding this highlighted passage:\n\n> ${h.text}\n\n`);
      view.setContextScope('pdf');
    }
  }

  // ─── Metadata display ──────────────────────────────────────────────────────

  private updateMetadataDisplay(): void {
    this.metadataEl.empty();
    if (!this.currentMetadata) {
      this.metadataEl.addClass('pcai-hidden');
      return;
    }

    const meta = this.currentMetadata;

    // If metadata has a title, show it instead of the filename
    if (meta.title) {
      this.filenameLabelEl.setText(meta.title);
    }

    // Build metadata subtitle line (author + year)
    const parts: string[] = [];
    if (meta.author) parts.push(meta.author);
    if (meta.creationDate) {
      const year = this.parseCreationDateYear(meta.creationDate);
      if (year) parts.push(year);
    }

    if (parts.length > 0) {
      this.metadataEl.removeClass('pcai-hidden');
      this.metadataEl.createSpan({
        cls: 'pcai-pdf-meta-text',
        text: parts.join(' \u00B7 '),
      });
    }

    // Tooltip with full metadata
    const tooltipParts: string[] = [];
    if (meta.title) tooltipParts.push(`Title: ${meta.title}`);
    if (meta.author) tooltipParts.push(`Author: ${meta.author}`);
    if (meta.subject) tooltipParts.push(`Subject: ${meta.subject}`);
    if (meta.keywords) tooltipParts.push(`Keywords: ${meta.keywords}`);
    if (meta.creationDate) tooltipParts.push(`Created: ${meta.creationDate}`);
    if (tooltipParts.length > 0) {
      this.metadataEl.title = tooltipParts.join('\n');
    }
  }

  private parseCreationDateYear(dateStr: string): string | null {
    // PDF date format: D:YYYYMMDDHHmmSS or just a date string
    const pdfDateMatch = dateStr.match(/D:(\d{4})/);
    if (pdfDateMatch) return pdfDateMatch[1];

    const yearMatch = dateStr.match(/(\d{4})/);
    return yearMatch ? yearMatch[1] : null;
  }

  // ─── Annotation export ──────────────────────────────────────────────────────

  private async exportAnnotations(): Promise<void> {
    if (!this.currentFile) {
      new Notice('No PDF open.');
      return;
    }

    const highlights = this.plugin.annotationStore.getForFile(this.currentFile.path);
    if (highlights.length === 0) {
      new Notice('No annotations to export.');
      return;
    }

    const labels = this.plugin.settings.colorLabels ?? DEFAULT_COLOR_LABELS;
    const title = this.currentMetadata?.title ?? this.currentFile.basename;
    const author = this.currentMetadata?.author;
    const totalPages = this.pdfDoc?.numPages ?? 0;

    // Group highlights by page
    const byPage = new Map<number, Highlight[]>();
    for (const h of highlights) {
      const arr = byPage.get(h.pageNumber) ?? [];
      arr.push(h);
      byPage.set(h.pageNumber, arr);
    }
    const sortedPages = [...byPage.keys()].sort((a, b) => a - b);

    // Build markdown content
    const lines: string[] = [];
    lines.push(`# Annotations: ${title}`);
    const metaParts: string[] = [];
    if (author) metaParts.push(`**Author:** ${author}`);
    if (totalPages > 0) metaParts.push(`**Pages:** ${totalPages}`);
    if (metaParts.length > 0) lines.push(metaParts.join('  '));
    lines.push('');

    const useCallouts = (this.plugin.settings.exportFormat ?? 'callout') === 'callout';

    for (const page of sortedPages) {
      lines.push(`## Page ${page}`);
      lines.push('');
      const pageHighlights = byPage.get(page)!;
      for (const h of pageHighlights) {
        const colorLabel = labels[h.color] ?? h.color;
        if (useCallouts) {
          lines.push(`> [!${colorLabel}] ${h.text}`);
          if (h.notePath) {
            const noteBasename = h.notePath.split('/').pop()?.replace(/\.md$/, '') ?? h.notePath;
            lines.push(`> Note: [[${noteBasename}]]`);
          } else if (h.note) {
            lines.push(`> *Note: ${h.note}*`);
          }
        } else {
          lines.push(`> ${h.text}`);
          lines.push(`> — **${colorLabel}**`);
          if (h.notePath) {
            const noteBasename = h.notePath.split('/').pop()?.replace(/\.md$/, '') ?? h.notePath;
            lines.push(`> Note: [[${noteBasename}]]`);
          } else if (h.note) {
            lines.push(`> *Note: ${h.note}*`);
          }
        }
        lines.push('');
      }
    }

    lines.push('---');
    const dateStr = new Date().toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
    lines.push(`*Exported on ${dateStr}*`);
    lines.push('');

    const content = lines.join('\n');

    // Determine output path: same directory as PDF, with " - Annotations.md" suffix
    const pdfDir = this.currentFile.parent?.path ?? '';
    const baseName = this.currentFile.basename;
    const mdPath = pdfDir
      ? `${pdfDir}/${baseName} - Annotations.md`
      : `${baseName} - Annotations.md`;

    try {
      // Check if file already exists
      const existing = this.app.vault.getAbstractFileByPath(mdPath);
      if (existing instanceof TFile) {
        await this.app.vault.modify(existing, content);
        new Notice(`Updated "${mdPath}".`);
        await this.app.workspace.getLeaf('tab').openFile(existing);
      } else {
        const newFile = await this.app.vault.create(mdPath, content);
        new Notice(`Created "${mdPath}".`);
        await this.app.workspace.getLeaf('tab').openFile(newFile);
      }
    } catch (err) {
      console.error('PDF Tools: annotation export error:', err);
      new Notice(`Export failed \u2014 ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  // ─── AI annotation summary ────────────────────────────────────────────────

  private async summarizeAnnotations(): Promise<void> {
    if (!this.currentFile) {
      new Notice('No PDF open.');
      return;
    }

    const highlights = this.plugin.annotationStore.getForFile(this.currentFile.path);
    if (highlights.length === 0) {
      new Notice('No annotations to summarize.');
      return;
    }

    const labels = this.plugin.settings.colorLabels ?? DEFAULT_COLOR_LABELS;
    const filename = this.currentMetadata?.title ?? this.currentFile.basename;

    // Build the annotation text for the AI
    const annoLines: string[] = [];
    for (const h of highlights) {
      const colorLabel = labels[h.color] ?? h.color;
      let line = `Page ${h.pageNumber} [${colorLabel}]: "${h.text}"`;
      if (h.notePath) {
        const noteFile = this.app.vault.getAbstractFileByPath(h.notePath);
        if (noteFile instanceof TFile) {
          try {
            const noteContent = await this.app.vault.cachedRead(noteFile);
            // Strip frontmatter and quote block, keep just the user's note
            const stripped = noteContent.replace(/^---[\s\S]*?---\s*/, '').replace(/^>.*\n?/gm, '').trim();
            if (stripped) line += ` (Note: ${stripped})`;
          } catch {
            // Ignore read errors
          }
        }
      } else if (h.note) {
        line += ` (Note: ${h.note})`;
      }
      annoLines.push(line);
    }

    const prompt = [
      `Summarize these annotations from "${filename}":`,
      '',
      ...annoLines,
      '',
      'Please provide a structured summary organized by theme, highlighting key concepts and connections between annotations.',
    ].join('\n');

    await this.plugin.activateAiSidebar();
    const view = this.plugin.getAiSidebarView();
    if (view) {
      if (this.currentFile) view.setCurrentPdf(this.currentFile);
      view.prefillQuestion(prompt);
      view.setContextScope('pdf');
    }
  }

  // ─── Cross-PDF annotation search ──────────────────────────────────────────

  private onAnnoSearchInput(): void {
    const query = this.annoSearchInputEl.value.trim().toLowerCase();
    const clearBtn = this.annoSearchInputEl.parentElement?.querySelector('.pcai-anno-search-clear') as HTMLElement | null;

    if (!query) {
      this.annoSearchActive = false;
      if (clearBtn) clearBtn.addClass('pcai-hidden');
      this.refreshAnnotations();
      return;
    }

    this.annoSearchActive = true;
    if (clearBtn) clearBtn.removeClass('pcai-hidden');

    const allHighlights = this.plugin.annotationStore.getAllHighlights();
    const matches = allHighlights.filter((h) =>
      h.text.toLowerCase().includes(query) ||
      (h.note?.toLowerCase().includes(query) ?? false) ||
      (h.notePath?.toLowerCase().includes(query) ?? false),
    );

    this.renderCrossPdfSearchResults(matches);
  }

  private renderCrossPdfSearchResults(highlights: Highlight[]): void {
    this.annotationsEl.empty();

    if (highlights.length === 0) {
      this.annotationsEl.createDiv({
        cls: 'pcai-anno-empty',
        text: 'No matching annotations found.',
      });
      return;
    }

    // Group by file
    const byFile = new Map<string, Highlight[]>();
    for (const h of highlights) {
      const arr = byFile.get(h.pdfPath) ?? [];
      arr.push(h);
      byFile.set(h.pdfPath, arr);
    }

    for (const [pdfPath, fileHighlights] of byFile) {
      const fileGroup = this.annotationsEl.createDiv('pcai-anno-search-group');
      const basename = pdfPath.split('/').pop()?.replace(/\.pdf$/i, '') ?? pdfPath;
      fileGroup.createDiv({
        cls: 'pcai-anno-search-file-label',
        text: basename,
        attr: { title: pdfPath },
      });

      for (const h of fileHighlights) {
        const card = fileGroup.createDiv('pcai-anno-card');
        card.addEventListener('click', () => {
          // Load the PDF and scroll to the page
          const file = this.app.vault.getAbstractFileByPath(h.pdfPath);
          if (file instanceof TFile) {
            void this.loadFile(file).then(() => {
              this.scrollToPage(h.pageNumber);
            }).catch((e: unknown) => {
              console.error('PDF Tools: cross-search navigate error:', e);
            });
          }
        });

        const row = card.createDiv('pcai-anno-row');
        const dot = row.createSpan('pcai-anno-dot');
        dot.setCssProps({ '--dot-color': COLOR_HEX[h.color] });

        const textCol = row.createDiv({ cls: 'pcai-anno-search-text-col' });
        textCol.createSpan({
          cls: 'pcai-anno-text',
          text: h.text.length > 100 ? h.text.slice(0, 100) + '\u2026' : h.text,
        });
        textCol.createSpan({
          cls: 'pcai-anno-search-page',
          text: `p. ${h.pageNumber}`,
        });
      }
    }
  }

  // ─── Canvas quick-link ────────────────────────────────────────────────────

  private addCurrentPdfToCanvas(): void {
    if (!this.currentFile) {
      new Notice('No PDF open.');
      return;
    }
    try {
      this.plugin.addToCanvas(this.currentFile);
    } catch (e: unknown) {
      console.error('PDF Tools: addToCanvas error:', e);
      new Notice('Failed to add PDF to canvas.');
    }
  }

  // ─── Dictionary ──────────────────────────────────────────────────────────

  private static readonly DICT_DOWNLOAD_URL =
    'https://github.com/voidash/obsidian-pdf-tools/releases/latest/download/dictionary.json';

  /**
   * Loads the WordNet dictionary.json — tries local cache first, then
   * auto-downloads from the GitHub release and saves to the plugin dir.
   */
  private async loadLocalDictionary(): Promise<LocalDict> {
    if (this.localDict) return this.localDict;
    if (this.localDictLoading) {
      // Wait for the in-flight load to complete
      return new Promise((resolve) => {
        const check = setInterval(() => {
          if (this.localDict) {
            clearInterval(check);
            resolve(this.localDict);
          }
        }, 50);
      });
    }

    this.localDictLoading = true;
    const adapter = this.app.vault.adapter;
    const dictPath = `${this.app.vault.configDir}/plugins/pdf-tools/dictionary.json`;

    // 1. Try reading from local cache
    try {
      const raw = await adapter.read(dictPath);
      if (raw) {
        this.localDict = JSON.parse(raw) as LocalDict;
        console.debug(`PDF Tools — loaded dictionary (${Object.keys(this.localDict).length} words) from cache`);
        this.localDictLoading = false;
        return this.localDict;
      }
    } catch {
      // File doesn't exist yet — will download below
    }

    // 2. Download from GitHub release
    try {
      new Notice('Downloading dictionary (first-time setup)…');
      const response = await requestUrl({ url: PdfViewerView.DICT_DOWNLOAD_URL });
      if (response.status === 200 && response.text) {
        // Save to plugin dir for future use
        await adapter.write(dictPath, response.text);
        this.localDict = JSON.parse(response.text) as LocalDict;
        const count = Object.keys(this.localDict).length;
        console.debug(`PDF Tools — downloaded dictionary (${count} words)`);
        new Notice(`Dictionary ready (${count.toLocaleString()} words).`);
        this.localDictLoading = false;
        return this.localDict;
      }
    } catch (err) {
      console.warn('PDF Tools — dictionary download failed:', err);
      new Notice('Dictionary download failed. Using online fallback.');
    }

    this.localDictLoading = false;
    // Return empty dict — API fallback will handle lookups
    const empty = Object.create(null) as LocalDict;
    this.localDict = empty;
    return empty;
  }

  private async lookupWord(word: string): Promise<void> {
    this.dictionaryResultsEl.empty();
    this.dictionaryResultsEl.createDiv({ cls: 'pcai-dict-loading', text: 'Looking up\u2026' });

    // Expand dictionary section and scroll into view
    this.dictionarySectionEl.removeClass('pcai-collapsed');
    this.dictionarySectionEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    const normalized = word.toLowerCase().trim();
    const dictSource = this.plugin.settings.dictionarySource ?? 'auto';

    // 1. Try embedded dictionary first (unless set to 'api' only)
    if (dictSource !== 'api') {
      try {
        const dict = await this.loadLocalDictionary();
        const entries = dict[normalized];
        if (entries && entries.length > 0) {
          this.dictionaryResultsEl.empty();
          this.renderLocalDictResults(normalized, entries);
          return;
        }
      } catch (err) {
        console.warn('PDF Tools — local dictionary lookup error:', err);
      }
    }

    // 2. Fallback to free API (unless set to 'local' only)
    if (dictSource === 'local') {
      this.dictionaryResultsEl.empty();
      this.dictionaryResultsEl.createDiv({
        cls: 'pcai-dict-empty',
        text: `No definition found for "${word}" in the local dictionary`,
      });
      return;
    }
    try {
      const response = await requestUrl({
        url: `https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(normalized)}`,
      });
      this.dictionaryResultsEl.empty();

      if (response.status !== 200) {
        this.dictionaryResultsEl.createDiv({
          cls: 'pcai-dict-empty',
          text: `No definition found for "${word}"`,
        });
        return;
      }

      const apiEntries = response.json as DictApiEntry[];
      if (!Array.isArray(apiEntries) || apiEntries.length === 0) {
        this.dictionaryResultsEl.createDiv({
          cls: 'pcai-dict-empty',
          text: `No definition found for "${word}"`,
        });
        return;
      }

      this.renderApiDictResults(apiEntries);
    } catch {
      this.dictionaryResultsEl.empty();
      this.dictionaryResultsEl.createDiv({
        cls: 'pcai-dict-empty',
        text: `No definition found for "${word}"`,
      });
    }
  }

  /**
   * Renders results from the embedded WordNet dictionary.
   * Format: [["n", "definition"], ["v", "definition"], ...]
   */
  private renderLocalDictResults(word: string, entries: [string, string][]): void {
    const header = this.dictionaryResultsEl.createDiv({ cls: 'pcai-dict-word-header' });
    header.createSpan({ cls: 'pcai-dict-word', text: word });
    header.createSpan({ cls: 'pcai-dict-source', text: ' (WordNet)' });

    // Group by part of speech
    const byPos = new Map<string, string[]>();
    for (const [pos, def] of entries) {
      const arr = byPos.get(pos) ?? [];
      arr.push(def);
      byPos.set(pos, arr);
    }

    for (const [pos, defs] of byPos) {
      const meaningDiv = this.dictionaryResultsEl.createDiv({ cls: 'pcai-dict-meaning' });
      meaningDiv.createDiv({ cls: 'pcai-dict-pos', text: POS_LABELS[pos] ?? pos });

      const list = meaningDiv.createEl('ol', { cls: 'pcai-dict-defs' });
      for (const def of defs) {
        const li = list.createEl('li');
        li.createSpan({ cls: 'pcai-dict-def-text', text: def });
      }
    }
  }

  /**
   * Renders results from the free dictionary API.
   */
  private renderApiDictResults(entries: DictApiEntry[]): void {
    for (const entry of entries) {
      const wordDiv = this.dictionaryResultsEl.createDiv({ cls: 'pcai-dict-word-header' });
      wordDiv.createSpan({ cls: 'pcai-dict-word', text: entry.word ?? '' });
      if (entry.phonetic) {
        wordDiv.createSpan({ cls: 'pcai-dict-phonetic', text: ` ${entry.phonetic}` });
      }

      if (!Array.isArray(entry.meanings)) continue;

      for (const meaning of entry.meanings) {
        const meaningDiv = this.dictionaryResultsEl.createDiv({ cls: 'pcai-dict-meaning' });
        meaningDiv.createDiv({ cls: 'pcai-dict-pos', text: meaning.partOfSpeech ?? '' });

        if (!Array.isArray(meaning.definitions)) continue;

        const list = meaningDiv.createEl('ol', { cls: 'pcai-dict-defs' });
        for (const def of meaning.definitions.slice(0, 5)) {
          const li = list.createEl('li');
          li.createSpan({ cls: 'pcai-dict-def-text', text: def.definition ?? '' });
          if (def.example) {
            li.createDiv({ cls: 'pcai-dict-example', text: `"${def.example}"` });
          }
        }
      }
    }
  }

  // ─── File picker ───────────────────────────────────────────────────────────

  private openFilePicker(): void {
    new PdfFileSuggestModal(this.app, (file) => {
      void this.loadFile(file).catch((e: unknown) =>
        console.error('PDF Tools \u2014 loadFile error:', e),
      );
    }).open();
  }
}

// ─── Fuzzy PDF picker modal ─────────────────────────────────────────────────

class PdfFileSuggestModal extends FuzzySuggestModal<TFile> {
  private onChoose: (file: TFile) => void;

  constructor(app: App, onChoose: (file: TFile) => void) {
    super(app);
    this.onChoose = onChoose;
    this.setPlaceholder('Type to search PDF files\u2026');
  }

  getItems(): TFile[] {
    return this.app.vault.getFiles().filter((f) => f.extension === 'pdf');
  }

  getItemText(file: TFile): string {
    return file.path;
  }

  onChooseItem(file: TFile): void {
    this.onChoose(file);
  }
}
