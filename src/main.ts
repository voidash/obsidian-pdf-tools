import { Plugin, Notice, TFile, Menu, ItemView, Modal } from 'obsidian';
import type { EventRef } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import pdfjsWorkerText from 'pdfjs-worker-inline';

import { DEFAULT_SETTINGS, PdfCanvasAiSettingTab } from './settings';
import type { PluginSettings } from './settings';
import { AnnotationStore } from './store/annotationStore';
import { ChatStore } from './store/chatStore';
import { AiService } from './services/aiService';
import { PdfService } from './services/pdfService';
import { ProxyManager } from './services/proxyManager';
import { AI_SIDEBAR_VIEW_TYPE, AiSidebarView } from './views/AiSidebarView';
import { AI_CHAT_VIEW_TYPE, AiChatView } from './views/AiChatView';
import { PDF_VIEWER_VIEW_TYPE, PdfViewerView } from './views/PdfViewerView';
import { getSelectedPdfNodes, getAllCanvasPdfNodes } from './canvas/canvasUtils';
import { CanvasPdfInjector } from './canvas/canvasPdfInjector';

type SpreadDirection = 'down' | 'right';

/** Minimal shape of canvas node serialized data (internal Obsidian API). */
interface CanvasNodeData {
  type?: string;
  text?: string;
  url?: string;
  label?: string;
  file?: string;
  x?: number;
  y?: number;
  width?: number;
  id?: string;
  fromNode?: string;
  toNode?: string;
}

interface CanvasNode {
  file?: TFile;
  x?: number;
  y?: number;
  width?: number;
  id?: string;
  contentEl?: HTMLElement;
  getData?: () => CanvasNodeData;
}

interface CanvasEdge {
  getData?: () => CanvasNodeData;
  fromNode?: string;
  toNode?: string;
}

/** Minimal typed surface for the internal Obsidian canvas object. */
interface CanvasApi {
  nodes?: Map<string, CanvasNode>;
  edges?: Map<string, CanvasEdge> | CanvasEdge[];
  data?: { edges?: Map<string, CanvasEdge> | CanvasEdge[] };
  selection?: Set<unknown>;
  removeNode?: (node: unknown) => void;
  createTextNode?: (opts: Record<string, unknown>) => CanvasNode | null;
  createFileNode?: (opts: Record<string, unknown>) => CanvasNode | null;
  createEdge?: (opts: Record<string, unknown>) => unknown;
  requestSave?: () => void;
  getViewportCenter?: () => { x: number; y: number };
}

class PageJumpModal extends Modal {
  private currentPage: number;
  private numPages: number;
  private onNavigate: (page: number) => void;

  constructor(
    app: Modal['app'],
    currentPage: number,
    numPages: number,
    onNavigate: (page: number) => void,
  ) {
    super(app);
    this.currentPage = currentPage;
    this.numPages = numPages;
    this.onNavigate = onNavigate;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.addClass('pcai-page-jump-modal');
    contentEl.createEl('h3', { text: 'Go to page' });

    const form = contentEl.createEl('form');
    const row = form.createDiv({ cls: 'pcai-page-jump-row' });

    const input = row.createEl('input', {
      type: 'number',
      attr: {
        min: '1',
        max: String(this.numPages),
        value: String(this.currentPage),
        placeholder: `1–${this.numPages}`,
      },
      cls: 'pcai-page-jump-input',
    });

    row.createSpan({
      cls: 'pcai-page-jump-total',
      text: `/ ${this.numPages}`,
    });

    form.createEl('button', {
      text: 'Go',
      cls: 'mod-cta pcai-page-jump-btn',
      type: 'submit',
    });

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const n = parseInt(input.value, 10);
      if (!isNaN(n) && n >= 1 && n <= this.numPages) {
        this.close();
        this.onNavigate(n);
      } else {
        input.addClass('pcai-input-error');
        setTimeout(() => input.removeClass('pcai-input-error'), 600);
      }
    });

    // Also navigate on Enter key directly
    input.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      }
    });

    // Focus and select input after modal opens
    setTimeout(() => { input.focus(); input.select(); }, 50);
  }

  onClose(): void {
    this.contentEl.empty();
  }
}

/** Outline item shape from pdfjs (items is recursive but typed loosely). */
type OutlineItem = { title?: string; dest?: unknown; items?: OutlineItem[] };

class OutlineModal extends Modal {
  private outline: OutlineItem[];
  private onNavigate: (entry: { dest?: unknown }) => void;

  constructor(
    app: Modal['app'],
    outline: OutlineItem[],
    onNavigate: (entry: { dest?: unknown }) => void,
  ) {
    super(app);
    this.outline = outline;
    this.onNavigate = onNavigate;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.addClass('pcai-outline-modal');
    contentEl.createEl('h3', { text: 'Outline' });

    const listEl = contentEl.createDiv({ cls: 'pcai-outline-modal-list' });
    this.renderItems(listEl, this.outline, 0);
  }

  private renderItems(container: HTMLElement, items: OutlineItem[], depth: number): void {
    for (const entry of items) {
      const title = entry.title?.trim() || '(untitled)';
      const row = container.createDiv({ cls: 'pcai-outline-modal-item' });
      row.setCssStyles({ paddingLeft: `${12 + depth * 18}px` });
      row.textContent = title;
      row.addEventListener('click', () => {
        this.close();
        this.onNavigate(entry);
      });

      if (Array.isArray(entry.items) && entry.items.length > 0) {
        this.renderItems(container, entry.items as OutlineItem[], depth + 1);
      }
    }
  }

  onClose(): void {
    this.contentEl.empty();
  }
}

class SpreadOptionsModal extends Modal {
  private onChoose: (direction: SpreadDirection) => void;

  constructor(app: Modal['app'], onChoose: (direction: SpreadDirection) => void) {
    super(app);
    this.onChoose = onChoose;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.addClass('pcai-spread-modal');
    contentEl.createEl('h3', { text: 'Spread PDF pages' });

    const row = contentEl.createDiv({ cls: 'pcai-spread-modal-row' });

    const downBtn = row.createEl('button', { cls: 'pcai-spread-modal-btn' });
    downBtn.createSpan({ text: '\u2193' }); // ↓
    downBtn.createEl('br');
    downBtn.createSpan({ cls: 'pcai-spread-modal-label', text: 'Down' });
    downBtn.addEventListener('click', () => {
      this.close();
      this.onChoose('down');
    });

    const rightBtn = row.createEl('button', { cls: 'pcai-spread-modal-btn' });
    rightBtn.createSpan({ text: '\u2192' }); // →
    rightBtn.createEl('br');
    rightBtn.createSpan({ cls: 'pcai-spread-modal-label', text: 'Right' });
    rightBtn.addEventListener('click', () => {
      this.close();
      this.onChoose('right');
    });
  }

  onClose(): void {
    this.contentEl.empty();
  }
}

export default class PdfCanvasAiPlugin extends Plugin {
  settings!: PluginSettings;
  annotationStore!: AnnotationStore;
  chatStore!: ChatStore;          // null when AI disabled
  aiService!: AiService;          // null when AI disabled
  pdfService!: PdfService;
  proxyManager!: ProxyManager;    // null when AI disabled
  canvasInjector!: CanvasPdfInjector;

  async onload(): Promise<void> {
    await this.loadSettings();

    this.annotationStore = new AnnotationStore(this);
    await this.annotationStore.load();

    this.pdfService = new PdfService(this.app);

    // AI-related services: only initialize when AI is enabled
    if (this.settings.enableAi) {
      this.chatStore = new ChatStore(this.app, this.manifest.dir ?? '');
      await this.chatStore.load();
      this.aiService = new AiService(this.settings);
      this.proxyManager = new ProxyManager(this.settings.baseUrl);
    }

    this.setupPdfjsWorker();
    this.registerViews();
    this.addCommands();
    this.addRibbonIcons();
    this.addSettingTab(new PdfCanvasAiSettingTab(this.app, this));
    this.canvasInjector = new CanvasPdfInjector(this);
    this.canvasInjector.start();
    this.registerCanvasMenu();
    this.registerPdfIntercept();
    this.registerVaultEvents();

    // Start the local proxy only if opt-in via settings
    if (this.settings.enableAi && this.settings.proxyAutoStart && this.settings.provider === 'local-proxy') {
      void this.proxyManager.ensureRunning().catch((e: unknown) => {
        console.error('PDF Tools: proxyManager.ensureRunning error', e);
      });
    }

    console.debug('PDF Tools: loaded');
  }

  onunload(): void {
    this.canvasInjector.stop();
    this.annotationStore.destroy();
    if (this.chatStore) {
      void this.chatStore.flush().then(() => {
        this.chatStore.destroy();
      });
    }
    this.proxyManager?.stop();
    console.debug('PDF Tools: unloaded');
  }

  async loadSettings(): Promise<void> {
    const raw = await this.loadData() as Record<string, unknown> | null;
    this.settings = Object.assign({}, DEFAULT_SETTINGS, raw ?? {});
    // Deep-merge colorLabels so partial saves don't drop defaults
    const savedLabels = raw != null && typeof raw === 'object' && 'colorLabels' in raw
      ? raw.colorLabels as Record<string, string>
      : {};
    this.settings.colorLabels = Object.assign(
      {},
      DEFAULT_SETTINGS.colorLabels,
      savedLabels,
    );

    // Migration: fix baseUrl for local-proxy missing /v1 suffix (from older bug)
    if (
      this.settings.provider === 'local-proxy' &&
      this.settings.baseUrl &&
      !this.settings.baseUrl.endsWith('/v1')
    ) {
      this.settings.baseUrl = this.settings.baseUrl.replace(/\/+$/, '') + '/v1';
    }
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
    this.aiService?.updateSettings(this.settings);

    // Keep proxy manager in sync and start proxy on demand
    if (this.proxyManager && this.settings.provider === 'local-proxy') {
      this.proxyManager.reset(this.settings.baseUrl);
      void this.proxyManager.ensureRunning().catch((e: unknown) => {
        console.error('PDF Tools: proxyManager.ensureRunning error', e);
      });
    }
  }

  // ─── pdfjs worker setup ────────────────────────────────────────────────────

  private setupPdfjsWorker(): void {
    // Worker source is inlined at build time by the pdfjs-worker-inline esbuild plugin.
    // We create a Blob URL so pdfjs can spawn the worker without a separate file.
    const blob = new Blob([pdfjsWorkerText], { type: 'application/javascript' });
    const workerUrl = URL.createObjectURL(blob);
    pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;
  }

  // ─── Views ─────────────────────────────────────────────────────────────────

  private registerViews(): void {
    if (this.settings.enableAi) {
      this.registerView(AI_SIDEBAR_VIEW_TYPE, (leaf) => new AiSidebarView(leaf, this));
      this.registerView(AI_CHAT_VIEW_TYPE, (leaf) => new AiChatView(leaf, this));
    }
    this.registerView(PDF_VIEWER_VIEW_TYPE, (leaf) => new PdfViewerView(leaf, this));
  }

  // ─── Ribbon ────────────────────────────────────────────────────────────────

  private addRibbonIcons(): void {
    if (this.settings.enableAi) {
      this.addRibbonIcon('bot', 'Open PDF tools sidebar', () => {
        void this.activateAiSidebar().catch((e: unknown) => {
          console.error('PDF Tools: activateAiSidebar error', e);
        });
      });

      this.addRibbonIcon('message-square', 'Open AI chat', () => {
        void this.activateAiChat().catch((e: unknown) => {
          console.error('PDF Tools: activateAiChat error', e);
        });
      });
    }

    this.addRibbonIcon('file-text', 'Open PDF viewer', () => {
      void this.activatePdfViewer().catch((e: unknown) => {
        console.error('PDF Tools: activatePdfViewer error', e);
      });
    });
  }

  // ─── Commands ──────────────────────────────────────────────────────────────

  private addCommands(): void {
    if (this.settings.enableAi) {
      this.addCommand({
        id: 'open-ai-sidebar',
        name: 'Open AI sidebar',
        callback: () => {
          void this.activateAiSidebar().catch((e: unknown) => console.error(e));
        },
      });

      this.addCommand({
        id: 'open-ai-chat',
        name: 'Open AI chat (full window)',
        callback: () => {
          void this.activateAiChat().catch((e: unknown) => console.error(e));
        },
      });

      this.addCommand({
        id: 'ask-selected-pdfs',
        name: 'Ask AI about selected canvas pdfs',
        callback: () => {
          void this.askAboutPdfs('selected').catch((e: unknown) => console.error(e));
        },
      });

      this.addCommand({
        id: 'ask-all-canvas-pdfs',
        name: 'Ask AI about all canvas pdfs',
        callback: () => {
          void this.askAboutPdfs('all').catch((e: unknown) => console.error(e));
        },
      });
    }

    this.addCommand({
      id: 'open-pdf-viewer',
      name: 'Open PDF viewer pane',
      callback: () => {
        void this.activatePdfViewer().catch((e: unknown) => console.error(e));
      },
    });

    this.addCommand({
      id: 'open-selected-pdf',
      name: 'Open selected canvas PDF in viewer',
      callback: () => {
        void this.openSelectedCanvasPdf().catch((e: unknown) => console.error(e));
      },
    });
  }

  // ─── Canvas node context menu ──────────────────────────────────────────────

  private registerCanvasMenu(): void {
    // `canvas:node-menu` fires when the user right-clicks a canvas node.
    // It is undocumented but has been stable across community plugins since v1.4.
    // Cast workspace — the event name is not in Obsidian's public types.
    const ws = this.app.workspace as unknown as {
      on(name: string, callback: (menu: Menu, node: unknown) => void): EventRef;
    };
    const ref: EventRef = ws.on('canvas:node-menu', (menu: Menu, node: unknown) => {
      const file = this.resolveCanvasNodeFile(node);
      if (!file || file.extension !== 'pdf') return;

      menu.addSeparator();

      // ── Page navigation ──
      const renderer = this.canvasInjector.getRendererForNode(node);
      if (renderer) {
        const numPages = renderer.getNumPages();
        const currentPage = renderer.getCurrentVisiblePage();
        menu.addItem((item) =>
          item
            .setTitle(`Go to page\u2026 (${currentPage} / ${numPages})`)
            .setIcon('hash')
            .onClick(() => {
              new PageJumpModal(
                this.app,
                currentPage,
                numPages,
                (page) => renderer.scrollToPage(page),
              ).open();
            }),
        );
        const outline = renderer.getOutline();
        if (outline && outline.length > 0) {
          menu.addItem((item) =>
            item
              .setTitle('Show outline')
              .setIcon('list')
              .onClick(() => {
                new OutlineModal(
                  this.app,
                  outline as OutlineItem[],
                  (entry) => void renderer.navigateToOutlineItem(entry),
                ).open();
              }),
          );
        }
        menu.addSeparator();
      }

      menu.addItem((item) =>
        item
          .setTitle('Open in PDF viewer')
          .setIcon('file-text')
          .onClick(() => {
            void this.activatePdfViewer()
              .then(() => this.openFileInViewer(file))
              .catch((e: unknown) => console.error(e));
          }),
      );
      if (this.settings.enableAi) {
        menu.addItem((item) =>
          item
            .setTitle('Ask AI about this PDF')
            .setIcon('bot')
            .onClick(() => {
              void this.openFileInViewerAndAsk(file).catch((e: unknown) => console.error(e));
            }),
        );
      }
      menu.addItem((item) =>
        item
          .setTitle('Spread PDF pages')
          .setIcon('layout-grid')
          .onClick(() => {
            new SpreadOptionsModal(this.app, (direction) => {
              void this.spreadPdfPages(file, node, direction).catch((e: unknown) => {
                console.error('PDF Tools — spreadPdfPages error:', e);
                new Notice('Failed to spread PDF pages.');
              });
            }).open();
          }),
      );
      menu.addItem((item) =>
        item
          .setTitle('Extract current page')
          .setIcon('scissors')
          .onClick(() => {
            void this.extractCurrentPage(file, node).catch((e: unknown) => {
              console.error('PDF Tools — extractCurrentPage error:', e);
              new Notice('Failed to extract page.');
            });
          }),
      );
    });
    this.registerEvent(ref);
  }

  /**
   * Extract a TFile from a canvas node object, handling all node shapes seen
   * across Obsidian versions:
   *   - `node.file` is a TFile object  (most versions, v1.4+)
   *   - `node.filePath` is a string    (older versions)
   *   - `node.file` is a string path   (some intermediate versions)
   *   - `node.getData().file` is a string (serialized canvas data shape)
   */
  private resolveCanvasNodeFile(node: unknown): TFile | null {
    if (!node || typeof node !== 'object') return null;
    const n = node as Record<string, unknown>;

    // Shape 1: node.file is already a TFile
    if (n.file instanceof TFile) return n.file;

    // Shape 2: node.filePath is a string
    if (typeof n.filePath === 'string') {
      const f = this.app.vault.getAbstractFileByPath(n.filePath);
      return f instanceof TFile ? f : null;
    }

    // Shape 3: node.file is a string path
    if (typeof n.file === 'string') {
      const f = this.app.vault.getAbstractFileByPath(n.file);
      return f instanceof TFile ? f : null;
    }

    // Shape 4: serialized canvas data via getData()
    if (typeof n.getData === 'function') {
      try {
        const data = (n.getData as () => Record<string, unknown>)();
        if (typeof data.file === 'string') {
          const f = this.app.vault.getAbstractFileByPath(data.file);
          return f instanceof TFile ? f : null;
        }
      } catch {
        // getData() failed — ignore
      }
    }

    return null;
  }

  // ─── PDF intercept: redirect native viewer to ours ─────────────────────────

  private originalPdfViewType: string | undefined;

  private registerPdfIntercept(): void {
    // Override the extension → view-type mapping so all PDF opens use our viewer.
    // viewRegistry is internal but widely used by community plugins (pdf++, excalidraw, etc.)
    const registry = (this.app as unknown as { viewRegistry?: { typeByExtension?: Record<string, string> } }).viewRegistry;
    const typeByExt = registry?.typeByExtension;
    if (!typeByExt) {
      console.warn('PDF Tools: viewRegistry not found, PDF intercept unavailable');
      return;
    }

    this.originalPdfViewType = typeByExt['pdf'];
    typeByExt['pdf'] = PDF_VIEWER_VIEW_TYPE;

    // Restore on unload so disabling the plugin reverts to native behavior
    this.register(() => {
      if (this.originalPdfViewType !== undefined) {
        typeByExt['pdf'] = this.originalPdfViewType;
      }
    });
  }

  // ─── Vault events ──────────────────────────────────────────────────────────

  private registerVaultEvents(): void {
    this.registerEvent(
      this.app.vault.on('modify', (file) => {
        if (file instanceof TFile && file.extension === 'pdf') {
          this.pdfService.invalidateFile(file.path);
        }
      }),
    );

    this.registerEvent(
      this.app.vault.on('rename', (file, oldPath) => {
        if (file instanceof TFile && file.extension === 'pdf') {
          this.annotationStore.renameFile(oldPath, file.path);
          this.pdfService.invalidateFile(oldPath);
        }
        // Track renames of annotation note files
        if (file instanceof TFile && file.extension === 'md') {
          this.annotationStore.renameNoteFile(oldPath, file.path);
        }
      }),
    );
  }

  // ─── View activation ───────────────────────────────────────────────────────

  async activateAiSidebar(): Promise<void> {
    const existing = this.app.workspace.getLeavesOfType(AI_SIDEBAR_VIEW_TYPE);
    if (existing.length > 0) {
      void this.app.workspace.revealLeaf(existing[0]);
      return;
    }
    const leaf = this.app.workspace.getRightLeaf(false);
    if (!leaf) return;
    await leaf.setViewState({ type: AI_SIDEBAR_VIEW_TYPE, active: true });
    const leaves = this.app.workspace.getLeavesOfType(AI_SIDEBAR_VIEW_TYPE);
    if (leaves.length > 0) void this.app.workspace.revealLeaf(leaves[0]);
  }

  async activateAiChat(): Promise<void> {
    const existing = this.app.workspace.getLeavesOfType(AI_CHAT_VIEW_TYPE);
    if (existing.length > 0) {
      void this.app.workspace.revealLeaf(existing[0]);
      return;
    }
    const leaf = this.app.workspace.getLeaf('tab');
    await leaf.setViewState({ type: AI_CHAT_VIEW_TYPE, active: true });
    void this.app.workspace.revealLeaf(leaf);
  }

  async activatePdfViewer(): Promise<void> {
    const existing = this.app.workspace.getLeavesOfType(PDF_VIEWER_VIEW_TYPE);
    if (existing.length > 0) {
      void this.app.workspace.revealLeaf(existing[0]);
      return;
    }
    // Open as a new tab (not a split — splits cause janky open-then-close behavior)
    const leaf = this.app.workspace.getLeaf('tab');
    await leaf.setViewState({ type: PDF_VIEWER_VIEW_TYPE, active: true });
    void this.app.workspace.revealLeaf(leaf);
  }

  async openFileInViewer(file: TFile, attempt = 0): Promise<void> {
    const leaves = this.app.workspace.getLeavesOfType(PDF_VIEWER_VIEW_TYPE);
    const leaf = leaves[0];
    if (!leaf) {
      if (attempt >= 2) {
        new Notice('Could not open PDF viewer pane.');
        return;
      }
      await this.activatePdfViewer();
      return this.openFileInViewer(file, attempt + 1);
    }
    const view = leaf.view as PdfViewerView;
    await view.loadFile(file);
    void this.app.workspace.revealLeaf(leaf);
  }

  getAiSidebarView(): AiSidebarView | null {
    const leaves = this.app.workspace.getLeavesOfType(AI_SIDEBAR_VIEW_TYPE);
    return leaves.length > 0 ? (leaves[0].view as AiSidebarView) : null;
  }

  /** Returns the canvas object from any open canvas leaf, or null. */
  getActiveCanvas(): CanvasApi | null {
    const active = this.app.workspace.getActiveViewOfType(ItemView) as ItemView & { canvas?: CanvasApi } | null;
    if (active?.getViewType?.() === 'canvas' && active.canvas) {
      return active.canvas;
    }
    const canvasLeaves = this.app.workspace.getLeavesOfType('canvas');
    for (const leaf of canvasLeaves) {
      const view = leaf.view as ItemView & { canvas?: CanvasApi };
      if (view?.canvas) return view.canvas;
    }
    return null;
  }

  /** Returns the file currently open in the standalone PDF viewer, if any. */
  getViewerCurrentFile(): TFile | null {
    const leaves = this.app.workspace.getLeavesOfType(PDF_VIEWER_VIEW_TYPE);
    if (leaves.length === 0) return null;
    const view = leaves[0].view;
    if (view instanceof PdfViewerView) {
      return view.getCurrentFile();
    }
    return null;
  }

  // ─── PDF context gathering ─────────────────────────────────────────────────

  async gatherPdfContext(scope: 'selected' | 'all'): Promise<string> {
    let nodes =
      scope === 'selected' ? getSelectedPdfNodes(this.app) : getAllCanvasPdfNodes(this.app);

    // If selection yielded nothing, fall back to viewer's open file
    if (nodes.length === 0) {
      const viewerLeaves = this.app.workspace.getLeavesOfType(PDF_VIEWER_VIEW_TYPE);
      if (viewerLeaves.length > 0) {
        const view = viewerLeaves[0].view;
        if (view instanceof PdfViewerView) {
          const file = view.getCurrentFile();
          if (file) nodes = [{ file, node: null }];
        }
      }
    }

    if (nodes.length === 0) {
      return '[No PDF files found. Open a canvas with PDF nodes, or open a PDF in the viewer.]';
    }

    const notice = new Notice(`Extracting text from ${nodes.length} PDF(s)…`, 0);
    try {
      const parts = await Promise.all(
        nodes.map(async ({ file }) => {
          const text = await this.pdfService.extractText(file);
          return `=== ${file.name} ===\n\n${text}`;
        }),
      );
      return parts.join('\n\n---\n\n');
    } finally {
      notice.hide();
    }
  }

  /**
   * Gather full canvas context: PDFs, text cards, file embeds, connections.
   * Gives Claude a holistic understanding of the entire canvas workspace.
   */
  async gatherCanvasContext(): Promise<string> {
    const canvas = this.getActiveCanvas();
    if (!canvas) {
      return '[No active canvas found. Open a canvas to use full canvas context.]';
    }
    const parts: string[] = [];

    // Collect all node content
    if (canvas.nodes) {
      const notice = new Notice('Reading canvas content…', 0);
      try {
        for (const node of canvas.nodes.values()) {
          const nodeData: CanvasNodeData = typeof node.getData === 'function'
            ? node.getData()
            : (node as unknown as CanvasNodeData);
          const nodeType = nodeData.type ?? 'unknown';

          if (nodeType === 'text' && nodeData.text) {
            parts.push(`[Text Card]\n${nodeData.text}`);
          } else if (nodeType === 'file' && node.file instanceof TFile) {
            if (node.file.extension === 'pdf') {
              try {
                const text = await this.pdfService.extractText(node.file);
                parts.push(`[PDF: ${node.file.name}]\n${text}`);
              } catch {
                parts.push(`[PDF: ${node.file.name}] (text extraction failed)`);
              }
            } else if (node.file.extension === 'md') {
              try {
                const content = await this.app.vault.cachedRead(node.file);
                parts.push(`[Note: ${node.file.name}]\n${content}`);
              } catch {
                parts.push(`[Note: ${node.file.name}] (read failed)`);
              }
            } else {
              parts.push(`[File: ${node.file.name}]`);
            }
          } else if (nodeType === 'link' && nodeData.url) {
            parts.push(`[Link: ${nodeData.url}]`);
          } else if (nodeType === 'group') {
            const label = nodeData.label || '(unnamed group)';
            parts.push(`[Group: ${label}]`);
          }
        }
      } finally {
        notice.hide();
      }
    }

    // Collect connections/edges for spatial context
    const edges = canvas.edges ?? canvas.data?.edges;
    if (edges && (edges instanceof Map || Array.isArray(edges))) {
      const edgeList: string[] = [];
      const iter = edges instanceof Map ? edges.values() : edges;
      for (const edge of iter) {
        const data = typeof edge.getData === 'function' ? edge.getData() : edge;
        if (data.fromNode && data.toNode) {
          edgeList.push(`${data.fromNode} -> ${data.toNode}`);
        }
      }
      if (edgeList.length > 0) {
        parts.push(`[Connections]\n${edgeList.join('\n')}`);
      }
    }

    if (parts.length === 0) {
      return '[Canvas is empty.]';
    }

    return parts.join('\n\n---\n\n');
  }

  // ─── Compound actions ──────────────────────────────────────────────────────

  private async openSelectedCanvasPdf(): Promise<void> {
    const nodes = getSelectedPdfNodes(this.app);
    if (nodes.length === 0) {
      new Notice('No PDF node selected on canvas.');
      return;
    }
    await this.activatePdfViewer();
    await this.openFileInViewer(nodes[0].file);
  }

  async askAboutPdfs(_scope: 'selected' | 'all'): Promise<void> {
    await this.activateAiSidebar();
    const view = this.getAiSidebarView();
    if (view) view.setContextScope('pdf');
    new Notice('Context set. Type your question in the sidebar.');
  }

  /**
   * Spread a PDF into individual page nodes on the canvas.
   * Each page becomes a text node with a %%pcai-spread:...:N%% marker
   * that our injector detects and replaces with a pdfjs page renderer.
   * This avoids file nodes (which trigger the PDF viewer intercept).
   */
  private async spreadPdfPages(file: TFile, node: unknown, direction: SpreadDirection): Promise<void> {
    const canvas = this.getActiveCanvas();
    if (!canvas) {
      new Notice('No active canvas found.');
      return;
    }

    // Load PDF to get page count and dimensions
    const buffer = await this.app.vault.readBinary(file);
    const pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
    const numPages = pdfDoc.numPages;

    if (numPages === 0) {
      void pdfDoc.destroy();
      new Notice('PDF has no pages.');
      return;
    }

    // Compute page aspect ratio from first page
    const firstPage = await pdfDoc.getPage(1);
    const vp = firstPage.getViewport({ scale: 1 });
    const aspectRatio = vp.height / vp.width;
    void pdfDoc.destroy();

    // Get the original node position
    const n = node as CanvasNode;
    const nodeData = typeof n.getData === 'function' ? n.getData() : n;
    const originX: number = nodeData.x ?? n.x ?? 0;
    const originY: number = nodeData.y ?? n.y ?? 0;

    const pageWidth = 400;
    const pageHeight = Math.round(pageWidth * aspectRatio);
    const gap = 20;

    const notice = new Notice(`Spreading ${numPages} pages…`, 0);
    try {
      // Remove the original node
      if (typeof canvas.removeNode === 'function') {
        canvas.removeNode(node);
      }

      // Create one text node per page with spread marker
      for (let i = 0; i < numPages; i++) {
        const pageNum = i + 1;
        const x = direction === 'right'
          ? originX + i * (pageWidth + gap)
          : originX;
        const y = direction === 'down'
          ? originY + i * (pageHeight + gap)
          : originY;
        const markerText = `%%pcai-spread:${file.path}:${pageNum}%%`;

        if (typeof canvas.createTextNode === 'function') {
          canvas.createTextNode({
            pos: { x, y },
            size: { width: pageWidth, height: pageHeight },
            text: markerText,
            focus: false,
            save: false,
          });
        } else {
          new Notice('Canvas API not available.');
          return;
        }
      }

      if (typeof canvas.requestSave === 'function') {
        canvas.requestSave();
      }

      new Notice(`Spread ${numPages} pages on canvas.`);
    } finally {
      notice.hide();
    }
  }

  /**
   * Extract the currently visible page of a PDF node as a standalone
   * single-page spread node, positioned to the right of the source.
   */
  private async extractCurrentPage(file: TFile, node: unknown): Promise<void> {
    const canvas = this.getActiveCanvas();
    if (!canvas) {
      new Notice('No active canvas found.');
      return;
    }

    const renderer = this.canvasInjector.getRendererForNode(node);
    const pageNum = renderer?.getCurrentVisiblePage() ?? 1;

    // Get source node position & dimensions
    const n = node as CanvasNode;
    const nodeData = typeof n.getData === 'function' ? n.getData() : n;
    const originX: number = nodeData.x ?? n.x ?? 0;
    const originY: number = nodeData.y ?? n.y ?? 0;
    const nodeW: number = nodeData.width ?? n.width ?? 400;

    // Compute page aspect ratio
    const buffer = await this.app.vault.readBinary(file);
    const pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
    const page = await pdfDoc.getPage(pageNum);
    const vp = page.getViewport({ scale: 1 });
    const aspectRatio = vp.height / vp.width;
    void pdfDoc.destroy();

    const pageWidth = 400;
    const pageHeight = Math.round(pageWidth * aspectRatio);
    const x = originX + nodeW + 40;
    const y = originY;
    const markerText = `%%pcai-spread:${file.path}:${pageNum}%%`;

    if (typeof canvas.createTextNode === 'function') {
      canvas.createTextNode({
        pos: { x, y },
        size: { width: pageWidth, height: pageHeight },
        text: markerText,
        focus: false,
      });
    } else {
      new Notice('Canvas API not available.');
      return;
    }

    if (typeof canvas.requestSave === 'function') {
      canvas.requestSave();
    }

    new Notice(`Extracted page ${pageNum}.`);
  }

  addToCanvas(file: TFile): void {
    const canvas = this.getActiveCanvas();
    if (!canvas) {
      new Notice('No active canvas. Open a canvas first.');
      return;
    }

    if (typeof canvas.createFileNode !== 'function' || typeof canvas.createTextNode !== 'function') {
      new Notice('Canvas API not available.');
      return;
    }

    // Find a sensible position — center of the current viewport or (0,0)
    const vp = typeof canvas.getViewportCenter === 'function' ? canvas.getViewportCenter() : null;
    const cx = vp?.x ?? 0;
    const cy = vp?.y ?? 0;

    const pdfWidth = 400;
    const pdfHeight = 500;
    const textWidth = 300;
    const textHeight = 200;
    const gap = 40;

    // Create PDF file node
    const pdfNode = canvas.createFileNode({
      pos: { x: cx - pdfWidth / 2, y: cy - pdfHeight / 2 },
      size: { width: pdfWidth, height: pdfHeight },
      file,
      save: false,
    });

    // Create connected text node to the right
    const textNode = canvas.createTextNode({
      pos: { x: cx + pdfWidth / 2 + gap, y: cy - textHeight / 2 },
      size: { width: textWidth, height: textHeight },
      text: `Notes: ${file.basename}`,
      focus: false,
      save: false,
    });

    // Create edge between them if the API supports it
    if (typeof canvas.createEdge === 'function' && pdfNode && textNode) {
      try {
        const fromId = pdfNode.id ?? (typeof pdfNode.getData === 'function' ? pdfNode.getData().id : null);
        const toId = textNode.id ?? (typeof textNode.getData === 'function' ? textNode.getData().id : null);
        if (fromId && toId) {
          canvas.createEdge({
            fromNode: pdfNode,
            fromSide: 'right',
            toNode: textNode,
            toSide: 'left',
            save: false,
          });
        }
      } catch (err) {
        console.warn('PDF Tools: Could not create edge between nodes:', err);
      }
    }

    if (typeof canvas.requestSave === 'function') {
      canvas.requestSave();
    }

    new Notice(`Added "${file.basename}" to canvas.`);
  }

  private async openFileInViewerAndAsk(file: TFile): Promise<void> {
    await this.activatePdfViewer();
    await this.openFileInViewer(file);
    await this.activateAiSidebar();
    const view = this.getAiSidebarView();
    if (view) {
      view.setCurrentPdf(file);
      view.setContextScope('pdf');
    }
  }

  /**
   * Gather vault-wide context by searching for files matching the query.
   * Uses a simple keyword search across vault file names and markdown content.
   */
  async gatherVaultContext(query: string): Promise<string> {
    const parts: string[] = [];
    const keywords = query.toLowerCase().split(/\s+/).filter((w) => w.length > 2);
    if (keywords.length === 0) {
      return '[No search keywords found in your question.]';
    }

    const allFiles = this.app.vault.getFiles();
    const matchingFiles: TFile[] = [];

    // Score files by keyword matches in name/path
    for (const file of allFiles) {
      if (file.extension !== 'md' && file.extension !== 'pdf') continue;
      const pathLower = file.path.toLowerCase();
      const score = keywords.filter((k) => pathLower.includes(k)).length;
      if (score > 0) matchingFiles.push(file);
    }

    // Also search markdown file content for keyword matches (limit search)
    const mdFiles = allFiles.filter((f) => f.extension === 'md');
    const contentSearchLimit = Math.min(mdFiles.length, 200);
    for (let i = 0; i < contentSearchLimit; i++) {
      const file = mdFiles[i];
      if (matchingFiles.includes(file)) continue;
      try {
        const content = await this.app.vault.cachedRead(file);
        const contentLower = content.toLowerCase();
        const score = keywords.filter((k) => contentLower.includes(k)).length;
        if (score >= Math.max(1, Math.floor(keywords.length / 2))) {
          matchingFiles.push(file);
        }
      } catch {
        // Skip unreadable files
      }
    }

    // Cap results
    const filesToInclude = matchingFiles.slice(0, 10);

    const notice = new Notice(`Reading ${filesToInclude.length} vault files…`, 0);
    try {
      for (const file of filesToInclude) {
        try {
          if (file.extension === 'pdf') {
            const text = await this.pdfService.extractText(file);
            parts.push(`[PDF: ${file.path}]\n${text}`);
          } else {
            const content = await this.app.vault.cachedRead(file);
            parts.push(`[Note: ${file.path}]\n${content}`);
          }
        } catch {
          parts.push(`[${file.path}] (read failed)`);
        }
      }
    } finally {
      notice.hide();
    }

    if (parts.length === 0) {
      return '[No matching files found in vault.]';
    }

    return parts.join('\n\n---\n\n');
  }
}
