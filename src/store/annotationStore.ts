import type { Highlight, HighlightColor, PageRect } from '../types/annotations';
import type PdfCanvasAiPlugin from '../main';
import { nanoid } from 'nanoid';

export interface ReadingProgress {
  lastPage: number;
  totalPages: number;
  lastOpened: number; // timestamp
}

interface StoredData {
  version: 1;
  /** Map from vault-relative PDF path → array of highlights */
  highlights: Record<string, Highlight[]>;
  /** Map from vault-relative PDF path → reading progress */
  readingProgress: Record<string, ReadingProgress>;
}

export class AnnotationStore {
  private data: StoredData = { version: 1, highlights: {}, readingProgress: {} };
  private saveTimer: ReturnType<typeof setTimeout> | null = null;
  private plugin: PdfCanvasAiPlugin;

  constructor(plugin: PdfCanvasAiPlugin) {
    this.plugin = plugin;
  }

  async load(): Promise<void> {
    const raw = await this.plugin.loadData() as Record<string, unknown> | null;
    if (raw?.version === 1 && typeof raw.highlights === 'object') {
      this.data = raw as unknown as StoredData;
      // Ensure readingProgress exists (migration from older data)
      if (!this.data.readingProgress) {
        this.data.readingProgress = {};
      }
    }
  }

  private scheduleSave(): void {
    if (this.saveTimer) clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => {
      this.plugin.saveData(this.data).catch((e: unknown) => {
        console.error('PDF Tools: annotation save failed', e);
      });
    }, 500);
  }

  getForFile(pdfPath: string): Highlight[] {
    return this.data.highlights[pdfPath] ?? [];
  }

  getForPage(pdfPath: string, pageNumber: number): Highlight[] {
    return this.getForFile(pdfPath).filter((h) => h.pageNumber === pageNumber);
  }

  add(
    pdfPath: string,
    pageNumber: number,
    text: string,
    color: HighlightColor,
    rects: PageRect[],
  ): Highlight {
    const highlight: Highlight = {
      id: nanoid(),
      pdfPath,
      pageNumber,
      text,
      color,
      rects,
      createdAt: Date.now(),
    };

    if (!this.data.highlights[pdfPath]) {
      this.data.highlights[pdfPath] = [];
    }
    this.data.highlights[pdfPath].push(highlight);
    this.scheduleSave();
    return highlight;
  }

  remove(id: string): void {
    for (const path of Object.keys(this.data.highlights)) {
      const arr = this.data.highlights[path];
      const idx = arr.findIndex((h) => h.id === id);
      if (idx !== -1) {
        arr.splice(idx, 1);
        if (arr.length === 0) delete this.data.highlights[path];
        this.scheduleSave();
        return;
      }
    }
  }

  updateNote(id: string, note: string): void {
    for (const arr of Object.values(this.data.highlights)) {
      const h = arr.find((h) => h.id === id);
      if (h) {
        h.note = note;
        this.scheduleSave();
        return;
      }
    }
  }

  updateNotePath(id: string, notePath: string): void {
    for (const arr of Object.values(this.data.highlights)) {
      const h = arr.find((h) => h.id === id);
      if (h) {
        h.notePath = notePath || undefined;
        this.scheduleSave();
        return;
      }
    }
  }

  getById(id: string): Highlight | undefined {
    for (const arr of Object.values(this.data.highlights)) {
      const h = arr.find((h) => h.id === id);
      if (h) return h;
    }
    return undefined;
  }

  getAllHighlights(): Highlight[] {
    const all: Highlight[] = [];
    for (const arr of Object.values(this.data.highlights)) {
      all.push(...arr);
    }
    return all;
  }

  updateReadingProgress(pdfPath: string, lastPage: number, totalPages: number): void {
    this.data.readingProgress[pdfPath] = {
      lastPage,
      totalPages,
      lastOpened: Date.now(),
    };
    this.scheduleSave();
  }

  getReadingProgress(pdfPath: string): ReadingProgress | undefined {
    return this.data.readingProgress[pdfPath];
  }

  getAllReadingProgress(): Record<string, ReadingProgress> {
    return this.data.readingProgress;
  }

  destroy(): void {
    if (this.saveTimer !== null) {
      clearTimeout(this.saveTimer);
      this.saveTimer = null;
    }
  }

  /** Update notePath references when a note file is renamed */
  renameNoteFile(oldNotePath: string, newNotePath: string): void {
    let changed = false;
    for (const arr of Object.values(this.data.highlights)) {
      for (const h of arr) {
        if (h.notePath === oldNotePath) {
          h.notePath = newNotePath;
          changed = true;
        }
      }
    }
    if (changed) this.scheduleSave();
  }

  renameFile(oldPath: string, newPath: string): void {
    let changed = false;
    if (this.data.highlights[oldPath]) {
      this.data.highlights[newPath] = this.data.highlights[oldPath].map((h) => ({
        ...h,
        pdfPath: newPath,
      }));
      delete this.data.highlights[oldPath];
      changed = true;
    }
    if (this.data.readingProgress[oldPath]) {
      this.data.readingProgress[newPath] = this.data.readingProgress[oldPath];
      delete this.data.readingProgress[oldPath];
      changed = true;
    }
    if (changed) this.scheduleSave();
  }
}
