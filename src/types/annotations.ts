export type HighlightColor = 'yellow' | 'green' | 'blue' | 'pink' | 'red';

export interface PageRect {
  /** 0–1, relative to page width */
  x: number;
  /** 0–1, relative to page height */
  y: number;
  width: number;
  height: number;
}

export interface Highlight {
  id: string;
  /** Vault-relative path to the PDF file */
  pdfPath: string;
  pageNumber: number;
  text: string;
  color: HighlightColor;
  /** One rect per text line, in page-relative coordinates */
  rects: PageRect[];
  createdAt: number;
  note?: string;
  /** Vault-relative path to the markdown note file for this annotation */
  notePath?: string;
}

export const HIGHLIGHT_COLORS: HighlightColor[] = ['yellow', 'green', 'blue', 'pink', 'red'];

export const COLOR_LABELS: Record<HighlightColor, string> = {
  yellow: '🟡',
  green: '🟢',
  blue: '🔵',
  pink: '🌸',
  red: '🔴',
};

export const COLOR_HEX: Record<HighlightColor, string> = {
  yellow: '#FFE066',
  green: '#66FF99',
  blue: '#66B2FF',
  pink: '#FF99CC',
  red: '#FF7777',
};
