/**
 * Column/row unit conversions (Calibri 11: max digit width 7px, 96 DPI),
 * matching SheetJS so widths round-trip identically.
 */
export const MAX_DIGIT_WIDTH = 7;

/** Characters ('wch') -> stored column width. */
export function charsToWidth(wch: number): number {
  return Math.round(((wch * MAX_DIGIT_WIDTH + 5) / MAX_DIGIT_WIDTH) * 256) / 256;
}

/** Pixels -> characters. */
export function pixelsToChars(px: number): number {
  return Math.floor(((px - 5) / MAX_DIGIT_WIDTH) * 100 + 0.5) / 100;
}

/** Stored column width -> pixels. */
export function widthToPixels(width: number): number {
  return Math.floor((width + Math.round(128 / MAX_DIGIT_WIDTH) / 256) * MAX_DIGIT_WIDTH);
}

export function widthToChars(width: number): number {
  return pixelsToChars(widthToPixels(width));
}

export const pointsToPixels = (pt: number): number => (pt * 96) / 72;
export const pixelsToPoints = (px: number): number => (px * 72) / 96;

/** Default column width in pixels / default row height in pixels. */
export const DEFAULT_COL_WIDTH_PX = 64;
export const DEFAULT_ROW_HEIGHT_PX = 20;

/** EMUs per pixel (DrawingML). */
export const EMU_PER_PIXEL = 9525;
