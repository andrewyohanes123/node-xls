/**
 * Drawing anchors. Objects sized in pixels are converted to two-cell anchors
 * by walking the sheet's column widths and row heights, so every reader
 * places them identically.
 */
import { CellAddress, DrawingAnchor, WorkSheet } from '../types.js';
import { decode_cell } from '../utils/address.js';
import { DEFAULT_COL_WIDTH_PX, DEFAULT_ROW_HEIGHT_PX, EMU_PER_PIXEL, charsToWidth, pointsToPixels, widthToPixels } from '../utils/units.js';

export interface Marker {
  col: number;
  colOff: number; // px
  row: number;
  rowOff: number; // px
}

export interface ResolvedAnchor {
  from: Marker;
  to: Marker;
  /** Object size in px. */
  width: number;
  height: number;
}

const addr = (a: string | CellAddress): CellAddress => (typeof a === 'string' ? decode_cell(a) : a);

export function columnWidthPx(ws: WorkSheet, c: number): number {
  const col = ws['!cols']?.[c];
  if (!col) return DEFAULT_COL_WIDTH_PX;
  if (col.hidden) return 0;
  if (col.wpx !== undefined) return col.wpx;
  if (col.width !== undefined) return widthToPixels(col.width);
  if (col.wch !== undefined) return widthToPixels(charsToWidth(col.wch));
  return DEFAULT_COL_WIDTH_PX;
}

export function rowHeightPx(ws: WorkSheet, r: number): number {
  const row = ws['!rows']?.[r];
  if (!row) return DEFAULT_ROW_HEIGHT_PX;
  if (row.hidden) return 0;
  if (row.hpx !== undefined) return row.hpx;
  if (row.hpt !== undefined) return pointsToPixels(row.hpt);
  return DEFAULT_ROW_HEIGHT_PX;
}

/** Walks cells of the given sizes from (start, startOff) by `length` px; returns the end cell and offset. */
function advance(start: number, startOff: number, length: number, sizeOf: (i: number) => number, limit: number): [number, number] {
  let i = start;
  let remaining = startOff + length;
  while (i < limit) {
    const size = sizeOf(i);
    if (size > 0 && remaining < size) return [i, remaining];
    remaining -= size;
    i++;
  }
  return [limit - 1, sizeOf(limit - 1)];
}

const inGrid = (a: CellAddress) => a.c >= 0 && a.c <= 16383 && a.r >= 0 && a.r <= 1048575;

export function resolveAnchor(ws: WorkSheet, anchor: DrawingAnchor, natural?: { width: number; height: number }): ResolvedAnchor {
  const f = addr(anchor.from);
  if (!inGrid(f)) throw new Error(`Drawing anchor ${JSON.stringify(anchor.from)} is outside the sheet`);
  const from: Marker = { col: f.c, colOff: anchor.offsetX ?? 0, row: f.r, rowOff: anchor.offsetY ?? 0 };
  if (anchor.to) {
    const t = addr(anchor.to);
    const to: Marker = { col: t.c, colOff: anchor.toOffsetX ?? 0, row: t.r, rowOff: anchor.toOffsetY ?? 0 };
    if (!inGrid(t)) throw new Error(`Drawing anchor ${JSON.stringify(anchor.to)} is outside the sheet`);
    if (to.col < from.col || to.row < from.row || (to.col === from.col && to.colOff < from.colOff) || (to.row === from.row && to.rowOff < from.rowOff)) {
      throw new Error(`Drawing anchor "to" (${JSON.stringify(anchor.to)}) lies before "from" (${JSON.stringify(anchor.from)})`);
    }
    let width = to.colOff - from.colOff;
    for (let c = f.c; c < t.c; c++) width += columnWidthPx(ws, c);
    let height = to.rowOff - from.rowOff;
    for (let r = f.r; r < t.r; r++) height += rowHeightPx(ws, r);
    return { from, to, width: Math.max(1, width), height: Math.max(1, height) };
  }
  const width = Math.round(anchor.width ?? natural?.width ?? 480);
  const height = Math.round(anchor.height ?? natural?.height ?? 288);
  const [col, colOff] = advance(from.col, from.colOff, width, (c) => columnWidthPx(ws, c), 16384);
  const [row, rowOff] = advance(from.row, from.rowOff, height, (r) => rowHeightPx(ws, r), 1048576);
  return { from, to: { col, colOff, row, rowOff }, width, height };
}

export const emu = (px: number) => Math.round(px * EMU_PER_PIXEL);

function markerXml(tag: string, m: Marker): string {
  return `<xdr:${tag}><xdr:col>${m.col}</xdr:col><xdr:colOff>${emu(m.colOff)}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>${emu(
    m.rowOff
  )}</xdr:rowOff></xdr:${tag}>`;
}

/** Wraps an object (pic / graphicFrame) in a two-cell anchor. */
export function twoCellAnchorXml(a: ResolvedAnchor, objectXml: string, editAs: 'oneCell' | 'twoCell' | 'absolute'): string {
  return `<xdr:twoCellAnchor editAs="${editAs}">${markerXml('from', a.from)}${markerXml('to', a.to)}${objectXml}<xdr:clientData/></xdr:twoCellAnchor>`;
}

