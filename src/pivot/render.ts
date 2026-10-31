/**
 * Writes a pivot layout as ordinary cells (Excel tabular form), so the
 * summary is visible in every viewer, including ones that ignore pivot parts.
 * Excel re-renders the same area when it refreshes the pivot on open.
 */
import { CellObject, CellStyle, PivotTableDefinition, Range, WorkSheet } from '../types.js';
import { decode_range, encode_cell, encode_range } from '../utils/address.js';
import { PivotLayout, itemLabel } from './layout.js';
import { PivotValue } from './source.js';

const BOLD: CellStyle = { font: { bold: true } };
const HEADER: CellStyle = { font: { bold: true }, fill: { fgColor: { theme: 4, tint: 0.8 } }, border: { bottom: { style: 'thin', color: { theme: 4, tint: 0.4 } } } };
const TOTAL: CellStyle = { font: { bold: true }, fill: { fgColor: { theme: 4, tint: 0.8 } }, border: { top: { style: 'thin', color: { theme: 4, tint: 0.4 } } } };

function itemCell(layout: PivotLayout, field: number, v: PivotValue, style?: CellStyle): CellObject {
  const { source } = layout;
  if (v === null) return { t: 's', v: '(blank)', ...(style ? { s: style } : {}) };
  if (typeof v === 'number') {
    const z = source.dateFields.has(field) ? (source.formats[field] ?? 14) : source.formats[field];
    return { t: 'n', v, ...(z !== undefined ? { z } : {}), ...(style ? { s: style } : {}) };
  }
  if (typeof v === 'boolean') return { t: 'b', v, ...(style ? { s: style } : {}) };
  return { t: 's', v, ...(style ? { s: style } : {}) };
}

const text = (v: string, s?: CellStyle): CellObject => ({ t: 's', v, ...(s ? { s } : {}) });

function commonPrefix(a: number[], b: number[]): number {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return i;
}

export function renderPivot(ws: WorkSheet, def: PivotTableDefinition, layout: PivotLayout): Range {
  const { geometry: g, rowFields, colFields, dataFields, rowAxis, colAxis, valuesOnCols, items, source } = layout;
  const set = (r: number, c: number, cell: CellObject) => {
    ws[encode_cell({ r, c })] = cell;
  };
  const k = rowFields.length;
  const top = g.location.s.r;
  const left = g.location.s.c;
  const m = g.headerRows - 1;

  if (g.title && def.title) set(g.title.r, g.title.c, text(def.title, { font: { bold: true, sz: 13 } }));
  layout.pageFields.forEach((p, i) => {
    const pos = g.pages[i];
    set(pos.r, pos.c, text(source.fields[p.field], BOLD));
    if (p.selected === undefined) set(pos.r, pos.c + 1, text('(All)'));
    else set(pos.r, pos.c + 1, itemCell(layout, p.field, items.get(p.field)![p.selected]));
  });

  // header rows
  for (let r = 0; r <= m; r++) for (let c = 0; c < k + colAxis.length; c++) set(top + r, left + c, { t: 'z', s: HEADER });
  if (m === 0) {
    rowFields.forEach((f, i) => set(top, left + i, text(source.fields[f], HEADER)));
    set(top, left + k, text(dataFields[0].name, HEADER));
  } else {
    if (!valuesOnCols) set(top, left, text(dataFields[0].name, HEADER));
    colFields.forEach((f, j) => set(top, left + k + j, text(source.fields[f], HEADER)));
    if (valuesOnCols) set(top, left + k + colFields.length, text('Values', HEADER));
    rowFields.forEach((f, i) => set(top + m, left + i, text(source.fields[f], HEADER)));

    let prev: number[] = [];
    colAxis.forEach((e, ci) => {
      const col = left + k + ci;
      if (e.type === 'grand') {
        set(top + 1, col, text(valuesOnCols ? `Total ${dataFields[e.data].name}` : 'Grand Total', HEADER));
        prev = [];
        return;
      }
      if (e.type === 'subtotal') {
        const level = e.path.length - 1;
        const field = colFields[level];
        set(top + 1 + level, col, text(`${itemLabel(source, field, items.get(field)![e.path[level]])} Total`, HEADER));
        if (valuesOnCols) set(top + m, col, text(dataFields[e.data].name, HEADER));
        prev = [];
        return;
      }
      const from = commonPrefix(e.path, prev);
      e.path.forEach((idx, level) => {
        if (level < from) return;
        if (level < colFields.length) {
          const field = colFields[level];
          set(top + 1 + level, col, { ...itemCell(layout, field, items.get(field)![idx]), s: HEADER });
        } else {
          set(top + 1 + level, col, text(dataFields[idx].name, HEADER));
        }
      });
      prev = e.path;
    });
  }

  // body
  let prev: number[] = [];
  rowAxis.forEach((e, ri) => {
    const r = top + g.headerRows + ri;
    const style = e.type === 'item' ? undefined : TOTAL;
    if (style) for (let c = 0; c < k; c++) set(r, left + c, { t: 'z', s: style });
    if (e.type === 'grand') {
      set(r, left, text('Grand Total', TOTAL));
      prev = [];
    } else if (e.type === 'subtotal') {
      const level = e.path.length - 1;
      const field = rowFields[level];
      set(r, left + level, text(`${itemLabel(source, field, items.get(field)![e.path[level]])} Total`, TOTAL));
      prev = [];
    } else {
      const from = commonPrefix(e.path, prev);
      e.path.forEach((idx, level) => {
        if (level >= from) set(r, left + level, itemCell(layout, rowFields[level], items.get(rowFields[level])![idx]));
      });
      prev = e.path;
    }
    colAxis.forEach((ce, ci) => {
      const v = layout.values[ri][ci];
      const df = dataFields[ce.data];
      const s: CellStyle | undefined = style || ce.type !== 'item' ? { ...(style ?? BOLD), ...(df.numFmt ? { numFmt: df.numFmt } : {}) } : df.numFmt ? { numFmt: df.numFmt } : undefined;
      const cell: CellObject = v === null ? { t: 'z' } : { t: 'n', v };
      if (s) cell.s = s;
      if (v !== null || s) set(r, left + k + ci, cell);
    });
  });

  // extend !ref over everything written
  const area: Range = {
    s: { r: Math.min(g.location.s.r, g.title?.r ?? Infinity, g.pages[0]?.r ?? Infinity), c: g.location.s.c },
    e: { r: g.location.e.r, c: Math.max(g.location.e.c, g.pages.length ? g.location.s.c + 1 : 0) },
  };
  if (ws['!ref']) {
    const cur = decode_range(ws['!ref']);
    area.s.r = Math.min(area.s.r, cur.s.r);
    area.s.c = Math.min(area.s.c, cur.s.c);
    area.e.r = Math.max(area.e.r, cur.e.r);
    area.e.c = Math.max(area.e.c, cur.e.c);
  }
  ws['!ref'] = encode_range(area);
  return g.location;
}
