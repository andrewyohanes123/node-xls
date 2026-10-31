/**
 * Writes a pivot layout as ordinary cells (Excel tabular form), so the
 * summary is visible in every viewer, including ones that ignore pivot parts.
 * Excel re-renders the same area when it refreshes the pivot on open.
 */
import { CellObject, CellStyle, PivotTableDefinition, Range, WorkSheet } from '../types.js';
import { decode_range, encode_cell, encode_range } from '../utils/address.js';
import { ERROR_VALUES } from '../utils/format.js';
import { PivotLayout, itemLabel } from './layout.js';
import { PivotValue, isPivotError } from './source.js';

// Style factories: every rendered cell gets its own objects, so editing one never changes another.
const bold = (): CellStyle => ({ font: { bold: true } });
const header = (): CellStyle => ({
  font: { bold: true },
  fill: { fgColor: { theme: 4, tint: 0.8 } },
  border: { bottom: { style: 'thin', color: { theme: 4, tint: 0.4 } } },
});
const total = (): CellStyle => ({
  font: { bold: true },
  fill: { fgColor: { theme: 4, tint: 0.8 } },
  border: { top: { style: 'thin', color: { theme: 4, tint: 0.4 } } },
});

function errorCell(error: string, s?: CellStyle): CellObject {
  return { t: 'e', v: ERROR_VALUES[error] ?? error, w: error, ...(s ? { s } : {}) };
}

function itemCell(layout: PivotLayout, field: number, v: PivotValue, s?: CellStyle): CellObject {
  const style = s ? { s } : {};
  if (v === null) return { t: 's', v: '(blank)', ...style };
  if (isPivotError(v)) return errorCell(v.error, s);
  if (typeof v === 'number') {
    const z = layout.source.itemFormats[field]?.get(v);
    return { t: 'n', v, ...(z !== undefined ? { z } : {}), ...style };
  }
  if (typeof v === 'boolean') return { t: 'b', v, ...style };
  return { t: 's', v, ...style };
}

const text = (v: string, s?: CellStyle): CellObject => ({ t: 's', v, ...(s ? { s } : {}) });

function commonPrefix(a: number[], b: number[]): number {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return i;
}

/** Full sheet area a pivot occupies: optional title and page fields above the pivot location. */
export function pivotArea(layout: PivotLayout): Range {
  const g = layout.geometry;
  const top = Math.min(g.location.s.r, g.title?.r ?? Infinity, g.pages[0]?.r ?? Infinity);
  return { s: { r: top, c: g.location.s.c }, e: { r: g.location.e.r, c: Math.max(g.location.e.c, g.pages.length ? g.location.s.c + 1 : 0) } };
}

export function renderPivot(ws: WorkSheet, def: PivotTableDefinition, layout: PivotLayout): Range {
  const { geometry: g, rowFields, colFields, dataFields, rowAxis, colAxis, valuesOnCols, items, source, date1904 } = layout;
  const area = pivotArea(layout);
  // the pivot owns its area: clear stale cells first
  for (let r = area.s.r; r <= area.e.r; r++) for (let c = area.s.c; c <= area.e.c; c++) delete ws[encode_cell({ r, c })];

  const set = (r: number, c: number, cell: CellObject) => {
    ws[encode_cell({ r, c })] = cell;
  };
  const label = (field: number, idx: number) => itemLabel(source, field, items.get(field)![idx], date1904);
  const k = rowFields.length;
  const top = g.location.s.r;
  const left = g.location.s.c;
  const m = g.headerRows - 1;

  if (g.title && def.title) set(g.title.r, g.title.c, text(def.title, { font: { bold: true, sz: 13 } }));
  layout.pageFields.forEach((p, i) => {
    const pos = g.pages[i];
    set(pos.r, pos.c, text(source.fields[p.field], bold()));
    if (p.selected === undefined) set(pos.r, pos.c + 1, text('(All)'));
    else set(pos.r, pos.c + 1, itemCell(layout, p.field, items.get(p.field)![p.selected]));
  });

  // header rows
  for (let r = 0; r <= m; r++) for (let c = 0; c < k + colAxis.length; c++) set(top + r, left + c, { t: 'z', s: header() });
  if (m === 0) {
    rowFields.forEach((f, i) => set(top, left + i, text(source.fields[f], header())));
    set(top, left + k, text(dataFields[0].name, header()));
  } else {
    if (!valuesOnCols) set(top, left, text(dataFields[0].name, header()));
    colFields.forEach((f, j) => set(top, left + k + j, text(source.fields[f], header())));
    if (valuesOnCols) set(top, left + k + colFields.length, text('Values', header()));
    rowFields.forEach((f, i) => set(top + m, left + i, text(source.fields[f], header())));

    let prev: number[] = [];
    colAxis.forEach((e, ci) => {
      const col = left + k + ci;
      if (e.type === 'grand') {
        set(top + 1, col, text(valuesOnCols ? `Total ${dataFields[e.data].name}` : 'Grand Total', header()));
        prev = [];
        return;
      }
      if (e.type === 'subtotal') {
        const level = e.path.length - 1;
        const caption = valuesOnCols ? `${label(colFields[level], e.path[level])} ${dataFields[e.data].name}` : `${label(colFields[level], e.path[level])} Total`;
        set(top + 1 + level, col, text(caption, header()));
        prev = e.path;
        return;
      }
      const from = commonPrefix(e.path, prev);
      e.path.forEach((idx, level) => {
        if (level < from) return;
        if (level < colFields.length) set(top + 1 + level, col, itemCell(layout, colFields[level], items.get(colFields[level])![idx], header()));
        else set(top + 1 + level, col, text(dataFields[idx].name, header()));
      });
      prev = e.path;
    });
  }

  // body
  let prev: number[] = [];
  rowAxis.forEach((e, ri) => {
    const r = top + g.headerRows + ri;
    const isTotal = e.type !== 'item';
    if (isTotal) for (let c = 0; c < k; c++) set(r, left + c, { t: 'z', s: total() });
    if (e.type === 'grand') {
      set(r, left, text('Grand Total', total()));
      prev = [];
    } else if (e.type === 'subtotal') {
      const level = e.path.length - 1;
      set(r, left + level, text(`${label(rowFields[level], e.path[level])} Total`, total()));
      // the next item shares the outer levels, so their labels are not repeated
      prev = e.path;
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
      const base = isTotal ? total() : ce.type !== 'item' ? bold() : undefined;
      const s: CellStyle | undefined = base || df.numFmt ? { ...(base ?? {}), ...(df.numFmt ? { numFmt: df.numFmt } : {}) } : undefined;
      let cell: CellObject;
      if (v === null) cell = { t: 'z' };
      else if (isPivotError(v)) cell = errorCell(v.error);
      else cell = { t: 'n', v };
      if (s) cell.s = s;
      if (v !== null || s) set(r, left + k + ci, cell);
    });
  });

  // extend !ref over everything written
  const ref: Range = { s: { ...area.s }, e: { ...area.e } };
  if (ws['!ref']) {
    const cur = decode_range(ws['!ref']);
    ref.s.r = Math.min(ref.s.r, cur.s.r);
    ref.s.c = Math.min(ref.s.c, cur.s.c);
    ref.e.r = Math.max(ref.e.r, cur.e.r);
    ref.e.c = Math.max(ref.e.c, cur.e.c);
  }
  ws['!ref'] = encode_range(ref);
  return g.location;
}
