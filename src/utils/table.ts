/**
 * HTML table import. Reads cell text, `data-t`/`data-v`/`data-z` attributes
 * (as written by sheet_to_html), inline CSS styles, and colspan/rowspan merges.
 */
import { CellObject, Range, WorkBook, WorkSheet } from '../types.js';
import { encode_cell, encode_range } from './address.js';
import { book_new, book_append_sheet } from './book.js';
import { cssToStyle } from '../styles/css.js';
import { ERROR_VALUES } from './format.js';
import { parsePlainNumber } from '../csv/index.js';

export interface Table2SheetOpts {
  /** Keep every value as text (no number detection). */
  raw?: boolean;
  /** Sheet name for table_to_book. */
  sheet?: string;
  /** Ignore inline styles. */
  cellStyles?: boolean;
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e.toLowerCase()] ?? m;
  });
}

function parseAttrs(s: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of s.matchAll(/([\w:-]+)\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))/g)) {
    out[m[1].toLowerCase()] = decodeEntities(m[3] ?? m[4] ?? m[5] ?? '');
  }
  return out;
}

function htmlOf(table: string | any): string {
  if (typeof table === 'string') return table;
  if (table && typeof table.outerHTML === 'string') return table.outerHTML;
  if (table && typeof table.innerHTML === 'string') return `<table>${table.innerHTML}</table>`;
  return '';
}

export function table_to_sheet(table: string | any, opts?: Table2SheetOpts): WorkSheet {
  let html = htmlOf(table);
  const start = html.search(/<table[\s>]/i);
  if (start !== -1) {
    const end = html.search(/<\/table\s*>/i);
    html = html.slice(start, end === -1 ? undefined : end);
  }
  const ws: WorkSheet = {};
  const merges: Range[] = [];
  const occupied = new Set<string>();
  let maxR = -1;
  let maxC = -1;
  let R = -1;
  for (const tr of html.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr\s*>/gi)) {
    R++;
    let C = 0;
    for (const td of tr[1].matchAll(/<(td|th)\b([^>]*)>([\s\S]*?)<\/\1\s*>/gi)) {
      while (occupied.has(`${R},${C}`)) C++;
      const a = parseAttrs(td[2]);
      const colspan = Math.max(1, Number(a.colspan) || 1);
      const rowspan = Math.max(1, Number(a.rowspan) || 1);
      const text = decodeEntities(td[3].replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ''));
      const cell: CellObject = {};
      const t = a['data-t'];
      const dv = a['data-v'];
      if (t === 'n' && dv !== undefined && Number.isFinite(Number(dv))) Object.assign(cell, { t: 'n', v: Number(dv) });
      else if (t === 'b') Object.assign(cell, { t: 'b', v: dv === 'true' || dv === '1' || text.toUpperCase() === 'TRUE' });
      else if (t === 'e') Object.assign(cell, { t: 'e', v: ERROR_VALUES[text.trim()] ?? 0x2a, w: text.trim() });
      else if (t === 'd' && dv) Object.assign(cell, { t: 'd', v: new Date(dv) });
      else if (t === 's') Object.assign(cell, { t: 's', v: dv ?? text });
      else {
        const trimmed = text.trim();
        if (trimmed === '') Object.assign(cell, { t: 'z' });
        else if (!opts?.raw && parsePlainNumber(trimmed) !== undefined) Object.assign(cell, { t: 'n', v: parsePlainNumber(trimmed) });
        else Object.assign(cell, { t: 's', v: text.replace(/^\s+|\s+$/g, '') });
      }
      if (a['data-z']) cell.z = a['data-z'];
      if (a.style && opts?.cellStyles !== false) {
        const s = cssToStyle(a.style);
        if (s) cell.s = s;
      }
      const link = /<a\b([^>]*)>/i.exec(td[3]);
      if (link) {
        const la = parseAttrs(link[1]);
        if (la.href) cell.l = { Target: la.href, ...(la.title ? { Tooltip: la.title } : {}) };
      }
      if (cell.t !== 'z' || cell.s) {
        ws[encode_cell({ r: R, c: C })] = cell;
      }
      if (colspan > 1 || rowspan > 1) merges.push({ s: { r: R, c: C }, e: { r: R + rowspan - 1, c: C + colspan - 1 } });
      for (let dr = 0; dr < rowspan; dr++) for (let dc = 0; dc < colspan; dc++) occupied.add(`${R + dr},${C + dc}`);
      maxR = Math.max(maxR, R + rowspan - 1);
      maxC = Math.max(maxC, C + colspan - 1);
      C += colspan;
    }
  }
  if (maxR >= 0 && maxC >= 0) ws['!ref'] = encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  if (merges.length) ws['!merges'] = merges;
  return ws;
}

export function table_to_book(table: string | any, opts?: Table2SheetOpts): WorkBook {
  const wb = book_new();
  book_append_sheet(wb, table_to_sheet(table, opts), opts?.sheet || 'Sheet1');
  return wb;
}
