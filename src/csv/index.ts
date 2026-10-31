import { WorkBook, ParsingOptions, WritingOptions, WorkSheet, CellObject, Sheet2CSVOpts } from '../types.js';
import { sheet_to_csv, sheet_to_txt, aoa_to_sheet, sheet_to_html } from '../utils/sheet.js';
import { book_new, book_append_sheet } from '../utils/book.js';

/** RFC 4180 parser (quoted fields, escaped quotes, CRLF/LF/CR line breaks). */
export function parseCsv(text: string, fs = ','): string[][] {
  const lines: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  const len = text.length;
  for (let i = 0; i < len; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && field === '') quoted = true;
    else if (ch === fs) {
      row.push(field);
      field = '';
    } else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      lines.push(row);
      row = [];
      field = '';
    } else field += ch;
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    lines.push(row);
  }
  return lines;
}

/** Picks the separator occurring most often outside quotes in the first record. */
function detectSeparator(text: string): string {
  const counts: Record<string, number> = { ',': 0, '\t': 0, ';': 0, '|': 0 };
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') quoted = !quoted;
    else if (!quoted && (ch === '\n' || ch === '\r')) break;
    else if (!quoted && ch in counts) counts[ch]++;
  }
  const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
  return best[1] > 0 ? best[0] : ',';
}

/**
 * Plain decimal numbers ("1.50", "-3", "1e3", ".5"). Codes with leading zeros
 * ("007"), hex and thousands separators stay text (SheetJS converts them).
 */
export function parsePlainNumber(s: string): number | undefined {
  const t = s.trim();
  if (!/^[+-]?(?:(?:0|[1-9]\d*)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(t)) return undefined;
  const n = Number(t);
  if (!Number.isFinite(n)) return undefined;
  return n === 0 ? 0 : n;
}

/** Typed cell for an imported text value; `w` keeps the source text (as SheetJS does). */
export function textToCell(val: string, raw?: boolean): CellObject | undefined {
  if (raw) return { t: 's', v: val };
  const trimmed = val.trim();
  if (trimmed === '') return undefined;
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') return { t: 'b', v: upper === 'TRUE', w: trimmed };
  const n = parsePlainNumber(trimmed);
  if (n !== undefined) return { t: 'n', v: n, w: trimmed };
  return { t: 's', v: val };
}

export function readCsv(data: string, opts?: ParsingOptions): WorkBook {
  let text = data.charCodeAt(0) === 0xfeff ? data.slice(1) : data;
  // Excel's "sep=;" directive line
  const sep = /^sep=(.)\r?\n/i.exec(text);
  if (sep) text = text.slice(sep[0].length);
  const fs = opts?.FS || sep?.[1] || detectSeparator(text);
  let rows = parseCsv(text, fs);
  if (opts?.sheetRows && opts.sheetRows > 0) rows = rows.slice(0, opts.sheetRows);
  const converted = rows.map((row) => row.map((val) => textToCell(val, opts?.raw)));
  const wb = book_new();
  book_append_sheet(wb, aoa_to_sheet(converted), 'Sheet1');
  return wb;
}

export function pickSheet(wb: WorkBook, sheet?: string | number): WorkSheet {
  const name = typeof sheet === 'number' ? wb.SheetNames[sheet] : (sheet ?? wb.SheetNames[0]);
  if (sheet !== undefined && (name === undefined || !wb.Sheets[name])) throw new Error(`Cannot find sheet |${sheet}|`);
  return (name !== undefined && wb.Sheets[name]) || {};
}

export function writeCsv(wb: WorkBook, opts?: WritingOptions): string {
  const ws = pickSheet(wb, opts?.sheet);
  const o: Sheet2CSVOpts = {
    FS: opts?.FS,
    RS: opts?.RS,
    strip: opts?.strip,
    blankrows: opts?.blankrows,
    forceQuotes: opts?.forceQuotes,
    skipHidden: opts?.skipHidden,
    dateNF: opts?.dateNF,
    rawNumbers: opts?.rawNumbers,
  };
  return opts?.bookType === 'txt' ? sheet_to_txt(ws, o) : sheet_to_csv(ws, o);
}

export function writeHtml(wb: WorkBook, opts?: WritingOptions): string {
  const ws = pickSheet(wb, opts?.sheet);
  return sheet_to_html(ws, {
    header: '<!DOCTYPE html><html><head><meta charset="utf-8"/><title>node-xls export</title></head><body>',
    footer: '</body></html>',
  });
}
