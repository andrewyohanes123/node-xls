import { WorkBook, ParsingOptions, WritingOptions, WorkSheet } from '../types.js';
import { sheet_to_csv, aoa_to_sheet, sheet_to_html } from '../utils/sheet.js';
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

function detectSeparator(text: string): string {
  const firstLine = text.slice(0, text.search(/\r|\n/) === -1 ? undefined : text.search(/\r|\n/));
  const counts = [',', '\t', ';', '|'].map((s) => [s, firstLine.split(s).length - 1] as const);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ',';
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

export function readCsv(data: string, opts?: ParsingOptions): WorkBook {
  const text = data.charCodeAt(0) === 0xfeff ? data.slice(1) : data;
  const fs = opts?.FS || detectSeparator(text);
  let rows = parseCsv(text, fs);
  if (opts?.sheetRows && opts.sheetRows > 0) rows = rows.slice(0, opts.sheetRows);
  const converted = rows.map((row) =>
    row.map((val) => {
      if (opts?.raw) return val;
      const trimmed = val.trim();
      if (trimmed === '') return undefined;
      const upper = trimmed.toUpperCase();
      if (upper === 'TRUE') return true;
      if (upper === 'FALSE') return false;
      return parsePlainNumber(trimmed) ?? val;
    })
  );
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
  return opts?.bookType === 'txt' ? sheet_to_csv(ws, { FS: '\t' }) : sheet_to_csv(ws);
}

export function writeHtml(wb: WorkBook, opts?: WritingOptions): string {
  const ws = pickSheet(wb, opts?.sheet);
  return sheet_to_html(ws, {
    header: '<!DOCTYPE html><html><head><meta charset="utf-8"/><title>node-xls export</title></head><body>',
    footer: '</body></html>',
  });
}
