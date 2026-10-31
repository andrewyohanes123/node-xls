/**
 * Pivot source data: header names (made unique, as Excel requires for cache
 * fields) and typed record values read from a worksheet range.
 */
import { CellObject, Range, WorkSheet } from '../types.js';
import { MAX_ROW, decode_range, encode_cell, encode_col } from '../utils/address.js';
import { errorText, format_cell, isDateFormat } from '../utils/format.js';
import { dateToSerial } from '../utils/dates.js';

/** An error value in the source (#N/A, #DIV/0!, ...). */
export interface PivotError {
  error: string;
}

/** A cache value; null is a blank. */
export type PivotValue = string | number | boolean | null | PivotError;

export const isPivotError = (v: unknown): v is PivotError => typeof v === 'object' && v !== null && 'error' in v;

export interface SourceData {
  fields: string[];
  /** records[row][field] */
  records: PivotValue[][];
  /** Number format of the first data cell per field (cache field format, date detection). */
  formats: Array<string | undefined>;
  /** Number format of the first cell holding each numeric item, per field (item labels). */
  itemFormats: Array<Map<number, string | number>>;
  /** Fields whose values are dates (stored as serial numbers). */
  dateFields: Set<number>;
  /** Range actually used: header row through the last non-empty row. */
  range: Range;
}

function cellValue(cell: CellObject | undefined, date1904: boolean): PivotValue {
  if (!cell || cell.t === 'z') return null;
  if (cell.t === 'e') return { error: errorText(cell.v ?? cell.w) };
  if (cell.v === undefined || cell.v === null) return null;
  if (cell.v instanceof Date) return dateToSerial(cell.v, { date1904 });
  if (typeof cell.v === 'string') return cell.v === '' ? null : cell.v;
  if (typeof cell.v === 'number' || typeof cell.v === 'boolean') return cell.v;
  return String(cell.v);
}

export function readSource(ws: WorkSheet, sourceRange: Range, opts: { date1904?: boolean } = {}): SourceData {
  const date1904 = Boolean(opts.date1904);
  // whole-column sources ('A:D') end at the sheet's used range
  let range = sourceRange;
  if (ws['!ref']) {
    const used = decode_range(ws['!ref']);
    range = { s: sourceRange.s, e: { r: Math.min(sourceRange.e.r, Math.max(used.e.r, sourceRange.s.r)), c: Math.min(sourceRange.e.c, Math.max(used.e.c, sourceRange.s.c)) } };
  }
  const fields: string[] = [];
  const seen = new Set<string>();
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[encode_cell({ r: range.s.r, c })] as CellObject | undefined;
    // Excel matches cache fields to the header text exactly (spaces included) when it refreshes
    const text = cell && cell.v !== undefined && cell.v !== null ? format_cell(cell) : '';
    if (text.trim() === '') {
      throw new Error(`Pivot source header in column ${encode_col(c)} is blank; Excel requires a name for every field in the source range`);
    }
    let name = text;
    for (let n = 2; seen.has(name.toLowerCase()); n++) name = `${text}${n}`;
    seen.add(name.toLowerCase());
    fields.push(name);
  }

  let lastRow = range.s.r;
  for (let r = range.e.r; r > range.s.r; r--) {
    let any = false;
    for (let c = range.s.c; c <= range.e.c && !any; c++) any = cellValue(ws[encode_cell({ r, c })], date1904) !== null;
    if (any) {
      lastRow = r;
      break;
    }
  }

  const records: PivotValue[][] = [];
  const formats: Array<string | undefined> = fields.map(() => undefined);
  const itemFormats = fields.map(() => new Map<number, string | number>());
  const dateFields = new Set<number>();
  for (let r = range.s.r + 1; r <= lastRow; r++) {
    const rec: PivotValue[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const i = c - range.s.c;
      const cell = ws[encode_cell({ r, c })] as CellObject | undefined;
      const v = cellValue(cell, date1904);
      rec.push(v);
      if (v === null || isPivotError(v)) continue;
      const z = cell?.s?.numFmt ?? cell?.z;
      const isDate = cell?.t === 'd' || cell?.v instanceof Date;
      if (formats[i] === undefined) {
        formats[i] = typeof z === 'string' ? z : undefined;
        if (isDate || (typeof v === 'number' && isDateFormat(z))) dateFields.add(i);
      }
      if (typeof v === 'number' && !itemFormats[i].has(v)) {
        const fmt = z ?? (isDate ? 14 : undefined);
        if (fmt !== undefined) itemFormats[i].set(v, fmt);
      }
    }
    records.push(rec);
  }
  // A whole-column source also covers the empty rows below the data: Excel shows them as one "(blank)" item.
  if (sourceRange.e.r === MAX_ROW && lastRow < MAX_ROW) records.push(fields.map(() => null));
  return { fields, records, formats, itemFormats, dateFields, range: { s: range.s, e: { r: lastRow, c: range.e.c } } };
}
