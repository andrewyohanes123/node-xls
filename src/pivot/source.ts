/**
 * Pivot source data: header names (made unique, as Excel requires for cache
 * fields) and typed record values read from a worksheet range.
 */
import { CellObject, Range, WorkSheet } from '../types.js';
import { encode_cell } from '../utils/address.js';
import { format_cell, isDateFormat } from '../utils/format.js';
import { dateToSerial } from '../utils/dates.js';

/** A cache value; null is a blank. */
export type PivotValue = string | number | boolean | null;

export interface SourceData {
  fields: string[];
  /** records[row][field] */
  records: PivotValue[][];
  /** Number format of the first data cell per field (for labels and date handling). */
  formats: Array<string | undefined>;
  /** Fields whose values are dates (stored as serial numbers). */
  dateFields: Set<number>;
  /** Range actually used: header row through the last non-empty row. */
  range: Range;
}

function cellValue(cell: CellObject | undefined): PivotValue {
  if (!cell || cell.v === undefined || cell.v === null || cell.t === 'z' || cell.t === 'e') return null;
  if (cell.v instanceof Date) return dateToSerial(cell.v);
  if (typeof cell.v === 'string') return cell.v === '' ? null : cell.v;
  if (typeof cell.v === 'number' || typeof cell.v === 'boolean') return cell.v;
  return String(cell.v);
}

export function readSource(ws: WorkSheet, range: Range): SourceData {
  const fields: string[] = [];
  const seen = new Set<string>();
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[encode_cell({ r: range.s.r, c })] as CellObject | undefined;
    const text = cell && cell.v !== undefined && cell.v !== null ? format_cell(cell).trim() : '';
    const base = text || `Column${c - range.s.c + 1}`;
    let name = base;
    for (let n = 2; seen.has(name.toLowerCase()); n++) name = `${base}${n}`;
    seen.add(name.toLowerCase());
    fields.push(name);
  }

  let lastRow = range.s.r;
  for (let r = range.e.r; r > range.s.r; r--) {
    let any = false;
    for (let c = range.s.c; c <= range.e.c && !any; c++) any = cellValue(ws[encode_cell({ r, c })]) !== null;
    if (any) {
      lastRow = r;
      break;
    }
  }

  const records: PivotValue[][] = [];
  const formats: Array<string | undefined> = fields.map(() => undefined);
  const dateFields = new Set<number>();
  for (let r = range.s.r + 1; r <= lastRow; r++) {
    const rec: PivotValue[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const i = c - range.s.c;
      const cell = ws[encode_cell({ r, c })] as CellObject | undefined;
      const v = cellValue(cell);
      rec.push(v);
      if (v !== null && formats[i] === undefined) {
        const z = cell?.s?.numFmt ?? cell?.z;
        formats[i] = typeof z === 'string' ? z : undefined;
        if (cell?.t === 'd' || cell?.v instanceof Date || (typeof v === 'number' && isDateFormat(z))) dateFields.add(i);
      }
    }
    records.push(rec);
  }
  return { fields, records, formats, dateFields, range: { s: range.s, e: { r: lastRow, c: range.e.c } } };
}
