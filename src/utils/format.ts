/**
 * Number-format rendering (cell.w) backed by SSF, SheetJS's format engine.
 */
import * as SSF from 'ssf';
import { CellObject } from '../types.js';
import { dateToSerial } from './dates.js';

/** Excel built-in number formats (ids 0-49) as used by SheetJS. */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = Object.freeze({ ...(SSF.get_table() as Record<number, string>) });

const BUILTIN_IDS = new Map<string, number>();
for (const [id, code] of Object.entries(BUILTIN_FORMATS)) {
  if (!BUILTIN_IDS.has(code)) BUILTIN_IDS.set(code, Number(id));
}

/** Default format for Date values (built-in 14, locale short date in Excel). */
export const DEFAULT_DATE_FORMAT = BUILTIN_FORMATS[14];

export const ERROR_CODES: Readonly<Record<number, string>> = Object.freeze({
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
  0x2b: '#GETTING_DATA',
});

export const ERROR_VALUES: Readonly<Record<string, number>> = Object.freeze(
  Object.fromEntries(Object.entries(ERROR_CODES).map(([k, v]) => [v, Number(k)]))
);

/** Error text for a cell value (numeric code or string); newer errors (#SPILL!, #CALC!, ...) pass through. */
export function errorText(v: unknown): string {
  if (typeof v === 'number') return ERROR_CODES[v] ?? '#N/A';
  const s = String(v ?? '');
  return ERROR_VALUES[s] !== undefined || /^#[A-Z0-9/_.]+[!?]?$/i.test(s) ? s : '#N/A';
}

/** Returns the built-in id for a format code, if it is one. */
export function builtinFormatId(code: string): number | undefined {
  return code.toLowerCase() === 'general' ? 0 : BUILTIN_IDS.get(code);
}

/** Built-in ids without a fixed code (locale / CJK variants) and their equivalents (as SheetJS). */
const BUILTIN_EQUIVALENTS: Readonly<Record<number, number>> = {
  5: 37, 6: 38, 7: 39, 8: 40, 23: 0, 24: 0, 25: 0, 26: 0, 27: 14, 28: 14, 29: 14, 30: 14, 31: 14,
  50: 14, 51: 14, 52: 14, 53: 14, 54: 14, 55: 14, 56: 14, 57: 14, 58: 14, 59: 1, 60: 2, 61: 3, 62: 4,
  67: 9, 68: 10, 69: 12, 70: 13, 71: 14, 72: 14, 73: 15, 74: 16, 75: 17, 76: 20, 77: 21, 78: 22,
  79: 45, 80: 46, 81: 47, 82: 0,
};
const BUILTIN_CURRENCY: Readonly<Record<number, string>> = {
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
};

/** Format code for a built-in id (including ids Excel resolves per locale), or undefined. */
export function builtinFormatCode(id: number): string | undefined {
  return BUILTIN_FORMATS[id] ?? BUILTIN_FORMATS[BUILTIN_EQUIVALENTS[id]] ?? BUILTIN_CURRENCY[id];
}

/** Resolves a format id or code to a format code. */
export function resolveFormat(fmt: string | number | undefined | null): string {
  if (fmt === undefined || fmt === null) return 'General';
  if (typeof fmt === 'number') return builtinFormatCode(fmt) ?? 'General';
  return fmt;
}

export function isDateFormat(fmt: string | number | undefined | null): boolean {
  if (fmt === undefined || fmt === null) return false;
  try {
    return SSF.is_date(resolveFormat(fmt));
  } catch {
    return false;
  }
}

/** Formats a value with an Excel number format; falls back to String(v). */
export function formatValue(fmt: string | number | undefined | null, value: unknown, opts?: { date1904?: boolean }): string {
  let v: unknown = value;
  if (v instanceof Date) v = dateToSerial(v, { date1904: opts?.date1904 });
  try {
    return SSF.format(resolveFormat(fmt), v, { date1904: opts?.date1904 });
  } catch {
    return String(v);
  }
}

/** SheetJS-compatible format_cell: the display text of a cell. */
export function format_cell(cell: CellObject | undefined | null, v?: unknown, opts?: { dateNF?: string; date1904?: boolean }): string {
  if (!cell || cell.t === undefined || cell.t === 'z') return '';
  if (cell.w !== undefined) return cell.w;
  if (cell.t === 'e') return errorText(cell.v);
  const value = v === undefined || v === null ? cell.v : v;
  if (value === undefined || value === null) return '';
  if (cell.t === 'b' || typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (cell.t === 's' && cell.z === undefined) return String(value);
  let fmt = cell.z ?? cell.s?.numFmt;
  if (cell.t === 'd' && fmt === undefined) fmt = opts?.dateNF ?? 14;
  return formatValue(fmt, value, { date1904: opts?.date1904 });
}
