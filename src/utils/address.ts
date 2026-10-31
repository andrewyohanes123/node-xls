import { CellAddress, Range } from '../types.js';

export function encode_col(col: number): string {
  if (col < 0) throw new Error(`invalid column ${col}`);
  let s = '';
  for (++col; col; col = Math.floor((col - 1) / 26)) {
    s = String.fromCharCode(((col - 1) % 26) + 65) + s;
  }
  return s;
}

export function decode_col(colStr: string): number {
  let col = 0;
  const clean = colStr.replace(/\$/g, '').toUpperCase();
  for (let i = 0; i < clean.length; ++i) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function encode_row(row: number): string {
  return `${row + 1}`;
}

export function decode_row(rowStr: string): number {
  const clean = rowStr.replace(/\$/g, '');
  return parseInt(clean, 10) - 1;
}

/** Splits 'A1' into ['A', '1'] (SheetJS split_cell). */
export function split_cell(cstr: string): string[] {
  return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, '$1,$2').split(',');
}

export function encode_cell(cell: CellAddress): string {
  return `${encode_col(cell.c)}${encode_row(cell.r)}`;
}

export function decode_cell(ref: string): CellAddress {
  const clean = ref.replace(/\$/g, '');
  const match = clean.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  return {
    c: decode_col(match[1]),
    r: decode_row(match[2]),
  };
}

export function encode_range(range: Range): string;
export function encode_range(s: CellAddress | string, e: CellAddress | string): string;
export function encode_range(a: Range | CellAddress | string, b?: CellAddress | string): string {
  const range = b !== undefined ? { s: a as CellAddress | string, e: b } : (a as Range);
  const s = typeof range.s === 'string' ? range.s : encode_cell(range.s);
  const e = typeof range.e === 'string' ? range.e : encode_cell(range.e);
  return s === e ? s : `${s}:${e}`;
}

/** !merges entries as Range objects (strings such as 'A1:B2' are accepted), without no-op single-cell merges. */
export function normaliseMerges(merges: Array<Range | string> | undefined): Range[] {
  return (merges ?? [])
    .map((m) => orderRange(typeof m === 'string' ? decode_range(m) : m))
    .filter((m) => m.s.r !== m.e.r || m.s.c !== m.e.c);
}

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

/**
 * Decodes 'A1:B2', 'B2', whole columns ('A:C') and whole rows ('2:5').
 * (SheetJS returns -1 rows/columns for the latter; node-xls spans the grid.)
 */
export function decode_range(rangeStr: string): Range {
  const whole = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(rangeStr) ?? /^\$?(\d+):\$?(\d+)$/.exec(rangeStr);
  if (whole) {
    if (/\d/.test(whole[1])) return { s: { c: 0, r: decode_row(whole[1]) }, e: { c: MAX_COL, r: decode_row(whole[2]) } };
    return { s: { c: decode_col(whole[1]), r: 0 }, e: { c: decode_col(whole[2]), r: MAX_ROW } };
  }
  const idx = rangeStr.indexOf(':');
  if (idx === -1) {
    const pt = decode_cell(rangeStr);
    return { s: { ...pt }, e: { ...pt } };
  }
  return {
    s: decode_cell(rangeStr.substring(0, idx)),
    e: decode_cell(rangeStr.substring(idx + 1)),
  };
}

export function parse_range_spec(spec?: string | Range, defaultRange?: Range): Range {
  if (!spec) {
    if (defaultRange) return defaultRange;
    throw new Error('Range specification is required');
  }
  if (typeof spec === 'string') {
    return decode_range(spec);
  }
  return spec;
}

/** Throws unless the range lies inside Excel's grid (A1:XFD1048576). */
export function checkRangeInGrid(range: Range, what: string): void {
  const ok = (a: CellAddress) => Number.isInteger(a.r) && Number.isInteger(a.c) && a.r >= 0 && a.r <= MAX_ROW && a.c >= 0 && a.c <= MAX_COL;
  if (!ok(range.s) || !ok(range.e)) throw new Error(`${what} is outside Excel's grid (A1:XFD1048576)`);
}

/** A range with start/end swapped where needed (s top-left, e bottom-right). */
export function orderRange(range: Range): Range {
  return {
    s: { r: Math.min(range.s.r, range.e.r), c: Math.min(range.s.c, range.e.c) },
    e: { r: Math.max(range.s.r, range.e.r), c: Math.max(range.s.c, range.e.c) },
  };
}

/** Space-separated range list ("A1:B2 D4") validated against the grid. */
export function checkSqref(sqref: string, what: string): void {
  for (const part of sqref.trim().split(/\s+/)) {
    let r: Range;
    try {
      r = decode_range(part.replace(/\$/g, ''));
    } catch {
      throw new Error(`${what}: invalid range "${part}"`);
    }
    checkRangeInGrid(r, `${what} range ${part}`);
  }
}

/** Excel's formula length limit. */
export const MAX_FORMULA_LENGTH = 8192;
export function checkFormula(f: string, what: string): void {
  if (f.length > MAX_FORMULA_LENGTH) throw new Error(`${what}: formula longer than ${MAX_FORMULA_LENGTH} characters`);
}

/** Quotes a sheet name for use in formulas: 'My Sheet' -> "'My Sheet'". */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

/** Whole-column ('A:C') or whole-row ('2:5') form of a range, if it spans the grid; else undefined. */
export function wholeRangeRef(range: Range, absolute = false): string | undefined {
  const $ = absolute ? '$' : '';
  if (range.s.r === 0 && range.e.r === MAX_ROW) return `${$}${encode_col(range.s.c)}:${$}${encode_col(range.e.c)}`;
  if (range.s.c === 0 && range.e.c === MAX_COL) return `${$}${range.s.r + 1}:${$}${range.e.r + 1}`;
  return undefined;
}

/** encode_range that keeps whole-column / whole-row ranges compact ('A:B'). */
export function encodeRangeCompact(range: Range): string {
  return wholeRangeRef(range) ?? encode_range(range);
}

/** Absolute reference like 'Sheet 1'!$A$1:$C$9 (or 'Sheet 1'!$A:$C for whole columns). */
export function absoluteRef(sheet: string, range: Range): string {
  const whole = wholeRangeRef(range, true);
  if (whole) return `${quoteSheetName(sheet)}!${whole}`;
  const cell = (a: CellAddress) => `$${encode_col(a.c)}$${a.r + 1}`;
  const s = cell(range.s);
  const e = cell(range.e);
  return `${quoteSheetName(sheet)}!${s === e ? s : `${s}:${e}`}`;
}

/** Intersection of a range with another (e.g. clamp a whole column to a sheet's used range). */
export function intersectRange(a: Range, b: Range): Range | undefined {
  const s = { r: Math.max(a.s.r, b.s.r), c: Math.max(a.s.c, b.s.c) };
  const e = { r: Math.min(a.e.r, b.e.r), c: Math.min(a.e.c, b.e.c) };
  return s.r <= e.r && s.c <= e.c ? { s, e } : undefined;
}

/**
 * Parses a sheet-qualified reference: "'My Sheet'!$A$1:$B$4", "Data!A1:A9",
 * or "A1:B2" (sheet undefined).
 */
export function parseSheetRef(ref: string): { sheet?: string; range: Range } {
  const s = ref.trim().replace(/^=/, '');
  const bang = s.lastIndexOf('!');
  if (bang === -1) return { range: decode_range(s) };
  let sheet = s.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, range: decode_range(s.slice(bang + 1)) };
}
