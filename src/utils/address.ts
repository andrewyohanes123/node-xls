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
export function encode_range(s: CellAddress, e: CellAddress): string;
export function encode_range(a: Range | CellAddress, b?: CellAddress): string {
  const range: Range = b ? { s: a as CellAddress, e: b } : (a as Range);
  const s = encode_cell(range.s);
  const e = encode_cell(range.e);
  return s === e ? s : `${s}:${e}`;
}

export function decode_range(rangeStr: string): Range {
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

/** Quotes a sheet name for use in formulas: 'My Sheet' -> "'My Sheet'". */
export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

/** Absolute reference like 'Sheet 1'!$A$1:$C$9. */
export function absoluteRef(sheet: string, range: Range): string {
  const cell = (a: CellAddress) => `$${encode_col(a.c)}$${a.r + 1}`;
  const s = cell(range.s);
  const e = cell(range.e);
  return `${quoteSheetName(sheet)}!${s === e ? s : `${s}:${e}`}`;
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
