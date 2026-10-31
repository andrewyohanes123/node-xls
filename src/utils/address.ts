import { CellAddress, Range } from '../types.js';

export function encode_col(col: number): string {
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

export function encode_range(range: Range): string {
  return `${encode_cell(range.s)}:${encode_cell(range.e)}`;
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
