/**
 * Formula helpers.
 */
import { encode_col, decode_col } from './address.js';

// Order matters: literals and quoted names are matched (and skipped) first.
const TOKEN =
  /("(?:[^"]|"")*")|('(?:[^']|'')*')|(\[[^\]]*\])|(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![A-Za-z0-9_(])|(?<![A-Za-z0-9_.$:])(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_.(:])|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})(?![A-Za-z0-9_(.])/g;

/**
 * Shifts the relative A1 references of a formula by (dr, dc), as Excel does
 * when expanding shared formulas or copying cells. Absolute ($) parts, string
 * literals, quoted sheet names and structured references are untouched.
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  const col = (abs: string, c: string) => (abs ? `$${c}` : encode_col(Math.max(0, decode_col(c) + dc)));
  const row = (abs: string, r: string) => (abs ? `$${r}` : String(Math.max(1, Number(r) + dr)));
  return formula.replace(TOKEN, (m, str, quoted, bracket, ca1, c1, ca2, c2, ra1, r1, ra2, r2, cAbs, c, rAbs, r) => {
    if (str || quoted || bracket) return m;
    if (c1 !== undefined) return `${col(ca1, c1)}:${col(ca2, c2)}`;
    if (r1 !== undefined) return `${row(ra1, r1)}:${row(ra2, r2)}`;
    return `${col(cAbs, c)}${row(rAbs, r)}`;
  });
}
