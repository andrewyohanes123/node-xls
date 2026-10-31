/**
 * Worksheet construction and conversion utilities. Semantics follow SheetJS
 * CE (verified by differential tests) unless documented otherwise.
 */
import {
  WorkSheet,
  CellObject,
  Range,
  CellStyle,
  AOA2SheetOpts,
  JSON2SheetOpts,
  Sheet2JSONOpts,
  Sheet2CSVOpts,
  Sheet2HTMLOpts,
  Comment,
  Comments,
} from '../types.js';
import { encode_cell, encode_col, encode_row, decode_cell, decode_range, encode_range, parse_range_spec, wholeRangeRef, intersectRange, normaliseMerges } from './address.js';
import { escapeAttr, escapeXml } from '../xlsx/xml.js';
import { dateToSerial, serialToDate, isValidDate } from './dates.js';
import { DEFAULT_DATE_FORMAT, format_cell, formatValue, isDateFormat } from './format.js';
import { styleToCss } from '../styles/css.js';

export { format_cell } from './format.js';

/** Non-enumerable marker the reader puts on sheets of 1904-date workbooks (used by sheet_to_json). */
export const DATE1904 = '!date1904';

function isPlainObject(v: unknown): v is Record<string, any> {
  return !!v && typeof v === 'object' && !Array.isArray(v) && !(v instanceof Date);
}

/** Builds a cell from a JS value (SheetJS aoa_to_sheet rules). Returns undefined to skip. */
export function create_cell(val: any, opts?: AOA2SheetOpts): CellObject | undefined {
  if (val === undefined) return undefined;
  if (isPlainObject(val)) return val as CellObject;
  let v = val;
  let f: string | undefined;
  if (Array.isArray(val)) {
    v = val[0];
    f = val[1];
  }
  const cell: CellObject = {};
  if (f !== undefined) cell.f = f;
  if (v === null || v === undefined) {
    if (f) return { ...cell, t: 'n' };
    if (opts?.nullError) return { t: 'e', v: 0 };
    if (opts?.sheetStubs) return { t: 'z', v: null };
    return undefined;
  }
  if (typeof v === 'number') {
    if (Number.isFinite(v)) return { ...cell, t: 'n', v };
    return { ...cell, t: 'e', v: Number.isNaN(v) ? 0x0f : 0x07 };
  }
  if (typeof v === 'boolean') return { ...cell, t: 'b', v };
  if (v instanceof Date) {
    if (!isValidDate(v)) return { ...cell, t: 'e', v: 0x0f };
    const z = opts?.dateNF || DEFAULT_DATE_FORMAT;
    const serial = dateToSerial(v, { utc: opts?.UTC, date1904: opts?.date1904 });
    const w = formatValue(z, serial, { date1904: opts?.date1904 });
    if (opts?.cellDates) return { ...cell, t: 'd', v: opts?.UTC ? serialToDate(serial, { date1904: opts?.date1904 }) : v, z, w };
    return { ...cell, t: 'n', v: serial, z, w };
  }
  return { ...cell, t: 's', v: String(v) };
}

function resolveOrigin(ws: WorkSheet, origin: AOA2SheetOpts['origin']): { r: number; c: number } {
  if (origin === undefined || origin === null) return { r: 0, c: 0 };
  if (typeof origin === 'number') {
    if (origin === -1) return { r: ws['!ref'] ? decode_range(ws['!ref']).e.r + 1 : 0, c: 0 };
    return { r: origin, c: 0 };
  }
  if (typeof origin === 'string') return decode_cell(origin);
  return { r: origin.r, c: origin.c };
}

export function sheet_new(): WorkSheet {
  return {};
}

export function aoa_to_sheet<T = any>(data: T[][], opts?: AOA2SheetOpts): WorkSheet {
  return sheet_add_aoa({}, data, opts);
}

export function sheet_add_aoa<T = any>(ws: WorkSheet, data: T[][], opts?: AOA2SheetOpts): WorkSheet {
  const target = ws ?? {};
  const origin = resolveOrigin(target, opts?.origin);
  const range: Range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
  if (target['!ref']) {
    const r = decode_range(target['!ref']);
    range.s = { ...r.s };
    range.e = { ...r.e };
  }
  let seen = false;
  for (let R = 0; R < data.length; R++) {
    const row = data[R];
    if (!row) continue;
    if (!Array.isArray(row)) throw new Error('aoa_to_sheet expects an array of arrays');
    const r = origin.r + R;
    for (let C = 0; C < row.length; C++) {
      if (row[C] === undefined) continue;
      const c = origin.c + C;
      // SheetJS extends the range for null values even when no cell is created
      if (range.s.r > r) range.s.r = r;
      if (range.s.c > c) range.s.c = c;
      if (range.e.r < r) range.e.r = r;
      if (range.e.c < c) range.e.c = c;
      seen = true;
      const cell = create_cell(row[C], opts);
      if (!cell) continue;
      const ref = encode_cell({ c, r });
      const prev = target[ref] as CellObject | undefined;
      // filling a template keeps the existing number format (as SheetJS) and style
      if (prev && cell !== row[C]) {
        if (prev.z !== undefined) cell.z = prev.z;
        if (prev.s !== undefined && cell.s === undefined) cell.s = prev.s;
      }
      target[ref] = cell;
    }
  }
  if (seen || target['!ref']) target['!ref'] = encode_range(range);
  return target;
}

export function json_to_sheet<T = any>(data: T[], opts?: JSON2SheetOpts): WorkSheet {
  return sheet_add_json({}, data, opts);
}

export function sheet_add_json<T = any>(ws: WorkSheet, data: T[], opts?: JSON2SheetOpts): WorkSheet {
  const target = ws ?? {};
  const rows = Array.isArray(data) ? data : [];
  const offset = opts?.skipHeader ? 0 : 1;
  const origin = resolveOrigin(target, opts?.origin);
  const range: Range = { s: { c: 0, r: 0 }, e: { c: origin.c, r: origin.r + rows.length - 1 + offset } };
  if (target['!ref']) {
    const r = decode_range(target['!ref']);
    range.e.c = Math.max(range.e.c, r.e.c);
    range.e.r = Math.max(range.e.r, r.e.r);
  }
  const hdr: string[] = opts?.header ? [...opts.header] : [];
  rows.forEach((obj, R) => {
    if (!obj || typeof obj !== 'object') return;
    for (const k of Object.keys(obj)) {
      let C = hdr.indexOf(k);
      if (C === -1) C = hdr.push(k) - 1;
      const ref = encode_col(origin.c + C) + encode_row(origin.r + R + offset);
      const v = (obj as any)[k];
      if (isPlainObject(v)) {
        target[ref] = v;
        continue;
      }
      const existing = target[ref] as CellObject | undefined;
      let t: CellObject['t'] = 'z';
      let value: any = v;
      let z: string | number | undefined;
      if (typeof v === 'number') t = Number.isFinite(v) ? 'n' : 'e';
      else if (typeof v === 'boolean') t = 'b';
      else if (typeof v === 'string') t = 's';
      else if (v instanceof Date) {
        z = existing?.z !== undefined && isDateFormat(existing.z) ? existing.z : opts?.dateNF || DEFAULT_DATE_FORMAT;
        if (opts?.cellDates) t = 'd';
        else {
          t = 'n';
          value = dateToSerial(v, { utc: opts?.UTC, date1904: opts?.date1904 });
        }
      } else if (v === null && opts?.nullError) {
        t = 'e';
        value = 0;
      }
      if (t === 'e' && typeof v === 'number') value = Number.isNaN(v) ? 0x0f : 0x07;
      if (existing) {
        existing.t = t;
        existing.v = value;
        delete existing.w;
        if (z !== undefined) existing.z = z;
      } else {
        target[ref] = z !== undefined ? { t, v: value, z } : { t, v: value };
      }
    }
  });
  range.e.c = Math.max(range.e.c, origin.c + hdr.length - 1);
  if (offset) {
    hdr.forEach((h, C) => {
      target[encode_col(origin.c + C) + encode_row(origin.r)] = { t: 's', v: h };
    });
  }
  target['!ref'] = encode_range(range);
  return target;
}

function resolveJsonRange(ws: WorkSheet, range: Sheet2JSONOpts['range']): Range {
  const ref = decode_range(ws['!ref']!);
  if (range === undefined || range === null) return ref;
  if (typeof range === 'number') return { s: { c: ref.s.c, r: range }, e: ref.e };
  if (typeof range === 'string') return decode_range(range);
  return range;
}

export function sheet_to_json<T = any>(ws: WorkSheet, opts?: Sheet2JSONOpts): T[] {
  if (!ws || !ws['!ref']) return [];
  const o = opts ?? {};
  let r = resolveJsonRange(ws, o.range);
  if (wholeRangeRef(r)) r = intersectRange(r, decode_range(ws['!ref'])) ?? r;
  const date1904 = o.date1904 ?? Boolean((ws as any)[DATE1904]);
  const header = o.header === 1 ? 1 : o.header === 'A' ? 2 : Array.isArray(o.header) ? 3 : 0;
  const offset = header > 0 ? 0 : 1;
  // like SheetJS: raw defaults to true only when the option is absent
  const raw = 'raw' in o ? Boolean(o.raw) : true;
  const colInfo = (o.skipHidden && ws['!cols']) || [];
  const rowInfo = (o.skipHidden && ws['!rows']) || [];

  const cols: string[] = [];
  const hdr: Array<string | number | undefined> = [];
  const seen: Record<string, number> = {};
  for (let C = r.s.c; C <= r.e.c; C++) {
    if (colInfo[C]?.hidden) continue;
    cols[C] = encode_col(C);
    const cell = ws[cols[C] + encode_row(r.s.r)] as CellObject | undefined;
    switch (header) {
      case 1:
        hdr[C] = C - r.s.c;
        break;
      case 2:
        hdr[C] = cols[C];
        break;
      case 3:
        hdr[C] = (o.header as string[])[C - r.s.c];
        break;
      default: {
        // styled blanks read from files are stub cells; treat them like missing cells (as SheetJS sees them)
        // styled blanks read from files are stub cells SheetJS would not create: treat them as missing
        const styledBlank = cell != null && cell.t === 'z' && cell.v === undefined && !cell.c && !cell.l;
        const name = cell == null || styledBlank ? '__EMPTY' : format_cell(cell, undefined, o);
        let candidate = name;
        let counter = seen[name] || 0;
        if (!counter) seen[name] = 1;
        else {
          do candidate = `${name}_${counter++}`;
          while (seen[candidate]);
          seen[name] = counter;
          seen[candidate] = 1;
        }
        hdr[C] = candidate;
      }
    }
  }

  const out: any[] = [];
  for (let R = r.s.r + offset; R <= r.e.r; R++) {
    if (rowInfo[R]?.hidden) continue;
    const row: any = header === 1 ? [] : {};
    if (header !== 1) Object.defineProperty(row, '__rowNum__', { value: R, enumerable: false });
    let isEmpty = true;
    for (let C = r.s.c; C <= r.e.c; C++) {
      if (cols[C] === undefined) continue;
      const cell = ws[cols[C] + encode_row(R)] as CellObject | undefined;
      const key = hdr[C];
      if (!cell || cell.t === undefined) {
        if (o.defval !== undefined && key != null) row[key] = o.defval;
        continue;
      }
      let v: any = cell.v;
      switch (cell.t) {
        case 'z':
          if (v == null) break;
          continue;
        case 'e':
          v = v === 0 ? null : undefined;
          break;
        case 'n':
          // like SheetJS, raw:false keeps the UTC-field date
          if (cell.z !== undefined && isDateFormat(cell.z) && typeof v === 'number') v = serialToDate(v, { utc: o.UTC || o.raw === false, date1904 });
          break;
        case 'd':
          if (o.UTC && v instanceof Date) v = new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds(), v.getMilliseconds()));
          break;
      }
      if (key == null) continue;
      if (v === null || v === undefined) {
        if (cell.t === 'e' && v === null) row[key] = null;
        else if (o.defval !== undefined) row[key] = o.defval;
        else if (raw && v === null) row[key] = null;
        else continue;
      } else {
        const useRaw = cell.t === 'n' && typeof o.rawNumbers === 'boolean' ? o.rawNumbers : raw;
        row[key] = useRaw ? v : format_cell(cell, undefined, o);
      }
      if (v !== null && v !== undefined) isEmpty = false;
    }
    if (!isEmpty || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out.push(row);
  }
  return out as T[];
}

export function sheet_to_csv(ws: WorkSheet, opts?: Sheet2CSVOpts): string {
  if (!ws || !ws['!ref']) return '';
  const o = opts ?? {};
  const r = decode_range(ws['!ref']);
  const FS = o.FS ?? ',';
  const RS = o.RS ?? '\n';
  const fs = FS.charCodeAt(0);
  const rs = RS.charCodeAt(0);
  const colInfo = (o.skipHidden && ws['!cols']) || [];
  const rowInfo = (o.skipHidden && ws['!rows']) || [];
  const cols: string[] = [];
  for (let C = r.s.c; C <= r.e.c; C++) if (!colInfo[C]?.hidden) cols[C] = encode_col(C);

  const out: string[] = [];
  let written = 0;
  for (let R = r.s.r; R <= r.e.r; R++) {
    if (rowInfo[R]?.hidden) continue;
    const fields: string[] = [];
    let isEmpty = true;
    for (let C = r.s.c; C <= r.e.c; C++) {
      if (cols[C] === undefined) continue;
      const cell = ws[cols[C] + encode_row(R)] as CellObject | undefined;
      let txt = '';
      if (cell && cell.v !== undefined && cell.v !== null) {
        isEmpty = false;
        txt = String(o.rawNumbers && cell.t === 'n' ? cell.v : format_cell(cell, undefined, o));
        for (let i = 0; i < txt.length; i++) {
          const cc = txt.charCodeAt(i);
          if (cc === fs || cc === rs || cc === 34 || o.forceQuotes) {
            txt = `"${txt.replace(/"/g, '""')}"`;
            break;
          }
        }
        if (txt === 'ID' && written === 0 && fields.length === 0) txt = '"ID"';
      } else if (cell && cell.f !== undefined && cell.f !== null && !cell.F) {
        isEmpty = false;
        txt = `=${cell.f}`;
        if (txt.includes(',')) txt = `"${txt.replace(/"/g, '""')}"`;
      }
      fields.push(txt);
    }
    if (o.strip) while (fields.length && fields[fields.length - 1] === '') fields.length--;
    if (o.blankrows === false && isEmpty) continue;
    out.push((written++ ? RS : '') + fields.join(FS));
  }
  return out.join('');
}

export function sheet_to_txt(ws: WorkSheet, opts?: Sheet2CSVOpts): string {
  return sheet_to_csv(ws, { ...opts, FS: '\t', RS: '\n' });
}

/** HTML table with inline CSS for cell styles; merges become colspan/rowspan. */
export function sheet_to_html(ws: WorkSheet, opts?: Sheet2HTMLOpts): string {
  const o = opts ?? {};
  const head = o.header ?? '';
  const foot = o.footer ?? '';
  if (!ws || !ws['!ref']) return `${head}<table${o.id ? ` id="${escapeAttr(o.id)}"` : ''}></table>${foot}`;
  const r = decode_range(ws['!ref']);
  const merges = normaliseMerges(ws['!merges']);
  const covered = new Set<string>();
  const spans = new Map<string, { rowspan: number; colspan: number }>();
  for (const m of merges) {
    spans.set(encode_cell(m.s), { rowspan: m.e.r - m.s.r + 1, colspan: m.e.c - m.s.c + 1 });
    for (let R = m.s.r; R <= m.e.r; R++) for (let C = m.s.c; C <= m.e.c; C++) if (R !== m.s.r || C !== m.s.c) covered.add(encode_cell({ r: R, c: C }));
  }
  const rows: string[] = [];
  for (let R = r.s.r; R <= r.e.r; R++) {
    const cells: string[] = [];
    for (let C = r.s.c; C <= r.e.c; C++) {
      const ref = encode_cell({ r: R, c: C });
      if (covered.has(ref)) continue;
      const cell = ws[ref] as CellObject | undefined;
      const span = spans.get(ref);
      let attrsStr = ` id="${o.id ? escapeAttr(`${o.id}-${ref}`) : `sjs-${ref}`}"`;
      if (span?.rowspan && span.rowspan > 1) attrsStr += ` rowspan="${span.rowspan}"`;
      if (span?.colspan && span.colspan > 1) attrsStr += ` colspan="${span.colspan}"`;
      let content = '';
      if (cell) {
        const text = format_cell(cell);
        content = escapeXml(text).replace(/\n/g, '<br/>');
        if (cell.t && cell.t !== 'z') attrsStr += ` data-t="${cell.t}"`;
        if (cell.v !== undefined && cell.v !== null && cell.t !== 'z') {
          const dv = cell.v instanceof Date ? cell.v.toISOString() : String(cell.v);
          attrsStr += ` data-v="${escapeAttr(dv)}"`;
        }
        if (cell.z !== undefined && typeof cell.z === 'string' && cell.z !== 'General') attrsStr += ` data-z="${escapeAttr(cell.z)}"`;
        const css = styleToCss(cell.s);
        if (css) attrsStr += ` style="${escapeAttr(css)}"`;
        if (cell.l?.Target) {
          const href = cell.l.Target;
          const safe = /^\s*javascript:/i.test(href) ? '#' : href;
          content = `<a href="${escapeAttr(safe)}"${cell.l.Tooltip ? ` title="${escapeAttr(cell.l.Tooltip)}"` : ''}>${content}</a>`;
        }
      }
      if (o.editable) attrsStr += ' contenteditable="true"';
      cells.push(`<td${attrsStr}>${content}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }
  return `${head}<table${o.id ? ` id="${escapeAttr(o.id)}"` : ''}>${rows.join('')}</table>${foot}`;
}

export function sheet_to_formulae(ws: WorkSheet, opts?: { values?: boolean }): string[] {
  if (!ws || !ws['!ref']) return [];
  const r = decode_range(ws['!ref']);
  const out: string[] = [];
  for (let R = r.s.r; R <= r.e.r; R++) {
    for (let C = r.s.c; C <= r.e.c; C++) {
      let ref = encode_cell({ r: R, c: C });
      const x = ws[ref] as CellObject | undefined;
      if (x === undefined) continue;
      let val = '';
      if (x.F != null) {
        ref = x.F;
        if (!x.f) continue;
        if (!ref.includes(':')) ref = `${ref}:${ref}`;
      }
      if (x.f != null) val = x.f;
      else if (opts?.values === false) continue;
      else if (x.t === 'z') continue;
      else if (x.t === 'n' && x.v != null) val = String(x.v);
      else if (x.t === 'b') val = x.v ? 'TRUE' : 'FALSE';
      else if (x.w !== undefined) val = `'${x.w}`;
      else if (x.v === undefined) continue;
      else if (x.t === 's') val = `'${x.v}`;
      else val = String(x.v);
      out.push(`${ref}=${val}`);
    }
  }
  return out;
}

function mergeStyle(base: CellStyle | undefined, patch: CellStyle): CellStyle {
  const out: CellStyle = { ...(base ?? {}) };
  for (const key of Object.keys(patch) as Array<keyof CellStyle>) {
    const v = patch[key];
    if (v === undefined) continue;
    if (key === 'border' && typeof v === 'object') {
      const b: any = { ...(base?.border ?? {}) };
      for (const [side, val] of Object.entries(v)) b[side] = val && typeof val === 'object' ? { ...(b[side] ?? {}), ...val } : val;
      out.border = b;
    } else if (v && typeof v === 'object' && !Array.isArray(v)) {
      (out as any)[key] = { ...((base as any)?.[key] ?? {}), ...v };
    } else {
      (out as any)[key] = v;
    }
  }
  return out;
}

function extendRef(ws: WorkSheet, r: number, c: number) {
  if (!ws['!ref']) {
    ws['!ref'] = encode_cell({ r, c });
    return;
  }
  const range = decode_range(ws['!ref']);
  if (r < range.s.r || r > range.e.r || c < range.s.c || c > range.e.c) {
    range.s.r = Math.min(range.s.r, r);
    range.s.c = Math.min(range.s.c, c);
    range.e.r = Math.max(range.e.r, r);
    range.e.c = Math.max(range.e.c, c);
    ws['!ref'] = encode_range(range);
  }
}

/** Merges `style` into a cell's style (nested objects are merged, not replaced). */
export function set_cell_style(ws: WorkSheet, cellRef: string, style: CellStyle): void {
  const addr = decode_cell(cellRef);
  const ref = encode_cell(addr);
  let cell = ws[ref] as CellObject | undefined;
  if (!cell) {
    cell = { t: 'z' };
    ws[ref] = cell;
    extendRef(ws, addr.r, addr.c);
  }
  cell.s = mergeStyle(cell.s, style);
}

export function set_range_style(ws: WorkSheet, rangeSpec: string | Range, style: CellStyle): void {
  const range = parse_range_spec(rangeSpec);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      set_cell_style(ws, encode_cell({ c, r }), style);
    }
  }
}

/** Returns the cell at an address, creating a stub if needed (SheetJS sheet_get_cell). */
export function sheet_get_cell(ws: WorkSheet, addr: string | { r: number; c: number } | number, c?: number): CellObject {
  const ref = typeof addr === 'string' ? addr : typeof addr === 'number' ? encode_cell({ r: addr, c: c ?? 0 }) : encode_cell(addr);
  return (ws[ref] ??= { t: 'z' });
}

export function cell_set_number_format(cell: CellObject, fmt: string | number): CellObject {
  cell.z = fmt;
  return cell;
}

export function cell_set_hyperlink(cell: CellObject, target?: string, tooltip?: string): CellObject {
  if (!target) delete cell.l;
  else cell.l = tooltip ? { Target: target, Tooltip: tooltip } : { Target: target };
  return cell;
}

export function cell_set_internal_link(cell: CellObject, range: string, tooltip?: string): CellObject {
  return cell_set_hyperlink(cell, `#${range}`, tooltip);
}

export function cell_add_comment(cell: CellObject, text: string, author?: string): void {
  if (!cell.c) cell.c = [] as Comment[] as Comments;
  cell.c.push({ t: text, a: author || 'SheetJS' });
}

export function sheet_set_array_formula(ws: WorkSheet, range: string | Range, formula: string, dynamic?: boolean): WorkSheet {
  const rng = typeof range === 'string' ? decode_range(range) : range;
  const ref = typeof range === 'string' ? range : encode_range(range);
  for (let R = rng.s.r; R <= rng.e.r; R++) {
    for (let C = rng.s.c; C <= rng.e.c; C++) {
      const cell = sheet_get_cell(ws, { r: R, c: C });
      cell.t = 'n';
      cell.F = ref;
      delete cell.v;
      if (R === rng.s.r && C === rng.s.c) {
        cell.f = formula;
        if (dynamic) cell.D = true;
      }
      extendRef(ws, R, C);
    }
  }
  return ws;
}
