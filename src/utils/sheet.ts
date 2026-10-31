import {
  WorkSheet,
  CellObject,
  Range,
  CellAddress,
  CellStyle,
  AOA2SheetOpts,
  JSON2SheetOpts,
  Sheet2JSONOpts,
  Sheet2CSVOpts,
  Sheet2HTMLOpts,
} from '../types.js';
import {
  encode_cell,
  decode_cell,
  encode_range,
  decode_range,
  decode_col,
  decode_row,
  parse_range_spec,
} from './address.js';
import { escapeXml } from '../styles/index.js';

function isDate(val: any): val is Date {
  return val instanceof Date && !isNaN(val.getTime());
}

function dateToExcelSerial(date: Date): number {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  const diff = date.getTime() - epoch.getTime();
  return diff / (24 * 60 * 60 * 1000);
}

export function create_cell(val: any, opts?: AOA2SheetOpts): CellObject {
  if (val === null || val === undefined) {
    return { t: 'z' };
  }
  if (typeof val === 'boolean') {
    return { t: 'b', v: val };
  }
  if (typeof val === 'number') {
    return { t: 'n', v: val };
  }
  if (isDate(val)) {
    if (opts?.cellDates) {
      return { t: 'd', v: val, z: opts?.dateNF || 'yyyy-mm-dd' };
    }
    return { t: 'n', v: dateToExcelSerial(val), z: opts?.dateNF || 'yyyy-mm-dd' };
  }
  return { t: 's', v: String(val) };
}

export function aoa_to_sheet<T = any>(data: T[][], opts?: AOA2SheetOpts): WorkSheet {
  const ws: WorkSheet = {};
  sheet_add_aoa(ws, data, opts);
  return ws;
}

export function sheet_add_aoa<T = any>(
  ws: WorkSheet,
  data: T[][],
  opts?: AOA2SheetOpts
): WorkSheet {
  let originC = 0;
  let originR = 0;

  if (opts?.origin !== undefined) {
    if (typeof opts.origin === 'number') {
      originR = opts.origin < 0 ? 0 : opts.origin;
    } else if (typeof opts.origin === 'string') {
      const pt = decode_cell(opts.origin);
      originC = pt.c;
      originR = pt.r;
    } else {
      originC = opts.origin.c;
      originR = opts.origin.r;
    }
  }

  let minC = originC;
  let minR = originR;
  let maxC = originC;
  let maxR = originR;

  if (ws['!ref']) {
    const existingRange = decode_range(ws['!ref']);
    minC = Math.min(minC, existingRange.s.c);
    minR = Math.min(minR, existingRange.s.r);
    maxC = Math.max(maxC, existingRange.e.c);
    maxR = Math.max(maxR, existingRange.e.r);
  }

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    const targetR = originR + r;
    if (targetR > maxR) maxR = targetR;

    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      const targetC = originC + c;
      if (targetC > maxC) maxC = targetC;

      if (val === undefined && !opts?.sheetStubs) {
        continue;
      }

      const cellRef = encode_cell({ c: targetC, r: targetR });
      ws[cellRef] = create_cell(val, opts);
    }
  }

  if (data.length > 0) {
    ws['!ref'] = encode_range({ s: { c: minC, r: minR }, e: { c: maxC, r: maxR } });
  } else if (!ws['!ref']) {
    ws['!ref'] = 'A1:A1';
  }

  return ws;
}

export function json_to_sheet<T = any>(data: T[], opts?: JSON2SheetOpts): WorkSheet {
  const ws: WorkSheet = {};
  sheet_add_json(ws, data, opts);
  return ws;
}

export function sheet_add_json<T = any>(
  ws: WorkSheet,
  data: T[],
  opts?: JSON2SheetOpts
): WorkSheet {
  if (!Array.isArray(data) || data.length === 0) {
    if (!ws['!ref']) ws['!ref'] = 'A1:A1';
    return ws;
  }

  let header = opts?.header;
  if (!header) {
    const headerSet = new Set<string>();
    for (const item of data) {
      if (item && typeof item === 'object') {
        Object.keys(item).forEach((k) => headerSet.add(k));
      }
    }
    header = Array.from(headerSet);
  }

  const aoa: any[][] = [];
  if (!opts?.skipHeader) {
    aoa.push(header);
  }

  for (const item of data) {
    const row: any[] = [];
    for (const h of header) {
      row.push((item as any)?.[h]);
    }
    aoa.push(row);
  }

  return sheet_add_aoa(ws, aoa, opts);
}

export function sheet_to_json<T = any>(ws: WorkSheet, opts?: Sheet2JSONOpts): T[] {
  if (!ws['!ref']) return [];
  const range = opts?.range
    ? typeof opts.range === 'number'
      ? { ...decode_range(ws['!ref']), s: { c: decode_range(ws['!ref']).s.c, r: opts.range } }
      : parse_range_spec(opts.range)
    : decode_range(ws['!ref']);

  const raw = opts?.raw ?? true;
  const defval = opts?.defval;
  const headerMode = opts?.header;

  // Header 1: return array of arrays
  if (headerMode === 1) {
    const result: any[][] = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row: any[] = [];
      let rowEmpty = true;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[encode_cell({ c, r })];
        const val = cell ? (raw ? cell.v : cell.w || cell.v) : defval;
        if (val !== undefined) rowEmpty = false;
        row.push(val);
      }
      if (!rowEmpty || opts?.blankrows !== false) {
        result.push(row);
      }
    }
    return result as T[];
  }

  // Column letters header 'A'
  if (headerMode === 'A') {
    const result: any[] = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const rowObj: Record<string, any> = {};
      let rowEmpty = true;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const colLetter = encode_cell({ c, r: 0 }).replace(/\d+$/, '');
        const cell = ws[encode_cell({ c, r })];
        const val = cell ? (raw ? cell.v : cell.w || cell.v) : defval;
        if (val !== undefined) rowEmpty = false;
        rowObj[colLetter] = val;
      }
      if (!rowEmpty || opts?.blankrows !== false) {
        result.push(rowObj);
      }
    }
    return result as T[];
  }

  // Explicit headers or first row headers
  const explicitHeaders = Array.isArray(headerMode) ? headerMode : null;
  const headers: string[] = [];
  const startRow = explicitHeaders ? range.s.r : range.s.r + 1;

  if (explicitHeaders) {
    headers.push(...explicitHeaders);
  } else {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[encode_cell({ c, r: range.s.r })];
      const hVal = cell && cell.v !== undefined ? String(cell.v) : `__EMPTY_${c}`;
      headers.push(hVal);
    }
  }

  const result: any[] = [];
  for (let r = startRow; r <= range.e.r; r++) {
    const rowObj: Record<string, any> = {};
    let rowEmpty = true;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const headerIndex = c - range.s.c;
      if (headerIndex >= headers.length) break;
      const headerKey = headers[headerIndex];
      const cell = ws[encode_cell({ c, r })];
      const val = cell ? (raw ? cell.v : cell.w || cell.v) : defval;
      if (val !== undefined) rowEmpty = false;
      if (val !== undefined || defval !== undefined) {
        rowObj[headerKey] = val;
      }
    }
    if (!rowEmpty || opts?.blankrows !== false) {
      result.push(rowObj);
    }
  }

  return result as T[];
}

export function sheet_to_csv(ws: WorkSheet, opts?: Sheet2CSVOpts): string {
  if (!ws['!ref']) return '';
  const range = decode_range(ws['!ref']);
  const fs = opts?.FS ?? ',';
  const rs = opts?.RS ?? '\n';
  const forceQuotes = opts?.forceQuotes ?? false;

  const lines: string[] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const rowVals: string[] = [];
    let rowEmpty = true;

    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[encode_cell({ c, r })];
      let valStr = '';
      if (cell && cell.v !== undefined && cell.v !== null) {
        valStr = String(cell.w || cell.v);
        rowEmpty = false;
      }

      if (forceQuotes || valStr.includes(fs) || valStr.includes('"') || valStr.includes('\n') || valStr.includes('\r')) {
        valStr = `"${valStr.replace(/"/g, '""')}"`;
      }
      rowVals.push(valStr);
    }

    if (!rowEmpty || opts?.blankrows !== false) {
      lines.push(rowVals.join(fs));
    }
  }

  return lines.join(rs);
}

export function sheet_to_txt(ws: WorkSheet, opts?: Sheet2CSVOpts): string {
  return sheet_to_csv(ws, { ...opts, FS: '\t' });
}

export function sheet_to_html(ws: WorkSheet, opts?: Sheet2HTMLOpts): string {
  if (!ws['!ref']) return '<table></table>';
  const range = decode_range(ws['!ref']);
  const tableId = opts?.id ? ` id="${escapeXml(opts.id)}"` : '';
  const editable = opts?.editable ? ' contenteditable="true"' : '';

  const rowsHtml: string[] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cellsHtml: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = encode_cell({ c, r });
      const cell = ws[cellRef];
      let text = '';
      let styleAttr = '';

      if (cell) {
        text = escapeXml(String(cell.w || cell.v || ''));
        if (cell.s) {
          const styles: string[] = [];
          if (cell.s.font?.bold) styles.push('font-weight: bold');
          if (cell.s.font?.italic) styles.push('font-style: italic');
          if (cell.s.font?.sz) styles.push(`font-size: ${cell.s.font.sz}pt`);
          if (cell.s.font?.name) styles.push(`font-family: ${cell.s.font.name}`);
          const color = typeof cell.s.font?.color === 'string' ? cell.s.font.color : cell.s.font?.color?.rgb;
          if (color) styles.push(`color: #${color.replace(/^#/, '')}`);

          const bgColor = typeof cell.s.fill?.fgColor === 'string' ? cell.s.fill.fgColor : cell.s.fill?.fgColor?.rgb;
          if (bgColor) styles.push(`background-color: #${bgColor.replace(/^#/, '')}`);

          if (cell.s.alignment?.horizontal) styles.push(`text-align: ${cell.s.alignment.horizontal}`);
          if (cell.s.alignment?.vertical) styles.push(`vertical-align: ${cell.s.alignment.vertical}`);

          if (styles.length > 0) {
            styleAttr = ` style="${styles.join('; ')}"`;
          }
        }
      }

      cellsHtml.push(`<td id="${cellRef}"${styleAttr}>${text}</td>`);
    }
    rowsHtml.push(`<tr>${cellsHtml.join('')}</tr>`);
  }

  return `<table${tableId}${editable}>${rowsHtml.join('')}</table>`;
}

export function sheet_to_formulae(ws: WorkSheet): string[] {
  if (!ws['!ref']) return [];
  const range = decode_range(ws['!ref']);
  const formulae: string[] = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const ref = encode_cell({ c, r });
      const cell = ws[ref];
      if (cell) {
        if (cell.f) {
          formulae.push(`${ref}=${cell.f}`);
        } else if (cell.v !== undefined) {
          formulae.push(`${ref}=${typeof cell.v === 'string' ? `'${cell.v}'` : cell.v}`);
        }
      }
    }
  }

  return formulae;
}

export function format_cell(cell: CellObject, v?: any): string {
  if (cell.w !== undefined) return cell.w;
  const val = v !== undefined ? v : cell.v;
  if (val === null || val === undefined) return '';
  return String(val);
}

export function set_cell_style(ws: WorkSheet, cellRef: string, style: CellStyle): void {
  let cell = ws[cellRef];
  if (!cell) {
    cell = { t: 'z', v: '' };
    ws[cellRef] = cell;
  }
  cell.s = { ...cell.s, ...style };
}

export function set_range_style(ws: WorkSheet, rangeSpec: string | Range, style: CellStyle): void {
  const range = parse_range_spec(rangeSpec);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const ref = encode_cell({ c, r });
      set_cell_style(ws, ref, style);
    }
  }
}
