/**
 * Worksheet part writer.
 *
 * CT_Worksheet is an xsd:sequence, so child order is significant: Excel
 * "repairs" (drops) files with misplaced elements. All sections are collected
 * into a map and serialised in WORKSHEET_CHILD_ORDER.
 */
import {
  CellObject,
  CellStyle,
  ColProperties,
  Comments,
  Hyperlink,
  RowProperties,
  WorkSheet,
  Range,
  FreezePane,
} from '../types.js';
import { WriteContext } from './context.js';
import { attrs, encodeXstring, escapeAttr, escapeXml, num } from './xml.js';
import { colorAttrs } from '../styles/index.js';
import { decode_col, decode_cell, decode_range, encode_cell, encode_range, absoluteRef } from '../utils/address.js';
import { dateToSerial, isValidDate, isoToSerial } from '../utils/dates.js';
import { errorText } from '../utils/format.js';
import { charsToWidth, pixelsToChars, pixelsToPoints } from '../utils/units.js';
import { sheetProtectionXml } from '../features/protection.js';
import { hyperlinksXml } from '../features/hyperlinks.js';
import { writeComments, CommentEntry } from '../features/comments.js';
import { conditionalFormattingXml } from '../features/conditional.js';
import { dataValidationsXml } from '../features/validation.js';
import { writeTables, tableCellOverrides } from '../features/tables.js';
import { writeDrawing } from '../drawing/drawing.js';

export const WORKSHEET_CHILD_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetProtection',
  'autoFilter',
  'mergeCells',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'drawing',
  'legacyDrawing',
  'tableParts',
] as const;
type Section = (typeof WORKSHEET_CHILD_ORDER)[number];

const CELL_REF = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/;
const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

interface CellEntry {
  c: number;
  key: string;
  cell: CellObject;
}

function inferType(v: unknown): string {
  if (v === undefined || v === null) return 'z';
  if (typeof v === 'boolean') return 'b';
  if (typeof v === 'number') return 'n';
  if (v instanceof Date) return 'd';
  return 's';
}

function toDate(v: unknown): Date | undefined {
  if (isValidDate(v)) return v;
  if (typeof v === 'string' || typeof v === 'number') {
    const d = new Date(v);
    return isValidDate(d) ? d : undefined;
  }
  return undefined;
}

/** cell.s plus number formats implied by cell.z or a Date value. */
function effectiveStyle(cell: CellObject, isDate: boolean): CellStyle | undefined {
  let s = cell.s;
  if (s?.numFmt === undefined) {
    const z = cell.z;
    if (z !== undefined && z !== null && z !== 'General' && z !== 0) s = { ...s, numFmt: z };
    else if (isDate) s = { ...s, numFmt: 14 };
  }
  return s;
}

function cellXml(ctx: WriteContext, ref: string, cell: CellObject): string | null {
  const v = cell.v;
  const t = cell.t ?? inferType(v);
  let formula: string | undefined;
  if (typeof cell.f === 'string' && cell.f.length > 0) {
    formula = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
    ctx.hasFormulas = true;
  }
  const fXml =
    formula === undefined ? '' : cell.F ? `<f t="array" ref="${escapeAttr(cell.F)}">${escapeXml(formula)}</f>` : `<f>${escapeXml(formula)}</f>`;

  let type: string | undefined;
  let vXml = '';
  let isDate = false;

  const writeNumber = (n: number) => {
    if (Number.isFinite(n)) vXml = `<v>${num(n)}</v>`;
    else {
      type = 'e';
      vXml = '<v>#NUM!</v>';
    }
  };
  const writeDate = (d: Date) => {
    isDate = true;
    writeNumber(dateToSerial(d, { date1904: ctx.date1904 }));
  };
  const writeString = (s: string) => {
    if (s.length > 32767) throw new Error(`Cell ${ref}: text exceeds Excel's 32767 character limit`);
    if (formula !== undefined) {
      type = 'str';
      vXml = `<v>${encodeXstring(s)}</v>`;
    } else if (cell.richText && cell.richText.length > 0) {
      type = 's';
      vXml = `<v>${ctx.sst.getRichIndex(cell.richText, cell.s?.font)}</v>`;
    } else {
      type = 's';
      vXml = `<v>${ctx.sst.getIndex(s)}</v>`;
    }
  };

  switch (t) {
    case 'z':
      break;
    case 'b':
      if (v !== undefined && v !== null) {
        type = 'b';
        vXml = `<v>${v ? 1 : 0}</v>`;
      }
      break;
    case 'e':
      type = 'e';
      vXml = `<v>${escapeXml(errorText(v ?? cell.w))}</v>`;
      break;
    case 'd': {
      const iso = typeof v === 'string' ? isoToSerial(v, { date1904: ctx.date1904 }) : undefined;
      const d = iso === undefined ? toDate(v) : undefined;
      if (iso !== undefined) {
        isDate = true;
        writeNumber(iso);
      } else if (d) writeDate(d);
      else if (v !== undefined && v !== null) writeString(String(v));
      break;
    }
    case 'n':
      if (v instanceof Date) writeDate(v);
      else if (typeof v === 'number') writeNumber(v);
      else if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v))) writeNumber(Number(v));
      else if (v !== undefined && v !== null && v !== '') writeString(String(v));
      break;
    default:
      if (v !== undefined && v !== null) writeString(typeof v === 'string' ? v : String(v));
      else if (cell.richText && cell.richText.length > 0) writeString(cell.richText.map((r) => r.text).join(''));
  }

  const style = effectiveStyle(cell, isDate);
  const s = style ? ctx.styles.getStyleIndex(style) : 0;
  if (!vXml && !fXml) {
    return s > 0 ? `<c r="${ref}" s="${s}"/>` : null;
  }
  return `<c${attrs([
    ['r', ref],
    ['s', s > 0 ? s : undefined],
    ['t', type],
  ])}>${fXml}${vXml}</c>`;
}

function colWidth(col: ColProperties): number | undefined {
  if (col.width !== undefined) return col.width;
  if (col.wch !== undefined) return charsToWidth(col.wch);
  if (col.wpx !== undefined) return charsToWidth(pixelsToChars(col.wpx));
  return undefined;
}

function colsXml(ctx: WriteContext, cols: ColProperties[] | undefined): { xml: string; maxLevel: number } {
  if (!cols || cols.length === 0) return { xml: '', maxLevel: 0 };
  let maxLevel = 0;
  const items: string[] = [];
  for (let i = 0; i < cols.length; i++) {
    const col = cols[i];
    if (!col) continue;
    const width = colWidth(col);
    const level = col.level ?? col.outlineLevel;
    const style = col.s ? ctx.styles.getStyleIndex(col.s) : 0;
    if (width === undefined && !col.hidden && !level && !style) continue;
    if (level) maxLevel = Math.max(maxLevel, level);
    items.push(
      `<col${attrs([
        ['min', i + 1],
        ['max', i + 1],
        ['width', width ?? 9.140625],
        ['style', style || undefined],
        ['hidden', col.hidden],
        ['customWidth', width !== undefined],
        ['outlineLevel', level || undefined],
      ])}/>`
    );
  }
  return { xml: items.length ? `<cols>${items.join('')}</cols>` : '', maxLevel };
}

function rowAttrs(ctx: WriteContext, r: number, prop?: RowProperties): string {
  const list: Array<[string, string | number | boolean | undefined]> = [['r', r + 1]];
  if (prop) {
    const style = prop.s ? ctx.styles.getStyleIndex(prop.s) : 0;
    const ht = prop.hpt ?? (prop.hpx !== undefined ? pixelsToPoints(prop.hpx) : undefined);
    const level = prop.level ?? prop.outlineLevel;
    list.push(
      ['s', style || undefined],
      ['customFormat', style > 0],
      ['ht', ht],
      ['hidden', prop.hidden],
      ['customHeight', ht !== undefined],
      ['outlineLevel', level || undefined]
    );
  }
  return attrs(list);
}

function normaliseFreeze(f: string | FreezePane | undefined): Required<FreezePane> | undefined {
  if (!f) return undefined;
  let xSplit: number;
  let ySplit: number;
  let topLeftCell: string | undefined;
  if (typeof f === 'string') {
    const a = decode_cell(f);
    xSplit = a.c;
    ySplit = a.r;
    topLeftCell = f.replace(/\$/g, '').toUpperCase();
  } else {
    xSplit = f.xSplit ?? 0;
    ySplit = f.ySplit ?? 0;
    topLeftCell = f.topLeftCell;
  }
  if (xSplit <= 0 && ySplit <= 0) return undefined;
  return { xSplit, ySplit, topLeftCell: topLeftCell ?? encode_cell({ c: Math.max(0, xSplit), r: Math.max(0, ySplit) }) };
}

function sheetViewsXml(ctx: WriteContext, ws: WorkSheet, index: number): string {
  const view = ws['!views'] ?? {};
  const rtl = view.rightToLeft ?? (index === 0 ? ctx.wb.Workbook?.Views?.[0]?.RTL : undefined);
  const freeze = normaliseFreeze(ws['!freeze']);
  let children = '';
  if (freeze) {
    const pane = freeze.xSplit > 0 && freeze.ySplit > 0 ? 'bottomRight' : freeze.ySplit > 0 ? 'bottomLeft' : 'topRight';
    children =
      `<pane${attrs([
        ['xSplit', freeze.xSplit > 0 ? freeze.xSplit : undefined],
        ['ySplit', freeze.ySplit > 0 ? freeze.ySplit : undefined],
        ['topLeftCell', freeze.topLeftCell],
        ['activePane', pane],
        ['state', 'frozen'],
      ])}/>` + `<selection pane="${pane}" activeCell="${freeze.topLeftCell}" sqref="${freeze.topLeftCell}"/>`;
  }
  const a = attrs([
    ['tabSelected', index === ctx.activeTab],
    ['showGridLines', view.showGridLines === false ? '0' : undefined],
    ['showRowColHeaders', view.showRowColHeaders === false ? '0' : undefined],
    ['showZeros', view.showZeros === false ? '0' : undefined],
    ['rightToLeft', rtl === true],
    ['zoomScale', view.zoomScale],
    ['zoomScaleNormal', view.zoomScale],
    ['workbookViewId', 0],
  ]);
  return `<sheetViews><sheetView${a}${children ? `>${children}</sheetView>` : '/>'}</sheetViews>`;
}

function mergeCellsXml(merges: Range[] | undefined): string {
  if (!merges || merges.length === 0) return '';
  const norm = merges.map((m) => (typeof m === 'string' ? decode_range(m) : m));
  for (let i = 0; i < norm.length; i++) {
    for (let j = i + 1; j < norm.length; j++) {
      const a = norm[i];
      const b = norm[j];
      if (a.s.r <= b.e.r && b.s.r <= a.e.r && a.s.c <= b.e.c && b.s.c <= a.e.c) {
        throw new Error(`Overlapping merged ranges ${encode_range(a)} and ${encode_range(b)}`);
      }
    }
  }
  return `<mergeCells count="${norm.length}">${norm.map((m) => `<mergeCell ref="${encode_range(m)}"/>`).join('')}</mergeCells>`;
}

export function writeWorksheet(ctx: WriteContext, sheetName: string, ws: WorkSheet, index: number, part: string): void {
  const sections: Partial<Record<Section, string>> = {};
  const rows = new Map<number, CellEntry[]>();
  const links: Array<{ ref: string; link: Hyperlink }> = [];
  const comments: CommentEntry[] = [];
  let minC = Infinity;
  let minR = Infinity;
  let maxC = -1;
  let maxR = -1;

  const overrides = tableCellOverrides(ws);
  const keys = overrides.size ? new Set([...Object.keys(ws), ...overrides.keys()]) : Object.keys(ws);
  for (const key of keys) {
    if (key.charCodeAt(0) === 33 /* '!' */) continue;
    const m = CELL_REF.exec(key);
    const cell = overrides.get(key) ?? ws[key];
    if (!m || !cell || typeof cell !== 'object') continue;
    const c = decode_col(m[1]);
    const r = Number(m[2]) - 1;
    if (c > 16383 || r > 1048575) throw new Error(`Cell ${key} is outside Excel's grid`);
    let list = rows.get(r);
    if (!list) rows.set(r, (list = []));
    list.push({ c, key, cell });
    if (c < minC) minC = c;
    if (r < minR) minR = r;
    if (c > maxC) maxC = c;
    if (r > maxR) maxR = r;
    if (cell.l && cell.l.Target) links.push({ ref: key, link: cell.l });
    if (Array.isArray(cell.c) && cell.c.length > 0) comments.push({ ref: key, addr: { c, r }, comments: cell.c as Comments });
  }
  if (ws['!ref']) {
    try {
      const ref = decode_range(ws['!ref']);
      minC = Math.min(minC, ref.s.c);
      minR = Math.min(minR, ref.s.r);
      maxC = Math.max(maxC, ref.e.c);
      maxR = Math.max(maxR, ref.e.r);
    } catch {
      /* ignore malformed !ref */
    }
  }

  // sheetPr
  const tabColor = colorAttrs(ws['!tabColor']);
  const setup = ws['!pageSetup'];
  const fitToPage = setup && (setup.fitToWidth !== undefined || setup.fitToHeight !== undefined);
  if (tabColor || fitToPage) {
    sections.sheetPr = `<sheetPr>${tabColor ? `<tabColor${attrs(tabColor)}/>` : ''}${fitToPage ? '<pageSetUpPr fitToPage="1"/>' : ''}</sheetPr>`;
  }

  sections.dimension = `<dimension ref="${
    maxC >= 0 && maxR >= 0
      ? minC === maxC && minR === maxR
        ? encode_cell({ c: minC, r: minR })
        : encode_range({ s: { c: minC, r: minR }, e: { c: maxC, r: maxR } })
      : 'A1'
  }"/>`;
  sections.sheetViews = sheetViewsXml(ctx, ws, index);

  // cols (before sheetFormatPr so outline levels are known)
  const cols = colsXml(ctx, ws['!cols']);

  // sheetData
  const rowProps = ws['!rows'] ?? [];
  const rowIndexes = new Set<number>(rows.keys());
  let maxRowLevel = 0;
  rowProps.forEach((p, i) => {
    if (!p) return;
    const level = p.level ?? p.outlineLevel ?? 0;
    maxRowLevel = Math.max(maxRowLevel, level);
    if (p.hpt !== undefined || p.hpx !== undefined || p.hidden || level || p.s) rowIndexes.add(i);
  });
  const rowXml: string[] = [];
  for (const r of [...rowIndexes].sort((a, b) => a - b)) {
    const cells = (rows.get(r) ?? []).sort((a, b) => a.c - b.c);
    const cellXmls: string[] = [];
    for (const { key, cell } of cells) {
      const x = cellXml(ctx, key, cell);
      if (x) cellXmls.push(x);
    }
    const prop = rowProps[r];
    if (cellXmls.length === 0 && !prop) continue;
    rowXml.push(cellXmls.length ? `<row${rowAttrs(ctx, r, prop)}>${cellXmls.join('')}</row>` : `<row${rowAttrs(ctx, r, prop)}/>`);
  }

  sections.sheetFormatPr = `<sheetFormatPr${attrs([
    ['defaultRowHeight', 15],
    ['outlineLevelRow', maxRowLevel || undefined],
    ['outlineLevelCol', cols.maxLevel || undefined],
  ])}/>`;
  if (cols.xml) sections.cols = cols.xml;
  sections.sheetData = rowXml.length ? `<sheetData>${rowXml.join('')}</sheetData>` : '<sheetData/>';

  if (ws['!protect']) sections.sheetProtection = sheetProtectionXml(ws['!protect']);

  if (ws['!autofilter']?.ref) {
    const ref = String(ws['!autofilter'].ref).replace(/\$/g, '');
    sections.autoFilter = `<autoFilter ref="${escapeAttr(ref)}"/>`;
    ctx.autoNames.push({ name: '_xlnm._FilterDatabase', sheet: index, ref: absoluteRef(sheetName, decode_range(ref)) });
  }

  const merges = mergeCellsXml(ws['!merges']);
  if (merges) sections.mergeCells = merges;

  if (ws['!condfmt']?.length) sections.conditionalFormatting = conditionalFormattingXml(ws['!condfmt'], ctx.styles);
  if (ws['!dataValidation']?.length) sections.dataValidations = dataValidationsXml(ws['!dataValidation'], ctx);

  const hl = hyperlinksXml(links, ctx.pkg, part);
  if (hl) sections.hyperlinks = hl;

  const mg = { ...DEFAULT_MARGINS, ...(ws['!margins'] ?? {}) };
  sections.pageMargins = `<pageMargins${attrs([
    ['left', mg.left],
    ['right', mg.right],
    ['top', mg.top],
    ['bottom', mg.bottom],
    ['header', mg.header],
    ['footer', mg.footer],
  ])}/>`;
  if (setup) {
    const ps = attrs([
      ['paperSize', setup.paperSize],
      ['scale', setup.scale],
      ['fitToWidth', setup.fitToWidth],
      ['fitToHeight', setup.fitToHeight],
      ['orientation', setup.orientation],
    ]);
    if (ps) sections.pageSetup = `<pageSetup${ps}/>`;
  }

  const drawing = writeDrawing(ctx, ws, part);
  if (drawing) sections.drawing = drawing;

  if (comments.length) {
    comments.sort((a, b) => a.addr.r - b.addr.r || a.addr.c - b.addr.c);
    sections.legacyDrawing = writeComments(comments, ctx.pkg, part, ++ctx.seq.comments);
  }

  const tables = writeTables(ctx, ws, sheetName, part);
  if (tables) sections.tableParts = tables;

  const body = WORKSHEET_CHILD_ORDER.map((k) => sections[k] ?? '').join('');
  ctx.pkg.addPart(
    part,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</worksheet>`,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
  );
}
