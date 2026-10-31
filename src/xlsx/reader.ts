/**
 * XLSX reader. Part locations are resolved through relationships (never
 * hard-coded), all text is read verbatim (no type coercion or trimming), and
 * the output mirrors SheetJS: v/t/w/z/f/F on cells, '!ref', '!cols', ... on
 * sheets, plus node-xls extensions (styles, tables, validation, ...).
 */
import { unzipSync, strFromU8, unzip } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import {
  CellObject,
  ColProperties,
  Comments,
  ParsingOptions,
  RichTextRun,
  RowProperties,
  SheetProtection,
  WorkBook,
  WorkSheet,
} from '../types.js';
import { decode_cell, decode_range, encode_cell, encode_range } from '../utils/address.js';
import { parseStyles, parseFontNode, parseColorNode, ParsedStyles } from '../styles/parser.js';
import { parsePivotTableXml } from '../pivot/parser.js';
import { REL_TYPES, decodeXstring, unescapeXml } from './xml.js';
import { resolveTarget, relsPathFor } from './package.js';
import { ERROR_VALUES, format_cell, isDateFormat } from '../utils/format.js';
import { serialToDate, isoToSerial } from '../utils/dates.js';
import { shiftFormula } from '../utils/formula.js';
import { widthToChars, widthToPixels, pointsToPixels } from '../utils/units.js';
import {
  parseCommentsXml,
  parseConditionalFormatting,
  parseDataValidations,
  parseTableXml,
  textOf,
} from './read-features.js';
import { readDrawing } from '../drawing/read.js';

const base = { ignoreAttributes: false, attributeNamePrefix: '@_', parseTagValue: false, trimValues: false } as const;

const relsParser = new XMLParser({ ...base, isArray: (n) => n === 'Relationship' });
const wbParser = new XMLParser({ ...base, isArray: (n) => n === 'sheet' || n === 'definedName' || n === 'workbookView' });
const sstParser = new XMLParser({ ...base, isArray: (n) => n === 'si' || n === 'r' });
const sheetParser = new XMLParser({
  ...base,
  isArray: (n) =>
    [
      'row',
      'c',
      'col',
      'mergeCell',
      'hyperlink',
      'conditionalFormatting',
      'cfRule',
      'cfvo',
      'color',
      'formula',
      'dataValidation',
      'tablePart',
      'sheetView',
      'selection',
      'r',
    ].includes(n),
});

interface Rel {
  type: string;
  target: string;
  external: boolean;
}

type Files = Record<string, Uint8Array>;

function readText(files: Files, path: string): string | undefined {
  const bytes = files[path];
  return bytes ? strFromU8(bytes) : undefined;
}

function readRels(files: Files, source: string): Map<string, Rel> {
  const out = new Map<string, Rel>();
  const xml = readText(files, relsPathFor(source));
  if (!xml) return out;
  for (const r of relsParser.parse(xml).Relationships?.Relationship ?? []) {
    const external = r['@_TargetMode'] === 'External';
    out.set(r['@_Id'], {
      type: r['@_Type'] ?? '',
      target: external ? r['@_Target'] : resolveTarget(source, r['@_Target'] ?? ''),
      external,
    });
  }
  return out;
}

function findRel(rels: Map<string, Rel>, type: string): Rel | undefined {
  for (const r of rels.values()) if (r.type === type || r.type.endsWith('/' + type.split('/').pop())) return r;
  return undefined;
}

const bool = (v: unknown) => v === '1' || v === 'true';
const asArray = <T>(v: T | T[] | undefined): T[] => (v === undefined ? [] : Array.isArray(v) ? v : [v]);

/** Deep copy for plain JSON style objects so cells never share mutable styles. */
function cloneStyle<T>(o: T): T {
  return o && typeof o === 'object' ? (JSON.parse(JSON.stringify(o)) as T) : o;
}

interface SharedString {
  text: string;
  runs?: RichTextRun[];
}

function parseSi(si: any): SharedString {
  if (si === undefined || si === null) return { text: '' };
  if (typeof si !== 'object') return { text: decodeXstring(String(si)) };
  const runNodes = asArray<any>(si.r);
  if (runNodes.length) {
    const runs: RichTextRun[] = runNodes.map((r) => {
      const run: RichTextRun = { text: decodeXstring(textOf(r.t)) };
      if (r.rPr && typeof r.rPr === 'object') {
        const font = parseFontNode(r.rPr);
        if (Object.keys(font).length) run.font = font;
      }
      return run;
    });
    const text = (si.t !== undefined ? decodeXstring(textOf(si.t)) : '') + runs.map((r) => r.text).join('');
    return runs.some((r) => r.font) ? { text, runs } : { text };
  }
  return { text: decodeXstring(textOf(si.t)) };
}

export function parseSharedStrings(xml: string): string[] {
  return (sstParser.parse(xml).sst?.si ?? []).map((si: any) => parseSi(si).text);
}

function sheetSelected(opts: ParsingOptions | undefined, name: string, index: number): boolean {
  if (opts?.sheets === undefined) return true;
  const list = Array.isArray(opts.sheets) ? opts.sheets : [opts.sheets];
  return list.some((s) => s === name || s === index);
}

interface SheetContext {
  files: Files;
  opts?: ParsingOptions;
  styles: ParsedStyles;
  sst: SharedString[];
  date1904: boolean;
  wb: WorkBook;
}

function parseWorksheet(ctx: SheetContext, sheetName: string, part: string): WorkSheet {
  const { opts, styles, sst, date1904 } = ctx;
  const xml = readText(ctx.files, part) ?? '';
  const root = sheetParser.parse(xml).worksheet ?? {};
  const rels = readRels(ctx.files, part);
  const ws: WorkSheet = {};
  const keepStyles = opts?.cellStyles !== false;
  const keepFormulas = opts?.cellFormula !== false;
  const keepFormats = opts?.cellNF !== false;
  const keepText = opts?.cellText !== false;
  const sheetRows = opts?.sheetRows && opts.sheetRows > 0 ? opts.sheetRows : Infinity;

  // sheetPr / views
  const tabColor = parseColorNode(root.sheetPr?.tabColor);
  if (tabColor) ws['!tabColor'] = tabColor;
  const view = asArray<any>(root.sheetViews?.sheetView)[0];
  if (view) {
    const v: Record<string, any> = {};
    if (view['@_showGridLines'] === '0' || view['@_showGridLines'] === 'false') v.showGridLines = false;
    if (view['@_showRowColHeaders'] === '0') v.showRowColHeaders = false;
    if (view['@_showZeros'] === '0') v.showZeros = false;
    if (bool(view['@_rightToLeft'])) v.rightToLeft = true;
    if (view['@_zoomScale'] !== undefined && Number(view['@_zoomScale']) !== 100) v.zoomScale = Number(view['@_zoomScale']);
    if (Object.keys(v).length) ws['!views'] = v;
    const pane = view.pane;
    if (pane && pane['@_state'] && pane['@_state'] !== 'split') {
      ws['!freeze'] = {
        xSplit: Number(pane['@_xSplit'] ?? 0),
        ySplit: Number(pane['@_ySplit'] ?? 0),
        ...(pane['@_topLeftCell'] ? { topLeftCell: pane['@_topLeftCell'] } : {}),
      };
    }
  }

  // cols
  const cols: ColProperties[] = [];
  for (const col of asArray<any>(root.cols?.col)) {
    const min = Number(col['@_min'] ?? 1);
    const max = Math.min(Number(col['@_max'] ?? min), 16384);
    const p: ColProperties = {};
    const custom = bool(col['@_customWidth']);
    const styleIdx = col['@_style'] !== undefined ? Number(col['@_style']) : 0;
    if (col['@_width'] !== undefined && (custom || !styleIdx)) {
      p.width = Number(col['@_width']);
      p.wpx = widthToPixels(p.width);
      p.wch = widthToChars(p.width);
    }
    if (bool(col['@_hidden'])) p.hidden = true;
    if (col['@_outlineLevel'] !== undefined && Number(col['@_outlineLevel']) > 0) p.level = Number(col['@_outlineLevel']);
    if (styleIdx > 0 && keepStyles && styles.cellStyles[styleIdx]) p.s = cloneStyle(styles.cellStyles[styleIdx]);
    if (!Object.keys(p).length) continue;
    for (let c = min; c <= max; c++) cols[c - 1] = { ...p };
  }
  if (cols.length) ws['!cols'] = cols;

  // rows & cells
  const rows: RowProperties[] = [];
  const shared = new Map<string, { f: string; r: number; c: number }>();
  let minC = Infinity;
  let minR = Infinity;
  let maxC = -1;
  let maxR = -1;
  let prevRow = -1;
  for (const row of asArray<any>(root.sheetData?.row)) {
    const rIdx = row['@_r'] !== undefined ? Number(row['@_r']) - 1 : prevRow + 1;
    prevRow = rIdx;
    if (rIdx >= sheetRows) break;
    const rp: RowProperties = {};
    if (row['@_ht'] !== undefined && bool(row['@_customHeight'] ?? '1')) {
      rp.hpt = Number(row['@_ht']);
      rp.hpx = pointsToPixels(rp.hpt);
    }
    if (bool(row['@_hidden'])) rp.hidden = true;
    if (row['@_outlineLevel'] !== undefined && Number(row['@_outlineLevel']) > 0) rp.level = Number(row['@_outlineLevel']);
    if (bool(row['@_customFormat']) && row['@_s'] !== undefined && keepStyles) {
      const st = styles.cellStyles[Number(row['@_s'])];
      if (st && Object.keys(st).length) rp.s = cloneStyle(st);
    }
    if (Object.keys(rp).length) rows[rIdx] = rp;

    let prevCol = -1;
    for (const c of asArray<any>(row.c)) {
      const addr = c['@_r'] !== undefined ? decode_cell(c['@_r']) : { r: rIdx, c: prevCol + 1 };
      prevCol = addr.c;
      const ref = encode_cell(addr);
      const t = c['@_t'] ?? 'n';
      const styleIdx = c['@_s'] !== undefined ? Number(c['@_s']) : 0;
      const fmt = styles.cellFormats[styleIdx];
      const raw = c.v !== undefined ? textOf(c.v) : undefined;
      const cell: CellObject = {};

      switch (t) {
        case 's': {
          const item = sst[Number(raw)];
          cell.t = 's';
          cell.v = item?.text ?? '';
          if (item?.runs) cell.richText = item.runs;
          break;
        }
        case 'inlineStr': {
          const item = parseSi(c.is);
          cell.t = 's';
          cell.v = item.text;
          if (item.runs) cell.richText = item.runs;
          break;
        }
        case 'str':
          cell.t = 's';
          cell.v = decodeXstring(raw ?? '');
          break;
        case 'b':
          cell.t = 'b';
          cell.v = raw === '1' || raw === 'true' || raw === 'TRUE';
          break;
        case 'e':
          cell.t = 'e';
          cell.v = ERROR_VALUES[raw ?? ''] ?? raw;
          cell.w = raw;
          break;
        case 'd': {
          const serial = raw ? isoToSerial(raw, { date1904 }) : undefined;
          if (serial !== undefined) {
            cell.t = 'n';
            cell.v = serial;
            if (!fmt) cell.z = 14;
          }
          break;
        }
        default:
          if (raw !== undefined && raw.trim() !== '') {
            const n = Number(raw);
            if (Number.isFinite(n)) {
              cell.t = 'n';
              cell.v = n;
            } else {
              cell.t = 's';
              cell.v = raw;
            }
          } else cell.t = 'z';
      }

      if (c.f !== undefined && keepFormulas) {
        const fNode = c.f;
        const text = textOf(fNode);
        const kind = typeof fNode === 'object' ? fNode['@_t'] : undefined;
        if (kind === 'shared') {
          const si = fNode['@_si'];
          if (text) {
            shared.set(si, { f: text, r: addr.r, c: addr.c });
            cell.f = text;
          } else {
            const anchor = shared.get(si);
            if (anchor) cell.f = shiftFormula(anchor.f, addr.r - anchor.r, addr.c - anchor.c);
          }
        } else if (kind === 'array') {
          cell.f = text;
          cell.F = fNode['@_ref'] ?? ref;
        } else if (kind !== 'dataTable' && text) {
          cell.f = text;
        }
        if (cell.t === 'z' && cell.f) cell.t = 'n';
      }

      if (styleIdx > 0 && keepStyles) {
        const st = styles.cellStyles[styleIdx];
        if (st && Object.keys(st).length) cell.s = cloneStyle(st);
      }
      if (fmt && fmt !== 'General' && keepFormats && cell.t !== 'z' && cell.z === undefined) cell.z = fmt;

      if (cell.t === 'n' && opts?.cellDates && isDateFormat(cell.z ?? fmt)) {
        cell.t = 'd';
        cell.v = serialToDate(cell.v, { date1904 });
      }
      if (keepText && cell.t !== 'z' && cell.w === undefined && cell.v !== undefined) {
        const w = format_cell(cell);
        if (w !== '' || cell.t === 's') cell.w = w;
      }
      if (cell.t === 'z' && !cell.s && !cell.f) continue;

      ws[ref] = cell;
      if (addr.c < minC) minC = addr.c;
      if (addr.r < minR) minR = addr.r;
      if (addr.c > maxC) maxC = addr.c;
      if (addr.r > maxR) maxR = addr.r;
    }
  }
  if (rows.length) ws['!rows'] = rows;

  const dim = root.dimension?.['@_ref'];
  if (dim && sheetRows === Infinity) {
    try {
      const d = decode_range(dim);
      minC = Math.min(minC, d.s.c);
      minR = Math.min(minR, d.s.r);
      maxC = Math.max(maxC, d.e.c);
      maxR = Math.max(maxR, d.e.r);
    } catch {
      /* ignore */
    }
  }
  if (maxC >= 0 && maxR >= 0) ws['!ref'] = encode_range({ s: { c: minC, r: minR }, e: { c: maxC, r: maxR } });

  // merges / autofilter
  const merges = asArray<any>(root.mergeCells?.mergeCell).map((m) => decode_range(m['@_ref']));
  if (merges.length) ws['!merges'] = merges;
  if (root.autoFilter?.['@_ref']) ws['!autofilter'] = { ref: root.autoFilter['@_ref'] };

  // protection
  const sp = root.sheetProtection;
  if (sp && bool(sp['@_sheet'])) {
    const p: SheetProtection = {};
    for (const k of [
      'objects',
      'scenarios',
      'selectLockedCells',
      'selectUnlockedCells',
      'formatCells',
      'formatColumns',
      'formatRows',
      'insertColumns',
      'insertRows',
      'insertHyperlinks',
      'deleteColumns',
      'deleteRows',
      'sort',
      'autoFilter',
      'pivotTables',
    ] as const) {
      if (sp[`@_${k}`] !== undefined) p[k] = bool(sp[`@_${k}`]);
    }
    if (sp['@_hashValue']) {
      p.algorithmName = sp['@_algorithmName'];
      p.hashValue = sp['@_hashValue'];
      p.saltValue = sp['@_saltValue'];
      p.spinCount = Number(sp['@_spinCount']);
    } else if (sp['@_password']) p.passwordHash = sp['@_password'];
    ws['!protect'] = p;
  }

  // conditional formatting / data validation
  const cf = asArray<any>(root.conditionalFormatting);
  if (cf.length) {
    const parsed = parseConditionalFormatting(cf, styles.dxfs);
    if (parsed.length) ws['!condfmt'] = parsed;
  }
  const dv = asArray<any>(root.dataValidations?.dataValidation);
  if (dv.length) ws['!dataValidation'] = parseDataValidations(dv);

  // hyperlinks
  for (const h of asArray<any>(root.hyperlinks?.hyperlink)) {
    const rel = h['@_r:id'] ? rels.get(h['@_r:id']) : undefined;
    let target = rel ? rel.target : '';
    if (h['@_location']) target = rel ? `${target}#${h['@_location']}` : `#${h['@_location']}`;
    if (!target) continue;
    const link = { Target: target, ...(h['@_tooltip'] ? { Tooltip: h['@_tooltip'] } : {}) };
    const range = decode_range(h['@_ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const ref = encode_cell({ r, c });
        if (!ws[ref]) ws[ref] = { t: 'z' };
        ws[ref].l = link;
      }
    }
  }

  // margins & page setup
  const pm = root.pageMargins;
  if (pm) {
    ws['!margins'] = {};
    for (const k of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) {
      if (pm[`@_${k}`] !== undefined) ws['!margins'][k] = Number(pm[`@_${k}`]);
    }
  }
  const ps = root.pageSetup;
  if (ps) {
    const setup: Record<string, any> = {};
    if (ps['@_orientation'] && ps['@_orientation'] !== 'default') setup.orientation = ps['@_orientation'];
    for (const k of ['paperSize', 'scale', 'fitToWidth', 'fitToHeight']) if (ps[`@_${k}`] !== undefined) setup[k] = Number(ps[`@_${k}`]);
    if (Object.keys(setup).length) ws['!pageSetup'] = setup;
  }

  // related parts: comments, tables, drawings, pivots
  for (const [, rel] of rels) {
    if (rel.external) continue;
    if (rel.type === REL_TYPES.comments) {
      const xmlC = readText(ctx.files, rel.target);
      const vmlRel = findRel(rels, REL_TYPES.vmlDrawing);
      const vml = vmlRel ? readText(ctx.files, vmlRel.target) : undefined;
      if (!xmlC) continue;
      for (const [ref, { comment, visible }] of parseCommentsXml(xmlC, vml)) {
        if (!ws[ref]) ws[ref] = { t: 'z' };
        const list = (ws[ref].c ?? []) as Comments;
        list.push(comment);
        if (visible) list.hidden = false;
        else list.hidden = true;
        ws[ref].c = list;
      }
    } else if (rel.type === REL_TYPES.table) {
      const xmlT = readText(ctx.files, rel.target);
      if (xmlT) (ws['!tables'] ??= []).push(parseTableXml(xmlT));
    } else if (rel.type === REL_TYPES.drawing) {
      readDrawing(ctx.files, rel.target, ws);
    } else if (rel.type === REL_TYPES.pivotTable) {
      const pivotXml = readText(ctx.files, rel.target);
      const cacheRel = findRel(readRels(ctx.files, rel.target), REL_TYPES.pivotCacheDefinition);
      const cacheXml = cacheRel ? readText(ctx.files, cacheRel.target) : undefined;
      if (pivotXml && cacheXml) {
        const def = parsePivotTableXml(pivotXml, cacheXml, sheetName);
        (ws['!pivots'] ??= []).push(def);
        ctx.wb.PivotTables!.push(def);
      }
    }
  }
  if (ws['!tables'] && ws['!autofilter']) {
    // Excel keeps table filters inside the table part; a sheet-level filter there is redundant.
    delete ws['!autofilter'];
  }
  return ws;
}

function parseDocProps(files: Files, wb: WorkBook, rootRels: Map<string, Rel>) {
  const props: Record<string, any> = {};
  const coreRel = findRel(rootRels, REL_TYPES.coreProps);
  const core = coreRel ? readText(files, coreRel.target) : undefined;
  if (core) {
    const get = (tag: string) => {
      const m = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`).exec(core);
      return m ? unescapeXml(m[1]) : undefined;
    };
    const map: Record<string, string> = {
      Title: 'dc:title',
      Subject: 'dc:subject',
      Author: 'dc:creator',
      Keywords: 'cp:keywords',
      Comments: 'dc:description',
      LastAuthor: 'cp:lastModifiedBy',
      Category: 'cp:category',
    };
    for (const [k, tag] of Object.entries(map)) {
      const v = get(tag);
      if (v !== undefined) props[k] = v;
    }
    const created = get('dcterms:created');
    if (created) props.CreatedDate = new Date(created);
    const modified = get('dcterms:modified');
    if (modified) props.ModifiedDate = new Date(modified);
  }
  const appRel = findRel(rootRels, REL_TYPES.extendedProps);
  const app = appRel ? readText(files, appRel.target) : undefined;
  if (app) {
    for (const k of ['Company', 'Manager']) {
      const m = new RegExp(`<${k}>([\\s\\S]*?)</${k}>`).exec(app);
      if (m) props[k] = unescapeXml(m[1]);
    }
  }
  wb.Props = props;
}

export function parseXlsxFromFiles(files: Files, opts?: ParsingOptions): WorkBook {
  const rootRels = readRels(files, '');
  const officeDoc = findRel(rootRels, REL_TYPES.officeDocument);
  const wbPart = officeDoc?.target ?? 'xl/workbook.xml';
  const wbXml = readText(files, wbPart);
  if (!wbXml) throw new Error(`Invalid XLSX: missing workbook part ${wbPart}`);
  const wbRels = readRels(files, wbPart);
  const wbRoot = wbParser.parse(wbXml).workbook ?? {};

  const stylesRel = findRel(wbRels, REL_TYPES.styles);
  const stylesXml = stylesRel ? readText(files, stylesRel.target) : undefined;
  const styles: ParsedStyles = stylesXml ? parseStyles(stylesXml) : { cellStyles: [], cellFormats: [], dxfs: [] };
  const sstRel = findRel(wbRels, REL_TYPES.sharedStrings);
  const sstXml = sstRel ? readText(files, sstRel.target) : undefined;
  const sst: SharedString[] = sstXml ? (sstParser.parse(sstXml).sst?.si ?? []).map(parseSi) : [];

  const date1904 = bool(wbRoot.workbookPr?.['@_date1904']);
  const wb: WorkBook = { SheetNames: [], Sheets: {}, PivotTables: [], Workbook: { Sheets: [], Names: [], WBProps: { date1904 } } };
  parseDocProps(files, wb, rootRels);

  const view = asArray<any>(wbRoot.bookViews?.workbookView)[0];
  if (view?.['@_activeTab'] !== undefined) wb.Workbook!.Views = [{ activeTab: Number(view['@_activeTab']) }];
  const wp = wbRoot.workbookProtection;
  if (wp) {
    wb.Workbook!.Protection = {
      lockStructure: bool(wp['@_lockStructure']),
      ...(bool(wp['@_lockWindows']) ? { lockWindows: true } : {}),
      ...(wp['@_workbookHashValue']
        ? {
            algorithmName: wp['@_workbookAlgorithmName'],
            hashValue: wp['@_workbookHashValue'],
            saltValue: wp['@_workbookSaltValue'],
            spinCount: Number(wp['@_workbookSpinCount']),
          }
        : {}),
    };
  }
  for (const n of asArray<any>(wbRoot.definedNames?.definedName)) {
    wb.Workbook!.Names!.push({
      Name: n['@_name'],
      Ref: textOf(n),
      ...(n['@_localSheetId'] !== undefined ? { Sheet: Number(n['@_localSheetId']) } : {}),
      ...(bool(n['@_hidden']) ? { Hidden: true } : {}),
      ...(n['@_comment'] ? { Comment: n['@_comment'] } : {}),
    });
  }

  const ctx: SheetContext = { files, opts, styles, sst, date1904, wb };
  asArray<any>(wbRoot.sheets?.sheet).forEach((s, index) => {
    const name = s['@_name'];
    const rel = wbRels.get(s['@_r:id'] ?? s['@_id']);
    const state = s['@_state'];
    wb.Workbook!.Sheets!.push({ name, Hidden: state === 'hidden' ? 1 : state === 'veryHidden' ? 2 : 0 });
    wb.SheetNames.push(name);
    if (!rel || !files[rel.target] || !sheetSelected(opts, name, index)) {
      wb.Sheets[name] = {};
      return;
    }
    wb.Sheets[name] = parseWorksheet(ctx, name, rel.target);
  });
  return wb;
}

export function readXlsxSync(data: Uint8Array, opts?: ParsingOptions): WorkBook {
  return parseXlsxFromFiles(unzipSync(data), opts);
}

export async function readXlsxAsync(data: Uint8Array, opts?: ParsingOptions): Promise<WorkBook> {
  return new Promise((resolve, reject) => {
    unzip(data, (err, files) => {
      if (err) return reject(err);
      try {
        resolve(parseXlsxFromFiles(files, opts));
      } catch (e) {
        reject(e);
      }
    });
  });
}
