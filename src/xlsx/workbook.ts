/**
 * workbook.xml and document property parts.
 * CT_Workbook child order: workbookPr, workbookProtection, bookViews, sheets,
 * definedNames, calcPr, pivotCaches.
 */
import { WorkBook } from '../types.js';
import { WriteContext } from './context.js';
import { CONTENT_TYPES, NS, REL_TYPES, XML_HEADER, attrs, escapeXml } from './xml.js';
import { workbookProtectionXml } from '../features/protection.js';

const INVALID_SHEET_CHARS = /[\\/?*[\]:]/;

/** Throws when a sheet name is not acceptable to Excel (mirrors SheetJS check_ws_name). */
export function checkSheetName(name: string): void {
  if (typeof name !== 'string' || name === '') throw new Error('Sheet name cannot be blank');
  if (name.length > 31) throw new Error(`Sheet name cannot exceed 31 chars: "${name}"`);
  if (name.startsWith("'") || name.endsWith("'")) throw new Error(`Sheet name cannot start or end with an apostrophe: "${name}"`);
  if (name.toLowerCase() === 'history') throw new Error("Sheet name cannot be 'History'");
  if (INVALID_SHEET_CHARS.test(name)) throw new Error(`Sheet name cannot contain : \\ / ? * [ ]: "${name}"`);
}

export function checkSheetNames(names: string[]): void {
  const seen = new Set<string>();
  for (const n of names) {
    checkSheetName(n);
    const key = n.toLowerCase();
    if (seen.has(key)) throw new Error(`Duplicate sheet name: "${n}"`);
    seen.add(key);
  }
}

export interface SheetEntry {
  name: string;
  rid: string;
  hidden: 0 | 1 | 2;
}

export function buildWorkbookXml(ctx: WriteContext, sheets: SheetEntry[], pivotCaches: Array<{ id: number; rid: string }>): string {
  const wb = ctx.wb;
  const wbPr = `<workbookPr${attrs([['date1904', ctx.date1904]])}/>`;
  const protection = workbookProtectionXml(wb.Workbook?.Protection);
  const firstVisible = Math.max(
    0,
    sheets.findIndex((s) => s.hidden === 0)
  );
  const bookViews = `<bookViews><workbookView${attrs([
    ['firstSheet', firstVisible || undefined],
    ['activeTab', ctx.activeTab || undefined],
  ])}/></bookViews>`;
  const sheetsXml = sheets
    .map(
      (s, i) =>
        `<sheet${attrs([
          ['name', s.name],
          ['sheetId', i + 1],
          ['state', s.hidden === 1 ? 'hidden' : s.hidden === 2 ? 'veryHidden' : undefined],
          ['r:id', s.rid],
        ])}/>`
    )
    .join('');

  const names: Array<{ name: string; sheet?: number; ref: string; hidden?: boolean; comment?: string }> = [];
  for (const n of wb.Workbook?.Names ?? []) {
    if (!n || !n.Name || !n.Ref) continue;
    names.push({ name: n.Name, sheet: n.Sheet ?? undefined, ref: String(n.Ref).replace(/^=/, ''), hidden: n.Hidden, comment: n.Comment });
  }
  for (const a of ctx.autoNames) {
    if (!names.some((n) => n.name === a.name && n.sheet === a.sheet)) names.push({ name: a.name, sheet: a.sheet, ref: a.ref, hidden: true });
  }
  const definedNames = names.length
    ? `<definedNames>${names
        .map(
          (n) =>
            `<definedName${attrs([
              ['name', n.name],
              ['comment', n.comment],
              ['localSheetId', n.sheet],
              ['hidden', n.hidden],
            ])}>${escapeXml(n.ref)}</definedName>`
        )
        .join('')}</definedNames>`
    : '';
  const calcPr = ctx.hasFormulas ? '<calcPr calcId="191029" fullCalcOnLoad="1"/>' : '<calcPr calcId="191029"/>';
  const caches = pivotCaches.length
    ? `<pivotCaches>${pivotCaches.map((p) => `<pivotCache cacheId="${p.id}" r:id="${p.rid}"/>`).join('')}</pivotCaches>`
    : '';
  return `${XML_HEADER}<workbook xmlns="${NS.main}" xmlns:r="${NS.r}">${wbPr}${protection}${bookViews}<sheets>${sheetsXml}</sheets>${definedNames}${calcPr}${caches}</workbook>`;
}

function w3cdtf(d: unknown): string {
  const date = d instanceof Date ? d : d ? new Date(d as string) : new Date();
  return (isNaN(date.getTime()) ? new Date() : date).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function writeDocProps(ctx: WriteContext): void {
  const p: Record<string, any> = { ...(ctx.wb.Props ?? {}), ...(ctx.opts?.Props ?? {}) };
  const tag = (name: string, v: unknown) => (v === undefined || v === null || v === '' ? '' : `<${name}>${escapeXml(v)}</${name}>`);
  const core =
    `${XML_HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    tag('dc:title', p.Title) +
    tag('dc:subject', p.Subject) +
    tag('dc:creator', p.Author ?? 'node-xls') +
    tag('cp:keywords', p.Keywords) +
    tag('dc:description', p.Comments) +
    tag('cp:lastModifiedBy', p.LastAuthor ?? p.Author ?? 'node-xls') +
    tag('cp:category', p.Category) +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${w3cdtf(p.CreatedDate)}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${w3cdtf(p.ModifiedDate)}</dcterms:modified>` +
    '</cp:coreProperties>';
  const app =
    `${XML_HEADER}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    '<Application>node-xls</Application>' +
    tag('Manager', p.Manager) +
    tag('Company', p.Company) +
    '</Properties>';
  ctx.pkg.addPart('docProps/core.xml', core, CONTENT_TYPES.coreProps);
  ctx.pkg.addPart('docProps/app.xml', app, CONTENT_TYPES.extendedProps);
  ctx.pkg.addRel('', REL_TYPES.coreProps, 'docProps/core.xml');
  ctx.pkg.addRel('', REL_TYPES.extendedProps, 'docProps/app.xml');
}

export function sheetVisibility(wb: WorkBook, i: number): 0 | 1 | 2 {
  const h = wb.Workbook?.Sheets?.[i]?.Hidden;
  return h === 1 || h === 2 ? h : 0;
}
