/**
 * XLSX package writer: orchestrates workbook, worksheets, shared strings,
 * styles, theme, document properties and pivot tables.
 */
import { zipSync, zip } from 'fflate';
import { WorkBook, WritingOptions, PivotTableDefinition, OutputType } from '../types.js';
import { createContext } from './context.js';
import { CONTENT_TYPES, REL_TYPES } from './xml.js';
import { writeWorksheet } from './worksheet.js';
import { buildWorkbookXml, checkSheetNames, sheetVisibility, writeDocProps } from './workbook.js';
import { buildThemeXml } from '../styles/theme.js';
import { generateOpenXmlPivotParts } from '../pivot/openxml.js';

export { SharedStringsBuilder } from './sst.js';

/**
 * Pivots from the sheets' '!pivots' and wb.PivotTables. A pivot is identified
 * by target sheet + name (Excel requires unique names per sheet), so the same
 * definition listed twice, or a cloned copy, is written once; the sheet's
 * entry wins.
 */
function collectPivots(wb: WorkBook, names: string[]): PivotTableDefinition[] {
  const out = new Map<string, PivotTableDefinition>();
  const add = (p: PivotTableDefinition, overwrite: boolean) => {
    if (!p) return;
    const key = `${p.targetSheet.toLowerCase()}\u0000${p.name.toLowerCase()}`;
    if (overwrite || !out.has(key)) out.set(key, p);
  };
  for (const n of names) (wb.Sheets[n]?.['!pivots'] ?? []).forEach((p) => add(p, true));
  (wb.PivotTables ?? []).forEach((p) => add(p, false));
  return [...out.values()];
}

export function buildXlsxArchiveFiles(wb: WorkBook, opts?: WritingOptions): Record<string, Uint8Array> {
  const names = wb.SheetNames && wb.SheetNames.length > 0 ? wb.SheetNames : ['Sheet1'];
  checkSheetNames(names);
  const ctx = createContext(wb, opts);
  const { pkg } = ctx;

  const visibility = names.map((_, i) => sheetVisibility(wb, i));
  if (!visibility.includes(0)) throw new Error('At least one sheet must be visible');
  let active = wb.Workbook?.Views?.[0]?.activeTab ?? 0;
  if (!(active >= 0 && active < names.length) || visibility[active] !== 0) active = visibility.indexOf(0);
  ctx.activeTab = active;

  pkg.addRel('', REL_TYPES.officeDocument, 'xl/workbook.xml');
  writeDocProps(ctx);

  const sheetParts = names.map((name, i) => {
    const part = `xl/worksheets/sheet${i + 1}.xml`;
    writeWorksheet(ctx, name, wb.Sheets?.[name] ?? {}, i, part);
    return part;
  });
  const sheets = names.map((name, i) => ({
    name,
    hidden: visibility[i],
    rid: pkg.addRel('xl/workbook.xml', REL_TYPES.worksheet, sheetParts[i]),
  }));

  const caches: Array<{ id: number; rid: string }> = [];
  collectPivots(wb, names).forEach((pivot, i) => {
    const id = i + 1;
    const source = names.includes(pivot.sourceSheet) ? wb.Sheets[pivot.sourceSheet] : undefined;
    if (!source) throw new Error(`Pivot table ${pivot.name}: source sheet "${pivot.sourceSheet}" not found`);
    const target = names.indexOf(pivot.targetSheet);
    if (target === -1) throw new Error(`Pivot table ${pivot.name}: target sheet "${pivot.targetSheet}" not found`);
    const cachePart = `xl/pivotCache/pivotCacheDefinition${id}.xml`;
    const recordsPart = `xl/pivotCache/pivotCacheRecords${id}.xml`;
    const tablePart = `xl/pivotTables/pivotTable${id}.xml`;
    const recordsRid = pkg.addRel(cachePart, REL_TYPES.pivotCacheRecords, recordsPart);
    const parts = generateOpenXmlPivotParts(source, pivot, id, recordsRid, ctx.styles, ctx.date1904);
    pkg.addPart(cachePart, parts.cacheDefinitionXml, CONTENT_TYPES.pivotCacheDefinition);
    pkg.addPart(recordsPart, parts.cacheRecordsXml, CONTENT_TYPES.pivotCacheRecords);
    pkg.addPart(tablePart, parts.pivotTableXml, CONTENT_TYPES.pivotTable);
    pkg.addRel(tablePart, REL_TYPES.pivotCacheDefinition, cachePart);
    pkg.addRel(sheetParts[target], REL_TYPES.pivotTable, tablePart);
    caches.push({ id, rid: pkg.addRel('xl/workbook.xml', REL_TYPES.pivotCacheDefinition, cachePart) });
  });

  pkg.addPart('xl/theme/theme1.xml', buildThemeXml(), CONTENT_TYPES.theme);
  pkg.addRel('xl/workbook.xml', REL_TYPES.theme, 'xl/theme/theme1.xml');
  pkg.addPart('xl/styles.xml', ctx.styles.buildXml(), CONTENT_TYPES.styles);
  pkg.addRel('xl/workbook.xml', REL_TYPES.styles, 'xl/styles.xml');
  if (ctx.sst.count > 0) {
    pkg.addPart('xl/sharedStrings.xml', ctx.sst.buildXml(), CONTENT_TYPES.sharedStrings);
    pkg.addRel('xl/workbook.xml', REL_TYPES.sharedStrings, 'xl/sharedStrings.xml');
  }
  // .xlsm needs the macro-enabled content type (Excel rejects an extension/content mismatch);
  // node-xls writes no VBA project, so the file simply contains no macros.
  const wbType = opts?.bookType === 'xlsm' ? CONTENT_TYPES.workbookMacro : CONTENT_TYPES.workbook;
  pkg.addPart('xl/workbook.xml', buildWorkbookXml(ctx, sheets, caches), wbType);

  return pkg.build();
}

export function writeXlsxSync(wb: WorkBook, opts?: WritingOptions): Uint8Array {
  return zipSync(buildXlsxArchiveFiles(wb, opts), { level: opts?.compression === false ? 0 : 6 });
}

export async function writeXlsxAsync(wb: WorkBook, opts?: WritingOptions): Promise<Uint8Array> {
  const files = buildXlsxArchiveFiles(wb, opts);
  return new Promise((resolve, reject) => {
    zip(files, { level: opts?.compression === false ? 0 : 6 }, (err, data) => (err ? reject(err) : resolve(data)));
  });
}

export function convertOutputType(bytes: Uint8Array, type?: OutputType): any {
  const outType = type || (typeof Buffer !== 'undefined' ? 'buffer' : 'array');
  switch (outType) {
    case 'array':
      return bytes;
    case 'binary': {
      let binary = '';
      for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
      return binary;
    }
    case 'base64': {
      if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
      let binary = '';
      for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
      return btoa(binary);
    }
    case 'string':
      return new TextDecoder('utf-8').decode(bytes);
    case 'buffer':
    case 'file':
    default:
      return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
}
