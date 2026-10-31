import { WorkBook, WorkSheet, PivotTableDefinition, PivotTableOptions, PivotDataField, PivotFilter, Range, CellAddress } from '../types.js';
import { defaultDataFieldName } from '../pivot/aggregate.js';
import { decode_range, decode_cell, encode_range, normaliseMerges, parse_range_spec } from './address.js';
import { readSource } from '../pivot/source.js';
import { buildLayout } from '../pivot/layout.js';
import { renderPivot, pivotArea } from '../pivot/render.js';
import { checkSheetName } from '../xlsx/workbook.js';

/** New workbook, optionally with a first sheet (SheetJS signature). */
export function book_new(ws?: WorkSheet, wsname?: string): WorkBook {
  const wb: WorkBook = { SheetNames: [], Sheets: {}, Props: {}, Custprops: {}, PivotTables: [] };
  if (ws) book_append_sheet(wb, ws, wsname || 'Sheet1');
  return wb;
}

/**
 * Appends a sheet and returns its name. With `roll`, a taken name gets a
 * numeric suffix bumped ("Sheet1" -> "Sheet2", "Data" -> "Data1").
 */
export function book_append_sheet(wb: WorkBook, ws: WorkSheet, name?: string, roll?: boolean): string {
  let sheetName = name;
  const taken = (n: string) => wb.SheetNames.some((s) => s.toLowerCase() === n.toLowerCase());
  if (!sheetName) {
    let i = 1;
    while (taken(`Sheet${i}`)) i++;
    sheetName = `Sheet${i}`;
  }
  if (roll && taken(sheetName)) {
    const m = /\d+$/.exec(sheetName);
    let i = m ? Number(m[0]) : 0;
    const root = m ? sheetName.slice(0, m.index) : sheetName;
    do sheetName = `${root}${++i}`;
    while (taken(sheetName));
  }
  checkSheetName(sheetName);
  if (taken(sheetName)) throw new Error(`Worksheet with name |${sheetName}| already exists!`);
  wb.SheetNames.push(sheetName);
  wb.Sheets[sheetName] = ws;
  return sheetName;
}

/** 0 = visible, 1 = hidden, 2 = very hidden. */
export function book_set_sheet_visibility(wb: WorkBook, sheet: number | string, visibility: 0 | 1 | 2): void {
  const idx = typeof sheet === 'number' ? sheet : wb.SheetNames.indexOf(sheet);
  if (idx < 0 || idx >= wb.SheetNames.length) throw new Error(`Cannot find sheet |${sheet}|`);
  if (![0, 1, 2].includes(visibility)) throw new Error(`Bad sheet visibility setting ${visibility}`);
  wb.Workbook ??= {};
  wb.Workbook.Sheets ??= [];
  wb.Workbook.Sheets[idx] = { ...(wb.Workbook.Sheets[idx] ?? {}), Hidden: visibility };
}

function normaliseDataField(v: string | PivotDataField): PivotDataField {
  if (typeof v === 'string') return { field: v, subtotal: 'sum', name: defaultDataFieldName('sum', v) };
  const subtotal = v.subtotal || 'sum';
  return { ...v, subtotal, name: v.name || defaultDataFieldName(subtotal, v.field) };
}

export function add_pivot_table(wb: WorkBook, options: PivotTableOptions): PivotTableDefinition {
  const sourceSheet = wb.Sheets[options.sourceSheet];
  if (!sourceSheet) throw new Error(`Source worksheet not found: ${options.sourceSheet}`);

  const sourceRange: Range = options.sourceRange
    ? parse_range_spec(options.sourceRange)
    : sourceSheet['!ref']
      ? decode_range(sourceSheet['!ref'])
      : { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

  let targetSheetName = options.targetSheet || `${options.sourceSheet} Pivot`.slice(0, 31);
  let targetSheet = wb.Sheets[targetSheetName];
  if (!targetSheet) {
    targetSheet = {};
    targetSheetName = book_append_sheet(wb, targetSheet, targetSheetName, true);
  }

  const targetCell: CellAddress = options.targetCell
    ? typeof options.targetCell === 'string'
      ? decode_cell(options.targetCell)
      : options.targetCell
    : { c: 0, r: 2 };

  const filters: PivotFilter[] = (options.filters ?? []).map((f) => (typeof f === 'string' ? { field: f } : f));
  const count = wb.PivotTables?.length ?? 0;
  const pivotDef: PivotTableDefinition = {
    id: count + 1,
    name: options.name || `PivotTable${count + 1}`,
    sourceSheet: options.sourceSheet,
    sourceRange,
    targetSheet: targetSheetName,
    targetCell,
    rows: options.rows || [],
    cols: options.cols || [],
    values: (options.values || []).map(normaliseDataField),
    filters,
    showRowTotals: options.showRowTotals ?? true,
    showColTotals: options.showColTotals ?? true,
    showSubtotals: options.showSubtotals ?? true,
    styleName: options.styleName || 'PivotStyleLight16',
    renderCells: options.renderCells ?? true,
    ...(options.title !== undefined ? { title: options.title } : {}),
  };

  // validates fields/filters up front and fixes the pivot area
  const date1904 = Boolean(wb.Workbook?.WBProps?.date1904);
  const layout = buildLayout(pivotDef, readSource(sourceSheet, sourceRange, { date1904 }), { date1904 });
  pivotDef.targetRange = layout.geometry.location;
  const area = pivotArea(layout);
  const overlaps = (a: Range, b: Range) => a.s.r <= b.e.r && b.s.r <= a.e.r && a.s.c <= b.e.c && b.s.c <= a.e.c;
  const where = `Pivot table ${pivotDef.name}`;
  const pivots = new Set<PivotTableDefinition>(wb.PivotTables ?? []);
  for (const n of wb.SheetNames) for (const p of wb.Sheets[n]?.['!pivots'] ?? []) pivots.add(p);
  for (const other of pivots) {
    if (!other.targetRange) continue;
    // an existing pivot owns its title/filter rows too
    const otherArea: Range = { s: { r: Math.min(other.targetCell.r, other.targetRange.s.r), c: other.targetRange.s.c }, e: other.targetRange.e };
    if (other.targetSheet === targetSheetName && overlaps(area, otherArea)) throw new Error(`${where} overlaps pivot table ${other.name}`);
    if (other.targetSheet === options.sourceSheet && overlaps(layout.source.range, otherArea)) {
      throw new Error(`${where}: source range includes the cells of pivot table ${other.name}`);
    }
  }
  if (targetSheetName === options.sourceSheet && overlaps(area, layout.source.range)) {
    throw new Error(`${where} would overwrite its own source range; place it elsewhere or on another sheet`);
  }
  for (const t of targetSheet['!tables'] ?? []) {
    if (overlaps(area, decode_range(t.ref))) throw new Error(`${where} overlaps table ${t.name}`);
  }
  for (const m of normaliseMerges(targetSheet['!merges'])) {
    if (overlaps(area, m)) throw new Error(`${where} overlaps merged cells ${encode_range(m)}`);
  }
  const af = targetSheet['!autofilter']?.ref;
  if (af && overlaps(area, decode_range(af))) throw new Error(`${where} overlaps the sheet autofilter`);
  if (pivotDef.renderCells !== false) renderPivot(targetSheet, pivotDef, layout);

  (wb.PivotTables ??= []).push(pivotDef);
  (targetSheet['!pivots'] ??= []).push(pivotDef);
  return pivotDef;
}
