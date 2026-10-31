import { WorkBook, WorkSheet, PivotTableDefinition, PivotTableOptions, PivotDataField, PivotFilter, Range, CellAddress } from '../types.js';
import { defaultDataFieldName } from '../pivot/aggregate.js';
import { decode_range, decode_cell, parse_range_spec } from './address.js';
import { readSource } from '../pivot/source.js';
import { buildLayout } from '../pivot/layout.js';
import { renderPivot } from '../pivot/render.js';
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
  const layout = buildLayout(pivotDef, readSource(sourceSheet, sourceRange));
  pivotDef.targetRange = layout.geometry.location;
  for (const other of [...(wb.PivotTables ?? []), ...(targetSheet['!pivots'] ?? [])]) {
    const o = other.targetRange;
    const a = layout.geometry.location;
    if (other.targetSheet === targetSheetName && o && a.s.r <= o.e.r && o.s.r <= a.e.r && a.s.c <= o.e.c && o.s.c <= a.e.c) {
      throw new Error(`Pivot table ${pivotDef.name} overlaps pivot table ${other.name}`);
    }
  }
  if (pivotDef.renderCells !== false) renderPivot(targetSheet, pivotDef, layout);

  (wb.PivotTables ??= []).push(pivotDef);
  (targetSheet['!pivots'] ??= []).push(pivotDef);
  return pivotDef;
}
