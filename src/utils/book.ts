import {
  WorkBook,
  WorkSheet,
  PivotTableDefinition,
  PivotTableOptions,
  PivotDataField,
  Range,
  CellAddress,
} from '../types.js';
import { decode_range, decode_cell, parse_range_spec } from './address.js';
import { extractSheetData, computePivotData, renderPivotToWorksheet } from '../pivot/compute.js';

export function book_new(): WorkBook {
  return {
    SheetNames: [],
    Sheets: {},
    Props: {},
    Custprops: {},
    PivotTables: [],
  };
}

export function book_append_sheet(
  wb: WorkBook,
  ws: WorkSheet,
  name?: string,
  roll?: boolean
): void {
  let sheetName = name || 'Sheet1';
  if (!name) {
    let idx = 1;
    while (wb.SheetNames.includes(`Sheet${idx}`)) {
      idx++;
    }
    sheetName = `Sheet${idx}`;
  }

  if (wb.SheetNames.includes(sheetName)) {
    if (roll) {
      let idx = 2;
      while (wb.SheetNames.includes(`${sheetName} ${idx}`)) {
        idx++;
      }
      sheetName = `${sheetName} ${idx}`;
    } else {
      throw new Error(`Sheet name already exists: ${sheetName}`);
    }
  }

  wb.SheetNames.push(sheetName);
  wb.Sheets[sheetName] = ws;
}

export function add_pivot_table(wb: WorkBook, options: PivotTableOptions): PivotTableDefinition {
  const sourceSheet = wb.Sheets[options.sourceSheet];
  if (!sourceSheet) {
    throw new Error(`Source worksheet not found: ${options.sourceSheet}`);
  }

  const sourceRange: Range = options.sourceRange
    ? parse_range_spec(options.sourceRange)
    : sourceSheet['!ref']
    ? decode_range(sourceSheet['!ref'])
    : { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

  let targetSheetName = options.targetSheet;
  if (!targetSheetName) {
    targetSheetName = `${options.sourceSheet} Pivot`;
  }

  let targetSheet = wb.Sheets[targetSheetName];
  if (!targetSheet) {
    targetSheet = {};
    book_append_sheet(wb, targetSheet, targetSheetName, true);
    // Might have been rolled if name collision
    targetSheetName = wb.SheetNames[wb.SheetNames.length - 1];
  }

  const targetCell: CellAddress = options.targetCell
    ? typeof options.targetCell === 'string'
      ? decode_cell(options.targetCell)
      : options.targetCell
    : { c: 0, r: 2 }; // A3 default

  const values: PivotDataField[] = (options.values || []).map((v) => {
    if (typeof v === 'string') {
      return { field: v, subtotal: 'sum', name: `Sum of ${v}` };
    }
    const subtotal = v.subtotal || 'sum';
    const subLabel = subtotal.charAt(0).toUpperCase() + subtotal.slice(1);
    return {
      subtotal,
      name: v.name || `${subLabel} of ${v.field}`,
      ...v,
    };
  });

  const pivotDef: PivotTableDefinition = {
    id: (wb.PivotTables?.length || 0) + 1,
    name: options.name || `PivotTable${(wb.PivotTables?.length || 0) + 1}`,
    sourceSheet: options.sourceSheet,
    sourceRange,
    targetSheet: targetSheetName,
    targetCell,
    rows: options.rows || [],
    cols: options.cols || [],
    values,
    filters: options.filters || [],
    showRowTotals: options.showRowTotals ?? true,
    showColTotals: options.showColTotals ?? true,
    styleName: options.styleName || 'PivotStyleLight16',
    renderCells: options.renderCells ?? true,
  };

  if (pivotDef.renderCells !== false) {
    const { rows } = extractSheetData(sourceSheet, sourceRange);
    const comp = computePivotData(rows, {
      rows: pivotDef.rows,
      cols: pivotDef.cols,
      values: pivotDef.values,
      filters: pivotDef.filters,
      showRowTotals: pivotDef.showRowTotals,
      showColTotals: pivotDef.showColTotals,
    });
    const renderedRange = renderPivotToWorksheet(targetSheet, pivotDef, comp);
    pivotDef.targetRange = renderedRange;
  }

  if (!wb.PivotTables) {
    wb.PivotTables = [];
  }
  wb.PivotTables.push(pivotDef);

  if (!targetSheet['!pivots']) {
    targetSheet['!pivots'] = [];
  }
  targetSheet['!pivots'].push(pivotDef);

  return pivotDef;
}
