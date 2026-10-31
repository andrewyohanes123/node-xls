import {
  PivotTableDefinition,
  PivotTableOptions,
  PivotDataField,
  PivotSubtotal,
  WorkSheet,
  WorkBook,
  CellObject,
  Range,
  CellAddress,
} from '../types.js';
import { encode_cell, decode_range, parse_range_spec } from '../utils/address.js';

export interface PivotComputationResult {
  rowKeys: string[][];
  colKeys: string[][];
  matrix: Record<string, Record<string, Record<string, number>>>; // [rowKeyStr][colKeyStr][valName]
  rowTotals: Record<string, Record<string, number>>; // [rowKeyStr][valName]
  colTotals: Record<string, Record<string, number>>; // [colKeyStr][valName]
  grandTotal: Record<string, number>; // [valName]
  dataFields: PivotDataField[];
  renderedRange: Range;
}

function calculateAggregation(values: any[], subtotal: PivotSubtotal = 'sum'): number {
  const numVals = values
    .map((v) => (typeof v === 'number' ? v : parseFloat(v)))
    .filter((v) => !isNaN(v) && isFinite(v));

  if (subtotal === 'count') {
    return values.filter((v) => v !== null && v !== undefined && v !== '').length;
  }
  if (subtotal === 'countNums') {
    return numVals.length;
  }
  if (numVals.length === 0) {
    return 0;
  }

  const sum = numVals.reduce((acc, v) => acc + v, 0);

  switch (subtotal) {
    case 'sum':
      return sum;
    case 'average':
      return sum / numVals.length;
    case 'max':
      return Math.max(...numVals);
    case 'min':
      return Math.min(...numVals);
    case 'product':
      return numVals.reduce((acc, v) => acc * v, 1);
    case 'var':
    case 'stdDev': {
      if (numVals.length < 2) return 0;
      const mean = sum / numVals.length;
      const variance = numVals.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0) / (numVals.length - 1);
      return subtotal === 'var' ? variance : Math.sqrt(variance);
    }
    case 'varp':
    case 'stdDevp': {
      const mean = sum / numVals.length;
      const variance = numVals.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0) / numVals.length;
      return subtotal === 'varp' ? variance : Math.sqrt(variance);
    }
    default:
      return sum;
  }
}

export function extractSheetData(
  sheet: WorkSheet,
  range: Range
): { headers: string[]; rows: Array<Record<string, any>> } {
  const headers: string[] = [];
  const headerRow = range.s.r;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const cellKey = encode_cell({ c, r: headerRow });
    const cell = sheet[cellKey];
    const val = cell && cell.v !== undefined ? String(cell.v).trim() : `Col${c + 1}`;
    headers.push(val);
  }

  const dataRows: Array<Record<string, any>> = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const rowObj: Record<string, any> = {};
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellKey = encode_cell({ c, r });
      const cell = sheet[cellKey];
      const val = cell ? cell.v : undefined;
      const colName = headers[c - range.s.c];
      rowObj[colName] = val;
      if (val !== undefined && val !== null && val !== '') {
        hasData = true;
      }
    }
    if (hasData) {
      dataRows.push(rowObj);
    }
  }

  return { headers, rows: dataRows };
}

export function computePivotData(
  dataRows: Array<Record<string, any>>,
  options: {
    rows?: string[];
    cols?: string[];
    values?: (string | PivotDataField)[];
    filters?: string[];
    showRowTotals?: boolean;
    showColTotals?: boolean;
  }
): {
  rowKeys: string[][];
  colKeys: string[][];
  matrix: Record<string, Record<string, Record<string, number>>>;
  rowTotals: Record<string, Record<string, number>>;
  colTotals: Record<string, Record<string, number>>;
  grandTotal: Record<string, number>;
  dataFields: PivotDataField[];
} {
  const rows = options.rows || [];
  const cols = options.cols || [];
  const dataFields: PivotDataField[] = (options.values || []).map((v) => {
    if (typeof v === 'string') {
      return {
        field: v,
        subtotal: 'sum',
        name: `Sum of ${v}`,
      };
    }
    const subtotal = v.subtotal || 'sum';
    const subLabel = subtotal.charAt(0).toUpperCase() + subtotal.slice(1);
    return {
      subtotal,
      name: v.name || `${subLabel} of ${v.field}`,
      ...v,
    };
  });

  if (dataFields.length === 0 && dataRows.length > 0) {
    const firstKey = Object.keys(dataRows[0])[0];
    dataFields.push({
      field: firstKey,
      subtotal: 'count',
      name: `Count of ${firstKey}`,
    });
  }

  const rowKeyMap = new Map<string, string[]>();
  const colKeyMap = new Map<string, string[]>();

  // Group items
  const bucketMap = new Map<string, Array<Record<string, any>>>();

  for (const item of dataRows) {
    const rKeyArr = rows.map((rf) => (item[rf] !== undefined && item[rf] !== null ? String(item[rf]) : ''));
    const cKeyArr = cols.map((cf) => (item[cf] !== undefined && item[cf] !== null ? String(item[cf]) : ''));

    const rKeyStr = rKeyArr.join('\u0000');
    const cKeyStr = cKeyArr.join('\u0000');

    if (!rowKeyMap.has(rKeyStr)) rowKeyMap.set(rKeyStr, rKeyArr);
    if (!colKeyMap.has(cKeyStr)) colKeyMap.set(cKeyStr, cKeyArr);

    const fullKey = `${rKeyStr}\u0001${cKeyStr}`;
    let bucket = bucketMap.get(fullKey);
    if (!bucket) {
      bucket = [];
      bucketMap.set(fullKey, bucket);
    }
    bucket.push(item);
  }

  // If no cols, add empty col key
  if (cols.length === 0) {
    colKeyMap.set('', []);
  }
  // If no rows, add empty row key
  if (rows.length === 0) {
    rowKeyMap.set('', []);
  }

  const rowKeys = Array.from(rowKeyMap.values()).sort((a, b) => a.join(' / ').localeCompare(b.join(' / ')));
  const colKeys = Array.from(colKeyMap.values()).sort((a, b) => a.join(' / ').localeCompare(b.join(' / ')));

  const matrix: Record<string, Record<string, Record<string, number>>> = {};
  const rowTotals: Record<string, Record<string, number>> = {};
  const colTotals: Record<string, Record<string, number>> = {};
  const grandTotal: Record<string, number> = {};

  // Initialize
  for (const df of dataFields) {
    grandTotal[df.name!] = 0;
  }

  for (const rKeyArr of rowKeys) {
    const rKeyStr = rKeyArr.join('\u0000');
    matrix[rKeyStr] = {};
    rowTotals[rKeyStr] = {};
    for (const df of dataFields) {
      rowTotals[rKeyStr][df.name!] = 0;
    }

    const rowItems: Array<Record<string, any>> = [];

    for (const cKeyArr of colKeys) {
      const cKeyStr = cKeyArr.join('\u0000');
      const fullKey = `${rKeyStr}\u0001${cKeyStr}`;
      const items = bucketMap.get(fullKey) || [];
      rowItems.push(...items);

      matrix[rKeyStr][cKeyStr] = {};
      for (const df of dataFields) {
        const vals = items.map((it) => it[df.field]);
        const agg = calculateAggregation(vals, df.subtotal);
        matrix[rKeyStr][cKeyStr][df.name!] = agg;
      }
    }

    // Row total
    for (const df of dataFields) {
      const vals = rowItems.map((it) => it[df.field]);
      rowTotals[rKeyStr][df.name!] = calculateAggregation(vals, df.subtotal);
    }
  }

  // Column totals
  for (const cKeyArr of colKeys) {
    const cKeyStr = cKeyArr.join('\u0000');
    colTotals[cKeyStr] = {};
    const colItems: Array<Record<string, any>> = [];

    for (const rKeyArr of rowKeys) {
      const rKeyStr = rKeyArr.join('\u0000');
      const fullKey = `${rKeyStr}\u0001${cKeyStr}`;
      const items = bucketMap.get(fullKey) || [];
      colItems.push(...items);
    }

    for (const df of dataFields) {
      const vals = colItems.map((it) => it[df.field]);
      colTotals[cKeyStr][df.name!] = calculateAggregation(vals, df.subtotal);
    }
  }

  // Grand total
  for (const df of dataFields) {
    const vals = dataRows.map((it) => it[df.field]);
    grandTotal[df.name!] = calculateAggregation(vals, df.subtotal);
  }

  return {
    rowKeys,
    colKeys,
    matrix,
    rowTotals,
    colTotals,
    grandTotal,
    dataFields,
  };
}

export function renderPivotToWorksheet(
  ws: WorkSheet,
  pivot: PivotTableDefinition,
  comp: ReturnType<typeof computePivotData>
): Range {
  const { rowKeys, colKeys, matrix, rowTotals, colTotals, grandTotal, dataFields } = comp;
  const startRow = pivot.targetCell.r;
  const startCol = pivot.targetCell.c;

  const headerStyle = {
    font: { bold: true, color: { rgb: '1F497D' }, sz: 11, name: 'Calibri' },
    fill: { fgColor: { rgb: 'DCE6F1' }, patternType: 'solid' as const },
    border: {
      bottom: { style: 'thin' as const, color: { rgb: '95B3D7' } },
      top: { style: 'thin' as const, color: { rgb: '95B3D7' } },
      left: { style: 'thin' as const, color: { rgb: '95B3D7' } },
      right: { style: 'thin' as const, color: { rgb: '95B3D7' } },
    },
    alignment: { horizontal: 'center' as const, vertical: 'center' as const },
  };

  const titleStyle = {
    font: { bold: true, color: { rgb: '1F497D' }, sz: 12, name: 'Calibri' },
    alignment: { horizontal: 'left' as const, vertical: 'center' as const },
  };

  const totalStyle = {
    font: { bold: true, sz: 11, name: 'Calibri' },
    fill: { fgColor: { rgb: 'F2F2F2' }, patternType: 'solid' as const },
    border: {
      top: { style: 'thin' as const, color: { rgb: '95B3D7' } },
      bottom: { style: 'double' as const, color: { rgb: '95B3D7' } },
    },
  };

  const dataStyle = {
    font: { sz: 11, name: 'Calibri' },
    border: {
      bottom: { style: 'hair' as const, color: { rgb: 'D9D9D9' } },
      right: { style: 'hair' as const, color: { rgb: 'D9D9D9' } },
    },
    alignment: { horizontal: 'right' as const },
  };

  const rowLabelStyle = {
    font: { bold: true, sz: 11, name: 'Calibri' },
    border: {
      bottom: { style: 'hair' as const, color: { rgb: 'D9D9D9' } },
      right: { style: 'thin' as const, color: { rgb: '95B3D7' } },
    },
    alignment: { horizontal: 'left' as const },
  };

  let currRow = startRow;

  // Title
  const titleCell = encode_cell({ c: startCol, r: currRow });
  ws[titleCell] = { v: pivot.name || 'Pivot Table', t: 's', s: titleStyle };
  currRow += 2;

  // Column headers
  const colDimCount = pivot.cols.length;
  const rowDimCount = Math.max(pivot.rows.length, 1);

  // Row header label
  for (let rIdx = 0; rIdx < rowDimCount; rIdx++) {
    const colName = pivot.rows[rIdx] || 'Row Labels';
    const cRef = encode_cell({ c: startCol + rIdx, r: currRow });
    ws[cRef] = { v: colName, t: 's', s: headerStyle };
  }

  // Column header values
  let valColOffset = startCol + rowDimCount;
  for (const cKey of colKeys) {
    for (const df of dataFields) {
      const cRef = encode_cell({ c: valColOffset, r: currRow });
      const label = cKey.length > 0 ? `${cKey.join(' - ')} (${df.name})` : df.name!;
      ws[cRef] = { v: label, t: 's', s: headerStyle };
      valColOffset++;
    }
  }

  if (pivot.showRowTotals !== false) {
    for (const df of dataFields) {
      const cRef = encode_cell({ c: valColOffset, r: currRow });
      ws[cRef] = { v: `Total ${df.name}`, t: 's', s: headerStyle };
      valColOffset++;
    }
  }

  const maxCol = valColOffset - 1;
  currRow++;

  // Data rows
  for (const rKey of rowKeys) {
    const rKeyStr = rKey.join('\u0000');

    // Row headers
    for (let rIdx = 0; rIdx < rowDimCount; rIdx++) {
      const cRef = encode_cell({ c: startCol + rIdx, r: currRow });
      ws[cRef] = { v: rKey[rIdx] || '', t: 's', s: rowLabelStyle };
    }

    let cCol = startCol + rowDimCount;
    for (const cKey of colKeys) {
      const cKeyStr = cKey.join('\u0000');
      for (const df of dataFields) {
        const val = matrix[rKeyStr]?.[cKeyStr]?.[df.name!] ?? 0;
        const cRef = encode_cell({ c: cCol, r: currRow });
        ws[cRef] = { v: val, t: 'n', s: dataStyle };
        cCol++;
      }
    }

    // Row total
    if (pivot.showRowTotals !== false) {
      for (const df of dataFields) {
        const val = rowTotals[rKeyStr]?.[df.name!] ?? 0;
        const cRef = encode_cell({ c: cCol, r: currRow });
        ws[cRef] = { v: val, t: 'n', s: totalStyle };
        cCol++;
      }
    }

    currRow++;
  }

  // Grand Total row
  if (pivot.showColTotals !== false) {
    const gtLabelCell = encode_cell({ c: startCol, r: currRow });
    ws[gtLabelCell] = { v: 'Grand Total', t: 's', s: totalStyle };

    for (let rIdx = 1; rIdx < rowDimCount; rIdx++) {
      const cRef = encode_cell({ c: startCol + rIdx, r: currRow });
      ws[cRef] = { v: '', t: 's', s: totalStyle };
    }

    let cCol = startCol + rowDimCount;
    for (const cKey of colKeys) {
      const cKeyStr = cKey.join('\u0000');
      for (const df of dataFields) {
        const val = colTotals[cKeyStr]?.[df.name!] ?? 0;
        const cRef = encode_cell({ c: cCol, r: currRow });
        ws[cRef] = { v: val, t: 'n', s: totalStyle };
        cCol++;
      }
    }

    if (pivot.showRowTotals !== false) {
      for (const df of dataFields) {
        const val = grandTotal[df.name!] ?? 0;
        const cRef = encode_cell({ c: cCol, r: currRow });
        ws[cRef] = { v: val, t: 'n', s: totalStyle };
        cCol++;
      }
    }

    currRow++;
  }

  // Update sheet range !ref
  const renderedRange: Range = {
    s: { c: startCol, r: startRow },
    e: { c: maxCol, r: currRow - 1 },
  };

  const existingRef = ws['!ref'] ? decode_range(ws['!ref']) : undefined;
  if (!existingRef) {
    ws['!ref'] = `${encode_cell(renderedRange.s)}:${encode_cell(renderedRange.e)}`;
  } else {
    ws['!ref'] = `${encode_cell({
      c: Math.min(existingRef.s.c, renderedRange.s.c),
      r: Math.min(existingRef.s.r, renderedRange.s.r),
    })}:${encode_cell({
      c: Math.max(existingRef.e.c, renderedRange.e.c),
      r: Math.max(existingRef.e.r, renderedRange.e.r),
    })}`;
  }

  return renderedRange;
}
