/**
 * Standalone pivot summary (utils.compute_pivot_table): aggregates an array
 * of records by row/column keys without touching a workbook.
 */
import { PivotDataField, PivotFilter, Range, WorkSheet } from '../types.js';
import { readSource, PivotValue } from './source.js';
import { aggregate, defaultDataFieldName } from './aggregate.js';
import { compareValues, itemKey } from './layout.js';

export interface PivotSummaryOptions {
  rows?: string[];
  cols?: string[];
  values?: (string | PivotDataField)[];
  filters?: Array<string | PivotFilter>;
}

export interface PivotSummary {
  rowKeys: string[][];
  colKeys: string[][];
  /** matrix[rowKey][colKey][dataFieldName]; keys join levels with "\u0000". */
  matrix: Record<string, Record<string, Record<string, number>>>;
  rowTotals: Record<string, Record<string, number>>;
  colTotals: Record<string, Record<string, number>>;
  grandTotal: Record<string, number>;
  dataFields: PivotDataField[];
}

const SEP = '\u0000';

function normalise(v: unknown): PivotValue {
  if (v === undefined || v === null || v === '') return null;
  if (v instanceof Date) return v.getTime();
  if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string') return v;
  return String(v);
}

export function computePivotData(dataRows: Array<Record<string, any>>, options: PivotSummaryOptions): PivotSummary {
  const rows = options.rows ?? [];
  const cols = options.cols ?? [];
  const dataFields: PivotDataField[] = (options.values ?? []).map((v) => {
    if (typeof v === 'string') return { field: v, subtotal: 'sum', name: defaultDataFieldName('sum', v) };
    const subtotal = v.subtotal ?? 'sum';
    return { ...v, subtotal, name: v.name ?? defaultDataFieldName(subtotal, v.field) };
  });
  if (dataFields.length === 0 && dataRows.length > 0) {
    const first = rows[0] ?? Object.keys(dataRows[0])[0];
    dataFields.push({ field: first, subtotal: 'count', name: defaultDataFieldName('count', first) });
  }

  const filters = (options.filters ?? []).map((f) => (typeof f === 'string' ? { field: f } : f));
  const records = dataRows.filter((r) =>
    filters.every((f) => f.selected === undefined || itemKey(normalise(r[f.field])) === itemKey(normalise(f.selected)))
  );

  const keyOf = (r: Record<string, any>, fields: string[]) => fields.map((f) => normalise(r[f]));
  const label = (v: PivotValue) => (v === null ? '' : String(v));
  const uniq = (fields: string[]) => {
    const map = new Map<string, PivotValue[]>();
    for (const r of records) {
      const k = keyOf(r, fields);
      map.set(k.map(itemKey).join(SEP), k);
    }
    if (fields.length === 0) map.set('', []);
    return [...map.values()].sort((a, b) => {
      for (let i = 0; i < a.length; i++) {
        const d = compareValues(a[i], b[i]);
        if (d) return d;
      }
      return 0;
    });
  };
  const rowKeys = uniq(rows);
  const colKeys = uniq(cols);

  const groups = new Map<string, Array<Record<string, any>>>();
  const push = (key: string, r: Record<string, any>) => {
    let g = groups.get(key);
    if (!g) groups.set(key, (g = []));
    g.push(r);
  };
  const rk = (r: Record<string, any>) => keyOf(r, rows).map(label).join(SEP);
  const ck = (r: Record<string, any>) => keyOf(r, cols).map(label).join(SEP);
  for (const r of records) {
    push(`${rk(r)}\u0001${ck(r)}`, r);
    push(`${rk(r)}\u0001*`, r);
    push(`*\u0001${ck(r)}`, r);
  }

  const agg = (list: Array<Record<string, any>> | undefined, df: PivotDataField) =>
    aggregate(
      (list ?? []).map((r) => normalise(r[df.field])),
      df.subtotal
    ) ?? 0;

  const matrix: PivotSummary['matrix'] = {};
  const rowTotals: PivotSummary['rowTotals'] = {};
  const colTotals: PivotSummary['colTotals'] = {};
  const grandTotal: PivotSummary['grandTotal'] = {};
  const rowLabels = rowKeys.map((k) => k.map(label));
  const colLabels = colKeys.map((k) => k.map(label));
  for (const r of rowLabels) {
    const rs = r.join(SEP);
    matrix[rs] = {};
    rowTotals[rs] = {};
    for (const c of colLabels) {
      const cs = c.join(SEP);
      matrix[rs][cs] = {};
      for (const df of dataFields) matrix[rs][cs][df.name!] = agg(groups.get(`${rs}\u0001${cs}`), df);
    }
    for (const df of dataFields) rowTotals[rs][df.name!] = agg(groups.get(`${rs}\u0001*`), df);
  }
  for (const c of colLabels) {
    const cs = c.join(SEP);
    colTotals[cs] = {};
    for (const df of dataFields) colTotals[cs][df.name!] = agg(groups.get(`*\u0001${cs}`), df);
  }
  for (const df of dataFields) grandTotal[df.name!] = agg(records, df);

  return { rowKeys: rowLabels, colKeys: colLabels, matrix, rowTotals, colTotals, grandTotal, dataFields };
}

/** Source range as header names + row objects (compatibility helper). */
export function extractSheetData(sheet: WorkSheet, range: Range): { headers: string[]; rows: Array<Record<string, any>> } {
  const src = readSource(sheet, range);
  return {
    headers: src.fields,
    rows: src.records
      .filter((rec) => rec.some((v) => v !== null))
      .map((rec) => Object.fromEntries(src.fields.map((f, i) => [f, rec[i] ?? undefined]))),
  };
}
