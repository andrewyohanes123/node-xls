/**
 * Pivot layout engine (Excel tabular form, values on columns).
 *
 * Produces the item order, row/column axis entries (items, subtotals, grand
 * totals), the aggregated value grid and the sheet geometry. The renderer and
 * the OpenXML writer both consume this structure, so the pre-rendered cells,
 * <rowItems>/<colItems> and <location> always agree.
 */
import { PivotSubtotal, PivotTableDefinition, Range } from '../types.js';
import { SourceData, PivotValue, PivotError, isPivotError } from './source.js';
import { aggregate, defaultDataFieldName } from './aggregate.js';
import { formatValue } from '../utils/format.js';

export interface AxisEntry {
  type: 'item' | 'subtotal' | 'grand';
  /** Display item indexes per field level; for the column axis with several data fields the last element is the data index. */
  path: number[];
  /** Data field index carried by this entry (column axis). */
  data: number;
}

export interface LayoutDataField {
  name: string;
  field: number;
  subtotal: PivotSubtotal;
  numFmt?: string;
}

export interface PivotGeometry {
  title?: { r: number; c: number };
  pages: Array<{ r: number; c: number }>;
  location: Range;
  headerRows: number;
  firstHeaderRow: number;
  firstDataRow: number;
  firstDataCol: number;
}

export interface PivotLayout {
  source: SourceData;
  rowFields: number[];
  colFields: number[];
  pageFields: Array<{ field: number; selected?: number }>;
  dataFields: LayoutDataField[];
  /** The "Values" pseudo-field is the last column field (2+ data fields). */
  valuesOnCols: boolean;
  /** Per axis/page field: all unique values in display (sorted) order. */
  items: Map<number, PivotValue[]>;
  /** Per axis/page field: all unique values in source order (cache shared items). */
  shared: Map<number, PivotValue[]>;
  /** item key -> display index / shared index, per field */
  displayIndex: Map<number, Map<string, number>>;
  sharedIndex: Map<number, Map<string, number>>;
  rowAxis: AxisEntry[];
  colAxis: AxisEntry[];
  /** values[rowEntry][colEntry]; null renders as an empty cell. */
  values: Array<Array<number | null | PivotError>>;
  date1904: boolean;
  showSubtotals: boolean;
  geometry: PivotGeometry;
}

/** Item identity: strings compare case-insensitively, as in Excel. */
export function itemKey(v: PivotValue): string {
  if (v === null) return 'm';
  if (isPivotError(v)) return `e:${v.error}`;
  if (typeof v === 'string') return `s:${v.toLowerCase()}`;
  return `${typeof v === 'number' ? 'n' : 'b'}:${v}`;
}

const typeRank = (v: PivotValue) =>
  typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : isPivotError(v) ? 3 : 4;

/** Excel's ascending item order: numbers, text, booleans, errors, blanks. */
export function compareValues(a: PivotValue, b: PivotValue): number {
  const d = typeRank(a) - typeRank(b);
  if (d) return d;
  if (typeof a === 'number') return a - (b as number);
  if (typeof a === 'string') return a.localeCompare(b as string, undefined, { sensitivity: 'base', numeric: true });
  if (typeof a === 'boolean') return Number(a) - Number(b);
  if (isPivotError(a) && isPivotError(b)) return a.error.localeCompare(b.error);
  return 0;
}

/** Display label of an item (as Excel captions it). */
export function itemLabel(source: SourceData, field: number, v: PivotValue, date1904 = false): string {
  if (v === null) return '(blank)';
  if (isPivotError(v)) return v.error;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return formatValue(source.itemFormats[field]?.get(v) ?? 'General', v, { date1904 });
  return v;
}

interface Node {
  idx: number;
  children: Map<number, Node>;
}

function axisEntries(root: Node, levels: number, subtotals: boolean, grand: boolean): AxisEntry[] {
  const out: AxisEntry[] = [];
  const walk = (node: Node, path: number[]) => {
    const kids = [...node.children.values()].sort((a, b) => a.idx - b.idx);
    for (const kid of kids) {
      const p = [...path, kid.idx];
      if (p.length === levels) out.push({ type: 'item', path: p, data: 0 });
      else {
        walk(kid, p);
        if (subtotals) out.push({ type: 'subtotal', path: p, data: 0 });
      }
    }
  };
  walk(root, []);
  if (grand) out.push({ type: 'grand', path: [], data: 0 });
  return out;
}

export function buildLayout(def: PivotTableDefinition, source: SourceData, opts: { date1904?: boolean } = {}): PivotLayout {
  const where = `Pivot table "${def.name}"`;
  const fieldIndex = (name: string, role: string) => {
    const i = source.fields.findIndex((f) => f.toLowerCase() === String(name).toLowerCase());
    if (i < 0) throw new Error(`${where}: ${role} field "${name}" not found (fields: ${source.fields.join(', ')})`);
    return i;
  };
  const rowFields = def.rows.map((f) => fieldIndex(f, 'row'));
  const colFields = def.cols.map((f) => fieldIndex(f, 'column'));
  const filters = def.filters.map((f) => (typeof f === 'string' ? { field: f } : f));
  const pageIdx = filters.map((f) => fieldIndex(f.field, 'filter'));
  if (rowFields.length === 0) throw new Error(`${where}: at least one row field is required`);
  const axisFields = [...rowFields, ...colFields, ...pageIdx];
  if (new Set(axisFields).size !== axisFields.length) throw new Error(`${where}: a field can only be used once across rows, columns and filters`);

  const values = def.values.length ? def.values : [{ field: source.fields[rowFields[0]], subtotal: 'count' as const }];
  const names = new Set<string>();
  const dataFields: LayoutDataField[] = values.map((v) => {
    const field = fieldIndex(v.field, 'value');
    const subtotal = v.subtotal ?? 'sum';
    const name = v.name ?? defaultDataFieldName(subtotal, source.fields[field]);
    if (source.fields.some((f) => f.toLowerCase() === name.toLowerCase())) {
      throw new Error(`${where}: data field name "${name}" equals a source field name; Excel rejects this (e.g. use "Total ${name}")`);
    }
    if (names.has(name.toLowerCase())) throw new Error(`${where}: duplicate data field name "${name}"`);
    names.add(name.toLowerCase());
    return { name, field, subtotal, ...(v.numFmt ? { numFmt: v.numFmt } : {}) };
  });

  // items for axis and page fields over the whole cache
  const items = new Map<number, PivotValue[]>();
  const shared = new Map<number, PivotValue[]>();
  const displayIndex = new Map<number, Map<string, number>>();
  const sharedIndex = new Map<number, Map<string, number>>();
  for (const f of axisFields) {
    const seen = new Map<string, PivotValue>();
    for (const rec of source.records) {
      const k = itemKey(rec[f]);
      if (!seen.has(k)) seen.set(k, rec[f]);
    }
    const sharedList = [...seen.values()];
    const sorted = [...sharedList].sort(compareValues);
    shared.set(f, sharedList);
    items.set(f, sorted);
    sharedIndex.set(f, new Map(sharedList.map((v, i) => [itemKey(v), i])));
    displayIndex.set(f, new Map(sorted.map((v, i) => [itemKey(v), i])));
  }

  const pageFields = filters.map((flt, i) => {
    const field = pageIdx[i];
    if (flt.selected === undefined) return { field };
    const selected = displayIndex.get(field)!.get(itemKey(flt.selected as PivotValue));
    if (selected === undefined) throw new Error(`${where}: filter value "${flt.selected === null ? '(blank)' : flt.selected}" not found in field "${flt.field}"`);
    return { field, selected };
  });

  const records = source.records.filter((rec) =>
    pageFields.every((p) => p.selected === undefined || displayIndex.get(p.field)!.get(itemKey(rec[p.field])) === p.selected)
  );

  // trees and value buckets keyed by row/column path prefixes
  const rowRoot: Node = { idx: -1, children: new Map() };
  const colRoot: Node = { idx: -1, children: new Map() };
  const buckets = new Map<string, number[]>();
  records.forEach((rec, ri) => {
    const rp = rowFields.map((f) => displayIndex.get(f)!.get(itemKey(rec[f]))!);
    const cp = colFields.map((f) => displayIndex.get(f)!.get(itemKey(rec[f]))!);
    for (const [root, path] of [
      [rowRoot, rp],
      [colRoot, cp],
    ] as const) {
      let node = root;
      for (const idx of path) {
        let next = node.children.get(idx);
        if (!next) node.children.set(idx, (next = { idx, children: new Map() }));
        node = next;
      }
    }
    for (let i = 0; i <= rp.length; i++) {
      for (let j = 0; j <= cp.length; j++) {
        const key = `${rp.slice(0, i).join(',')}|${cp.slice(0, j).join(',')}`;
        let list = buckets.get(key);
        if (!list) buckets.set(key, (list = []));
        list.push(ri);
      }
    }
  });

  const showSubtotals = def.showSubtotals !== false;
  const rowAxis = axisEntries(rowRoot, rowFields.length, showSubtotals && rowFields.length > 1, def.showColTotals !== false);
  const valuesOnCols = dataFields.length > 1;
  let colAxis: AxisEntry[];
  if (colFields.length === 0) {
    colAxis = valuesOnCols ? dataFields.map((_, d) => ({ type: 'item', path: [d], data: d })) : [{ type: 'item', path: [], data: 0 }];
  } else {
    const base = axisEntries(colRoot, colFields.length, showSubtotals && colFields.length > 1, def.showRowTotals !== false);
    colAxis = [];
    for (const e of base) {
      dataFields.forEach((_, d) => {
        if (!valuesOnCols) colAxis.push(e);
        else colAxis.push({ type: e.type, path: e.type === 'item' ? [...e.path, d] : e.path, data: d });
      });
    }
  }

  const colPath = (e: AxisEntry) => (e.type === 'item' && valuesOnCols ? e.path.slice(0, -1) : e.path);
  const grid = rowAxis.map((re) =>
    colAxis.map((ce) => {
      const list = buckets.get(`${re.path.join(',')}|${colPath(ce).join(',')}`);
      if (!list) return null;
      const df = dataFields[ce.data];
      return aggregate(
        list.map((i) => records[i][df.field]),
        df.subtotal
      );
    })
  );

  // geometry
  const start = def.targetCell;
  let top = start.r;
  const title = def.title ? { r: top, c: start.c } : undefined;
  if (title) top += 2;
  const pages = pageFields.map((_, i) => ({ r: top + i, c: start.c }));
  if (pages.length) top += pages.length + 1;
  const headerRows = 1 + colFields.length + (valuesOnCols ? 1 : 0);
  const width = rowFields.length + colAxis.length;
  const height = headerRows + rowAxis.length;
  const geometry: PivotGeometry = {
    title,
    pages,
    location: { s: { r: top, c: start.c }, e: { r: top + height - 1, c: start.c + width - 1 } },
    headerRows,
    firstHeaderRow: 1,
    firstDataRow: headerRows,
    firstDataCol: rowFields.length,
  };

  return {
    source,
    rowFields,
    colFields,
    pageFields,
    dataFields,
    valuesOnCols,
    items,
    shared,
    displayIndex,
    sharedIndex,
    rowAxis,
    colAxis,
    values: grid,
    showSubtotals,
    geometry,
    date1904: Boolean(opts.date1904),
  };
}
