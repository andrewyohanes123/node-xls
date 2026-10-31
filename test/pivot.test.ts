/**
 * Pivot tables: summary API, rendered layout (Excel tabular form), OpenXML
 * parts (schema + internal consistency, the checks Excel performs before it
 * refreshes), filters, subtotals, multiple data fields and read-back.
 *
 * Real Excel remains the final judge: see out/excel-check (npm run excel-check).
 */
import { describe, it, expect } from 'vitest';
import XLSX, { utils, WorkBook, PivotTableOptions } from '../src/index.js';
import { roundtrip } from './helpers/util.js';
import type { XlsxReport } from './helpers/validate.js';

const sales = [
  { Region: 'East', Category: 'Furniture', Year: 2025, Sales: 100, Qty: 2 },
  { Region: 'East', Category: 'Technology', Year: 2026, Sales: 250, Qty: 1 },
  { Region: 'West', Category: 'Furniture', Year: 2025, Sales: 150, Qty: 3 },
  { Region: 'West', Category: 'Technology', Year: 2026, Sales: 300, Qty: 2 },
  { Region: 'West', Category: 'Supplies', Year: 2026, Sales: 50, Qty: 5 },
  { Region: 'north', Category: 'Supplies', Year: 2025, Sales: 70, Qty: 4 },
];

function pivotBook(opts: Partial<PivotTableOptions>, data: any[] = sales): WorkBook {
  const wb = utils.book_new();
  utils.book_append_sheet(wb, utils.json_to_sheet(data), 'Data');
  utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'Pivot', targetCell: 'A3', rows: ['Region'], values: [{ field: 'Sales', name: 'Total Sales' }], ...opts });
  return wb;
}

/** Cell texts of a rectangular area (formatted), '' for empty. */
function grid(ws: any, ref: string): string[][] {
  const r = utils.decode_range(ref);
  const out: string[][] = [];
  for (let R = r.s.r; R <= r.e.r; R++) {
    const row: string[] = [];
    for (let C = r.s.c; C <= r.e.c; C++) {
      const cell = ws[utils.encode_cell({ r: R, c: C })];
      row.push(cell && cell.t !== 'z' ? utils.format_cell(cell) : '');
    }
    out.push(row);
  }
  return out;
}

/**
 * Cross-checks the three pivot parts the way Excel does on load: records and
 * shared items agree, field item lists cover the shared items, row/col items
 * index valid field items, and the location size equals header + item rows.
 */
function checkPivotParts(rep: XlsxReport, n = 1) {
  const cache = rep.xml(`xl/pivotCache/pivotCacheDefinition${n}.xml`);
  const records = rep.xml(`xl/pivotCache/pivotCacheRecords${n}.xml`);
  const table = rep.xml(`xl/pivotTables/pivotTable${n}.xml`);
  const fields = [...cache.matchAll(/<cacheField [^>]*>([\s\S]*?)<\/cacheField>/g)].map((m) => {
    const listed = /<sharedItems[^>]*>([\s\S]*)<\/sharedItems>/.exec(m[1]);
    return listed ? (listed[1].match(/<(s|n|b|d|m|e)\b/g) ?? []).length : null;
  });
  const recs = [...records.matchAll(/<r>([\s\S]*?)<\/r>/g)];
  expect(Number(/recordCount="(\d+)"/.exec(cache)![1])).toBe(recs.length);
  for (const r of recs) {
    const values = [...r[1].matchAll(/<(x|s|n|b|d|m|e)\b(?:\s+v="([^"]*)")?/g)];
    expect(values.length).toBe(fields.length);
    values.forEach((v, f) => {
      if (fields[f] === null) expect(v[1]).not.toBe('x');
      else {
        expect(v[1]).toBe('x');
        expect(Number(v[2])).toBeLessThan(fields[f]!);
      }
    });
  }
  const pf = [...table.matchAll(/<pivotField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/pivotField>)/g)];
  expect(pf.length).toBe(fields.length);
  const itemCounts = pf.map((p, f) => {
    const items = [...(p[2] ?? '').matchAll(/<item x="(\d+)"\/>/g)].map((m) => Number(m[1]));
    if (/axis=/.test(p[1])) {
      expect(fields[f]).not.toBeNull();
      expect([...items].sort((a, b) => a - b)).toEqual(Array.from({ length: fields[f]! }, (_, i) => i));
    }
    return items.length;
  });
  const fieldXs = (tag: string) => [...(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`).exec(table)?.[1] ?? '').matchAll(/<field x="(-?\d+)"\/>/g)].map((m) => Number(m[1]));
  const dataCount = Number(/<dataFields count="(\d+)"/.exec(table)![1]);
  const checkItems = (tag: string, axisFields: number[]) => {
    const body = new RegExp(`<${tag} count="(\\d+)">([\\s\\S]*?)</${tag}>`).exec(table);
    if (!body) return 0;
    const is = [...body[2].matchAll(/<i\b([^>]*?)(?:\/>|>([\s\S]*?)<\/i>)/g)];
    expect(is.length).toBe(Number(body[1]));
    for (const i of is) {
      const r = Number(/ r="(\d+)"/.exec(i[1])?.[1] ?? 0);
      const xs = [...(i[2] ?? '').matchAll(/<x(?: v="(\d+)")?\/>/g)].map((m) => Number(m[1] ?? 0));
      if (/t="grand"/.test(i[1])) continue;
      xs.forEach((x, j) => {
        const f = axisFields[r + j];
        const limit = f === -2 ? dataCount : itemCounts[f];
        expect(x).toBeLessThan(limit);
      });
    }
    return is.length;
  };
  const rowCount = checkItems('rowItems', fieldXs('rowFields'));
  const colCount = checkItems('colItems', fieldXs('colFields'));
  const loc = /<location ref="([^"]+)" firstHeaderRow="(\d+)" firstDataRow="(\d+)" firstDataCol="(\d+)"/.exec(table)!;
  const range = utils.decode_range(loc[1]);
  expect(range.e.r - range.s.r + 1).toBe(Number(loc[3]) + rowCount);
  expect(range.e.c - range.s.c + 1).toBe(Number(loc[4]) + colCount);
  return { cache, records, table };
}

describe('compute_pivot_table (summary API)', () => {
  it('aggregates by row/column keys', () => {
    const comp = utils.compute_pivot_table(sales, { rows: ['Region'], cols: ['Category'], values: [{ field: 'Sales' }, { field: 'Qty' }] });
    expect(comp.rowKeys).toEqual([['East'], ['north'], ['West']]);
    expect(comp.colKeys).toEqual([['Furniture'], ['Supplies'], ['Technology']]);
    expect(comp.matrix.East.Furniture['Sum of Sales']).toBe(100);
    expect(comp.rowTotals.West['Sum of Sales']).toBe(500);
    expect(comp.colTotals.Supplies['Sum of Qty']).toBe(9);
    expect(comp.grandTotal['Sum of Sales']).toBe(920);
  });

  it('every summary function', () => {
    const fns = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const;
    const comp = utils.compute_pivot_table(sales, { rows: ['Region'], values: fns.map((subtotal) => ({ field: 'Sales', subtotal, name: subtotal })) });
    const west = comp.matrix.West[''];
    expect(west).toMatchObject({ sum: 500, count: 3, average: 500 / 3, max: 300, min: 50, product: 150 * 300 * 50, countNums: 3 });
    expect(west.var).toBeCloseTo(15833.333, 3);
    expect(west.varp).toBeCloseTo(10555.556, 3);
    expect(west.stdDev).toBeCloseTo(Math.sqrt(15833.333), 3);
  });

  it('numeric keys sort numerically and filters select records', () => {
    const data = [{ n: 10, v: 1 }, { n: 9, v: 2 }, { n: 100, v: 3 }, { n: 9, v: 4 }];
    expect(utils.compute_pivot_table(data, { rows: ['n'], values: ['v'] }).rowKeys).toEqual([['9'], ['10'], ['100']]);
    const filtered = utils.compute_pivot_table(sales, { rows: ['Region'], values: ['Sales'], filters: [{ field: 'Year', selected: 2026 }] });
    expect(filtered.grandTotal['Sum of Sales']).toBe(600);
  });
});

describe('pivot layout and OpenXML', () => {
  it('rows only: tabular layout, location and items', async () => {
    const wb = pivotBook({});
    expect(grid(wb.Sheets.Pivot, 'A3:B7')).toEqual([
      ['Region', 'Total Sales'],
      ['East', '350'],
      ['north', '70'],
      ['West', '500'],
      ['Grand Total', '920'],
    ]);
    const { rep } = await roundtrip(wb);
    const { table, cache } = checkPivotParts(rep);
    expect(table).toContain('<location ref="A3:B7" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>');
    expect(table).toContain('<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems>');
    expect(table).toContain('<colItems count="1"><i/></colItems>');
    expect(table).not.toContain('<colFields');
    // shared items in source order, pivot items in display order
    expect(cache).toContain('<sharedItems count="3"><s v="East"/><s v="West"/><s v="north"/></sharedItems>');
    expect(table).toContain('<items count="4"><item x="0"/><item x="2"/><item x="1"/><item t="default"/></items>');
    // data-only numeric field: flags, no listed items
    expect(cache).toMatch(/<cacheField name="Sales" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="50" maxValue="300"\/>/);
  });

  it('rows x columns with row/column grand totals', async () => {
    const wb = pivotBook({ cols: ['Category'] });
    expect(grid(wb.Sheets.Pivot, 'A3:E8')).toEqual([
      ['Total Sales', 'Category', '', '', ''],
      ['Region', 'Furniture', 'Supplies', 'Technology', 'Grand Total'],
      ['East', '100', '', '250', '350'],
      ['north', '', '70', '', '70'],
      ['West', '150', '50', '300', '500'],
      ['Grand Total', '250', '120', '550', '920'],
    ]);
    const { rep } = await roundtrip(wb);
    const { table } = checkPivotParts(rep);
    expect(table).toContain('<location ref="A3:E8" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>');
    expect(table).toContain('<colItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></colItems>');
  });

  it('two row fields with subtotals (r attributes and t="default")', async () => {
    const wb = pivotBook({ rows: ['Region', 'Category'] });
    expect(grid(wb.Sheets.Pivot, 'A3:C13')).toEqual([
      ['Region', 'Category', 'Total Sales'],
      ['East', 'Furniture', '100'],
      ['', 'Technology', '250'],
      ['East Total', '', '350'],
      ['north', 'Supplies', '70'],
      ['north Total', '', '70'],
      ['West', 'Furniture', '150'],
      ['', 'Supplies', '50'],
      ['', 'Technology', '300'],
      ['West Total', '', '500'],
      ['Grand Total', '', '920'],
    ]);
    const { rep } = await roundtrip(wb);
    const { table } = checkPivotParts(rep);
    expect(table).toContain(
      '<rowItems count="10"><i><x/><x/></i><i r="1"><x v="2"/></i><i t="default"><x/></i><i><x v="1"/><x v="1"/></i><i t="default"><x v="1"/></i><i><x v="2"/><x/></i><i r="1"><x v="1"/></i><i r="1"><x v="2"/></i><i t="default"><x v="2"/></i><i t="grand"><x/></i></rowItems>'
    );
  });

  it('subtotals and grand totals can be turned off', async () => {
    const wb = pivotBook({ rows: ['Region', 'Category'], showSubtotals: false, showColTotals: false });
    expect(grid(wb.Sheets.Pivot, 'A3:C9')).toEqual([
      ['Region', 'Category', 'Total Sales'],
      ['East', 'Furniture', '100'],
      ['', 'Technology', '250'],
      ['north', 'Supplies', '70'],
      ['West', 'Furniture', '150'],
      ['', 'Supplies', '50'],
      ['', 'Technology', '300'],
    ]);
    const { rep } = await roundtrip(wb);
    const { table } = checkPivotParts(rep);
    expect(table).toContain('colGrandTotals="0"');
    expect(table).toContain('defaultSubtotal="0"');
    expect(table).not.toContain('t="default"');
  });

  it('two data fields: the Values field (-2) on columns', async () => {
    const wb = pivotBook({ values: [{ field: 'Sales', name: 'Total Sales' }, { field: 'Qty', subtotal: 'average', name: 'Avg Qty', numFmt: '0.0' }] });
    expect(grid(wb.Sheets.Pivot, 'A3:C8')).toEqual([
      ['', 'Values', ''],
      ['Region', 'Total Sales', 'Avg Qty'],
      ['East', '350', '1.5'],
      ['north', '70', '4.0'],
      ['West', '500', '3.3'],
      ['Grand Total', '920', '2.8'],
    ]);
    const { rep } = await roundtrip(wb);
    const { table } = checkPivotParts(rep);
    expect(table).toContain('<colFields count="1"><field x="-2"/></colFields>');
    expect(table).toContain('<colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>');
    expect(table).toMatch(/<dataField name="Avg Qty" fld="4" subtotal="average" baseField="0" baseItem="0" numFmtId="\d+"\/>/);
  });

  it('columns x two data fields: grand total per data field', async () => {
    const wb = pivotBook({ cols: ['Year'], values: [{ field: 'Sales', name: 'S' }, { field: 'Qty', name: 'Q' }] });
    expect(grid(wb.Sheets.Pivot, 'A3:G9')).toEqual([
      ['', 'Year', 'Values', '', '', '', ''],
      ['', '2025', '', '2026', '', 'Total S', 'Total Q'],
      ['Region', 'S', 'Q', 'S', 'Q', '', ''],
      ['East', '100', '2', '250', '1', '350', '3'],
      ['north', '70', '4', '', '', '70', '4'],
      ['West', '150', '3', '350', '7', '500', '10'],
      ['Grand Total', '320', '9', '600', '8', '920', '17'],
    ]);
    const { rep } = await roundtrip(wb);
    const { table, cache } = checkPivotParts(rep);
    expect(table).toContain('<colItems count="6"><i><x/><x/></i><i r="1" i="1"><x v="1"/></i><i><x v="1"/><x/></i><i r="1" i="1"><x v="1"/></i><i t="grand"><x/></i><i t="grand" i="1"><x/></i></colItems>');
    expect(cache).toContain('<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="2025" maxValue="2026" count="2"><n v="2025"/><n v="2026"/></sharedItems>');
  });

  it('report filters (page fields) with and without a selection', async () => {
    const wb = pivotBook({ targetCell: 'A1', filters: [{ field: 'Year', selected: 2026 }, 'Category'] });
    const ws = wb.Sheets.Pivot;
    expect(grid(ws, 'A1:B2')).toEqual([
      ['Year', '2026'],
      ['Category', '(All)'],
    ]);
    expect(grid(ws, 'A4:B7')).toEqual([
      ['Region', 'Total Sales'],
      ['East', '250'],
      ['West', '350'],
      ['Grand Total', '600'],
    ]);
    const { rep, back } = await roundtrip(wb);
    const { table } = checkPivotParts(rep);
    expect(table).toContain('<location ref="A4:B7" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="2" colPageCount="1"/>');
    expect(table).toContain('<pageFields count="2"><pageField fld="2" item="1" hier="-1"/><pageField fld="1" hier="-1"/></pageFields>');
    expect(back.PivotTables![0]).toMatchObject({ filters: [{ field: 'Year', selected: 2026 }, { field: 'Category' }], targetCell: { r: 0, c: 0 } });
  });

  it('title is placed above the page fields, outside the pivot area', async () => {
    const wb = pivotBook({ targetCell: 'B2', title: 'Sales report', filters: ['Year'] });
    const ws = wb.Sheets.Pivot;
    expect(ws.B2.v).toBe('Sales report');
    expect(ws.B4.v).toBe('Year');
    const { rep } = await roundtrip(wb);
    expect(rep.xml('xl/pivotTables/pivotTable1.xml')).toContain('<location ref="B6:C10"');
  });

  it('dates, blanks, booleans and case-insensitive items', async () => {
    const data = [
      { When: new Date(2026, 0, 2), Who: 'ann', Ok: true, Amt: 1 },
      { When: new Date(2026, 0, 1), Who: 'Ann', Ok: false, Amt: 2 },
      { When: new Date(2026, 0, 2), Who: null, Ok: true, Amt: 4 },
    ];
    const wb = pivotBook({ rows: ['When', 'Who'], cols: ['Ok'], values: [{ field: 'Amt', name: 'Total' }], showSubtotals: false }, data);
    const { rep } = await roundtrip(wb);
    const { cache } = checkPivotParts(rep);
    expect(cache).toMatch(/<cacheField name="When" numFmtId="14"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="2026-01-01T00:00:00" maxDate="2026-01-03T00:00:00" count="2"><d v="2026-01-02T00:00:00"\/><d v="2026-01-01T00:00:00"\/><\/sharedItems>/);
    expect(cache).toContain('<sharedItems containsBlank="1" count="2"><s v="ann"/><m/></sharedItems>');
    // MS-OI29500: booleans count as text for containsString / containsSemiMixedTypes
    expect(cache).toContain('<sharedItems count="2"><b v="1"/><b v="0"/></sharedItems>');
    expect(grid(wb.Sheets.Pivot, 'A5:B7')).toEqual([
      ['1/1/26', 'ann'],
      ['1/2/26', 'ann'],
      ['', '(blank)'],
    ]);
    expect(grid(wb.Sheets.Pivot, 'C4:D4')).toEqual([['FALSE', 'TRUE']]);
  });

  it('number formats from the source are used for item labels; renderCells:false leaves the sheet empty', async () => {
    const wb = pivotBook({ renderCells: false });
    expect(Object.keys(wb.Sheets.Pivot).filter((k) => !k.startsWith('!'))).toEqual([]);
    const { rep } = await roundtrip(wb);
    checkPivotParts(rep);
  });

  it('reads pivot definitions back and can write them again', async () => {
    const wb = pivotBook({ rows: ['Region', 'Category'], cols: ['Year'], values: [{ field: 'Sales', name: 'S' }, { field: 'Qty', name: 'Q', subtotal: 'max' }], filters: ['Category'].slice(0, 0), styleName: 'PivotStyleMedium9', showRowTotals: false });
    const { back } = await roundtrip(wb);
    const p = back.PivotTables![0];
    expect(p).toMatchObject({
      name: 'PT',
      sourceSheet: 'Data',
      targetSheet: 'Pivot',
      rows: ['Region', 'Category'],
      cols: ['Year'],
      values: [
        { field: 'Sales', name: 'S', subtotal: 'sum' },
        { field: 'Qty', name: 'Q', subtotal: 'max' },
      ],
      showRowTotals: false,
      showColTotals: true,
      showSubtotals: true,
      styleName: 'PivotStyleMedium9',
    });
    const again = await roundtrip(back);
    checkPivotParts(again.rep);
  });

  it('several pivots in one workbook get their own caches', async () => {
    const wb = pivotBook({});
    utils.add_pivot_table(wb, { name: 'PT2', sourceSheet: 'Data', targetSheet: 'Pivot', targetCell: 'E3', rows: ['Category'], values: ['Qty'] });
    utils.add_pivot_table(wb, { name: 'PT3', sourceSheet: 'Data', rows: ['Year'], values: ['Sales'] });
    expect(wb.SheetNames).toEqual(['Data', 'Pivot', 'Data Pivot']);
    const { rep } = await roundtrip(wb);
    for (const n of [1, 2, 3]) checkPivotParts(rep, n);
    expect(rep.xml('xl/workbook.xml')).toContain('<pivotCaches><pivotCache cacheId="1" r:id="rId');
  });
});

describe('pivot validation', () => {
  const bad = (opts: Partial<PivotTableOptions>) => () => pivotBook(opts);
  it('rejects unknown fields, reused fields, missing rows, name clashes and overlaps', () => {
    expect(bad({ rows: ['Nope'] })).toThrow(/row field "Nope" not found \(fields: Region, Category, Year, Sales, Qty\)/);
    expect(bad({ rows: ['Region'], cols: ['Region'] })).toThrow(/only be used once/);
    expect(bad({ rows: [] })).toThrow(/at least one row field/);
    expect(bad({ values: [{ field: 'Sales', name: 'Sales' }] })).toThrow(/equals a source field name/);
    expect(bad({ values: [{ field: 'Sales', name: 'X' }, { field: 'Qty', name: 'x' }] })).toThrow(/duplicate data field name/);
    expect(bad({ filters: [{ field: 'Year', selected: 1999 }] })).toThrow(/filter value "1999" not found/);
    const wb = pivotBook({});
    expect(() => utils.add_pivot_table(wb, { name: 'Clash', sourceSheet: 'Data', targetSheet: 'Pivot', targetCell: 'B4', rows: ['Region'], values: ['Qty'] })).toThrow(/overlaps pivot table PT/);
  });

  it('duplicate headers are made unique (as Excel does); blank headers are rejected', async () => {
    const ws = utils.aoa_to_sheet([
      ['K', 'K', 'V '],
      ['a', 1, 2],
    ]);
    const wb = utils.book_new(ws, 'Data');
    utils.add_pivot_table(wb, { sourceSheet: 'Data', rows: ['K'], values: ['K2', 'V '] });
    const { rep } = await roundtrip(wb);
    // cache field names are the exact header text (trailing space kept) so a refresh still finds them
    expect(rep.xml('xl/pivotCache/pivotCacheDefinition1.xml')).toMatch(/name="K".*name="K2".*name="V "/);
    const blank = utils.book_new(utils.aoa_to_sheet([['K', null], ['a', 1]]), 'Data');
    expect(() => utils.add_pivot_table(blank, { sourceSheet: 'Data', rows: ['K'] })).toThrow(/header in column B is blank/);
  });
});
