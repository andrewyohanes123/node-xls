/**
 * Regression tests for defects found in the initial audit (see CLAUDE.md).
 * Each test pins the correct behaviour; all written files go through the
 * independent validator (schema + SheetJS CE + ExcelJS).
 */
import { describe, it, expect } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import * as SheetJS from 'xlsx';
import XLSX, { utils } from '../src/index.js';
import { validateXlsx } from './helpers/validate.js';

function bookWith(ws: any, name = 'S') {
  const wb = utils.book_new();
  utils.book_append_sheet(wb, ws, name);
  return wb;
}
const writeArr = (wb: any): Uint8Array => XLSX.write(wb, { type: 'array' });

describe('audit regressions: reader fidelity', () => {
  it('preserves string values exactly (no numeric/boolean coercion, no trimming)', async () => {
    const values = ['007', '  padded  ', '1e3', 'true', '0x1F', '12345678901234567890', ' ', 'a\nb'];
    const buf = writeArr(bookWith(utils.aoa_to_sheet([values])));
    await validateXlsx(buf);
    const ws = XLSX.read(buf).Sheets.S;
    expect(values.map((_, i) => ws[utils.encode_cell({ c: i, r: 0 })].v)).toEqual(values);
    // cross-check: SheetJS CE agrees
    const ce = SheetJS.read(buf).Sheets.S;
    expect(values.map((_, i) => ce[utils.encode_cell({ c: i, r: 0 })].v)).toEqual(values);
  });

  it('reads packages whose relationship targets are absolute (openpyxl style)', () => {
    const files = unzipSync(writeArr(bookWith(utils.aoa_to_sheet([['hi']]))));
    files['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(files['xl/_rels/workbook.xml.rels']).replace(/Target="(?!\/)/g, 'Target="/xl/')
    );
    const wb = XLSX.read(zipSync(files));
    expect(wb.SheetNames).toEqual(['S']);
    expect(wb.Sheets.S.A1.v).toBe('hi');
  });
});

describe('audit regressions: writer produces valid packages', () => {
  it('encodes XML-illegal control characters', async () => {
    const buf = writeArr(bookWith(utils.aoa_to_sheet([['bad\u0001char', 'tab\tok', '_x0041_ literal']])));
    const rep = await validateXlsx(buf);
    expect(rep.sheetjs.Sheets.S.A1.v).toBe('bad\u0001char');
    expect(rep.sheetjs.Sheets.S.C1.v).toBe('_x0041_ literal');
    const ws = XLSX.read(buf).Sheets.S;
    expect(ws.A1.v).toBe('bad\u0001char');
    expect(ws.B1.v).toBe('tab\tok');
    expect(ws.C1.v).toBe('_x0041_ literal');
  });

  it('orders autoFilter before mergeCells', async () => {
    const ws = utils.aoa_to_sheet([
      ['a', 'b'],
      [1, 2],
    ]);
    ws['!merges'] = [utils.decode_range('A1:B1')];
    ws['!autofilter'] = { ref: 'A1:B2' };
    await validateXlsx(writeArr(bookWith(ws)));
  });

  it('never writes NaN/Infinity as numbers', async () => {
    const buf = writeArr(bookWith(utils.aoa_to_sheet([[NaN, Infinity, -Infinity, 1]])));
    const rep = await validateXlsx(buf);
    expect(rep.sheetjs.Sheets.S.A1.t).toBe('e');
    expect(rep.sheetjs.Sheets.S.D1.v).toBe(1);
  });

  it('strips a leading "=" from formulas', async () => {
    const ws = utils.aoa_to_sheet([[1, 2]]);
    ws.C1 = { t: 'n', f: '=A1*2' };
    const rep = await validateXlsx(writeArr(bookWith(ws)));
    expect(rep.sheetjs.Sheets.S.C1.f).toBe('A1*2');
  });

  it('writes theme colors instead of dropping them (no black fills)', async () => {
    const ws = utils.aoa_to_sheet([['x']]);
    ws.A1.s = { font: { color: { theme: 4 } }, fill: { fgColor: { theme: 5, tint: 0.4 } } };
    const rep = await validateXlsx(writeArr(bookWith(ws)));
    const styles = rep.xml('xl/styles.xml');
    expect(styles).toContain('<color theme="4"/>');
    expect(styles).toMatch(/<fgColor theme="5" tint="0.4"\/>/);
    const cell = rep.exceljs.getWorksheet('S')!.getCell('A1');
    expect((cell.fill as any).fgColor).toMatchObject({ theme: 5 });
  });

  it('writes styled empty cells as blank cells and extends !ref', async () => {
    const ws = utils.aoa_to_sheet([['a']]);
    utils.set_cell_style(ws, 'C3', { fill: { fgColor: { rgb: 'FF0000' } } });
    expect(ws['!ref']).toBe('A1:C3');
    const rep = await validateXlsx(writeArr(bookWith(ws)));
    expect(rep.xml('xl/worksheets/sheet1.xml')).toMatch(/<c r="C3" s="\d+"\/>/);
  });

  it('handles sparse !cols entries', async () => {
    const ws = utils.aoa_to_sheet([['a', 'b', 'c']]);
    ws['!cols'] = [undefined as any, { wch: 20 }];
    await validateXlsx(writeArr(bookWith(ws)));
  });

  it('writes a pivot table with two data fields and a filter that passes schema validation', async () => {
    const wb = utils.book_new();
    utils.book_append_sheet(
      wb,
      utils.json_to_sheet([
        { R: 'E', C: 'x', D: 'd1', S: 1, Q: 2 },
        { R: 'W', C: 'y', D: 'd2', S: 3, Q: 4 },
      ]),
      'Data'
    );
    utils.add_pivot_table(wb, {
      sourceSheet: 'Data',
      targetSheet: 'P',
      rows: ['R'],
      cols: ['C'],
      filters: ['D'],
      values: [{ field: 'S' }, { field: 'Q' }],
    });
    const rep = await validateXlsx(writeArr(wb));
    expect(rep.xml('xl/pivotTables/pivotTable1.xml')).toContain('<field x="-2"/>');
  });
});

describe('audit regressions: dates and number formats', () => {
  it('converts Dates to whole-day serials using local time', () => {
    const ws = utils.aoa_to_sheet([[new Date(2026, 0, 15), new Date(2026, 0, 15, 12, 0, 0)]]);
    expect(ws.A1.v).toBe(46037);
    expect(ws.B1.v).toBe(46037.5);
  });

  it('writes Date cells with a date number format', async () => {
    const ws = utils.aoa_to_sheet([[new Date(2026, 0, 15)]]);
    const rep = await validateXlsx(writeArr(bookWith(ws)));
    const v = rep.exceljs.getWorksheet('S')!.getCell('A1').value;
    expect(v).toBeInstanceOf(Date);
    expect((v as Date).toISOString().slice(0, 10)).toBe('2026-01-15');
  });

  it('honours cell.z (SheetJS number format field)', async () => {
    const ws = utils.aoa_to_sheet([[0.25]]);
    ws.A1.z = '0.00%';
    const rep = await validateXlsx(writeArr(bookWith(ws)));
    expect(rep.exceljs.getWorksheet('S')!.getCell('A1').numFmt).toBe('0.00%');
    const back = XLSX.read(writeArr(bookWith(ws))).Sheets.S.A1;
    expect(back.z).toBe('0.00%');
    expect(back.w).toBe('25.00%');
  });
});

describe('audit regressions: utils', () => {
  it('set_cell_style deep-merges nested style objects', () => {
    const ws = utils.aoa_to_sheet([['a']]);
    utils.set_cell_style(ws, 'A1', { font: { bold: true } });
    utils.set_cell_style(ws, 'A1', { font: { color: { rgb: 'FF0000' } } });
    expect(ws.A1.s.font).toEqual({ bold: true, color: { rgb: 'FF0000' } });
  });

  it('sheet_to_html keeps 0 and false', () => {
    const html = utils.sheet_to_html(utils.aoa_to_sheet([[0, false]]));
    expect(html).toMatch(/>0</);
    expect(html).toMatch(/>FALSE</);
  });

  it('sheet_to_html escapes style values', () => {
    const ws = utils.aoa_to_sheet([['x']]);
    ws.A1.s = { font: { name: '"><script>alert(1)</script>' } };
    expect(utils.sheet_to_html(ws)).not.toContain('<script>');
  });

  it('rejects invalid sheet names like Excel does', () => {
    const wb = utils.book_new();
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'x'.repeat(32))).toThrow();
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'a/b')).toThrow();
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), "'quoted'")).toThrow();
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'History')).toThrow(); // reserved by Excel
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'Data')).not.toThrow();
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'data')).toThrow(); // names are case-insensitive
  });

  it('sheet_to_json skips blank rows in object mode like SheetJS', () => {
    const aoa = [['h'], [1], [], [2]];
    expect(utils.sheet_to_json(utils.aoa_to_sheet(aoa))).toEqual(SheetJS.utils.sheet_to_json(SheetJS.utils.aoa_to_sheet(aoa)));
  });

  it('sheet_to_json de-duplicates header names like SheetJS', () => {
    const aoa = [
      ['a', 'a', ''],
      [1, 2, 3],
    ];
    expect(utils.sheet_to_json(utils.aoa_to_sheet(aoa))).toEqual(SheetJS.utils.sheet_to_json(SheetJS.utils.aoa_to_sheet(aoa)));
  });

  it('sheet_add_aoa supports origin -1 (append)', () => {
    const ws = utils.aoa_to_sheet([['a'], ['b']]);
    utils.sheet_add_aoa(ws, [['c']], { origin: -1 });
    expect(ws.A3.v).toBe('c');
    expect(ws['!ref']).toBe('A1:A3');
  });

  it('strips a UTF-8 BOM when reading CSV', () => {
    const wb = XLSX.read('﻿Name,Age\nAl,3', { type: 'string' });
    expect(wb.Sheets.Sheet1.A1.v).toBe('Name');
  });
});
