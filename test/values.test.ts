/**
 * Cell values and formulas: written by node-xls, verified through SheetJS CE,
 * ExcelJS and node-xls' own reader.
 */
import { describe, it, expect } from 'vitest';
import XLSX, { utils } from '../src/index.js';
import { roundtripSheet, roundtrip, book } from './helpers/util.js';

describe('strings', () => {
  const values = [
    'plain',
    '',
    ' leading',
    'trailing ',
    'line1\nline2\r\nline3',
    'tab\there',
    'Ünïcödé — 日本語 — العربية — emoji 🎉👍🏽',
    '<xml> & "quotes" \'apos\'',
    '_x0041_ literal escape',
    'ctrl \u0001\u0002\u001F chars',
    '00123',
    '1e5',
    'TRUE',
    '=not a formula',
    'x'.repeat(32767),
  ];

  it('round-trip exactly through all readers', async () => {
    const { rep, ws } = await roundtripSheet(utils.aoa_to_sheet(values.map((v) => [v])));
    const ex = rep.exceljs.getWorksheet('S')!;
    values.forEach((v, i) => {
      const ref = `A${i + 1}`;
      expect(ws[ref]?.v ?? '', ref).toBe(v);
      if (v !== '') {
        // SheetJS normalises CRLF to LF when reading; the file itself keeps the CR
        expect(rep.sheetjs.Sheets.S[ref].v, `SheetJS ${ref}`).toBe(v.replace(/\r\n/g, '\n'));
        expect(ex.getCell(ref).value, `ExcelJS ${ref}`).toBe(v);
      }
    });
  });

  it('deduplicates shared strings and counts references', async () => {
    const { rep } = await roundtripSheet(utils.aoa_to_sheet([['a', 'b', 'a', 'a']]));
    const sst = rep.xml('xl/sharedStrings.xml');
    expect(sst).toContain('count="4" uniqueCount="2"');
  });

  it('rejects strings longer than Excel allows', () => {
    expect(() => XLSX.write(book(['S', utils.aoa_to_sheet([['x'.repeat(32768)]])]), { type: 'array' })).toThrow(/32767/);
  });
});

describe('numbers and booleans', () => {
  const nums = [0, -0, 1, -1, 0.1 + 0.2, 1 / 3, 123456789012345, 1e21, 1e-300, -1.7976931348623157e308, Number.MAX_SAFE_INTEGER];
  it('round-trip with full precision', async () => {
    const { rep, ws } = await roundtripSheet(utils.aoa_to_sheet([nums, [true, false]]));
    nums.forEach((n, i) => {
      const ref = utils.encode_cell({ r: 0, c: i });
      expect(ws[ref].v).toBe(n === 0 ? 0 : n);
      expect(rep.sheetjs.Sheets.S[ref].v).toBe(n === 0 ? 0 : n);
    });
    expect(ws.A2).toMatchObject({ t: 'b', v: true });
    expect(ws.B2).toMatchObject({ t: 'b', v: false });
    expect(rep.exceljs.getWorksheet('S')!.getCell('A2').value).toBe(true);
  });

  it('numeric strings in t:"n" cells are written as numbers', async () => {
    const { ws } = await roundtripSheet({ '!ref': 'A1:B1', A1: { t: 'n', v: '42' }, B1: { t: 'n', v: 'abc' } });
    expect(ws.A1).toMatchObject({ t: 'n', v: 42 });
    expect(ws.B1).toMatchObject({ t: 's', v: 'abc' });
  });

  it('non-finite numbers become errors', async () => {
    const { ws, rep } = await roundtripSheet({ '!ref': 'A1:C1', A1: { t: 'n', v: NaN }, B1: { t: 'n', v: Infinity }, C1: { v: -Infinity } });
    for (const ref of ['A1', 'B1', 'C1']) {
      expect(ws[ref].t).toBe('e');
      expect(rep.sheetjs.Sheets.S[ref].t).toBe('e');
    }
  });

  it('cell.w is the formatted value on read', async () => {
    const ws = utils.aoa_to_sheet([[1234.5, 0.25, 3]]);
    ws.A1.z = '#,##0.00';
    ws.B1.z = '0%';
    const back = (await roundtripSheet(ws)).ws;
    expect([back.A1.w, back.B1.w, back.C1.w]).toEqual(['1,234.50', '25%', '3']);
  });
});

describe('errors', () => {
  it('all Excel error values round-trip (SheetJS codes)', async () => {
    const errors = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
    const ws: any = { '!ref': `A1:A${errors.length}` };
    errors.forEach((e, i) => (ws[`A${i + 1}`] = { t: 'e', v: e }));
    const { ws: back, rep } = await roundtripSheet(ws);
    errors.forEach((e, i) => {
      expect(back[`A${i + 1}`].w).toBe(e);
      expect(rep.sheetjs.Sheets.S[`A${i + 1}`].w).toBe(e);
      expect(back[`A${i + 1}`].v).toBe(rep.sheetjs.Sheets.S[`A${i + 1}`].v);
    });
  });
});

describe('dates', () => {
  it('Dates become serial numbers with a date format; cellDates returns Dates', async () => {
    const d1 = new Date(2026, 0, 15);
    const d2 = new Date(2026, 6, 4, 18, 30, 15);
    const ws = utils.aoa_to_sheet([[d1, d2]]);
    ws.B1.z = 'yyyy-mm-dd hh:mm:ss';
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back.A1.t).toBe('n');
    expect(back.A1.v).toBe(46037);
    expect(back.B1.w).toBe('2026-07-04 18:30:15');
    const ex = rep.exceljs.getWorksheet('S')!;
    expect((ex.getCell('A1').value as Date).toISOString()).toBe('2026-01-15T00:00:00.000Z');
    const withDates = XLSX.read(XLSX.write(book(['S', ws]), { type: 'array' }), { cellDates: true }).Sheets.S;
    expect(withDates.A1.t).toBe('d');
    expect(withDates.A1.v).toEqual(d1);
    expect(withDates.B1.v).toEqual(d2);
  });

  it('t:"d" cells (JS Date or ISO string) are written as dates', async () => {
    const { ws } = await roundtripSheet({ '!ref': 'A1:B1', A1: { t: 'd', v: new Date(2025, 11, 31) }, B1: { t: 'd', v: '2024-02-29' } }, { cellDates: true });
    expect(ws.A1.v).toEqual(new Date(2025, 11, 31));
    expect((ws.B1.v as Date).getDate()).toBe(29);
  });

  it('date1904 workbooks', async () => {
    const wb = book(['S', utils.aoa_to_sheet([[new Date(2026, 0, 15)]], { cellDates: true })]);
    wb.Workbook = { WBProps: { date1904: true } };
    const { rep, back } = await roundtrip(wb, undefined, { cellDates: true });
    expect(rep.xml('xl/workbook.xml')).toContain('date1904="1"');
    expect(back.Workbook?.WBProps?.date1904).toBe(true);
    expect(back.Sheets.S.A1.v).toEqual(new Date(2026, 0, 15));
    expect(rep.sheetjs.Sheets.S.A1.v).toBe(46037 - 1462);
  });
});

describe('formulas', () => {
  it('writes formulas with/without cached values and enables full recalculation', async () => {
    const ws = utils.aoa_to_sheet([[2, 3]]);
    ws.C1 = { t: 'n', f: 'A1*B1', v: 6 };
    ws.D1 = { t: 'n', f: '=SUM(A1:C1)' };
    ws.E1 = { t: 's', f: 'A1&"x"', v: '2x' };
    ws.F1 = { t: 'b', f: 'A1>B1', v: false };
    ws['!ref'] = 'A1:F1';
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back.C1).toMatchObject({ f: 'A1*B1', v: 6 });
    expect(back.D1.f).toBe('SUM(A1:C1)');
    expect(back.E1).toMatchObject({ t: 's', f: 'A1&"x"', v: '2x' });
    expect(back.F1).toMatchObject({ t: 'b', f: 'A1>B1', v: false });
    expect(rep.xml('xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
    expect(rep.exceljs.getWorksheet('S')!.getCell('C1').value).toEqual({ formula: 'A1*B1', result: 6 });
  });

  it('array formulas', async () => {
    const ws = utils.aoa_to_sheet([[1], [2]]);
    utils.sheet_set_array_formula(ws, 'B1:B2', 'A1:A2*10');
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back.B1).toMatchObject({ f: 'A1:A2*10', F: 'B1:B2' });
    expect(rep.sheetjs.Sheets.S.B1.F).toBe('B1:B2');
  });

  it('cellFormula:false drops formulas on read', () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws.B1 = { t: 'n', f: 'A1+1', v: 2 };
    ws['!ref'] = 'A1:B1';
    const back = XLSX.read(XLSX.write(book(['S', ws]), { type: 'array' }), { cellFormula: false }).Sheets.S;
    expect(back.B1).toMatchObject({ t: 'n', v: 2 });
    expect(back.B1.f).toBeUndefined();
  });
});

describe('sparse, empty and large sheets', () => {
  it('empty workbook gets an empty Sheet1', async () => {
    const { back } = await roundtrip(utils.book_new());
    expect(back.SheetNames).toEqual(['Sheet1']);
  });

  it('cells at the grid limits', async () => {
    const ws: any = { '!ref': 'A1:XFD1048576', A1: { t: 's', v: 'first' }, XFD1048576: { t: 's', v: 'last' } };
    const { ws: back } = await roundtripSheet(ws);
    expect(back.XFD1048576.v).toBe('last');
    expect(() => XLSX.write(book(['S', { XFE1: { t: 'n', v: 1 } }]), { type: 'array' })).toThrow(/outside/);
  });

  it('100k cells write and read in reasonable time', () => {
    const aoa = Array.from({ length: 10_000 }, (_, r) => Array.from({ length: 10 }, (_, c) => (c % 2 ? r * c : `r${r}c${c}`)));
    const ws = utils.aoa_to_sheet(aoa);
    utils.set_range_style(ws, 'A1:J1', { font: { bold: true } });
    const t0 = performance.now();
    const buf = XLSX.write(book(['S', ws]), { type: 'array' });
    const t1 = performance.now();
    const back = XLSX.read(buf);
    const t2 = performance.now();
    expect(back.Sheets.S.J10000.v).toBe(9999 * 9);
    expect(t1 - t0).toBeLessThan(5000);
    expect(t2 - t1).toBeLessThan(8000);
  });
});
