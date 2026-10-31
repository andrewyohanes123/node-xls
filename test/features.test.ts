/**
 * Pro Basic features: hyperlinks, comments, protection, conditional
 * formatting, data validation and Excel tables. Each feature is validated
 * against the ECMA schema, read back by ExcelJS / SheetJS CE, and round-tripped.
 */
import { describe, it, expect } from 'vitest';
import XLSX, { utils, ConditionalFormatRule, DataValidation } from '../src/index.js';
import { hashPassword } from '../src/features/protection.js';
import { book, roundtrip, roundtripSheet } from './helpers/util.js';
// ExcelJS' own password hashing, used as an independent oracle
// eslint-disable-next-line @typescript-eslint/no-require-imports
const Encryptor = require('exceljs/lib/utils/encryptor.js');

const data = () =>
  utils.aoa_to_sheet([
    ['Name', 'Score', 'Due'],
    ['Al', 5, new Date(2026, 0, 2)],
    ['Bo', 9, new Date(2026, 0, 3)],
    ['Cy', 2, new Date(2026, 0, 4)],
    ['Di', 7, new Date(2026, 0, 5)],
  ]);

describe('hyperlinks', () => {
  it('external URLs, mailto, internal locations and tooltips', async () => {
    const ws = data();
    ws.A2.l = { Target: 'https://example.com/path?q=1&x=2', Tooltip: 'Example' };
    ws.A3.l = { Target: 'mailto:bo@example.com' };
    ws.A4.l = { Target: '#Other!B2' };
    ws.A5.l = { Target: "#'My Sheet'!A1" };
    const wb = book(['S', ws], ['Other', utils.aoa_to_sheet([[1]])], ['My Sheet', utils.aoa_to_sheet([[1]])]);
    const { back, rep } = await roundtrip(wb);
    const s = back.Sheets.S;
    expect(s.A2.l).toEqual({ Target: 'https://example.com/path?q=1&x=2', Tooltip: 'Example' });
    expect(s.A3.l).toEqual({ Target: 'mailto:bo@example.com' });
    expect(s.A4.l).toEqual({ Target: '#Other!B2' });
    expect(s.A5.l).toEqual({ Target: "#'My Sheet'!A1" });
    expect(rep.sheetjs.Sheets.S.A2.l?.Target).toBe('https://example.com/path?q=1&x=2');
    expect(rep.sheetjs.Sheets.S.A4.l?.Target).toBe('#Other!B2');
    expect(rep.sheetjs.Sheets.S.A2.l?.Tooltip).toBe('Example');
    expect(rep.exceljs.getWorksheet('S')!.getCell('A2').value).toMatchObject({ hyperlink: 'https://example.com/path?q=1&x=2' });
    const rels = rep.xml('xl/worksheets/_rels/sheet1.xml.rels');
    expect(rels).toContain('TargetMode="External"');
    expect(rels).not.toContain('Other!B2');
  });

  it('a link on an empty cell creates the cell', async () => {
    const ws = data();
    ws.E1 = { t: 'z', l: { Target: 'https://example.com' } };
    const back = (await roundtripSheet(ws)).ws;
    expect(back.E1.l.Target).toBe('https://example.com');
  });
});

describe('comments', () => {
  it('notes with authors, multiple comments per cell and visibility', async () => {
    const ws = data();
    utils.cell_add_comment(ws.A2, 'first note', 'Ann');
    utils.cell_add_comment(ws.B3, 'line 1\nline 2 <&>', 'Bo');
    ws.B3.c.hidden = false;
    utils.cell_add_comment(ws.C4, 'one', 'X');
    utils.cell_add_comment(ws.C4, 'two', 'Y');
    const { ws: back, rep } = await roundtripSheet(ws);
    expect([...back.A2.c]).toEqual([{ t: 'first note', a: 'Ann' }]);
    expect(back.A2.c.hidden).toBe(true);
    expect([...back.B3.c]).toEqual([{ t: 'line 1\nline 2 <&>', a: 'Bo' }]);
    expect(back.B3.c.hidden).toBe(false);
    expect(back.C4.c[0].t).toBe('X:\none\nY:\ntwo');
    expect(rep.sheetjs.Sheets.S.A2.c?.[0]).toMatchObject({ a: 'Ann', t: 'first note' });
    const note = rep.exceljs.getWorksheet('S')!.getCell('B3').note as any;
    expect(note.texts.map((t: any) => t.text).join('')).toBe('line 1\nline 2 <&>');
    expect(rep.files['xl/drawings/vmlDrawing1.vml']).toBeDefined();
  });

  it('comments on several sheets use separate parts and shape ids', async () => {
    const a = data();
    const b = data();
    utils.cell_add_comment(a.A1, 'a', 'x');
    utils.cell_add_comment(b.A1, 'b', 'x');
    const { rep, back } = await roundtrip(book(['A', a], ['B', b]));
    expect(rep.files['xl/comments2.xml']).toBeDefined();
    expect(rep.xml('xl/drawings/vmlDrawing2.vml')).toContain('data="2"');
    expect(back.Sheets.B.A1.c[0].t).toBe('b');
  });
});

describe('protection', () => {
  it('password hash matches ExcelJS (independent implementation)', () => {
    const salt = Buffer.from('0123456789abcdef').toString('base64');
    const ours = hashPassword('s3cr3t ünïcode', salt, 1000);
    expect(ours.hashValue).toBe(Encryptor.convertPasswordToHash('s3cr3t ünïcode', 'SHA512', salt, 1000));
  });

  it('sheet protection with password and allowed actions', async () => {
    const ws = data();
    ws['!protect'] = { password: 'pw', formatCells: false, sort: false, selectLockedCells: true };
    ws.B2.s = { protection: { locked: false } };
    const { ws: back, rep } = await roundtripSheet(ws);
    const xml = rep.xml('xl/worksheets/sheet1.xml');
    expect(xml).toMatch(/<sheetProtection algorithmName="SHA-512" hashValue="[^"]+" saltValue="[^"]+" spinCount="100000" sheet="1" selectLockedCells="1" formatCells="0" sort="0"\/>/);
    // ExcelJS models these flags as "allowed" (the inverse of the OOXML attribute)
    expect(rep.exceljs.getWorksheet('S')!.sheetProtection).toMatchObject({ sheet: true, formatCells: true, sort: true, algorithmName: 'SHA-512' });
    expect(back['!protect']).toMatchObject({ formatCells: false, sort: false, selectLockedCells: true, algorithmName: 'SHA-512', spinCount: 100000 });
    // re-writing a read file keeps the same password hash
    const again = await roundtripSheet(back);
    expect(again.ws['!protect']!.hashValue).toBe(back['!protect']!.hashValue);
  });

  it('workbook structure protection', async () => {
    const wb = book(['S', data()]);
    wb.Workbook = { Protection: { password: 'wb', lockWindows: true } };
    const { rep, back } = await roundtrip(wb);
    expect(rep.xml('xl/workbook.xml')).toMatch(/<workbookProtection workbookAlgorithmName="SHA-512"[^>]* lockStructure="1" lockWindows="1"\/>/);
    expect(back.Workbook?.Protection).toMatchObject({ lockStructure: true, lockWindows: true, algorithmName: 'SHA-512' });
  });
});

describe('conditional formatting', () => {
  const hi = { font: { bold: true, color: { rgb: '9C0006' } }, fill: { fgColor: { rgb: 'FFC7CE' } } };
  const rules: ConditionalFormatRule[] = [
    { type: 'cellIs', operator: 'greaterThan', formulae: [6], style: hi },
    { type: 'cellIs', operator: 'between', formulae: [2, 4], style: { fill: { fgColor: { theme: 9, tint: 0.6 } } } },
    { type: 'expression', formula: '$B2>MEDIAN($B$2:$B$5)', style: { font: { italic: true } }, stopIfTrue: true },
    { type: 'containsText', text: 'A"l', style: hi },
    { type: 'notContainsText', text: 'x', style: hi },
    { type: 'beginsWith', text: 'B', style: hi },
    { type: 'endsWith', text: 'y', style: hi },
    { type: 'top10', rank: 2, style: hi },
    { type: 'top10', rank: 10, percent: true, bottom: true, style: hi },
    { type: 'aboveAverage', style: hi },
    { type: 'aboveAverage', below: true, equalAverage: true, stdDev: 1, style: hi },
    { type: 'duplicateValues', style: hi },
    { type: 'uniqueValues', style: hi },
    { type: 'containsBlanks', style: hi },
    { type: 'notContainsBlanks', style: hi },
    { type: 'containsErrors', style: hi },
    { type: 'notContainsErrors', style: hi },
    { type: 'timePeriod', timePeriod: 'last7Days', style: { numFmt: 'yyyy-mm-dd', border: { bottom: { style: 'thin' } } } },
    { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], colors: [{ rgb: 'F8696B' }, { rgb: 'FFEB84' }, { rgb: '63BE7B' }] },
    { type: 'dataBar', color: { rgb: '638EC6' }, min: { type: 'num', value: 0 }, max: { type: 'num', value: 10 } },
    { type: 'iconSet', iconSet: '3TrafficLights1', reverse: true },
    { type: 'iconSet', iconSet: '5Arrows', showValue: false },
  ];

  it('every rule type is schema-valid, read by ExcelJS and round-trips', async () => {
    const ws = data();
    ws['!condfmt'] = [
      { ref: 'B2:B5', rules: rules.filter((r) => !['containsText', 'notContainsText', 'beginsWith', 'endsWith', 'timePeriod'].includes(r.type)) },
      { ref: 'A2:A5', rules: rules.filter((r) => ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(r.type)) },
      { ref: 'C2:C5', rules: rules.filter((r) => r.type === 'timePeriod') },
    ];
    const { ws: back, rep } = await roundtripSheet(ws);
    const ex = rep.exceljs.getWorksheet('S')!.conditionalFormattings as any[];
    const exRules = ex.flatMap((c) => c.rules);
    expect(exRules).toHaveLength(rules.length);
    expect(exRules.map((r: any) => r.priority)).toEqual(rules.map((_, i) => i + 1).sort((a, b) => a - b).slice(0, rules.length).map((p) => p));
    expect(exRules[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThan', style: { font: { bold: true }, fill: { bgColor: { argb: 'FFFFC7CE' } } } });

    // round trip: same rules (styles normalised) in the same ranges
    const flatten = (list: any[]) => list.flatMap((c) => c.rules.map((r: any) => r.type));
    expect(flatten(back['!condfmt']!)).toEqual(flatten(ws['!condfmt']));
    const first = back['!condfmt']![0].rules[0] as any;
    expect(first).toEqual({ type: 'cellIs', operator: 'greaterThan', formulae: [6], style: { font: { bold: true, color: { rgb: '9C0006' } }, fill: { patternType: 'solid', fgColor: { rgb: 'FFC7CE' } } } });
    const contains = back['!condfmt']![1].rules[0] as any;
    expect(contains).toMatchObject({ type: 'containsText', text: 'A"l' });
    expect(rep.xml('xl/worksheets/sheet1.xml')).toContain('<formula>NOT(ISERROR(SEARCH(&quot;A&quot;&quot;l&quot;,A2)))</formula>'.replace(/&quot;/g, '"'));
    const scale = back['!condfmt']![0].rules.find((r) => r.type === 'colorScale') as any;
    expect(scale.cfvo).toEqual([{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }]);
    expect(rep.xml('xl/styles.xml')).toMatch(/<dxfs count="\d+">/);
  });

  it('rejects malformed rules', () => {
    const bad = (rule: any) => {
      const ws = data();
      ws['!condfmt'] = [{ ref: 'B2:B5', rules: [rule] }];
      return () => XLSX.write(book(['S', ws]), { type: 'array' });
    };
    expect(bad({ type: 'cellIs', operator: 'between', formulae: [1], style: {} })).toThrow(/2 formulae/);
    expect(bad({ type: 'colorScale', cfvo: [{ type: 'min' }], colors: [{ rgb: 'FF0000' }] })).toThrow(/2 or 3/);
    expect(bad({ type: 'iconSet', iconSet: '4Arrows', cfvo: [{ type: 'min' }] })).toThrow(/4 cfvo/);
  });
});

describe('data validation', () => {
  const validations: DataValidation[] = [
    { ref: 'A2:A5', type: 'list', values: ['Al', 'Bo', 'Cy, Jr.', 'D"i'], prompt: { title: 'Name', message: 'Pick one' } },
    { ref: 'B2:B5', type: 'whole', operator: 'between', formula1: 0, formula2: 10, error: { title: 'Oops', message: '0..10 only', style: 'warning' } },
    { ref: 'C2:C5', type: 'date', operator: 'greaterThan', formula1: new Date(2026, 0, 1) },
    { ref: 'D2:D5', type: 'decimal', operator: 'lessThanOrEqual', formula1: 99.5, allowBlank: false },
    { ref: 'E2:E5', type: 'textLength', operator: 'lessThan', formula1: 10 },
    { ref: 'F2:F5', type: 'custom', formula1: 'ISNUMBER(F2)' },
    { ref: 'G2:G5', type: 'list', formula1: '$A$2:$A$5', hideDropDown: true },
    { ref: 'H2:H5', type: 'time', operator: 'between', formula1: 0.25, formula2: 0.75 },
    { ref: 'I2', type: 'any', prompt: { message: 'Just a hint' } },
  ];

  it('every type is schema-valid, read by ExcelJS and round-trips', async () => {
    const ws = data();
    ws['!dataValidation'] = validations;
    const { ws: back, rep } = await roundtripSheet(ws);
    const ex = rep.exceljs.getWorksheet('S')!.dataValidations.model as any;
    expect(ex.A2).toMatchObject({ type: 'list', formulae: ['"Al,Bo,Cy, Jr.,D""i"'], promptTitle: 'Name', prompt: 'Pick one' });
    expect(ex.B3).toMatchObject({ type: 'whole', operator: 'between', formulae: [0, 10], errorStyle: 'warning', error: '0..10 only' });
    expect(ex.C2).toMatchObject({ type: 'date', operator: 'greaterThan' });
    expect(ex.D2).toMatchObject({ type: 'decimal', operator: 'lessThanOrEqual' });
    expect(rep.xml('xl/worksheets/sheet1.xml')).toMatch(/<dataValidation type="decimal" operator="lessThanOrEqual" showInputMessage="1"/);
    expect(back['!dataValidation']).toEqual([
      { ref: 'A2:A5', type: 'list', values: ['Al', 'Bo', 'Cy', ' Jr.', 'D"i'], prompt: { title: 'Name', message: 'Pick one' } },
      { ref: 'B2:B5', type: 'whole', formula1: '0', formula2: '10', error: { title: 'Oops', message: '0..10 only', style: 'warning' } },
      { ref: 'C2:C5', type: 'date', formula1: '46023', operator: 'greaterThan' },
      { ref: 'D2:D5', type: 'decimal', formula1: '99.5', operator: 'lessThanOrEqual', allowBlank: false },
      { ref: 'E2:E5', type: 'textLength', formula1: '10', operator: 'lessThan' },
      { ref: 'F2:F5', type: 'custom', formula1: 'ISNUMBER(F2)' },
      { ref: 'G2:G5', type: 'list', formula1: '$A$2:$A$5', hideDropDown: true },
      { ref: 'H2:H5', type: 'time', formula1: '0.25', formula2: '0.75' },
      { ref: 'I2', type: 'any', prompt: { message: 'Just a hint' } },
    ]);
  });

  it('rejects lists over 255 characters and missing formulas', () => {
    const tooLong = data();
    tooLong['!dataValidation'] = [{ ref: 'A2', type: 'list', values: Array.from({ length: 100 }, (_, i) => `value${i}`) }];
    expect(() => XLSX.write(book(['S', tooLong]), { type: 'array' })).toThrow(/255/);
    const missing = data();
    missing['!dataValidation'] = [{ ref: 'A2', type: 'whole', operator: 'between', formula1: 1 }];
    expect(() => XLSX.write(book(['S', missing]), { type: 'array' })).toThrow(/formula2/);
  });
});

describe('Excel tables', () => {
  it('table with style, totals row (SUBTOTAL formulas) and header normalisation', async () => {
    const ws = utils.aoa_to_sheet([
      ['Item', 2025, 'Item', null],
      ['a', 1, 'x', 4],
      ['b', 2, 'y', 5],
      [null, null, null, null],
    ]);
    ws['!tables'] = [
      {
        name: 'Sales_2025',
        ref: 'A1:D4',
        style: 'TableStyleLight9',
        totalsRow: true,
        showColumnStripes: true,
        columns: [{ totalsRowLabel: 'Total' }, { totalsRowFunction: 'sum' }, {}, { totalsRowFunction: 'average' }],
      },
    ];
    const { ws: back, rep } = await roundtripSheet(ws);
    const table = rep.xml('xl/tables/table1.xml');
    expect(table).toContain('ref="A1:D4"');
    expect(table).toContain('<autoFilter ref="A1:D3"/>');
    expect(table).toContain('<tableColumn id="2" name="2025" totalsRowFunction="sum"/>');
    expect(table).toContain('<tableColumn id="3" name="Item2"/>');
    expect(table).toContain('<tableColumn id="4" name="Column4" totalsRowFunction="average"/>');
    // header cells are rewritten as the (unique) column names, totals get SUBTOTAL formulas
    expect(back.B1).toMatchObject({ t: 's', v: '2025' });
    expect(back.C1.v).toBe('Item2');
    expect(back.A4.v).toBe('Total');
    expect(back.B4.f).toBe('SUBTOTAL(109,Sales_2025[2025])');
    expect(back.D4.f).toBe('SUBTOTAL(101,Sales_2025[Column4])');
    const exTable = (rep.exceljs.getWorksheet('S') as any).getTable('Sales_2025').table;
    expect(exTable.columns.map((c: any) => c.name)).toEqual(['Item', '2025', 'Item2', 'Column4']);
    expect(back['!tables']).toEqual([
      {
        name: 'Sales_2025',
        ref: 'A1:D4',
        totalsRow: true,
        style: 'TableStyleLight9',
        showColumnStripes: true,
        columns: [
          { name: 'Item', totalsRowLabel: 'Total' },
          { name: '2025', totalsRowFunction: 'sum' },
          { name: 'Item2' },
          { name: 'Column4', totalsRowFunction: 'average' },
        ],
      },
    ]);
  });

  it('multiple tables across sheets get unique ids', async () => {
    const mk = (name: string) => {
      const ws = data();
      ws['!tables'] = [{ name, ref: 'A1:C5' }];
      return ws;
    };
    const { rep } = await roundtrip(book(['A', mk('TblA')], ['B', mk('TblB')]));
    expect(rep.xml('xl/tables/table2.xml')).toMatch(/<table [^>]*id="2" name="TblB"/);
  });

  it('validation: names, duplicates, overlaps', () => {
    const write = (ws: any, ws2?: any) => () => XLSX.write(ws2 ? book(['A', ws], ['B', ws2]) : book(['A', ws]), { type: 'array' });
    const t = (defs: any[], extra: any = {}) => Object.assign(data(), { '!tables': defs }, extra);
    expect(write(t([{ name: 'Bad Name', ref: 'A1:C5' }]))).toThrow(/Invalid table name/);
    expect(write(t([{ name: 'A1', ref: 'A1:C5' }]))).toThrow(/Invalid table name/);
    expect(write(t([{ name: 'T', ref: 'A1:C5' }]), t([{ name: 't', ref: 'A1:C5' }]))).toThrow(/Duplicate table name/);
    expect(write(t([{ name: 'T', ref: 'A1:B5' }, { name: 'U', ref: 'B1:C5' }]))).toThrow(/overlap/);
    expect(write(t([{ name: 'T', ref: 'A1:C5' }], { '!merges': [utils.decode_range('A2:B2')] }))).toThrow(/merged/);
    expect(write(t([{ name: 'T', ref: 'A1:C5' }], { '!autofilter': { ref: 'A1:C5' } }))).toThrow(/autofilter/);
  });
});
