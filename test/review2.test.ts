/**
 * Regression tests for the second independent review: findings are numbered
 * as in that review. Each test reproduces the original failure.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as SheetJS from 'xlsx';
import ExcelJS from 'exceljs';
import XLSX, { utils, WorkBook, WorkSheet } from '../src/index.js';
import { shiftFormula } from '../src/utils/formula.js';
import { imageSize } from '../src/drawing/image.js';
import { book, roundtrip, roundtripSheet } from './helpers/util.js';
import { validateXlsx } from './helpers/validate.js';
import { makePng } from './helpers/png.js';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const write = (wb: WorkBook) => XLSX.write(wb, { type: 'array' });
const one = (ws: WorkSheet) => book(['S', ws]);

/** Writes the workbook, then edits one XML part by hand (for constructs node-xls does not write itself). */
function patched(wb: WorkBook, part: string, edit: (xml: string) => string): Uint8Array {
  const files = unzipSync(write(wb) as Uint8Array);
  files[part] = strToU8(edit(strFromU8(files[part])));
  return zipSync(files);
}

describe('review round 2: Excel limits are enforced on write', () => {
  it('#41 sheet names with control characters', () => {
    const wb = utils.book_new();
    expect(() => {
      utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'a\u0001b');
      write(wb);
    }).toThrow(/control characters/);
  });

  it('#45 defined names: grammar, scope uniqueness, sheet index, formula length', () => {
    const named = (names: any[]) => {
      const wb = one(utils.aoa_to_sheet([[1]]));
      wb.Workbook = { Names: names };
      return wb;
    };
    expect(() => write(named([{ Name: 'B2', Ref: 'S!$A$1' }]))).toThrow(/defined name/);
    expect(() => write(named([{ Name: 'has space', Ref: 'S!$A$1' }]))).toThrow(/defined name/);
    expect(() => write(named([{ Name: 'x', Ref: 'S!$A$1' }, { Name: 'X', Ref: 'S!$A$1' }]))).toThrow(/Duplicate/);
    expect(() => write(named([{ Name: 'x', Ref: 'S!$A$1', Sheet: 3 }]))).toThrow(/not a sheet index/);
    expect(() => write(named([{ Name: 'x', Ref: '1+'.repeat(4097) + '1' }]))).toThrow(/8192/);
    // same name in different scopes, and built-in names, are fine
    const ok = named([
      { Name: 'x', Ref: 'S!$A$1' },
      { Name: 'x', Ref: 'S!$A$1', Sheet: 0 },
      { Name: '_xlnm.Print_Area', Ref: 'S!$A$1:$B$2', Sheet: 0 },
    ]);
    return roundtrip(ok);
  });

  it('#25/#26 merges: reversed ranges are normalised, out-of-grid ranges rejected', async () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws['!merges'] = [{ s: { r: 2, c: 2 }, e: { r: 0, c: 0 } }, utils.decode_range('B2:D4')];
    expect(() => write(one(ws))).toThrow(/overlap/i);
    ws['!merges'] = [{ s: { r: 2, c: 2 }, e: { r: 0, c: 0 } }];
    const { ws: back } = await roundtripSheet(ws);
    expect(back['!merges']).toEqual([utils.decode_range('A1:C3')]);
    ws['!merges'] = [{ s: { r: 0, c: 16383 }, e: { r: 0, c: 16384 } }];
    expect(() => write(one(ws))).toThrow(/grid/);
    ws['!merges'] = [{ s: { r: -1, c: 0 }, e: { r: 1, c: 0 } }];
    expect(() => write(one(ws))).toThrow(/grid/);
  });

  it('#47 grid limits for autofilter, conditional formats and validations', () => {
    const ws = () => utils.aoa_to_sheet([['h'], [1]]);
    const a = ws();
    a['!autofilter'] = { ref: 'A1:XFE2' };
    expect(() => write(one(a))).toThrow();
    const b = ws();
    b['!condfmt'] = [{ ref: 'A1:A1048577', rules: [{ type: 'expression', formula: 'TRUE', style: { font: { bold: true } } }] }];
    expect(() => write(one(b))).toThrow();
    const c = ws();
    c['!dataValidation'] = [{ ref: 'A0', type: 'whole', operator: 'greaterThan', formula1: 0 }];
    expect(() => write(one(c))).toThrow(/Data validation/);
  });

  it('#52 formulas longer than 8,192 characters', () => {
    const long = '1+'.repeat(4097) + '1';
    const ws = utils.aoa_to_sheet([[{ t: 'n', v: 0, f: long }]]);
    expect(() => write(one(ws))).toThrow(/8192/);
    const cf = utils.aoa_to_sheet([[1]]);
    cf['!condfmt'] = [{ ref: 'A1', rules: [{ type: 'expression', formula: long, style: { font: { bold: true } } }] }];
    expect(() => write(one(cf))).toThrow(/8192/);
    const dv = utils.aoa_to_sheet([[1]]);
    dv['!dataValidation'] = [{ ref: 'A1', type: 'custom', formula1: long }];
    expect(() => write(one(dv))).toThrow(/8192/);
  });

  it('#54 outline levels above 7', () => {
    const a = utils.aoa_to_sheet([[1], [2]]);
    a['!rows'] = [{ level: 8 }];
    expect(() => write(one(a))).toThrow(/outline level/);
    const b = utils.aoa_to_sheet([[1, 2]]);
    b['!cols'] = [{ level: 8 }];
    expect(() => write(one(b))).toThrow(/outline level/);
  });

  it('#56 zoom and print scale ranges, negative margins', () => {
    const a = utils.aoa_to_sheet([[1]]);
    a['!views'] = { zoomScale: 500 };
    expect(() => write(one(a))).toThrow(/zoomScale/);
    const b = utils.aoa_to_sheet([[1]]);
    b['!pageSetup'] = { scale: 5 };
    expect(() => write(one(b))).toThrow(/scale/);
    const c = utils.aoa_to_sheet([[1]]);
    c['!margins'] = { left: -1 };
    expect(() => write(one(c))).toThrow(/Margin/);
  });

  it('#57 hyperlink tooltip and target lengths', () => {
    const a = utils.aoa_to_sheet([['x']]);
    a.A1.l = { Target: 'https://example.com', Tooltip: 't'.repeat(256) };
    expect(() => write(one(a))).toThrow(/tooltip/);
    const b = utils.aoa_to_sheet([['x']]);
    b.A1.l = { Target: 'https://example.com/' + 'p'.repeat(2084) };
    expect(() => write(one(b))).toThrow(/2084/);
  });

  it('#58 style values outside Office ranges', () => {
    const style = (s: any) => {
      const ws = utils.aoa_to_sheet([['x']]);
      ws.A1.s = s;
      return () => write(one(ws));
    };
    expect(style({ font: { sz: 500 } })).toThrow(/font size/);
    expect(style({ alignment: { indent: 251 } })).toThrow(/indent/);
    expect(style({ numFmt: '0'.repeat(256) })).toThrow(/number format/);
    const ws = utils.aoa_to_sheet([['x']]);
    ws['!cols'] = [{ wch: 300 }];
    expect(() => write(one(ws))).toThrow(/width/);
  });

  it('#59 comment text and rich text over 32,767 characters', () => {
    const a = utils.aoa_to_sheet([['x']]);
    a.A1.c = [{ a: 'me', t: 'c'.repeat(32768) }];
    expect(() => write(one(a))).toThrow(/32767/);
    const b = utils.aoa_to_sheet([['x']]);
    b.A1.richText = [{ text: 'a'.repeat(20000) }, { text: 'b'.repeat(20000) }];
    expect(() => write(one(b))).toThrow(/32767/);
  });

  it('#42 data validation error message is capped at 224 characters', async () => {
    const dv = (n: number) => {
      const ws = utils.aoa_to_sheet([[1]]);
      ws['!dataValidation'] = [{ ref: 'A1', type: 'whole', operator: 'greaterThan', formula1: 0, error: { message: 'e'.repeat(n) } }];
      return ws;
    };
    expect(() => write(one(dv(225)))).toThrow(/224/);
    await roundtripSheet(dv(224));
  });
});

describe('review round 2: formulas and text', () => {
  it('#13 shared-formula shifting leaves names with non-ASCII letters, backslashes and escaped structured refs alone', () => {
    expect(shiftFormula('éA1+A1', 1, 1)).toBe('éA1+B2');
    expect(shiftFormula('a\\A1+A1', 1, 0)).toBe('a\\A1+A2');
    expect(shiftFormula("T[[#This Row],['[A1']]]+A1", 1, 0)).toBe("T[[#This Row],['[A1']]]+A2");
  });

  it('#97 rich-text runs do not inherit the cell font flags', async () => {
    const ws = utils.aoa_to_sheet([['x']]);
    ws.A1.s = { font: { bold: true, name: 'Arial', sz: 12 } };
    ws.A1.richText = [{ text: 'plain ', font: { italic: true } }, { text: 'bold', font: { bold: true } }];
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back.A1.richText[0].font).toMatchObject({ italic: true, name: 'Arial', sz: 12 });
    expect(back.A1.richText[0].font.bold).toBeUndefined();
    expect(back.A1.richText[1].font).toMatchObject({ bold: true, name: 'Arial', sz: 12 });
    const runs = (rep.exceljs.getWorksheet('S')!.getCell('A1').value as any).richText;
    expect(runs[0].font.bold).toBeUndefined();
    expect(runs[0].font.italic).toBe(true);
  });

  it('#105 cells sharing a rich string get independent richText arrays', async () => {
    const ws = utils.aoa_to_sheet([['x', 'x']]);
    ws.A1.richText = ws.B1.richText = [{ text: 'r', font: { bold: true } }, { text: 'x' }];
    const { ws: back } = await roundtripSheet(ws);
    back.A1.richText[0].text = 'changed';
    expect(back.B1.richText[0].text).toBe('r');
  });

  it('#109 a range hyperlink gives each cell its own link object', async () => {
    const ws = utils.aoa_to_sheet([['a', 'b']]);
    ws.A1.l = { Target: 'https://example.com' };
    const buf = patched(one(ws), 'xl/worksheets/sheet1.xml', (x) => x.replace(/<hyperlink ref="A1"/, '<hyperlink ref="A1:B1"'));
    await validateXlsx(buf);
    const back = XLSX.read(buf).Sheets.S;
    expect(back.B1.l.Target).toBe('https://example.com');
    back.A1.l.Target = 'https://changed.example';
    expect(back.B1.l.Target).toBe('https://example.com');
  });

  it('#15 a whole-column hyperlink is limited to the used cells', () => {
    const ws = utils.aoa_to_sheet([['a'], ['b']]);
    ws.A1.l = { Target: 'https://example.com' };
    const buf = patched(one(ws), 'xl/worksheets/sheet1.xml', (x) => x.replace(/<hyperlink ref="A1"/, '<hyperlink ref="A:A"'));
    const back = XLSX.read(buf).Sheets.S;
    expect(back.A2.l.Target).toBe('https://example.com');
    expect(Object.keys(back).filter((k) => !k.startsWith('!'))).toHaveLength(2);
  });
});

describe('review round 2: number formats (vendored SSF matches SheetJS)', () => {
  const cases: Array<[string | number, number]> = [
    ['hh:mm:ss', 0.99999999],
    ['h:mm', 0.0416666666],
    ['mm:ss.0', 0.000694444],
    ['[h]:mm:ss', 1.9999999],
    ['h:mm a/p', 0.75],
    ['yyyy-mm-dd hh:mm', 46037.99999999],
    [5, -1234.5],
    [6, 1234.5],
    [7, -0.5],
    [8, 99.999],
    [37, -1234],
    [38, 1234],
    [39, -1234.5],
    [40, 1234.5],
    ['0.00E+00', 123456],
    ['# ?/?', 0.333],
    ['@', 12],
    ['[Red]0.00;[Blue]-0.00', -3],
  ];
  for (const [fmt, v] of cases) {
    it(`#74/#83/#1 format ${JSON.stringify(fmt)} of ${v}`, () => {
      expect(XLSX.SSF.format(fmt, v)).toBe(SheetJS.SSF.format(fmt, v));
    });
  }

  it('#80 format_cell uses s.numFmt ahead of z, like the writer', () => {
    expect(utils.format_cell({ t: 'n', v: 0.5, z: '0.00', s: { numFmt: '0%' } })).toBe('50%');
  });
});

describe('review round 2: dates in the 1904 system', () => {
  it('#0/#30 sheet_to_json returns correct Dates from a date1904 workbook', () => {
    const wb = utils.book_new();
    utils.book_append_sheet(wb, utils.aoa_to_sheet([['d'], [new Date(2026, 0, 15)]], { cellDates: false, date1904: true } as any), 'S');
    wb.Workbook = { ...(wb.Workbook ?? {}), WBProps: { date1904: true } };
    const back = XLSX.read(write(wb));
    expect(back.Sheets.S.A2.w).toBe('1/15/26');
    const rows = utils.sheet_to_json<any>(back.Sheets.S, { UTC: false } as any);
    expect(rows[0].d).toEqual(new Date(2026, 0, 15));
  });

  it('#102/#115 pivots in a date1904 workbook label and cache dates correctly', async () => {
    const wb = utils.book_new();
    wb.Workbook = { WBProps: { date1904: true } };
    utils.book_append_sheet(
      wb,
      utils.aoa_to_sheet(
        [
          ['When', 'Amount'],
          [new Date(2026, 0, 15), 1],
          [new Date(2026, 0, 16), 2],
        ],
        { date1904: true } as any
      ),
      'Data'
    );
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: ['When'], values: ['Amount'] });
    const { rep, back } = await roundtrip(wb);
    expect(utils.format_cell(wb.Sheets.P.A2, undefined, { date1904: true })).toBe('1/15/26');
    expect(rep.xml('xl/pivotCache/pivotCacheDefinition1.xml')).toContain('2026-01-15T00:00:00');
    expect(back.Sheets.P.A2.w).toBe('1/15/26');
  });
});

describe('review round 2: CSV, HTML and I/O', () => {
  it('#22 CSV number cells keep their source text in w', () => {
    const ws = XLSX.read('a,b\n1.10,1e3', { type: 'string' }).Sheets.Sheet1;
    expect(ws.A2).toMatchObject({ t: 'n', v: 1.1, w: '1.10' });
    expect(utils.sheet_to_csv(ws)).toBe('a,b\n1.10,1e3');
  });

  it('#75 separator detection ignores separators inside quotes', () => {
    const ws = XLSX.read('"Name, first, last";amount\n"Smith, J";1\n"Doe, A";2', { type: 'string' }).Sheets.Sheet1;
    expect(ws.A1.v).toBe('Name, first, last');
    expect(ws.A2.v).toBe('Smith, J');
    expect(ws.B3.v).toBe(2);
  });

  it("#81 Excel's sep= directive picks the separator", () => {
    const ws = XLSX.read('sep=|\na|b\n1|2', { type: 'string' }).Sheets.Sheet1;
    expect(ws.A1.v).toBe('a');
    expect(ws.B2.v).toBe(2);
  });

  it('#78/#79 CSV output options and the UTF-8 BOM on byte output', () => {
    const wb = one(utils.aoa_to_sheet([['a', 'b'], [1, 'ü']]));
    expect(XLSX.write(wb, { bookType: 'csv', type: 'string', FS: ';', RS: '\r\n' } as any)).toBe('a;b\r\n1;ü');
    const bytes: Buffer = XLSX.write(wb, { bookType: 'csv', type: 'buffer' });
    expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(SheetJS.read(bytes).Sheets.Sheet1.B2.v).toBe('ü');
  });

  it("#77 read(path, {type: 'file'}) reads the file", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'node-xls-r2-'));
    try {
      const file = path.join(dir, 'f.xlsx');
      XLSX.writeFile(one(utils.aoa_to_sheet([[42]])), file);
      expect(XLSX.read(file, { type: 'file' }).Sheets.S.A1.v).toBe(42);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('#24 HTML with a leading meta tag, comment or div is read as HTML', () => {
    for (const html of ['<meta charset="utf-8"><table><tr><td>7</td></tr></table>', '<!-- x --><div><table><tr><td>7</td></tr></table></div>']) {
      expect(XLSX.read(html, { type: 'string' }).Sheets.Sheet1.A1.v).toBe(7);
    }
  });

  it("#89 xlsx output as type 'string' is refused instead of corrupted", () => {
    expect(() => XLSX.write(one(utils.aoa_to_sheet([[1]])), { type: 'string' })).toThrow();
  });

  it('#86 SheetJS top-level aliases exist', () => {
    for (const k of ['readFileSync', 'writeFileSync', 'writeXLSX', 'set_fs'] as const) expect(typeof (XLSX as any)[k]).toBe('function');
  });
});

describe('review round 2: sheet utilities match SheetJS CE', () => {
  it('#76 sheet_add_aoa keeps the number format of an overwritten cell', () => {
    const mk = (U: any) => {
      const ws = U.aoa_to_sheet([[1]]);
      ws.A1.z = '0.00';
      U.sheet_add_aoa(ws, [[2]], { origin: 'A1' });
      return ws.A1.z;
    };
    expect(mk(utils)).toBe(mk(SheetJS.utils));
  });

  it('#87/#88 sheet_to_json raw:undefined and null header entries', () => {
    const data = [
      ['a', 'b'],
      [1.5, 2],
    ];
    const ours = utils.aoa_to_sheet(data);
    ours.A2.z = '0.0';
    const theirs = SheetJS.utils.aoa_to_sheet(data);
    theirs.A2.z = '0.0';
    expect(utils.sheet_to_json(ours, { raw: undefined } as any)).toEqual(SheetJS.utils.sheet_to_json(theirs, { raw: undefined } as any));
    expect(utils.sheet_to_json(ours, { header: ['x', null as any] })).toEqual(SheetJS.utils.sheet_to_json(theirs, { header: ['x', null as any] }));
  });

  it('#16 sheet_to_json over a whole-column range stops at the used range', () => {
    const ws = utils.aoa_to_sheet([['a'], [1], [2]]);
    expect(utils.sheet_to_json(ws, { range: 'A:A' })).toHaveLength(2);
    expect(utils.sheet_to_json(ws, { range: 'A:A', header: 1, blankrows: true })).toHaveLength(3);
  });

  it('#91 encode_range with string endpoints', () => {
    expect(utils.encode_range('A1' as any, 'B2' as any)).toBe(SheetJS.utils.encode_range('A1' as any, 'B2' as any));
  });

  it('#92 sheet_set_array_formula dynamic flag', () => {
    const ours = utils.aoa_to_sheet([[1], [2]]);
    const theirs = SheetJS.utils.aoa_to_sheet([[1], [2]]);
    utils.sheet_set_array_formula(ours, 'B1:B2', 'A1:A2*2', true);
    SheetJS.utils.sheet_set_array_formula(theirs, 'B1:B2', 'A1:A2*2', true);
    expect(ours.B1).toEqual(theirs.B1);
  });

  it('#93 aoa_to_sheet on empty data leaves !ref unset', () => {
    expect(utils.aoa_to_sheet([])['!ref']).toBe(SheetJS.utils.aoa_to_sheet([])['!ref']);
  });
});

describe('review round 2: reader fidelity', () => {
  it('#90 sheetRows records the original dimension in !fullref', () => {
    const ws = XLSX.read(write(one(utils.aoa_to_sheet([[1], [2], [3]]))), { sheetRows: 1 }).Sheets.S;
    expect(ws['!ref']).toBe('A1');
    expect(ws['!fullref']).toBe('A1:A3');
  });

  it('#73/#112 custom document properties are written and read', async () => {
    const wb = one(utils.aoa_to_sheet([[1]]));
    wb.Custprops = { Project: 'Apollo', Build: 42, Released: true };
    const { rep, back } = await roundtrip(wb);
    expect(rep.sheetjs.Custprops).toMatchObject({ Project: 'Apollo' });
    expect(back.Custprops).toMatchObject({ Project: 'Apollo', Build: 42, Released: true });
  });

  it('#103 several comments on one cell (as SheetJS CE writes them) are all kept', () => {
    const ws = SheetJS.utils.aoa_to_sheet([['x']]);
    SheetJS.utils.cell_add_comment(ws.A1, 'first', 'Ann');
    SheetJS.utils.cell_add_comment(ws.A1, 'second', 'Bob');
    const back = XLSX.read(SheetJS.write(SheetJS.utils.book_new(ws, 'S'), { type: 'buffer', bookType: 'xlsx' })).Sheets.S;
    expect(back.A1.c.map((c: any) => c.t)).toEqual(['first', 'second']);
  });

  it('#113 the legacy workbookPassword survives read -> write', async () => {
    const wb = one(utils.aoa_to_sheet([[1]]));
    wb.Workbook = { ...(wb.Workbook ?? {}), Protection: { passwordHash: 'CC1A' } } as any;
    const { back } = await roundtrip(wb);
    expect(back.Workbook?.Protection).toMatchObject({ passwordHash: 'CC1A', lockStructure: true });
    const again = await roundtrip(back);
    expect(again.rep.xml('xl/workbook.xml')).toContain('workbookPassword="CC1A"');
  });

  it('#111 a sheet protection hash without spinCount is not rewritten as NaN', async () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws['!protect'] = { password: 'pw' };
    const buf = patched(one(ws), 'xl/worksheets/sheet1.xml', (x) => x.replace(/ spinCount="\d+"/, ''));
    const { rep } = await roundtrip(XLSX.read(buf));
    expect(rep.xml('xl/worksheets/sheet1.xml')).not.toContain('NaN');
  });

  it('#63 a containsText rule without a text attribute keeps its formula', async () => {
    const ws = utils.aoa_to_sheet([['abc']]);
    ws['!condfmt'] = [{ ref: 'A1', rules: [{ type: 'containsText', operator: 'containsText', text: 'b', style: { font: { bold: true } } }] }];
    const buf = patched(one(ws), 'xl/worksheets/sheet1.xml', (x) => x.replace(/ text="b"/, ''));
    const { rep } = await roundtrip(XLSX.read(buf));
    expect(rep.xml('xl/worksheets/sheet1.xml')).not.toContain('SEARCH(&quot;&quot;');
    expect(rep.xml('xl/worksheets/sheet1.xml')).toMatch(/<formula>[^<]*SEARCH\("b"/);
  });

  it('#1/#31 built-in format ids without a format code are read like SheetJS', () => {
    const ws = utils.aoa_to_sheet([[45306.75]]);
    ws.A1.z = '0';
    for (const id of [5, 7, 32, 36, 63, 66]) {
      const buf = patched(one(ws), 'xl/styles.xml', (x) => x.replace(/numFmtId="1"/g, `numFmtId="${id}"`));
      const ours = XLSX.read(buf, { cellDates: true }).Sheets.S.A1;
      const theirs = SheetJS.read(buf, { cellDates: true }).Sheets.S.A1;
      expect([id, ours.w, ours.t]).toEqual([id, theirs.w, theirs.t]);
    }
  });

  it('#107 a colourless gray125 fill is kept', async () => {
    const ws = utils.aoa_to_sheet([['x']]);
    ws.A1.s = { fill: { patternType: 'gray125' } };
    const { ws: back } = await roundtripSheet(ws);
    expect(back.A1.s.fill).toMatchObject({ patternType: 'gray125' });
  });

  it('#62 data validations with the alert or prompt turned off stay off', async () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws['!dataValidation'] = [{ ref: 'A1', type: 'whole', operator: 'greaterThan', formula1: 0, showErrorMessage: false, showInputMessage: false }];
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back['!dataValidation'][0]).toMatchObject({ showErrorMessage: false, showInputMessage: false });
    const dv = (rep.exceljs.getWorksheet('S') as any).dataValidations.model.A1;
    expect(dv.showErrorMessage).toBeFalsy();
  });

  it('#68 a table without a style stays unstyled', async () => {
    const ws = utils.aoa_to_sheet([['h'], [1]]);
    ws['!tables'] = [{ name: 'T', ref: 'A1:A2', style: null }];
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back['!tables'][0].style).toBeNull();
    expect(rep.xml('xl/tables/table1.xml')).not.toContain('TableStyleMedium2');
  });

  it('#40/#98 totals rows use displayName and keep cached values', async () => {
    const ws = utils.aoa_to_sheet([['h'], [1], [2], [3]]);
    ws['!tables'] = [{ name: 'T', displayName: 'Shown', ref: 'A1:A4', totalsRow: true, columns: [{ totalsRowFunction: 'sum' }] }];
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(rep.xml('xl/worksheets/sheet1.xml')).toContain('SUBTOTAL(109,Shown[h])');
    // the totals cell is written with its formula; ExcelJS sees the formula
    expect(back.A4.f).toBe('SUBTOTAL(109,Shown[h])');
  });

  it('#17/#70 an ExcelJS table with a custom totals function without formula can be written back', async () => {
    const ejs = new ExcelJS.Workbook();
    const sheet = ejs.addWorksheet('T');
    sheet.addTable({ name: 'Tb', ref: 'A1', totalsRow: true, columns: [{ name: 'K', totalsRowFunction: 'custom' } as any], rows: [[1], [2]] });
    const buf = new Uint8Array(await ejs.xlsx.writeBuffer());
    await validateXlsx(write(XLSX.read(buf)));
  });
});

describe('review round 2: images and charts', () => {
  it('#60/#95 images written by ExcelJS keep their size after read -> write', async () => {
    const ejs = new ExcelJS.Workbook();
    const sheet = ejs.addWorksheet('S');
    const id = ejs.addImage({ buffer: Buffer.from(makePng(40, 20, () => [255, 0, 0, 255])) as any, extension: 'png' });
    sheet.addImage(id, { tl: { col: 1, row: 1 }, ext: { width: 40, height: 20 } });
    const read = XLSX.read(new Uint8Array(await ejs.xlsx.writeBuffer()));
    const img = read.Sheets.S['!images'][0];
    expect(img.anchor.width).toBe(40);
    expect(img.anchor.height).toBe(20);
    const { back } = await roundtrip(read);
    expect(back.Sheets.S['!images'][0].anchor).toMatchObject({ width: 40, height: 20 });
  });

  it('#72 picture stacking order follows the document order (one- and two-cell anchors mixed)', async () => {
    const ejs = new ExcelJS.Workbook();
    const sheet = ejs.addWorksheet('S');
    const png = (rgb: number[]) => ejs.addImage({ buffer: Buffer.from(makePng(4, 4, () => [...rgb, 255])) as any, extension: 'png' });
    sheet.addImage(png([255, 0, 0]), { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 } });
    sheet.addImage(png([0, 255, 0]), { tl: { col: 0, row: 0 }, br: { col: 3, row: 3 } } as any);
    sheet.addImage(png([0, 0, 255]), { tl: { col: 1, row: 1 }, ext: { width: 10, height: 10 } });
    const imgs = XLSX.read(new Uint8Array(await ejs.xlsx.writeBuffer())).Sheets.S['!images'];
    expect(imgs.map((i: any) => [i.anchor.from, i.anchor.to === undefined])).toEqual([
      ['A1', true],
      ['A1', false],
      ['B2', true],
    ]);
  });

  it('#19 OS/2 bitmap headers are sized from 16-bit fields', () => {
    const bmp = new Uint8Array(26 + 12);
    bmp.set([0x42, 0x4d]);
    const dv = new DataView(bmp.buffer);
    dv.setUint32(10, 26, true);
    dv.setUint32(14, 12, true);
    dv.setUint16(18, 30, true);
    dv.setUint16(20, 20, true);
    dv.setUint16(22, 1, true);
    dv.setUint16(24, 24, true);
    expect(imageSize(bmp, 'bmp')).toEqual({ width: 30, height: 20 });
  });

  it('#99 charts sized with width/height keep their size after read -> write', async () => {
    const ws = utils.aoa_to_sheet([['a', 1], ['b', 2]]);
    ws['!charts'] = [{ type: 'column', series: [{ categories: 'S!A1:A2', values: 'S!B1:B2' }], anchor: { from: 'D2', width: 333, height: 222, offsetX: 5 } }];
    const first = await roundtripSheet(ws);
    const second = await roundtrip(first.back);
    const anchorXml = (x: string) => /<xdr:to>.*?<\/xdr:to>/s.exec(x)?.[0];
    expect(anchorXml(second.rep.xml('xl/drawings/drawing1.xml'))).toBe(anchorXml(first.rep.xml('xl/drawings/drawing1.xml')));
  });

  it('#117 chart caches keep error cells as errors, not error codes', async () => {
    const ws = utils.aoa_to_sheet([['a', 1], ['b', { t: 'e', v: 0x2a }], ['c', 3]]);
    ws['!charts'] = [{ type: 'line', series: [{ categories: 'S!A1:A3', values: 'S!B1:B3' }], anchor: { from: 'D2' } }];
    const { rep } = await roundtripSheet(ws);
    const xml = rep.xml('xl/charts/chart1.xml');
    expect(xml).not.toMatch(/<c:v>42<\/c:v>/);
  });

  it('#14 series over a defined name or multi-area range can be written', async () => {
    const ws = utils.aoa_to_sheet([['a', 1], ['b', 2]]);
    const wb = one(ws);
    wb.Workbook = { Names: [{ Name: 'Vals', Ref: 'S!$B$1:$B$2' }] };
    ws['!charts'] = [
      { type: 'column', series: [{ values: 'S!Vals' }], anchor: { from: 'D2' } },
      { type: 'column', series: [{ values: '(S!$B$1,S!$B$2)' }], anchor: { from: 'D20' } },
    ];
    await roundtrip(wb);
  });

  it('#133 a literal series name with "!" is not a reference', async () => {
    const ws = utils.aoa_to_sheet([['a', 1]]);
    ws['!charts'] = [{ type: 'column', series: [{ name: 'Wow!A1', values: [1, 2] }], anchor: { from: 'D2' } }];
    const { ws: back } = await roundtripSheet(ws);
    expect(back['!charts'][0].series[0].name).toBe('Wow!A1');
  });

  it('#130 out-of-range doughnut hole and bar gap are rejected', () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws['!charts'] = [{ type: 'doughnut', holeSize: 99, series: [{ values: [1, 2] }], anchor: { from: 'D2' } } as any];
    expect(() => write(one(ws))).toThrow();
    ws['!charts'] = [{ type: 'column', gapWidth: 600, series: [{ values: [1, 2] }], anchor: { from: 'D2' } } as any];
    expect(() => write(one(ws))).toThrow();
  });
});

describe('review round 2: pivots', () => {
  const data = [
    ['Region', 'Kind', 'Item', 'Sales'],
    ['East', 'A', 'x', 1],
    ['East', 'A', 'y', 2],
    ['East', 'B', 'x', 3],
    ['West', 'A', 'x', 4],
  ];
  const base = (opts: any = {}) => {
    const wb = utils.book_new();
    utils.book_append_sheet(wb, utils.aoa_to_sheet(data), 'Data');
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A3', rows: ['Region'], values: ['Sales'], ...opts });
    return wb;
  };

  it('#96 rendered cells do not share style objects', () => {
    const wb = base();
    const ws = wb.Sheets.P;
    ws.A3.s.font.color = { rgb: 'FF0000' };
    expect(ws.B3.s.font.color).toBeUndefined();
    const other = base();
    expect(other.Sheets.P.A3.s.font.color).toBeUndefined();
  });

  it('#114 large sources with numeric fields do not overflow the stack', async () => {
    const rows: any[][] = [['K', 'V']];
    for (let i = 0; i < 130000; i++) rows.push([i % 3 ? 'a' : 'b', i]);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, utils.aoa_to_sheet(rows), 'Data');
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: ['V'], values: [{ field: 'K', subtotal: 'count' }], renderCells: false });
    expect(() => write(wb)).not.toThrow();
  }, 60000);

  it('#116 error cells in the source propagate instead of reading as blank', () => {
    const wb = utils.book_new();
    utils.book_append_sheet(
      wb,
      utils.aoa_to_sheet([
        ['K', 'V'],
        ['a', 1],
        ['a', { t: 'e', v: 0x07 }],
      ]),
      'Data'
    );
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: ['K'], values: ['V'] });
    expect(wb.Sheets.P.B2).toMatchObject({ t: 'e', w: '#DIV/0!' });
  });

  it('#118/#43/#44/#119 placement: no overlap with other pivots (incl. page rows), tables, merges or the source', () => {
    const wb = base({ filters: ['Kind'] });
    // the first pivot's page field sits at A1:B1, above its location A3
    expect(() => utils.add_pivot_table(wb, { name: 'PT2', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'D1', rows: ['Region'], values: ['Sales'], filters: ['Item', 'Kind'] })).not.toThrow();
    expect(() => utils.add_pivot_table(wb, { name: 'PT3', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A2', rows: ['Region'], values: ['Sales'] })).toThrow(/overlap/i);
    expect(() => utils.add_pivot_table(wb, { name: 'PT4', sourceSheet: 'Data', targetSheet: 'Data', targetCell: 'B2', rows: ['Region'], values: ['Sales'] })).toThrow(/source/i);
    wb.Sheets.Data['!tables'] = [{ name: 'Tb', ref: 'A1:D5' }];
    const t = utils.aoa_to_sheet([['h'], [1]]);
    t['!tables'] = [{ name: 'Tb2', ref: 'A1:A2' }];
    utils.book_append_sheet(wb, t, 'T');
    expect(() => utils.add_pivot_table(wb, { name: 'PT5', sourceSheet: 'Data', targetSheet: 'T', targetCell: 'A1', rows: ['Region'], values: ['Sales'] })).toThrow(/table/i);
  });

  it('#120 compute_pivot_table groups case-insensitively without losing values', () => {
    const comp = utils.compute_pivot_table(
      [
        { K: 'a', V: 1 },
        { K: 'A', V: 2 },
      ],
      { rows: ['K'], values: [{ field: 'V' }] }
    );
    expect(comp.rowKeys).toHaveLength(1);
    expect(comp.grandTotal['Sum of V']).toBe(3);
    expect(comp.rowTotals[comp.rowKeys[0][0]]['Sum of V']).toBe(3);
  });

  it('#121 header names are kept exactly; blank headers are refused', () => {
    const wb = utils.book_new();
    utils.book_append_sheet(
      wb,
      utils.aoa_to_sheet([
        [' Region ', 'Sales'],
        ['a', 1],
      ]),
      'Data'
    );
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: [' Region '], values: ['Sales'] });
    const wb2 = utils.book_new();
    utils.book_append_sheet(
      wb2,
      utils.aoa_to_sheet([
        ['', 'Sales'],
        ['a', 1],
      ]),
      'Data'
    );
    expect(() => utils.add_pivot_table(wb2, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: ['Sales'], values: ['Sales'] })).toThrow(/blank/);
  });

  it('#123 groups without numbers render empty for sum, as Excel does', () => {
    const wb = utils.book_new();
    utils.book_append_sheet(
      wb,
      utils.aoa_to_sheet([
        ['K', 'V'],
        ['a', 1],
        ['b', null],
        ['b', 'text'],
      ]),
      'Data'
    );
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: ['K'], values: ['V'] });
    const ws = wb.Sheets.P;
    expect(ws.A3.v).toBe('b');
    expect(ws.B3?.v).toBeUndefined();
  });

  it('#124 stale cells inside the pivot area are cleared', () => {
    const wb = utils.book_new();
    utils.book_append_sheet(wb, utils.aoa_to_sheet(data), 'Data');
    // West has no Kind B, so that value cell (C4) stays empty in the pivot
    utils.book_append_sheet(wb, utils.aoa_to_sheet([[], [], [], [null, null, 'stale']]), 'P');
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A1', rows: ['Region'], cols: ['Kind'], values: ['Sales'] });
    expect(wb.Sheets.P.A4.v).toBe('West');
    expect(JSON.stringify(wb.Sheets.P)).not.toContain('stale');
  });

  it('#128 outer labels are not repeated after an inner subtotal', () => {
    const wb = base({ rows: ['Region', 'Kind', 'Item'] });
    const col = (r: number) => wb.Sheets.P[utils.encode_cell({ r, c: 0 })]?.v;
    const labels = Array.from({ length: 12 }, (_, i) => col(3 + i)).filter((v) => v === 'East');
    expect(labels).toHaveLength(1);
  });

  it('#94/#122 date and boolean filter selections survive read -> write', async () => {
    const wb = utils.book_new();
    utils.book_append_sheet(
      wb,
      utils.aoa_to_sheet(
        [
          ['When', 'Flag', 'V'],
          [new Date(2026, 0, 15), true, 1],
          [new Date(2026, 0, 16), false, 2],
        ],
        {}
      ),
      'Data'
    );
    const serial = wb.Sheets.Data.A2.v;
    utils.add_pivot_table(wb, { name: 'PT', sourceSheet: 'Data', targetSheet: 'P', targetCell: 'A5', rows: ['V'], values: ['V'], filters: [{ field: 'When', selected: serial }, { field: 'Flag', selected: true }] });
    const { back } = await roundtrip(wb);
    const def = back.Sheets.P['!pivots'][0];
    expect(def.filters).toEqual([
      { field: 'When', selected: serial },
      { field: 'Flag', selected: true },
    ]);
    await roundtrip(back);
  });

  it('#101 data-field number formats are read back', async () => {
    const { back } = await roundtrip(base({ values: [{ field: 'Sales', numFmt: '0.00' }] }));
    expect(back.Sheets.P['!pivots'][0].values[0].numFmt).toBe('0.00');
  });

  it('#104 a read pivot replaced on its sheet is written once, with the replacement', async () => {
    const { back } = await roundtrip(base());
    const old = back.Sheets.P['!pivots'][0];
    back.Sheets.P['!pivots'] = [{ ...old, values: [{ field: 'Sales', name: 'Total', subtotal: 'max' }] }];
    const { rep } = await roundtrip(back);
    expect(Object.keys(rep.files).filter((f) => /pivotTables\/pivotTable\d+\.xml$/.test(f))).toHaveLength(1);
    expect(rep.xml('xl/pivotTables/pivotTable1.xml')).toContain('subtotal="max"');
  });
});
