/**
 * Regression tests for defects found by the independent review (round 1).
 * Each test reproduces a verified finding.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import * as SheetJS from 'xlsx';
import XLSX, { utils, WorkBook } from '../src/index.js';
import { validateXlsx } from './helpers/validate.js';
import { book, roundtrip, roundtripSheet } from './helpers/util.js';

type Files = Record<string, Uint8Array>;
const unzip = (wb: WorkBook): Files => unzipSync(XLSX.write(wb, { type: 'array' }));
const edit = (files: Files, part: string, fn: (s: string) => string) => {
  files[part] = strToU8(fn(strFromU8(files[part])));
};
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Rewrites a SpreadsheetML part to use an explicit "x:" prefix (Open XML SDK style). */
function prefixMain(xml: string): string {
  return xml
    .replace(`xmlns="${MAIN}"`, `xmlns:x="${MAIN}"`)
    .replace(/<(\/?)([A-Za-z][\w.-]*)(?=[\s/>])/g, (m, slash, name) => `<${slash}x:${name}`);
}

describe('review round 1: reader robustness', () => {
  it('#1 reads workbooks whose SpreadsheetML elements use a namespace prefix', () => {
    const ws = utils.aoa_to_sheet([['name', 42]]);
    ws.A1.s = { font: { bold: true } };
    utils.cell_add_comment(ws.B1, 'note', 'me');
    const files = unzip(book(['S', ws]));
    for (const p of ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/comments1.xml']) edit(files, p, prefixMain);
    const wb = XLSX.read(zipSync(files));
    expect(wb.SheetNames).toEqual(['S']);
    expect(wb.Sheets.S.A1).toMatchObject({ v: 'name', s: { font: { bold: true } } });
    expect(wb.Sheets.S.B1.v).toBe(42);
    expect(wb.Sheets.S.B1.c[0].t).toBe('note');
  });

  it('#2 decodes numeric character references in text and attributes', async () => {
    const files = unzip(book(['S', utils.aoa_to_sheet([['x']])]));
    edit(files, 'xl/sharedStrings.xml', (s) => s.replace('<t>x</t>', '<t>a&#10;b &#x263A; &#65;</t>'));
    expect(XLSX.read(zipSync(files)).Sheets.S.A1.v).toBe('a\nb ☺ A');

    const ws = utils.aoa_to_sheet([['link']]);
    ws.A1.l = { Target: 'https://example.com', Tooltip: 'line1\nline2' };
    ws['!dataValidation'] = [{ ref: 'B1', type: 'any', prompt: { title: 'T', message: 'p1\np2' } }];
    const back = (await roundtripSheet(ws)).ws;
    expect(back.A1.l.Tooltip).toBe('line1\nline2');
    expect(back['!dataValidation']![0].prompt!.message).toBe('p1\np2');
  });

  it('#4 formats date1904 cells correctly', async () => {
    const wb = book(['S', utils.aoa_to_sheet([[new Date(2026, 0, 15)]], { date1904: true })]);
    wb.Workbook = { WBProps: { date1904: true } };
    const { back } = await roundtrip(wb);
    expect(back.Sheets.S.A1.w).toBe('1/15/26');
    expect(utils.sheet_to_json(back.Sheets.S, { header: 1, raw: false })).toEqual([['1/15/26']]);
  });

  it('#13 recognises built-in CJK date formats (ids 27-36, 50-58)', async () => {
    const files = unzip(book(['S', utils.aoa_to_sheet([[45306, 45306]])]));
    edit(files, 'xl/styles.xml', (s) => s.replace(/<cellXfs count="1">(.*?)<\/cellXfs>/, '<cellXfs count="3">$1<xf numFmtId="58" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="31" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'));
    edit(files, 'xl/worksheets/sheet1.xml', (s) => s.replace('<c r="A1">', '<c r="A1" s="1">').replace('<c r="B1">', '<c r="B1" s="2">'));
    const ws = XLSX.read(zipSync(files)).Sheets.S;
    const ref = SheetJS.read(zipSync(files)).Sheets.S;
    expect(ws.A1.w).toBe(ref.A1.w);
    expect(ws.B1.w).toBe(ref.B1.w);
    const dates = XLSX.read(zipSync(files), { cellDates: true }).Sheets.S;
    expect(dates.A1.t).toBe('d');
    // the original built-in id is kept for writing back
    const again = await roundtrip(XLSX.read(zipSync(files)));
    expect(again.rep.xml('xl/styles.xml')).toMatch(/<xf numFmtId="58" /);
  });

  it('#14 styled empty header cells read from a file behave like missing cells in sheet_to_json', async () => {
    const ws = utils.aoa_to_sheet([
      ['name', null, null],
      ['x', 5, null],
    ]);
    utils.set_range_style(ws, 'B1:C1', { fill: { fgColor: { rgb: 'FFFF00' } } });
    ws['!ref'] = 'A1:C2';
    const { buf, ws: back } = await roundtripSheet(ws);
    const ref = SheetJS.read(buf).Sheets.S;
    expect(utils.sheet_to_json(back, { defval: null })).toEqual(SheetJS.utils.sheet_to_json(ref, { defval: null }));
  });

  it('lower: showRowColHeaders/showZeros "false" spellings are honoured', () => {
    const files = unzip(book(['S', utils.aoa_to_sheet([[0]])]));
    edit(files, 'xl/worksheets/sheet1.xml', (s) => s.replace('<sheetView ', '<sheetView showRowColHeaders="false" showZeros="false" showGridLines="false" '));
    expect(XLSX.read(zipSync(files)).Sheets.S['!views']).toEqual({ showGridLines: false, showRowColHeaders: false, showZeros: false });
  });

  it('lower: relationship types from Strict OOXML are recognised', () => {
    const ws = utils.aoa_to_sheet([['a', 'b'], [1, 2]]);
    utils.cell_add_comment(ws.A2, 'strict note', 'me');
    ws['!tables'] = [{ name: 'StrictT', ref: 'A1:B2' }];
    const files = unzip(book(['S', ws]));
    edit(files, 'xl/worksheets/_rels/sheet1.xml.rels', (s) => s.replace(/http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\//g, 'http://purl.oclc.org/ooxml/officeDocument/relationships/'));
    const back = XLSX.read(zipSync(files)).Sheets.S;
    expect(back.A2.c[0].t).toBe('strict note');
    expect(back['!tables']![0].name).toBe('StrictT');
  });

  it('lower: array formulas set F on every cell of the range (SheetJS parity)', async () => {
    const ws = utils.aoa_to_sheet([[1], [2]]);
    utils.sheet_set_array_formula(ws, 'B1:B2', 'A1:A2*2');
    ws.B1.v = 2;
    ws.B2.v = 4;
    const { buf, ws: back } = await roundtripSheet(ws);
    expect(back.B2.F).toBe('B1:B2');
    expect(utils.sheet_to_formulae(back)).toEqual(SheetJS.utils.sheet_to_formulae(SheetJS.read(buf).Sheets.S));
  });

  it('lower: newer error values (#SPILL!, #CALC!) round-trip', async () => {
    const ws: any = { '!ref': 'A1:B1', A1: { t: 'e', v: '#SPILL!' }, B1: { t: 'e', v: '#CALC!' } };
    const back = (await roundtripSheet(ws)).ws;
    expect([back.A1.w, back.B1.w]).toEqual(['#SPILL!', '#CALC!']);
  });

  it('lower: a sheet autofilter beside (not over) a table is kept', async () => {
    const ws = utils.aoa_to_sheet([
      ['a', 'b', null, 'c', 'd'],
      [1, 2, null, 3, 4],
    ]);
    ws['!tables'] = [{ name: 'Left', ref: 'A1:B2' }];
    ws['!autofilter'] = { ref: 'D1:E2' };
    expect((await roundtripSheet(ws)).ws['!autofilter']).toEqual({ ref: 'D1:E2' });
  });
});

describe('review round 1: whole-column and whole-row ranges', () => {
  it('decode/encode column-only and row-only ranges', () => {
    expect(utils.decode_range('A:B')).toEqual({ s: { c: 0, r: 0 }, e: { c: 1, r: 1048575 } });
    expect(utils.decode_range('$2:$3')).toEqual({ s: { c: 0, r: 1 }, e: { c: 16383, r: 2 } });
  });

  it('#3 reads pivots whose cache source is whole columns, and writes them back', async () => {
    const wb = utils.book_new(utils.json_to_sheet([{ k: 'a', v: 1 }, { k: 'b', v: 2 }]), 'Src');
    utils.add_pivot_table(wb, { sourceSheet: 'Src', targetSheet: 'P', rows: ['k'], values: [{ field: 'v', name: 'Total' }] });
    const files = unzip(wb);
    edit(files, 'xl/pivotCache/pivotCacheDefinition1.xml', (s) => s.replace(/ref="A1:B3"/, 'ref="A:B"'));
    const back = XLSX.read(zipSync(files));
    expect(back.PivotTables![0].sourceRange).toEqual({ s: { c: 0, r: 0 }, e: { c: 1, r: 1048575 } });
    const again = await roundtrip(back);
    // whole-column sources also cover the empty rows below the data: one extra blank record, as Excel shows "(blank)"
    const cache = again.rep.xml('xl/pivotCache/pivotCacheDefinition1.xml');
    expect(cache).toContain('ref="A:B"');
    expect(cache).toContain('recordCount="3"');
  });

  it('#7 conditional formatting on whole columns and rows', async () => {
    const ws = utils.aoa_to_sheet([['x', 'y']]);
    ws['!condfmt'] = [
      { ref: 'A:A', rules: [{ type: 'containsText', text: 'x', style: { font: { bold: true } } }] },
      { ref: '2:2', rules: [{ type: 'containsBlanks', style: { font: { italic: true } } }] },
    ];
    const { ws: back } = await roundtripSheet(ws);
    expect(back['!condfmt']!.map((c) => c.ref)).toEqual(['A:A', '2:2']);
  });

  it('#9c chart series over whole columns cache only the used rows', async () => {
    const ws = utils.aoa_to_sheet([
      ['m', 'v'],
      ['a', 1],
      ['b', 2],
    ]);
    ws['!charts'] = [{ type: 'column', series: [{ categories: "'S'!$A:$A", values: "'S'!$B:$B" }], anchor: { from: 'D1' } }];
    const { rep, ws: back } = await roundtripSheet(ws);
    const xml = rep.xml('xl/charts/chart1.xml');
    expect(xml).toContain('<c:f>S!$B:$B</c:f>');
    expect(xml).toMatch(/<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"\/>/);
    expect(back['!charts']![0].series[0].values).toBe('S!$B:$B');
  });
});

describe('review round 1: writer fidelity', () => {
  it('#5 shared formulas do not shift 3D or sheet-qualified references as cells', () => {
    const files = unzip(book(['S', utils.aoa_to_sheet([[1]])]));
    edit(files, 'xl/worksheets/sheet1.xml', (s) =>
      s.replace(/<sheetData>.*<\/sheetData>/, '<sheetData><row r="1"><c r="B1"><f t="shared" ref="B1:C1" si="0">SUM(JAN:DEC!B5)+AB1!C1+B1</f><v>0</v></c><c r="C1"><f t="shared" si="0"/><v>0</v></c></row></sheetData>')
    );
    expect(XLSX.read(zipSync(files)).Sheets.S.C1.f).toBe('SUM(JAN:DEC!C5)+AB1!D1+C1');
  });

  it('#6 .xlsm files get the macro-enabled workbook content type', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'node-xls-xlsm-'));
    try {
      const file = path.join(dir, 'book.xlsm');
      XLSX.writeFile(book(['S', utils.aoa_to_sheet([[1]])]), file);
      const rep = await validateXlsx(fs.readFileSync(file));
      expect(rep.contentTypes['xl/workbook.xml']).toBe('application/vnd.ms-excel.sheet.macroEnabled.main+xml');
      expect(XLSX.readFile(file).Sheets.S.A1.v).toBe(1);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('#8 custom totals-row formulas and calculated columns survive', async () => {
    const ws = utils.aoa_to_sheet([
      ['Item', 'Sales', 'Double'],
      ['a', 1, null],
      ['b', 2, null],
      [null, null, null],
    ]);
    ws.C2 = { t: 'n', f: 'T[[#This Row],[Sales]]*2' };
    ws.C3 = { t: 'n', f: 'T[[#This Row],[Sales]]*2' };
    ws['!tables'] = [
      {
        name: 'T',
        ref: 'A1:C4',
        totalsRow: true,
        columns: [{ totalsRowLabel: 'Total' }, { totalsRowFunction: 'custom', totalsRowFormula: 'SUM(T[Sales])*2' }, { calculatedColumnFormula: 'T[[#This Row],[Sales]]*2' }],
      },
    ];
    const { rep, ws: back } = await roundtripSheet(ws);
    const table = rep.xml('xl/tables/table1.xml');
    expect(table).toContain('<tableColumn id="2" name="Sales" totalsRowFunction="custom"><totalsRowFormula>SUM(T[Sales])*2</totalsRowFormula></tableColumn>');
    expect(table).toContain('<calculatedColumnFormula>T[[#This Row],[Sales]]*2</calculatedColumnFormula>');
    expect(back.B4.f).toBe('SUM(T[Sales])*2');
    expect(back['!tables']![0].columns![1]).toEqual({ name: 'Sales', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(T[Sales])*2' });
    const again = await roundtripSheet(back);
    expect(again.ws.B4.f).toBe('SUM(T[Sales])*2');
  });

  it('lower: table header text with line breaks is kept (encoded in the column name)', async () => {
    const ws = utils.aoa_to_sheet([['Line1\nLine2', 'b'], [1, 2]]);
    ws['!tables'] = [{ name: 'NL', ref: 'A1:B2' }];
    const { rep, ws: back } = await roundtripSheet(ws);
    expect(back.A1.v).toBe('Line1\nLine2');
    expect(rep.xml('xl/tables/table1.xml')).toContain('name="Line1_x000A_Line2"');
    expect(back['!tables']![0].columns![0].name).toBe('Line1\nLine2');
  });

  it('#9a pivots whose source is a table name are read and written back', async () => {
    const data = utils.json_to_sheet([{ k: 'a', v: 1 }, { k: 'b', v: 2 }]);
    data['!tables'] = [{ name: 'SalesTbl', ref: 'A1:B3' }];
    const wb = utils.book_new(data, 'Src');
    utils.add_pivot_table(wb, { sourceSheet: 'Src', targetSheet: 'P', rows: ['k'], values: [{ field: 'v', name: 'Total' }] });
    const files = unzip(wb);
    edit(files, 'xl/pivotCache/pivotCacheDefinition1.xml', (s) => s.replace(/<worksheetSource [^>]*\/>/, '<worksheetSource name="SalesTbl"/>'));
    const back = XLSX.read(zipSync(files));
    expect(back.PivotTables![0]).toMatchObject({ sourceSheet: 'Src', sourceRange: utils.decode_range('A1:B3') });
    await roundtrip(back);
  });

  it('#9b non-PNG/JPEG/GIF images (EMF, BMP) are passed through', async () => {
    const emf = new Uint8Array(108);
    emf.set([1, 0, 0, 0], 0);
    emf.set([0x20, 0x45, 0x4d, 0x46], 40);
    const bmp = new Uint8Array(70);
    bmp.set([0x42, 0x4d], 0);
    new DataView(bmp.buffer).setInt32(18, 3, true);
    new DataView(bmp.buffer).setInt32(22, 2, true);
    const ws = utils.aoa_to_sheet([['img']]);
    ws['!images'] = [
      { data: emf, anchor: { from: 'A2', width: 40, height: 30 } },
      { data: bmp, anchor: { from: 'C2' } },
    ];
    const { rep, ws: back } = await roundtripSheet(ws);
    expect(rep.contentTypes['xl/media/image1.emf']).toBe('image/x-emf');
    expect(rep.contentTypes['xl/media/image2.bmp']).toBe('image/bmp');
    expect(back['!images']!.map((i) => [i.type, i.anchor.width, i.anchor.height])).toEqual([
      ['emf', 40, 30],
      ['bmp', 3, 2],
    ]);
    await roundtripSheet(back);
  });

  it('#9d chart series with literal values and categories', async () => {
    const ws = utils.aoa_to_sheet([['chart']]);
    ws['!charts'] = [{ type: 'line', series: [{ name: 'Lit', categories: ['Q1', 'Q2', 'Q3'], values: [3, 5, 4] }], anchor: { from: 'B2' } }];
    const { rep, ws: back } = await roundtripSheet(ws);
    const xml = rep.xml('xl/charts/chart1.xml');
    expect(xml).toContain('<c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt>');
    expect(xml).toContain('<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>3</c:v></c:pt>');
    expect(back['!charts']![0].series[0]).toMatchObject({ name: 'Lit', categories: ['Q1', 'Q2', 'Q3'], values: [3, 5, 4] });
    await roundtripSheet(back);
  });

  it('#10 inline list validation rejects values containing commas', () => {
    const ws = utils.aoa_to_sheet([['x']]);
    ws['!dataValidation'] = [{ ref: 'A1', type: 'list', values: ['Smith, John', 'Doe'] }];
    expect(() => XLSX.write(book(['S', ws]), { type: 'array' })).toThrow(/comma/);
  });

  it('#11 fit-to-page is only enabled when requested, and survives read -> write', async () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws['!pageSetup'] = { fitToWidth: 1, fitToHeight: 1, fitToPage: false };
    const { rep, ws: back } = await roundtripSheet(ws);
    expect(rep.xml('xl/worksheets/sheet1.xml')).not.toContain('fitToPage');
    expect(back['!pageSetup']).toMatchObject({ fitToPage: false, fitToWidth: 1, fitToHeight: 1 });
    const on = utils.aoa_to_sheet([[1]]);
    on['!pageSetup'] = { fitToWidth: 1, fitToHeight: 0 };
    const onBack = await roundtripSheet(on);
    expect(onBack.ws['!pageSetup']).toMatchObject({ fitToPage: true });
    expect((await roundtripSheet(onBack.ws)).rep.xml('xl/worksheets/sheet1.xml')).toContain('<pageSetUpPr fitToPage="1"/>');
  });

  it('#12 conditional formatting priorities are preserved across blocks', async () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws['!condfmt'] = [
      { ref: 'A1:A10', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [0], style: { font: { color: { rgb: 'FF0000' } } }, priority: 2 }] },
      { ref: 'A1:A20', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [0], style: { font: { color: { rgb: '00FF00' } } }, priority: 1, stopIfTrue: true }] },
    ];
    const { rep, ws: back } = await roundtripSheet(ws);
    const xml = rep.xml('xl/worksheets/sheet1.xml');
    expect(xml.indexOf('priority="2"')).toBeLessThan(xml.indexOf('priority="1"'));
    const again = await roundtripSheet(back);
    const x2 = again.rep.xml('xl/worksheets/sheet1.xml');
    expect(x2.indexOf('priority="2"')).toBeLessThan(x2.indexOf('priority="1"'));
  });

  it('lower: more than 1023 comments on several sheets keep VML shape ids unique', async () => {
    const big = utils.aoa_to_sheet(Array.from({ length: 1100 }, (_, i) => [i]));
    for (let r = 1; r <= 1100; r++) utils.cell_add_comment(big[`A${r}`], `n${r}`, 'x');
    const small = utils.aoa_to_sheet([[1]]);
    utils.cell_add_comment(small.A1, 'other', 'x');
    const { rep } = await roundtrip(book(['Big', big], ['Small', small]));
    const ids = (p: string) => [...rep.xml(p).matchAll(/<v:shape id="_x0000_s(\d+)"/g)].map((m) => Number(m[1]));
    const all = [...ids('xl/drawings/vmlDrawing1.vml'), ...ids('xl/drawings/vmlDrawing2.vml')];
    expect(new Set(all).size).toBe(all.length);
    const blocks = (p: string) => /data="([^"]+)"/.exec(rep.xml(p))![1].split(',').map(Number);
    for (const p of ['xl/drawings/vmlDrawing1.vml', 'xl/drawings/vmlDrawing2.vml']) {
      const allowed = blocks(p);
      expect(ids(p).every((id) => allowed.includes(Math.floor(id / 1024)))).toBe(true);
    }
  });

  it('lower: single-cell merges are dropped', async () => {
    const ws = utils.aoa_to_sheet([[1, 2]]);
    ws['!merges'] = [utils.decode_range('A1:A1'), utils.decode_range('A2:B2')];
    expect((await roundtripSheet(ws)).ws['!merges']).toEqual([utils.decode_range('A2:B2')]);
  });

  it('lower: external link targets are written as valid URIs', async () => {
    const ws = utils.aoa_to_sheet([['a', 'b']]);
    ws.A1.l = { Target: 'https://example.com/a file.pdf?q=ü' };
    ws.B1.l = { Target: 'https://example.com/already%20encoded' };
    const { rep, ws: back } = await roundtripSheet(ws);
    const rels = rep.xml('xl/worksheets/_rels/sheet1.xml.rels');
    expect(rels).toContain('Target="https://example.com/a%20file.pdf?q=%C3%BC"');
    expect(rels).toContain('Target="https://example.com/already%20encoded"');
    expect(back.A1.l.Target).toBe('https://example.com/a%20file.pdf?q=%C3%BC');
  });
});

describe('review round 1: text import number detection', () => {
  it('#15 CSV and HTML parse ordinary decimals but keep leading-zero codes as text', () => {
    const ws = XLSX.read('a,b,c,d,e,f\n1.50,1e3,-0,007,0x1F,19.90', { type: 'string' }).Sheets.Sheet1;
    expect(['A2', 'B2', 'C2', 'D2', 'E2', 'F2'].map((k) => ws[k].v)).toEqual([1.5, 1000, 0, '007', '0x1F', 19.9]);
    const h = utils.table_to_sheet('<table><tr><td>1.50</td><td>007</td><td>.5</td></tr></table>');
    expect([h.A1.v, h.B1.v, h.C1.v]).toEqual([1.5, '007', 0.5]);
  });
});
