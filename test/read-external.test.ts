/**
 * Reading files produced by other software: SheetJS CE, ExcelJS, and
 * hand-crafted packages that exercise spec features Excel and other tools
 * emit (inline strings, shared formulas, missing r attributes, t="d" cells,
 * absolute part names, rich text, pretty-printed XML, ...).
 */
import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import * as SheetJS from 'xlsx';
import ExcelJS from 'exceljs';
import XLSX from '../src/index.js';
import { validateXlsx } from './helpers/validate.js';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** Minimal package builder for hand-written parts. */
function pkg(sheetXml: string, extra: Record<string, string> = {}, opts: { absolute?: boolean; pretty?: boolean } = {}) {
  const t = (p: string) => (opts.absolute ? `/xl/${p}` : p);
  const files: Record<string, string> = {
    '[Content_Types].xml':
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels':
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${t(
      'worksheets/sheet1.xml'
    )}"/>${extra['xl/sharedStrings.xml'] ? `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="${t('sharedStrings.xml')}"/>` : ''}</Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${NS}>${sheetXml}</worksheet>`,
    ...extra,
  };
  const u8: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) u8[k] = strToU8(opts.pretty ? v.replace(/></g, '>\n  <') : v);
  return zipSync(u8);
}

describe('files from SheetJS CE', () => {
  it('values, formats, merges, formulas, comments, links', () => {
    const ws = SheetJS.utils.aoa_to_sheet([
      ['Name', 'Value', 'When'],
      ['a', 1.5, new Date(2026, 0, 15)],
      ['b', { t: 'n', v: 3, f: 'B2*2' }, null],
    ]);
    ws.B2.z = '0.00%';
    ws['!merges'] = [SheetJS.utils.decode_range('A4:C4')];
    SheetJS.utils.cell_add_comment(ws.A2, 'note', 'Ann');
    SheetJS.utils.cell_set_hyperlink(ws.A3, 'https://example.com');
    const wb = SheetJS.utils.book_new(ws, 'S');
    const back = XLSX.read(SheetJS.write(wb, { type: 'buffer', bookType: 'xlsx' })).Sheets.S;
    expect(back.B2.w).toBe('150.00%');
    expect(back.C2.w).toBe('1/15/26');
    expect(back.B3.f).toBe('B2*2');
    expect(back['!merges']).toEqual([SheetJS.utils.decode_range('A4:C4')]);
    expect(back.A2.c[0]).toMatchObject({ t: 'note' });
    expect(back.A3.l.Target).toBe('https://example.com');
  });
});

describe('files from ExcelJS', () => {
  it('styles, rich text, shared formulas, validation, CF, tables, protection', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data', { properties: { tabColor: { argb: 'FF00FF00' } } });
    ws.addRow(['Item', 'Qty', 'Double']);
    ws.addRow(['a', 1]);
    ws.addRow(['b', 2]);
    ws.getCell('C2').value = { formula: 'B2*2', result: 2 } as any;
    ws.getCell('C3').value = { sharedFormula: 'C2', result: 4 } as any;
    ws.getCell('A1').font = { bold: true, color: { theme: 4 } };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    ws.getCell('A1').border = { bottom: { style: 'double', color: { argb: 'FF0000FF' } } };
    ws.getCell('B2').numFmt = '0.0';
    ws.getCell('A5').value = { richText: [{ text: 'bold', font: { bold: true } }, { text: ' plain' }] } as any;
    ws.getCell('B5').dataValidation = { type: 'list', allowBlank: true, formulae: ['"x,y"'] };
    ws.addConditionalFormatting({ ref: 'B2:B3', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [1], priority: 1, style: { font: { italic: true } } }] });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.mergeCells('A7:C7');
    const ws2 = wb.addWorksheet('Tbl');
    ws2.addTable({ name: 'People', ref: 'A1', columns: [{ name: 'Who' }, { name: 'Age' }], rows: [['x', 1], ['y', 2]] });
    await ws2.protect('pw', { formatCells: true });
    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    await validateXlsx(buf);

    const back = XLSX.read(buf);
    const s = back.Sheets.Data;
    expect(s.A1.s).toMatchObject({ font: { bold: true, color: { theme: 4 } }, fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } }, border: { bottom: { style: 'double', color: { rgb: '0000FF' } } } });
    expect(s.B2.w).toBe('1.0');
    expect(s.C3.f).toBe('B3*2');
    expect(s.A5.v).toBe('bold plain');
    expect(s.A5.richText?.[0]).toEqual({ text: 'bold', font: expect.objectContaining({ bold: true }) });
    // ExcelJS omits showErrorMessage/showInputMessage, whose schema default is false
    expect(s['!dataValidation']).toEqual([{ ref: 'B5', type: 'list', values: ['x', 'y'], showErrorMessage: false, showInputMessage: false }]);
    expect(s['!condfmt']?.[0].rules[0]).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulae: [1], style: { font: { italic: true } } });
    expect(s['!freeze']).toMatchObject({ ySplit: 1 });
    expect(s['!tabColor']).toEqual({ rgb: '00FF00' });
    expect(s['!merges']).toEqual([{ s: { r: 6, c: 0 }, e: { r: 6, c: 2 } }]);
    const t = back.Sheets.Tbl;
    expect(t['!tables']?.[0]).toMatchObject({ name: 'People', ref: 'A1:B3', columns: [{ name: 'Who' }, { name: 'Age' }] });
    expect(t['!protect']).toMatchObject({ formatCells: false, algorithmName: 'SHA-512' });

    // and node-xls can write it back out as a valid file
    await validateXlsx(XLSX.write(back, { type: 'array' }));
  });
});

describe('hand-crafted spec features', () => {
  it('absolute relationship targets, inline strings, t="str"/"d"/"e", missing r attributes', () => {
    const sheet =
      '<sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> inline </t></is></c><c t="inlineStr"><is><r><rPr><b/></rPr><t>ri</t></r><r><t>ch</t></r></is></c></row>' +
      '<row><c t="str"><f>"a"&amp;"b"</f><v>ab</v></c><c t="d"><v>2026-01-15T00:00:00Z</v></c><c t="e"><v>#DIV/0!</v></c><c><v>  </v></c></row>' +
      '</sheetData>';
    const wb = XLSX.read(pkg(sheet, {}, { absolute: true }));
    const ws = wb.Sheets.Data;
    expect(ws.A1.v).toBe(' inline ');
    expect(ws.B1.v).toBe('rich');
    expect(ws.B1.richText).toEqual([{ text: 'ri', font: { bold: true } }, { text: 'ch' }]);
    expect(ws.A2).toMatchObject({ t: 's', v: 'ab', f: '"a"&"b"' });
    expect(ws.B2).toMatchObject({ t: 'n', v: 46037, w: '1/15/26' });
    expect(ws.C2).toMatchObject({ t: 'e', v: 0x07, w: '#DIV/0!' });
    expect(ws.D2).toBeUndefined();
    expect(ws['!ref']).toBe('A1:C2');
    const dates = XLSX.read(pkg(sheet), { cellDates: true }).Sheets.Data;
    expect(dates.B2.v).toEqual(new Date(2026, 0, 15));
  });

  it('shared formulas expand with relative references shifted', () => {
    const sheet =
      '<sheetData>' +
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:C3" si="0">A1*$A$1+SUM(A$1:A1)&amp;"A1"</f><v>2</v></c><c r="C1"><f t="shared" si="0"/><v>2</v></c></row>' +
      '<row r="3"><c r="B3"><f t="shared" si="0"/><v>0</v></c></row>' +
      '</sheetData>';
    const ws = XLSX.read(pkg(sheet)).Sheets.Data;
    expect(ws.B1.f).toBe('A1*$A$1+SUM(A$1:A1)&"A1"');
    expect(ws.C1.f).toBe('B1*$A$1+SUM(B$1:B1)&"A1"');
    expect(ws.B3.f).toBe('A3*$A$1+SUM(A$1:A3)&"A1"');
  });

  it('pretty-printed XML with whitespace between elements', () => {
    const sheet = '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>12</v></c></row></sheetData>';
    const sst = '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t xml:space="preserve">  keep  </t></si></sst>';
    const ws = XLSX.read(pkg(sheet, { 'xl/sharedStrings.xml': sst }, { pretty: true })).Sheets.Data;
    expect(ws.A1.v).toBe('  keep  ');
    expect(ws.B1.v).toBe(12);
  });

  it('rejects non-xlsx zip data and corrupt archives', () => {
    expect(() => XLSX.read(zipSync({ 'hello.txt': strToU8('hi') }))).toThrow(/workbook/);
    const good = XLSX.write(XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([[1]])), { type: 'array' }) as Uint8Array;
    expect(() => XLSX.read(good.slice(0, good.length / 2))).toThrow();
  });
});
