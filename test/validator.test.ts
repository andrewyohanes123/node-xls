/**
 * Sanity checks for the validation oracle itself: it must accept files from
 * independent, known-good writers and reject known-bad packages.
 */
import { describe, it, expect } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import * as SheetJS from 'xlsx';
import ExcelJS from 'exceljs';
import { validateXlsx, SCHEMA_VALIDATION } from './helpers/validate.js';

describe('validator oracle', () => {
  it('schema validation is enabled (run scripts/fetch-schemas.sh)', () => {
    expect(SCHEMA_VALIDATION).toBe(true);
  });

  it('accepts a SheetJS CE generated workbook', async () => {
    const wb = SheetJS.utils.book_new();
    const ws = SheetJS.utils.aoa_to_sheet([
      ['a', 'b', new Date(2026, 0, 1)],
      [1, true, '  x  '],
    ]);
    ws['!merges'] = [SheetJS.utils.decode_range('A3:B3')];
    SheetJS.utils.book_append_sheet(wb, ws, 'Sheet One');
    await validateXlsx(SheetJS.write(wb, { type: 'array', bookType: 'xlsx' }));
  });

  it('accepts an ExcelJS generated workbook with styles, validation and conditional formatting', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow(['Name', 'Score']);
    ws.addRow(['Al', 3]);
    ws.getCell('A1').font = { bold: true, color: { argb: 'FFFF0000' } };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    ws.getCell('B2').dataValidation = { type: 'whole', operator: 'between', formulae: [0, 10] };
    ws.addConditionalFormatting({
      ref: 'B2:B10',
      rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [5], priority: 1, style: { font: { bold: true } } }],
    });
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    await validateXlsx(new Uint8Array(await wb.xlsx.writeBuffer()));
  });

  it('rejects a package with a dangling relationship', async () => {
    const wb = SheetJS.utils.book_new();
    SheetJS.utils.book_append_sheet(wb, SheetJS.utils.aoa_to_sheet([[1]]), 'S');
    const files = unzipSync(new Uint8Array(SheetJS.write(wb, { type: 'array' })));
    files['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdX" Type="x" Target="nope.xml"/></Relationships>')
    );
    await expect(validateXlsx(zipSync(files))).rejects.toThrow(/missing part xl\/nope.xml/);
  });

  it('rejects schema-invalid element order', async () => {
    const wb = SheetJS.utils.book_new();
    SheetJS.utils.book_append_sheet(wb, SheetJS.utils.aoa_to_sheet([[1]]), 'S');
    const files = unzipSync(new Uint8Array(SheetJS.write(wb, { type: 'array' })));
    const sheet = strFromU8(files['xl/worksheets/sheet1.xml']);
    files['xl/worksheets/sheet1.xml'] = strToU8(sheet.replace('<sheetData>', '<mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells><sheetData>'));
    await expect(validateXlsx(zipSync(files))).rejects.toThrow(/XSD/);
  });
});
