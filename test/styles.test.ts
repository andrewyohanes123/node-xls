/**
 * Cell styling (the core "Pro" feature): every property is checked through
 * ExcelJS's independent reading of the file and node-xls' own round trip.
 */
import { describe, it, expect } from 'vitest';
import { utils, CellStyle, BorderStyle, FillPattern } from '../src/index.js';
import { roundtripSheet } from './helpers/util.js';

async function styled(style: CellStyle, value: any = 'x') {
  const ws = utils.aoa_to_sheet([[value]]);
  ws.A1.s = style;
  const rt = await roundtripSheet(ws);
  return { ...rt, ex: rt.rep.exceljs.getWorksheet('S')!.getCell('A1'), back: rt.ws.A1 };
}

describe('fonts', () => {
  it('name, size, bold, italic, strike, underline, color', async () => {
    const { ex, back } = await styled({
      font: { name: 'Arial', sz: 14, bold: true, italic: true, strike: true, underline: 'double', color: { rgb: 'C00000' } },
    });
    expect(ex.font).toMatchObject({ name: 'Arial', size: 14, bold: true, italic: true, strike: true, underline: 'double', color: { argb: 'FFC00000' } });
    expect(back.s.font).toEqual({ name: 'Arial', sz: 14, bold: true, italic: true, strike: true, underline: 'double', color: { rgb: 'C00000' } });
  });

  it('single underline, superscript', async () => {
    const { ex, back } = await styled({ font: { underline: true, vertAlign: 'superscript' } });
    expect(ex.font).toMatchObject({ underline: true, vertAlign: 'superscript' });
    expect(back.s.font).toMatchObject({ underline: true, vertAlign: 'superscript' });
  });

  it('theme, indexed and tinted colors', async () => {
    const { ex, back } = await styled({ font: { color: { theme: 4, tint: -0.25 } } });
    expect(ex.font.color).toEqual({ theme: 4, tint: -0.25 });
    expect(back.s.font.color).toEqual({ theme: 4, tint: -0.25 });
    const idx = await styled({ font: { color: { indexed: 10 } } });
    expect(idx.ex.font.color).toEqual({ indexed: 10 });
  });

  it('accepts #RRGGBB, RGB shorthand and ARGB strings', async () => {
    for (const [input, argb] of [
      ['#00ff00', 'FF00FF00'],
      ['0F0', 'FF00FF00'],
      ['8000FF00', '8000FF00'],
    ]) {
      const { ex } = await styled({ font: { color: { rgb: input } } });
      expect(ex.font.color).toEqual({ argb });
    }
    await expect(styled({ font: { color: { rgb: 'nothex' } } })).rejects.toThrow(/Invalid color/);
  });

  it('a partial font keeps Excel defaults (Calibri 11, theme minor)', async () => {
    const { ex, rep } = await styled({ font: { bold: true } });
    expect(ex.font).toMatchObject({ bold: true, size: 11, name: 'Calibri' });
    expect(rep.xml('xl/styles.xml')).toContain('<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');
  });
});

describe('fills', () => {
  it('solid fill from fgColor alone', async () => {
    const { ex, back } = await styled({ fill: { fgColor: { rgb: 'FFFF00' } } });
    expect(ex.fill).toMatchObject({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } });
    expect(back.s.fill).toEqual({ patternType: 'solid', fgColor: { rgb: 'FFFF00' } });
  });

  const patterns: FillPattern[] = [
    'darkGray',
    'mediumGray',
    'lightGray',
    'gray0625',
    'darkHorizontal',
    'darkVertical',
    'darkDown',
    'darkUp',
    'darkGrid',
    'darkTrellis',
    'lightHorizontal',
    'lightVertical',
    'lightDown',
    'lightUp',
    'lightGrid',
    'lightTrellis',
  ];
  it('all pattern types with fg/bg colors', async () => {
    for (const patternType of patterns) {
      const { ex, back } = await styled({ fill: { patternType, fgColor: { rgb: 'FF0000' }, bgColor: { theme: 0 } } });
      expect(ex.fill, patternType).toMatchObject({ pattern: patternType, fgColor: { argb: 'FFFF0000' }, bgColor: { theme: 0 } });
      expect(back.s.fill.patternType).toBe(patternType);
    }
  });

  it('theme fill colors stay theme colors (never black)', async () => {
    const { ex } = await styled({ fill: { fgColor: { theme: 5, tint: 0.6 } } });
    expect(ex.fill).toMatchObject({ pattern: 'solid', fgColor: { theme: 5, tint: 0.6 } });
  });

  it('patternType none / empty fill writes no fill', async () => {
    const { rep } = await styled({ fill: { patternType: 'none' } });
    expect(rep.xml('xl/styles.xml')).toContain('<fills count="2">');
  });
});

describe('borders', () => {
  const styles: BorderStyle[] = [
    'thin',
    'medium',
    'dashed',
    'dotted',
    'thick',
    'double',
    'hair',
    'mediumDashed',
    'dashDot',
    'mediumDashDot',
    'dashDotDot',
    'mediumDashDotDot',
    'slantDashDot',
  ];
  it('every border style on every side', async () => {
    for (const style of styles) {
      const side = { style, color: { rgb: '0000FF' } };
      const { ex, back } = await styled({ border: { top: side, bottom: side, left: side, right: side } });
      for (const k of ['top', 'bottom', 'left', 'right'] as const) {
        expect(ex.border[k], `${style} ${k}`).toEqual({ style, color: { argb: 'FF0000FF' } });
        expect(back.s.border[k]).toEqual({ style, color: { rgb: '0000FF' } });
      }
    }
  });

  it('diagonal borders and default (auto) color', async () => {
    const { ex, back } = await styled({ border: { diagonal: { style: 'thin' }, diagonalUp: true, diagonalDown: true } });
    expect(ex.border.diagonal).toMatchObject({ up: true, down: true, style: 'thin' });
    expect(back.s.border).toEqual({ diagonal: { style: 'thin' }, diagonalUp: true, diagonalDown: true });
  });
});

describe('alignment', () => {
  it('horizontal, vertical, wrap, indent, shrink, reading order', async () => {
    const { ex, back } = await styled({
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true, indent: 2, readingOrder: 2 },
    });
    expect(ex.alignment).toMatchObject({ horizontal: 'left', vertical: 'middle', wrapText: true, indent: 2, readingOrder: 'rtl' });
    expect(back.s.alignment).toEqual({ horizontal: 'left', vertical: 'center', wrapText: true, indent: 2, readingOrder: 2 });
    const shrink = await styled({ alignment: { horizontal: 'distributed', shrinkToFit: true } });
    expect(shrink.ex.alignment).toMatchObject({ horizontal: 'distributed', shrinkToFit: true });
  });

  it('text rotation including negative angles and vertical text', async () => {
    for (const [input, excel] of [
      [45, 45],
      [90, 90],
      [-45, 135],
      [-90, 180],
      [255, 255],
    ]) {
      const { rep, back } = await styled({ alignment: { textRotation: input } });
      expect(rep.xml('xl/styles.xml')).toContain(`textRotation="${excel}"`);
      expect(back.s.alignment.textRotation).toBe(input);
    }
  });
});

describe('number formats', () => {
  it('built-in formats are referenced by id, custom ones get ids >= 164', async () => {
    const ws = utils.aoa_to_sheet([[1, 2, 3, 4]]);
    ws.A1.s = { numFmt: '0.00' };
    ws.B1.s = { numFmt: 10 };
    ws.C1.s = { numFmt: '$#,##0.00' };
    ws.D1.z = '[Blue]0.0;[Red]-0.0';
    const { rep, ws: back } = await roundtripSheet(ws);
    const styles = rep.xml('xl/styles.xml');
    expect(styles).toMatch(/<numFmt numFmtId="164" formatCode="\$#,##0.00"\/>/);
    expect(styles).toMatch(/<xf numFmtId="2" /);
    expect(styles).toMatch(/<xf numFmtId="10" /);
    const ex = rep.exceljs.getWorksheet('S')!;
    expect(ex.getCell('A1').numFmt).toBe('0.00');
    expect(ex.getCell('B1').numFmt).toBe('0.00%');
    expect(ex.getCell('C1').numFmt).toBe('$#,##0.00');
    expect(ex.getCell('D1').numFmt).toBe('[Blue]0.0;[Red]-0.0');
    expect([back.A1.w, back.B1.w, back.C1.w, back.D1.w]).toEqual(['1.00', '200.00%', '$3.00', '4.0']);
    expect(back.C1.z).toBe('$#,##0.00');
  });

  it('numeric ids must be built-in', async () => {
    await expect(styled({ numFmt: 200 })).rejects.toThrow(/not a built-in format id/);
    const { rep } = await styled({ numFmt: 44 }, 1.5); // locale-dependent accounting format, implicit in Excel
    expect(rep.xml('xl/styles.xml')).toMatch(/<xf numFmtId="44" /);
  });

  it('s.numFmt takes precedence over cell.z', async () => {
    const ws = utils.aoa_to_sheet([[0.5]]);
    ws.A1.z = '0%';
    ws.A1.s = { numFmt: '0.000' };
    expect((await roundtripSheet(ws)).ws.A1.w).toBe('0.500');
  });
});

describe('protection flags', () => {
  it('locked / hidden', async () => {
    const { ex, back } = await styled({ protection: { locked: false, hidden: true } });
    expect(ex.protection).toEqual({ locked: false, hidden: true });
    expect(back.s.protection).toEqual({ locked: false, hidden: true });
  });
});

describe('style table', () => {
  it('identical styles share one cellXf', async () => {
    const ws = utils.aoa_to_sheet([[1, 2, 3]]);
    utils.set_range_style(ws, 'A1:C1', { font: { bold: true }, fill: { fgColor: { rgb: 'EEEEEE' } } });
    ws.D1 = { t: 'n', v: 4, s: { fill: { fgColor: { rgb: 'EEEEEE' } }, font: { bold: true } } };
    ws['!ref'] = 'A1:D1';
    const { rep } = await roundtripSheet(ws);
    expect(rep.xml('xl/styles.xml')).toContain('<cellXfs count="2">');
    expect(rep.xml('xl/worksheets/sheet1.xml').match(/ s="1"/g)).toHaveLength(4);
  });

  it('read styles are independent copies per cell', async () => {
    const ws = utils.aoa_to_sheet([[1, 2]]);
    utils.set_range_style(ws, 'A1:B1', { font: { bold: true } });
    const back = (await roundtripSheet(ws)).ws;
    back.A1.s.font.italic = true;
    expect(back.B1.s.font.italic).toBeUndefined();
  });

  it('cellStyles:false skips styles on read', async () => {
    const ws = utils.aoa_to_sheet([[1]]);
    ws.A1.s = { font: { bold: true } };
    const { ws: back } = await roundtripSheet(ws, { cellStyles: false });
    expect(back.A1.s).toBeUndefined();
  });
});

describe('rich text', () => {
  it('runs with different fonts in one cell', async () => {
    const runs = [
      { text: 'Bold ', font: { bold: true } },
      { text: 'red', font: { color: { rgb: 'FF0000' }, sz: 14 } },
      { text: ' plain' },
      { text: '2', font: { vertAlign: 'superscript' as const } },
    ];
    const ws = { '!ref': 'A1', A1: { t: 's' as const, v: 'Bold red plain2', richText: runs } };
    const { rep, ws: back } = await roundtripSheet(ws);
    const ex = rep.exceljs.getWorksheet('S')!.getCell('A1').value as any;
    expect(ex.richText.map((r: any) => r.text)).toEqual(['Bold ', 'red', ' plain', '2']);
    expect(ex.richText[0].font).toMatchObject({ bold: true });
    expect(ex.richText[1].font).toMatchObject({ color: { argb: 'FFFF0000' }, size: 14 });
    expect(rep.sheetjs.Sheets.S.A1.v).toBe('Bold red plain2');
    expect(back.A1.v).toBe('Bold red plain2');
    // runs carry complete fonts (defaults filled in, as Excel writes them)
    expect(back.A1.richText).toMatchObject([
      { text: 'Bold ', font: { bold: true, sz: 11, name: 'Calibri' } },
      { text: 'red', font: { color: { rgb: 'FF0000' }, sz: 14, name: 'Calibri' } },
      { text: ' plain' },
      { text: '2', font: { vertAlign: 'superscript', sz: 11 } },
    ]);
    expect(back.A1.richText![2].font).toBeUndefined();
  });

  it('runs inherit the cell font', async () => {
    const ws = { '!ref': 'A1', A1: { t: 's' as const, v: 'ab', s: { font: { name: 'Georgia', sz: 16 } }, richText: [{ text: 'a', font: { bold: true } }, { text: 'b' }] } };
    const back = (await roundtripSheet(ws)).ws;
    expect(back.A1.richText![0].font).toMatchObject({ name: 'Georgia', sz: 16, bold: true });
  });

  it('v is derived from the runs when missing', async () => {
    const ws = { '!ref': 'A1', A1: { t: 's' as const, richText: [{ text: 'a', font: { italic: true } }, { text: 'b' }] } };
    expect((await roundtripSheet(ws)).ws.A1.v).toBe('ab');
  });
});

describe('row and column styles', () => {
  it('apply to empty cells of a row/column', async () => {
    const ws = utils.aoa_to_sheet([['a']]);
    ws['!cols'] = [undefined as any, { s: { fill: { fgColor: { rgb: 'DDEBF7' } } } }];
    ws['!rows'] = [undefined as any, undefined as any, { s: { font: { bold: true } }, hpt: 24 }];
    const { rep, ws: back } = await roundtripSheet(ws);
    const sheet = rep.xml('xl/worksheets/sheet1.xml');
    expect(sheet).toMatch(/<col min="2" max="2" width="9.140625" style="\d+"\/>/);
    expect(sheet).toMatch(/<row r="3" s="\d+" customFormat="1" ht="24" customHeight="1"\/>/);
    const ex = rep.exceljs.getWorksheet('S')!;
    expect(ex.getColumn(2).fill).toMatchObject({ fgColor: { argb: 'FFDDEBF7' } });
    expect(ex.getRow(3).font).toMatchObject({ bold: true });
    expect(back['!cols']![1].s).toEqual({ fill: { patternType: 'solid', fgColor: { rgb: 'DDEBF7' } } });
    expect(back['!rows']![2]).toMatchObject({ hpt: 24, s: { font: { bold: true } } });
  });
});
