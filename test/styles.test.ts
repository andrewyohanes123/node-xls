import { describe, it, expect } from 'vitest';
import XLSX, { utils, CellStyle } from '../src/index.js';

describe('Styling Support (Fonts, Fills, Borders, Alignment, Number Formats)', () => {
  it('should write and read styled cells with full fidelity', () => {
    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet([
      ['Header 1', 'Header 2', 'Header 3'],
      ['Row 1', 1234.56, true],
    ]);

    const headerStyle: CellStyle = {
      font: {
        name: 'Arial',
        sz: 14,
        bold: true,
        italic: true,
        underline: true,
        color: { rgb: 'FF0000' }, // Red
      },
      fill: {
        patternType: 'solid',
        fgColor: { rgb: 'FFFF00' }, // Yellow background
      },
      border: {
        top: { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'double', color: { rgb: '0000FF' } },
        left: { style: 'medium', color: { rgb: '000000' } },
        right: { style: 'medium', color: { rgb: '000000' } },
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
    };

    const currencyStyle: CellStyle = {
      font: { name: 'Calibri', sz: 11, bold: true },
      numFmt: '$#,##0.00',
      alignment: { horizontal: 'right' },
    };

    // Apply styles
    utils.set_cell_style(ws, 'A1', headerStyle);
    utils.set_range_style(ws, 'B1:C1', headerStyle);
    utils.set_cell_style(ws, 'B2', currencyStyle);

    utils.book_append_sheet(wb, ws, 'StyledSheet');

    // Write to buffer
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);

    // Read back
    const readWb = XLSX.read(buffer, { type: 'buffer' });
    expect(readWb.SheetNames).toContain('StyledSheet');

    const readWs = readWb.Sheets['StyledSheet'];
    expect(readWs['A1'].v).toBe('Header 1');

    // Verify cell A1 styles
    const a1Style = readWs['A1'].s;
    expect(a1Style).toBeDefined();
    expect(a1Style?.font?.name).toBe('Arial');
    expect(a1Style?.font?.sz).toBe(14);
    expect(a1Style?.font?.bold).toBe(true);
    expect(a1Style?.font?.italic).toBe(true);
    expect(a1Style?.font?.underline).toBe(true);
    expect(a1Style?.font?.color).toEqual({ rgb: 'FF0000' });

    // Verify fill
    expect(a1Style?.fill?.patternType).toBe('solid');
    expect(a1Style?.fill?.fgColor).toEqual({ rgb: 'FFFF00' });

    // Verify border
    expect(a1Style?.border?.top?.style).toBe('thin');
    expect(a1Style?.border?.bottom?.style).toBe('double');
    expect(a1Style?.border?.bottom?.color).toEqual({ rgb: '0000FF' });

    // Verify alignment
    expect(a1Style?.alignment?.horizontal).toBe('center');
    expect(a1Style?.alignment?.vertical).toBe('center');
    expect(a1Style?.alignment?.wrapText).toBe(true);

    // Verify cell B2 currency format
    const b2Style = readWs['B2'].s;
    expect(b2Style?.numFmt).toBe('$#,##0.00');
    expect(b2Style?.font?.bold).toBe(true);
  });

  it('should support column widths and row heights with styles', () => {
    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet([['Wide Col', 'Tall Row']]);

    ws['!cols'] = [{ wch: 30 }, { width: 15 }];
    ws['!rows'] = [{ hpt: 40 }];

    utils.book_append_sheet(wb, ws, 'Dimensions');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const readWb = XLSX.read(buf, { type: 'buffer' });
    const readWs = readWb.Sheets['Dimensions'];

    expect(readWs['!cols']).toBeDefined();
    expect(readWs['!cols']![0].width).toBe(31);
    expect(readWs['!rows']).toBeDefined();
    expect(readWs['!rows']![0].hpt).toBe(40);
  });
});
