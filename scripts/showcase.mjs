#!/usr/bin/env node
// Generates one showcase workbook per feature area into out/showcase/ using the
// built package (dist/), i.e. exactly what consumers get. Inspect them with
// `npm run render` (PNG via LibreOffice), `npm run inspect -- <file>` (XML), or
// by opening them in Excel.
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import XLSX from '../dist/index.js';

const { utils } = XLSX;
const OUT = join('out', 'showcase');
mkdirSync(OUT, { recursive: true });

const only = process.argv.slice(2);
const sections = [];
const section = (file, build) => sections.push({ file, build });

const title = { font: { bold: true, sz: 14, color: { rgb: '1F4E78' } } };
const head = {
  font: { bold: true, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: '1F4E78' } },
  alignment: { horizontal: 'center', vertical: 'center' },
};
const note = { font: { italic: true, color: { rgb: '7F7F7F' } } };

// ---------------------------------------------------------------------------
section('01-values', () => {
  const ws = utils.aoa_to_sheet([
    ['Cell values & types', null, 'Expected'],
    ['Type', 'Value', 'What Excel should show'],
    ['string', 'Hello, world', 'Hello, world'],
    ['leading zeros', '007', '007 (text, not 7)'],
    ['padded', '  padded  ', 'two spaces each side kept'],
    ['unicode', 'Ünïcödé — 日本語 — emoji 🎉', 'all glyphs'],
    ['multi-line', 'line 1\nline 2', 'two lines (wrap on)'],
    ['control char', 'a\u0001b', 'a?b — file must still open'],
    ['integer', 42, '42'],
    ['float', 3.14159, '3.14159'],
    ['negative', -1234.5, '-1234.5'],
    ['large', 12345678901234, '1.23457E+13 or full'],
    ['boolean', true, 'TRUE'],
    ['NaN', NaN, '#NUM!'],
    ['formula', null, '=SUM(B9:B11) → -1189.35859'],
    ['formula w/ =', null, '=B9*2 → 84'],
    ['string formula', null, '=CONCATENATE("a","b") → ab'],
  ]);
  ws.B15 = { t: 'n', f: 'SUM(B9:B11)' };
  ws.B16 = { t: 'n', f: '=B9*2' };
  ws.B17 = { t: 's', f: 'CONCATENATE("a","b")' };
  ws['!cols'] = [{ wch: 16 }, { wch: 34 }, { wch: 34 }];
  utils.set_cell_style(ws, 'A1', title);
  utils.set_range_style(ws, 'A2:C2', head);
  utils.set_cell_style(ws, 'B7', { alignment: { wrapText: true } });
  utils.set_range_style(ws, 'C3:C17', note);
  return ws;
});

// ---------------------------------------------------------------------------
section('02-styles', () => {
  const rows = [['Fonts', 'Fills', 'Borders', 'Alignment']];
  const ws = utils.aoa_to_sheet(rows);
  const put = (ref, v, s) => {
    ws[ref] = { t: 's', v };
    utils.set_cell_style(ws, ref, s);
  };
  const fonts = [
    ['bold', { font: { bold: true } }],
    ['italic', { font: { italic: true } }],
    ['underline', { font: { underline: true } }],
    ['double underline', { font: { underline: 'double' } }],
    ['strike', { font: { strike: true } }],
    ['Arial 16', { font: { name: 'Arial', sz: 16 } }],
    ['red rgb', { font: { color: { rgb: 'C00000' } } }],
    ['theme 4 (accent1)', { font: { color: { theme: 4 } } }],
    ['Times 12 bold italic', { font: { name: 'Times New Roman', sz: 12, bold: true, italic: true } }],
  ];
  const fills = [
    ['solid yellow', { fill: { fgColor: { rgb: 'FFFF00' } } }],
    ['solid #4472C4', { fill: { patternType: 'solid', fgColor: { rgb: '#4472C4' } }, font: { color: { rgb: 'FFFFFF' } } }],
    ['theme 5 tint .4', { fill: { fgColor: { theme: 5, tint: 0.4 } } }],
    ['darkGray', { fill: { patternType: 'darkGray', fgColor: { rgb: 'FF0000' }, bgColor: { rgb: 'FFFFFF' } } }],
    ['lightGrid', { fill: { patternType: 'lightGrid', fgColor: { rgb: '0070C0' } } }],
    ['lightTrellis', { fill: { patternType: 'lightTrellis', fgColor: { rgb: '00B050' } } }],
    ['gray125', { fill: { patternType: 'gray125' } }],
  ];
  const borderStyles = ['thin', 'medium', 'thick', 'double', 'dashed', 'dotted', 'hair', 'mediumDashDot', 'slantDashDot'];
  const aligns = [
    ['left', { alignment: { horizontal: 'left' } }],
    ['center', { alignment: { horizontal: 'center' } }],
    ['right', { alignment: { horizontal: 'right' } }],
    ['top', { alignment: { vertical: 'top' } }],
    ['wrap: a long sentence that should wrap inside the cell', { alignment: { wrapText: true } }],
    ['rotated 45°', { alignment: { textRotation: 45 } }],
    ['indent 2', { alignment: { horizontal: 'left', indent: 2 } }],
    ['shrink to fit this long text', { alignment: { shrinkToFit: true } }],
  ];
  fonts.forEach(([v, s], i) => put(`A${i + 3}`, v, s));
  fills.forEach(([v, s], i) => put(`B${i + 3}`, v, s));
  borderStyles.forEach((b, i) =>
    put(`C${i + 3}`, b, { border: { top: { style: b }, bottom: { style: b }, left: { style: b }, right: { style: b, color: { rgb: 'FF0000' } } } })
  );
  aligns.forEach(([v, s], i) => put(`D${i + 3}`, v, s));
  ws['!cols'] = [{ wch: 24 }, { wch: 20 }, { wch: 18 }, { wch: 24 }];
  ws['!rows'] = [{ hpt: 22 }, undefined, ...Array(10).fill({ hpt: 30 })];
  utils.set_range_style(ws, 'A1:D1', head);
  return ws;
});

// ---------------------------------------------------------------------------
section('03-numfmt-dates', () => {
  const ws = utils.aoa_to_sheet([
    ['Number formats & dates', null, 'Expected display'],
    ['Format', 'Value', 'Expected'],
  ]);
  const cases = [
    ['General', 1234.5678, '1234.5678'],
    ['0.00', 1234.5678, '1234.57'],
    ['#,##0', 1234.5678, '1,235'],
    ['$#,##0.00', 1234.5678, '$1,234.57'],
    ['0%', 0.256, '26%'],
    ['0.00%', 0.256, '25.60%'],
    ['0.00E+00', 1234.5678, '1.23E+03'],
    ['# ?/?', 1.25, '1 1/4'],
    ['[Red]-0.00;[Blue]0.00', -5, '-5.00 in red'],
    ['yyyy-mm-dd', new Date(2026, 0, 15), '2026-01-15'],
    ['dd mmm yyyy', new Date(2026, 0, 15), '15 Jan 2026'],
    ['yyyy-mm-dd hh:mm', new Date(2026, 0, 15, 13, 45), '2026-01-15 13:45'],
    ['h:mm AM/PM', new Date(2026, 0, 15, 13, 45), '1:45 PM'],
    ['(default for Date)', new Date(2026, 0, 15), 'a date, NOT 46037'],
  ];
  cases.forEach(([fmt, v, expected], i) => {
    const r = i + 3;
    ws[`A${r}`] = { t: 's', v: fmt };
    ws[`B${r}`] = v instanceof Date ? { t: 'd', v } : { t: 'n', v };
    if (!fmt.startsWith('(')) ws[`B${r}`].z = fmt;
    ws[`C${r}`] = { t: 's', v: expected };
  });
  ws['!ref'] = `A1:C${cases.length + 2}`;
  ws['!cols'] = [{ wch: 26 }, { wch: 22 }, { wch: 26 }];
  utils.set_cell_style(ws, 'A1', title);
  utils.set_range_style(ws, 'A2:C2', head);
  utils.set_range_style(ws, `C3:C${cases.length + 2}`, note);
  return ws;
});

// ---------------------------------------------------------------------------
section('04-layout', () => {
  const ws = utils.aoa_to_sheet([
    ['Merged title across A1:E1'],
    ['Region', 'Q1', 'Q2', 'Q3', 'Q4'],
    ['North', 10, 20, 30, 40],
    ['South', 15, 25, 35, 45],
    ['East', 12, 22, 32, 42],
    ['(hidden row)', 0, 0, 0, 0],
    ['West', 11, 21, 31, 41],
  ]);
  ws['!merges'] = [utils.decode_range('A1:E1')];
  ws['!autofilter'] = { ref: 'A2:E7' };
  ws['!cols'] = [{ wch: 18 }, { width: 10 }, { wpx: 80 }, { wch: 10 }, { wch: 10 }];
  ws['!rows'] = [{ hpt: 32 }];
  ws['!rows'][5] = { hidden: true };
  utils.set_cell_style(ws, 'A1', { ...title, alignment: { horizontal: 'center', vertical: 'center' } });
  utils.set_range_style(ws, 'A2:E2', head);
  return ws;
});

// ---------------------------------------------------------------------------
section('13-pivots', (wb) => {
  const data = [];
  const regions = ['East', 'West', 'North'];
  const cats = ['Furniture', 'Technology', 'Supplies'];
  let i = 0;
  for (const r of regions) for (const c of cats) data.push({ Region: r, Category: c, Year: 2025 + (i % 2), Sales: 100 + 37 * i++, Units: i });
  const ws = utils.json_to_sheet(data);
  utils.book_append_sheet(wb, ws, 'Data');
  utils.add_pivot_table(wb, {
    name: 'SalesPivot',
    sourceSheet: 'Data',
    targetSheet: 'Pivot',
    targetCell: 'A3',
    rows: ['Region'],
    cols: ['Category'],
    values: [{ field: 'Sales', subtotal: 'sum', name: 'Total Sales' }],
  });
  return null;
});

// ---------------------------------------------------------------------------
section('Sales_Report', (wb) => {
  const data = [
    { Region: 'East', Category: 'Furniture', Sales: 1200, Profit: 300 },
    { Region: 'East', Category: 'Technology', Sales: 2500, Profit: 800 },
    { Region: 'West', Category: 'Furniture', Sales: 1800, Profit: 450 },
    { Region: 'West', Category: 'Technology', Sales: 3200, Profit: 950 },
    { Region: 'West', Category: 'Supplies', Sales: 600, Profit: 120 },
  ];
  const ws = utils.json_to_sheet(data);
  utils.set_range_style(ws, 'A1:D1', {
    font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '1F497D' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: { bottom: { style: 'medium', color: { rgb: '000000' } } },
  });
  utils.set_range_style(ws, 'C2:D6', { numFmt: '$#,##0.00', alignment: { horizontal: 'right' } });
  ws['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }];
  utils.book_append_sheet(wb, ws, 'Sales Data');
  utils.add_pivot_table(wb, {
    name: 'SalesByCategoryPivot',
    sourceSheet: 'Sales Data',
    targetSheet: 'Pivot Summary',
    targetCell: 'A3',
    rows: ['Region'],
    cols: ['Category'],
    values: [
      { field: 'Sales', subtotal: 'sum', name: 'Total Sales' },
      { field: 'Profit', subtotal: 'sum', name: 'Total Profit' },
    ],
  });
  return null;
});

// ---------------------------------------------------------------------------
let failed = 0;
for (const { file, build } of sections) {
  if (only.length && !only.some((o) => file.includes(o))) continue;
  try {
    const wb = utils.book_new();
    const ws = build(wb);
    if (ws) utils.book_append_sheet(wb, ws, file.replace(/^\d+-/, '').slice(0, 31));
    const path = join(OUT, `${file}.xlsx`);
    XLSX.writeFile(wb, path);
    console.log(`ok    ${path}`);
  } catch (e) {
    failed++;
    console.log(`FAIL  ${file}: ${e.message}`);
  }
}
process.exit(failed ? 1 : 0);
