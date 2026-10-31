#!/usr/bin/env node
// Generates one showcase workbook per feature area into out/showcase/ using the
// built package (dist/), i.e. exactly what consumers get. Inspect them with
// `npm run render` (PNG via LibreOffice), `npm run inspect -- <file>` (XML), or
// by opening them in Excel.
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';
import XLSX from '../dist/index.js';
import { makePng } from './png.mjs';

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
    ['[Blue]0.00;[Red]-0.00', -5, '-5.00 in red'],
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
  ws['!freeze'] = 'B3';
  ws['!tabColor'] = { theme: 5 };
  ws['!views'] = { zoomScale: 110 };
  ws['!margins'] = { left: 0.5, right: 0.5 };
  ws['!pageSetup'] = { orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 };
  utils.set_cell_style(ws, 'A1', { ...title, alignment: { horizontal: 'center', vertical: 'center' } });
  utils.set_range_style(ws, 'A2:E2', head);
  return ws;
});

// ---------------------------------------------------------------------------
section('05-rich-text', () => {
  const ws = utils.aoa_to_sheet([['Rich text runs inside single cells'], [], ['Example', 'Expected']]);
  ws.A4 = {
    t: 's',
    richText: [
      { text: 'Bold', font: { bold: true } },
      { text: ' + ' },
      { text: 'red italic', font: { italic: true, color: { rgb: 'C00000' } } },
      { text: ' + ' },
      { text: 'big', font: { sz: 16, name: 'Georgia' } },
    ],
  };
  ws.B4 = { t: 's', v: 'Bold, red italic, big Georgia in one cell' };
  ws.A5 = { t: 's', richText: [{ text: 'H' }, { text: '2', font: { vertAlign: 'subscript' } }, { text: 'O and x' }, { text: '2', font: { vertAlign: 'superscript' } }] };
  ws.B5 = { t: 's', v: 'H₂O and x²' };
  ws.A6 = { t: 's', richText: [{ text: 'underlined', font: { underline: true } }, { text: ' and ' }, { text: 'struck', font: { strike: true } }] };
  ws.B6 = { t: 's', v: 'underline + strikethrough' };
  ws['!ref'] = 'A1:B6';
  ws['!cols'] = [{ wch: 40 }, { wch: 40 }];
  utils.set_cell_style(ws, 'A1', title);
  utils.set_range_style(ws, 'A3:B3', head);
  utils.set_range_style(ws, 'B4:B6', note);
  return ws;
});

// ---------------------------------------------------------------------------
section('06-conditional', () => {
  const rows = [['Rep', 'Q1', 'Q2', 'Q3', 'Q4', 'Status', 'Due']];
  const names = ['Ann', 'Ben', 'Cat', 'Dan', 'Eve', 'Fay', 'Gus', 'Hal'];
  names.forEach((n, i) => rows.push([n, 20 + ((i * 37) % 80), 10 + ((i * 53) % 90), 5 + ((i * 29) % 95), (i * 61) % 100, i % 3 ? 'ok' : 'late', new Date(2026, 0, 1 + i * 9)]));
  const ws = utils.aoa_to_sheet(rows);
  const last = rows.length;
  ws['!condfmt'] = [
    { ref: `B2:B${last}`, rules: [{ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], colors: [{ rgb: 'F8696B' }, { rgb: 'FFEB84' }, { rgb: '63BE7B' }] }] },
    { ref: `C2:C${last}`, rules: [{ type: 'dataBar', color: { rgb: '638EC6' } }] },
    { ref: `D2:D${last}`, rules: [{ type: 'iconSet', iconSet: '3TrafficLights1' }] },
    { ref: `E2:E${last}`, rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [60], style: { font: { bold: true, color: { rgb: '9C0006' } }, fill: { fgColor: { rgb: 'FFC7CE' } } } }] },
    { ref: `F2:F${last}`, rules: [{ type: 'containsText', text: 'late', style: { font: { color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'C00000' } } } }] },
    { ref: `A2:G${last}`, rules: [{ type: 'expression', formula: '$E2<10', style: { font: { italic: true, strike: true } } }] },
  ];
  ws['!cols'] = [{ wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 12 }];
  utils.set_range_style(ws, 'A1:G1', head);
  utils.set_range_style(ws, `G2:G${last}`, { numFmt: 'dd mmm' });
  return ws;
});

// ---------------------------------------------------------------------------
section('07-validation', () => {
  const ws = utils.aoa_to_sheet([
    ['Field', 'Input (try it in Excel)', 'Rule'],
    ['Status', 'Open', 'dropdown: Open / Closed / On hold'],
    ['Quantity', 5, 'whole number 1..100 (stop alert)'],
    ['Discount', 0.1, 'decimal <= 0.5 (warning alert)'],
    ['Start date', new Date(2026, 0, 1), 'date >= 2026-01-01'],
    ['Code', 'AB12', 'text length exactly 4'],
    ['Owner', 'Ann', 'dropdown from range E2:E4'],
    ['Note', '', 'input prompt only'],
  ]);
  ws.E1 = { t: 's', v: 'Owners' };
  ws.E2 = { t: 's', v: 'Ann' };
  ws.E3 = { t: 's', v: 'Ben' };
  ws.E4 = { t: 's', v: 'Cat' };
  ws['!ref'] = 'A1:E8';
  ws['!dataValidation'] = [
    { ref: 'B2', type: 'list', values: ['Open', 'Closed', 'On hold'], prompt: { title: 'Status', message: 'Pick a status' } },
    { ref: 'B3', type: 'whole', operator: 'between', formula1: 1, formula2: 100, error: { title: 'Invalid', message: 'Enter 1 to 100' } },
    { ref: 'B4', type: 'decimal', operator: 'lessThanOrEqual', formula1: 0.5, error: { message: 'Max 50%', style: 'warning' } },
    { ref: 'B5', type: 'date', operator: 'greaterThanOrEqual', formula1: new Date(2026, 0, 1) },
    { ref: 'B6', type: 'textLength', operator: 'equal', formula1: 4 },
    { ref: 'B7', type: 'list', formula1: '$E$2:$E$4' },
    { ref: 'B8', type: 'any', prompt: { title: 'Note', message: 'Free text' } },
  ];
  ws['!cols'] = [{ wch: 12 }, { wch: 24 }, { wch: 36 }, { wch: 2 }, { wch: 10 }];
  ws.B4.z = '0%';
  ws.B5.z = 'yyyy-mm-dd';
  utils.set_range_style(ws, 'A1:C1', head);
  utils.set_cell_style(ws, 'E1', head);
  utils.set_range_style(ws, 'B2:B8', { fill: { fgColor: { rgb: 'FFF2CC' } }, protection: { locked: false } });
  utils.set_range_style(ws, 'C2:C8', note);
  return ws;
});

// ---------------------------------------------------------------------------
section('08-tables', () => {
  const ws = utils.aoa_to_sheet([
    ['Product', 'Region', 'Units', 'Revenue'],
    ['Chair', 'East', 12, 1440],
    ['Desk', 'West', 4, 1200],
    ['Lamp', 'East', 30, 750],
    ['Shelf', 'North', 7, 630],
    [null, null, null, null],
  ]);
  ws['!tables'] = [
    {
      name: 'SalesTable',
      ref: 'A1:D6',
      style: 'TableStyleMedium9',
      totalsRow: true,
      columns: [{ totalsRowLabel: 'Total' }, {}, { totalsRowFunction: 'sum' }, { totalsRowFunction: 'sum' }],
    },
  ];
  utils.set_range_style(ws, 'D2:D6', { numFmt: '$#,##0' });
  ws['!cols'] = [{ wch: 12 }, { wch: 10 }, { wch: 8 }, { wch: 12 }];
  return ws;
});

// ---------------------------------------------------------------------------
section('09-links-comments', (wb) => {
  const ws = utils.aoa_to_sheet([
    ['Kind', 'Cell', 'Expected'],
    ['External URL', 'example.com', 'click opens https://example.com'],
    ['Email', 'Write us', 'mailto link'],
    ['Internal', 'Go to Target!B2', 'jumps to sheet Target'],
    ['Note (hidden)', 'hover me', 'red triangle; note shows on hover'],
    ['Note (visible)', 'always shown', 'note box always visible'],
  ]);
  const linkStyle = { font: { color: { theme: 10 }, underline: true } };
  ws.B2.l = { Target: 'https://example.com', Tooltip: 'Example site' };
  ws.B3.l = { Target: 'mailto:hello@example.com' };
  ws.B4.l = { Target: '#Target!B2' };
  ['B2', 'B3', 'B4'].forEach((r) => utils.set_cell_style(ws, r, linkStyle));
  utils.cell_add_comment(ws.B5, 'Hidden note: hover to see me.', 'node-xls');
  utils.cell_add_comment(ws.B6, 'Visible note', 'node-xls');
  ws.B6.c.hidden = false;
  ws['!cols'] = [{ wch: 16 }, { wch: 18 }, { wch: 36 }];
  utils.set_range_style(ws, 'A1:C1', head);
  utils.set_range_style(ws, 'C2:C6', note);
  utils.book_append_sheet(wb, ws, 'Links');
  const target = utils.aoa_to_sheet([[], [null, 'You arrived at Target!B2']]);
  utils.set_cell_style(target, 'B2', { fill: { fgColor: { rgb: 'C6EFCE' } } });
  utils.book_append_sheet(wb, target, 'Target');
  return null;
});

// ---------------------------------------------------------------------------
section('10-protection', (wb) => {
  const ws = utils.aoa_to_sheet([
    ['Protected sheet (password: node-xls)'],
    ['Locked', 'cannot edit'],
    ['Unlocked', 'you can edit this cell'],
    ['Hidden formula', null],
  ]);
  ws.B4 = { t: 'n', f: 'LEN(B3)' };
  ws['!ref'] = 'A1:B4';
  ws['!protect'] = { password: 'node-xls', selectLockedCells: false, formatColumns: false };
  utils.set_cell_style(ws, 'B3', { protection: { locked: false }, fill: { fgColor: { rgb: 'FFF2CC' } } });
  utils.set_cell_style(ws, 'B4', { protection: { hidden: true } });
  utils.set_cell_style(ws, 'A1', title);
  ws['!cols'] = [{ wch: 16 }, { wch: 28 }];
  utils.book_append_sheet(wb, ws, 'Protected');
  wb.Workbook = { ...(wb.Workbook ?? {}), Protection: { password: 'node-xls' } };
  utils.book_append_sheet(wb, utils.aoa_to_sheet([['This sheet is hidden']]), 'Hidden');
  utils.book_set_sheet_visibility(wb, 'Hidden', 1);
  return null;
});

// ---------------------------------------------------------------------------
section('11-images', () => {
  // gradient "logo" with a circle, generated in memory
  const logo = makePng(160, 80, (x, y) => {
    const inCircle = (x - 40) ** 2 + (y - 40) ** 2 < 28 ** 2;
    return inCircle ? [255, 255, 255, 255] : [31 + x / 2, 78 + y, 120 + x / 3, 255];
  });
  const stripes = makePng(60, 60, (x, y) => (((x + y) >> 3) % 2 ? [237, 125, 49, 255] : [255, 242, 204, 255]));
  const jpeg = Buffer.from('/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAACAAMDAREAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAAB//EABQQAQAAAAAAAAAAAAAAAAAAAAD/xAAVAQEBAAAAAAAAAAAAAAAAAAAHCP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AC9UAkf/2Q==', 'base64');
  const ws = utils.aoa_to_sheet([
    ['Images', null, null, null, null, 'Expected'],
    [],
    [null, null, null, null, null, 'B3: 160x80 gradient logo with white circle (natural size)'],
    [null, null, null, null, null, 'B9: striped square stretched to B9:C14 (two-cell anchor)'],
    [null, null, null, null, null, 'E9: blue JPEG scaled to 90x60 px, offset 10px'],
  ]);
  ws['!images'] = [
    { data: logo, anchor: { from: 'B3' }, name: 'Logo', description: 'Generated gradient logo' },
    { data: stripes, anchor: { from: 'B9', to: 'D15' } },
    { data: jpeg, anchor: { from: 'E9', width: 90, height: 60, offsetX: 10, offsetY: 10 } },
  ];
  ws['!cols'] = [{ wch: 4 }, { wch: 12 }, { wch: 12 }, { wch: 4 }, { wch: 16 }, { wch: 52 }];
  utils.set_cell_style(ws, 'A1', title);
  utils.set_range_style(ws, 'F3:F5', note);
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
