#!/usr/bin/env node
// Builds out/excel-check/: small focused workbooks + CHECKLIST.md for manual
// verification in Microsoft Excel (the final judge for pivots and charts).
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import XLSX from '../dist/index.js';
import { makePng } from './png.mjs';

const { utils } = XLSX;
const OUT = join('out', 'excel-check');
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });

const head = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F4E78' } } };
const sales = [];
['East', 'West', 'North'].forEach((region, ri) =>
  ['Furniture', 'Technology', 'Supplies'].forEach((category, ci) =>
    [2025, 2026].forEach((year, yi) =>
      sales.push({ Region: region, Category: category, Year: year, Date: new Date(year, ci * 3 + ri, 10 + yi), Sales: 100 + ri * 50 + ci * 20 + yi * 7, Units: 1 + ((ri + ci + yi) % 5) })
    )
  )
);

const checks = [];
function pivotFile(file, expect, options) {
  const wb = utils.book_new();
  const ws = utils.json_to_sheet(sales, { dateNF: 'yyyy-mm-dd' });
  utils.set_range_style(ws, 'A1:F1', head);
  ws['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 6 }, { wch: 11 }, { wch: 7 }, { wch: 6 }];
  utils.book_append_sheet(wb, ws, 'Data');
  utils.add_pivot_table(wb, { name: 'Pivot1', sourceSheet: 'Data', targetSheet: 'Pivot', targetCell: 'A3', ...options });
  wb.Workbook = { Views: [{ activeTab: 1 }] };
  const path = join(OUT, `${file}.xlsx`);
  XLSX.writeFile(wb, path);
  checks.push({ file: `${file}.xlsx`, expect });
  return path;
}

pivotFile('pivot-1-rows', 'Rows = Region (East, North, West), one value "Total Sales", Grand Total row. Sales sum = 2,358.', {
  rows: ['Region'],
  values: [{ field: 'Sales', name: 'Total Sales' }],
});
pivotFile('pivot-2-rows-cols', 'Rows = Region, Columns = Year (2025, 2026), Grand Total row and column.', {
  rows: ['Region'],
  cols: ['Year'],
  values: [{ field: 'Sales', name: 'Total Sales', numFmt: '#,##0' }],
});
pivotFile('pivot-3-subtotals', 'Rows = Region > Category with "East Total" style subtotal rows (tabular form).', {
  rows: ['Region', 'Category'],
  values: [{ field: 'Sales', name: 'Total Sales' }],
});
pivotFile('pivot-4-two-values', 'Rows = Category, two values: "Revenue" ($ format) and "Avg Units" (average, 0.0). "Values" appears in the Columns area.', {
  rows: ['Category'],
  values: [
    { field: 'Sales', name: 'Revenue', numFmt: '$#,##0' },
    { field: 'Units', name: 'Avg Units', subtotal: 'average', numFmt: '0.0' },
  ],
});
pivotFile('pivot-5-full', 'Rows = Region > Category, Columns = Year, two values; subtotals; "Total Revenue"/"Total Units" columns.', {
  rows: ['Region', 'Category'],
  cols: ['Year'],
  values: [
    { field: 'Sales', name: 'Revenue', numFmt: '#,##0' },
    { field: 'Units', name: 'Total Units', subtotal: 'sum' },
  ],
});
pivotFile('pivot-6-filters', 'Report filter "Year" = 2026 above the table (B1), second filter "Category" = (All). Only 2026 rows are summed.', {
  targetCell: 'A1',
  rows: ['Region'],
  filters: [{ field: 'Year', selected: 2026 }, 'Category'],
  values: [{ field: 'Sales', name: 'Total Sales' }],
});
pivotFile('pivot-7-dates', 'Rows = Date (formatted yyyy-mm-dd, chronological), value = count of Region named "Rows".', {
  rows: ['Date'],
  values: [{ field: 'Region', name: 'Rows', subtotal: 'count' }],
});

// A/B diagnostic: same as pivot-2 but without rowItems/colItems (Excel must lay it out itself on refresh).
{
  const src = join(OUT, 'pivot-2-rows-cols.xlsx');
  const files = unzipSync(new Uint8Array(XLSX.write(XLSX.readFile(src), { type: 'array' })));
  const p = 'xl/pivotTables/pivotTable1.xml';
  files[p] = strToU8(strFromU8(files[p]).replace(/<rowItems[\s\S]*?<\/rowItems>/, '').replace(/<colItems[\s\S]*?<\/colItems>/, ''));
  writeFileSync(join(OUT, 'pivot-8-ab-no-items.xlsx'), zipSync(files));
  checks.push({ file: 'pivot-8-ab-no-items.xlsx', expect: 'Diagnostic twin of pivot-2 without rowItems/colItems. Only matters if pivot-2 fails: tell me whether THIS one opens cleanly.' });
}

// Pro features in one workbook
{
  const wb = utils.book_new();
  const ws = utils.aoa_to_sheet([
    ['Product', 'Region', 'Units', 'Price', 'Revenue', 'Status'],
    ['Chair', 'East', 12, 120, null, 'ok'],
    ['Desk', 'West', 4, 300, null, 'late'],
    ['Lamp', 'East', 30, 25, null, 'ok'],
    ['Shelf', 'North', 7, 90, null, 'ok'],
    [null, null, null, null, null, null],
  ]);
  for (let r = 2; r <= 5; r++) ws[`E${r}`] = { t: 'n', f: `C${r}*D${r}` };
  ws['!tables'] = [{ name: 'Sales', ref: 'A1:F6', totalsRow: true, style: 'TableStyleMedium2', columns: [{ totalsRowLabel: 'Total' }, {}, { totalsRowFunction: 'sum' }, { totalsRowFunction: 'average' }, { totalsRowFunction: 'sum' }, {}] }];
  utils.set_range_style(ws, 'D2:E6', { numFmt: '$#,##0.00' });
  ws['!condfmt'] = [
    { ref: 'C2:C5', rules: [{ type: 'dataBar', color: { rgb: '638EC6' } }] },
    { ref: 'F2:F5', rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"late"'], style: { font: { bold: true, color: { rgb: '9C0006' } }, fill: { fgColor: { rgb: 'FFC7CE' } } } }] },
  ];
  ws['!dataValidation'] = [{ ref: 'F2:F5', type: 'list', values: ['ok', 'late', 'cancelled'], prompt: { title: 'Status', message: 'Pick a status' } }];
  ws['!freeze'] = 'A2';
  ws['!cols'] = [{ wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 10 }];
  ws.A8 = { t: 's', v: 'Rich: ', richText: [{ text: 'bold', font: { bold: true } }, { text: ' and ' }, { text: 'red', font: { color: { rgb: 'FF0000' } } }] };
  ws.A9 = { t: 's', v: 'Link to example.com', l: { Target: 'https://example.com', Tooltip: 'opens browser' }, s: { font: { underline: true, color: { theme: 10 } } } };
  ws.A10 = { t: 's', v: 'Has a note', c: [{ a: 'node-xls', t: 'This is a note' }] };
  ws.A11 = { t: 'd', v: new Date(2026, 0, 15), z: 'dddd, mmmm d, yyyy' };
  ws['!ref'] = 'A1:F11';
  ws['!merges'] = [utils.decode_range('A11:D11')];
  ws['!images'] = [{ data: makePng(120, 60, (x, y) => [31 + x, 78 + y, 160, 255]), anchor: { from: 'H2' } }];
  ws['!charts'] = [{ type: 'column', title: 'Units by product', series: [{ name: 'Units', categories: 'A2:A5', values: 'C2:C5' }], anchor: { from: 'H7', width: 400, height: 240 } }];
  utils.book_append_sheet(wb, ws, 'Features');
  const prot = utils.aoa_to_sheet([['Locked sheet (password: test)'], ['editable ->', 'type here']]);
  prot.B2.s = { protection: { locked: false }, fill: { fgColor: { rgb: 'FFF2CC' } } };
  prot['!protect'] = { password: 'test' };
  utils.book_append_sheet(wb, prot, 'Protected');
  XLSX.writeFile(wb, join(OUT, 'features.xlsx'));
  checks.push({
    file: 'features.xlsx',
    expect:
      'Sheet "Features": styled table "Sales" with Total row (sum/average), Revenue formulas, data bars in Units, "late" highlighted red, Status dropdown with prompt, frozen header row, rich text in A8, link in A9, note on A10, long date in A11, picture at H2, column chart at H7. Sheet "Protected": only B2 editable (password "test").',
  });
}

// Charts dashboard (same as the showcase)
{
  const wb = XLSX.readFile(join('out', 'showcase', '12-charts.xlsx'));
  XLSX.writeFile(wb, join(OUT, 'charts.xlsx'));
  checks.push({ file: 'charts.xlsx', expect: 'Sheet "Dashboard": 8 charts (clustered column, stacked bar, line with markers, stacked area, pie with % labels, doughnut, scatter, 100% stacked column with value labels).' });
}

const md = `# Excel verification checklist

Generated by \`npm run excel-check\` on ${new Date().toISOString().slice(0, 10)}. Open each file in **Microsoft Excel** (desktop).

For every file, report:

1. **Did Excel show "We found a problem with some content…"?** If yes, click *Yes* to repair, then send me the text of the repair log (Excel shows a link to an XML log file listing the "Removed Part" / "Repaired Records").
2. **Does the sheet look like the description?**
3. For pivots: click inside the pivot, then *PivotTable Analyze → Refresh*. Does the result stay the same?

| File | What you should see | Opened cleanly? | Looks right? |
|---|---|---|---|
${checks.map((c) => `| \`${c.file}\` | ${c.expect} | ☐ | ☐ |`).join('\n')}

Notes
- Pivots are written with *refreshOnLoad*, so Excel rebuilds them from the "Data" sheet when the file opens; the values must match the pre-computed cells.
- \`pivot-8-ab-no-items.xlsx\` is a diagnostic variant; it only matters if \`pivot-2-rows-cols.xlsx\` fails.
`;
writeFileSync(join(OUT, 'CHECKLIST.md'), md);
console.log(`wrote ${checks.length} files + CHECKLIST.md to ${OUT}/`);
