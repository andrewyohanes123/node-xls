# node-xls

Read and write `.xlsx` files with a [SheetJS](https://www.npmjs.com/package/xlsx)-compatible API, plus the features SheetJS only offers in its paid **Pro** editions:

| Feature | SheetJS CE | SheetJS Pro | node-xls |
|---|:-:|:-:|:-:|
| Values, formulas, merges, column widths, CSV/HTML/JSON | ✓ | ✓ | ✓ |
| Cell styles: fonts, fills, borders, alignment, number formats, protection | – | Basic | ✓ |
| Rich text (several fonts in one cell) | – | Basic | ✓ |
| Row / column styles, freeze panes, tab colors, sheet views | – | Basic | ✓ |
| Conditional formatting (highlight rules, color scales, data bars, icon sets) | – | Basic | ✓ |
| Data validation (dropdowns, ranges, prompts, alerts) | – | Basic | ✓ |
| Excel tables (styled ListObjects with totals rows) | – | Basic | ✓ |
| Styled HTML export / import | – | Basic | ✓ |
| Images (PNG / JPEG / GIF) | – | Image | ✓ |
| Native charts (column, bar, line, area, pie, doughnut, scatter) | – | Chart | ✓ |
| Pivot tables | – | add-on | ✓ |
| Hyperlinks, comments, defined names, hidden sheets, sheet/workbook protection | partly | ✓ | ✓ |

> **Status: 0.1 (pre-release).** Every generated file is validated in the test suite against the ECMA-376 schemas, and read back by SheetJS CE, ExcelJS and LibreOffice. Pivot tables and charts are still waiting on verification in desktop Microsoft Excel; see [Verifying output](#verifying-output).

---

## Install

```bash
npm install node-xls
```

Node.js 18 or newer. node-xls works with CommonJS (`require('node-xls')`) and ESM (`import XLSX, { utils } from 'node-xls'`).

## Quick start

```ts
import XLSX, { utils } from 'node-xls';

const ws = utils.json_to_sheet([
  { Region: 'East', Category: 'Furniture', Sales: 1200, Profit: 300 },
  { Region: 'East', Category: 'Technology', Sales: 2500, Profit: 800 },
  { Region: 'West', Category: 'Furniture', Sales: 1800, Profit: 450 },
  { Region: 'West', Category: 'Technology', Sales: 3200, Profit: 950 },
]);

// header style + currency columns
utils.set_range_style(ws, 'A1:D1', {
  font: { bold: true, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: '1F497D' } },
  alignment: { horizontal: 'center' },
});
utils.set_range_style(ws, 'C2:D5', { numFmt: '$#,##0.00' });
ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
ws['!freeze'] = 'A2';

const wb = utils.book_new(ws, 'Sales');
utils.add_pivot_table(wb, {
  sourceSheet: 'Sales',
  targetSheet: 'Summary',
  rows: ['Region'],
  cols: ['Category'],
  values: [{ field: 'Sales', name: 'Total Sales', numFmt: '#,##0' }],
});

XLSX.writeFile(wb, 'report.xlsx');
```

---

## Reading and writing

```ts
const wb = XLSX.readFile('in.xlsx', { cellDates: true });
XLSX.writeFile(wb, 'out.xlsx');                       // format from extension: .xlsx .csv .txt/.tsv .html .json

const buf = XLSX.write(wb, { type: 'buffer' });       // 'buffer' | 'array' | 'base64' | 'binary' | 'string'
const wb2 = XLSX.read(buf);                            // Buffer, Uint8Array, ArrayBuffer, base64; CSV/HTML text too

await XLSX.writeFileAsync(wb, 'out.xlsx');            // async variants: readAsync, readFileAsync, writeAsync
```

| Read option | Default | Effect |
|---|---|---|
| `cellDates` | `false` | Date-formatted cells become `{ t: 'd', v: Date }` |
| `cellStyles` | `true` | Populate `cell.s` |
| `cellFormula` | `true` | Populate `cell.f` / `cell.F` (shared formulas are expanded) |
| `cellNF` / `cellText` | `true` | Populate `cell.z` (number format) / `cell.w` (formatted text) |
| `sheetRows` | – | Only read the first N rows |
| `sheets` | – | Only parse these sheets (names or indexes) |
| `raw`, `FS` | – | CSV: keep text as strings / field separator |

Write options: `bookType` (`xlsx` `csv` `txt` `html` `json`), `type`, `compression` (default on), `sheet` (for single-sheet formats), `Props`.

### The data model

Workbooks, sheets and cells use the SheetJS shapes. Sheet keys starting with `!` hold sheet-level settings.

```ts
ws.A1 = { t: 's', v: 'text' };                 // t: s n b d e z
ws.B1 = { t: 'n', v: 0.25, z: '0.0%' };        // z: number format (SheetJS convention)
ws.C1 = { t: 'n', f: 'A2*B2' };                // formula without '='; Excel recalculates on open
ws.D1 = { t: 'd', v: new Date(2026, 0, 15) };  // written as a date serial + date format
```

**Dates:** like SheetJS, a JS `Date` is read by its **local** calendar fields. So `new Date(2026, 0, 15)` is 15 Jan 2026 in every timezone. Set `UTC: true` in `aoa_to_sheet`/`json_to_sheet` to use the UTC fields instead.

---

## Styling

Set `cell.s` directly or merge styles in with `utils.set_cell_style(ws, 'A1', style)` / `utils.set_range_style(ws, 'A1:D1', style)`. Nested objects are merged, so a second call adds properties instead of replacing them.

```ts
interface CellStyle {
  font?: {
    name?: string; sz?: number;               // 'Arial', 14
    bold?: boolean; italic?: boolean; strike?: boolean;
    underline?: boolean | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
    vertAlign?: 'superscript' | 'subscript';
    color?: Color;
  };
  fill?: {
    patternType?: 'solid' | 'darkGray' | 'lightGrid' | ...;  // 18 Excel patterns; default solid when a color is set
    fgColor?: Color;                          // the cell color for solid fills
    bgColor?: Color;
  };
  border?: {
    top?: { style: BorderStyle; color?: Color };   // also bottom, left, right, diagonal
    diagonalUp?: boolean; diagonalDown?: boolean;
  };                                          // BorderStyle: thin medium thick double dashed dotted hair mediumDashed dashDot ...
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'justify' | 'fill' | 'centerContinuous' | 'distributed';
    vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
    wrapText?: boolean; shrinkToFit?: boolean; indent?: number;
    textRotation?: number;                    // -90..90, or 255 for vertical text
  };
  numFmt?: string | number;                   // '0.00%', '$#,##0', 'yyyy-mm-dd' or a built-in id (14)
  protection?: { locked?: boolean; hidden?: boolean };
}

type Color = string | { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean };
// rgb: 'FF0000', '#FF0000', 'F00' or ARGB '80FF0000'; theme: 0-11 (4-9 = accent1-6); tint: -1..1
```

**Rich text:** use `cell.richText` to put several fonts in one cell.

```ts
ws.A1 = { t: 's', richText: [{ text: 'Bold', font: { bold: true } }, { text: ' and ' }, { text: 'red', font: { color: { rgb: 'C00000' } } }] };
```

**Row and column styles** apply to the empty cells of a whole row or column:

```ts
ws['!cols'] = [{ wch: 20, s: { fill: { fgColor: { rgb: 'DDEBF7' } } } }];
ws['!rows'] = [{ hpt: 24, s: { font: { bold: true } } }];
```

## Sheet layout

```ts
ws['!merges'] = [utils.decode_range('A1:D1')];
ws['!cols'] = [{ wch: 20 }, { wpx: 120 }, { width: 15, hidden: true }, { level: 1 }];   // level = outline group
ws['!rows'] = [{ hpt: 30 }, { hpx: 40, hidden: true, level: 1 }];
ws['!freeze'] = 'B2';                              // or { xSplit: 1, ySplit: 1 }
ws['!views'] = { showGridLines: false, zoomScale: 120, rightToLeft: false };
ws['!tabColor'] = { theme: 5 };
ws['!autofilter'] = { ref: 'A1:D20' };
ws['!margins'] = { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 };
ws['!pageSetup'] = { orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0 };
```

Workbook level: `wb.Workbook.Names` (defined names), `utils.book_set_sheet_visibility(wb, 'Sheet', 1 | 2)` (hidden / very hidden), `wb.Workbook.Views = [{ activeTab: 1 }]`, `wb.Props` (Title, Author, Company, …).

## Hyperlinks and comments

```ts
ws.A1.l = { Target: 'https://example.com', Tooltip: 'Open' };   // also mailto:, file paths
ws.A2.l = { Target: "#'Other Sheet'!B2" };                       // internal link
utils.cell_add_comment(ws.A3, 'Check this value', 'Ann');        // ws.A3.c.hidden = false keeps the note visible
```

## Conditional formatting

```ts
ws['!condfmt'] = [
  { ref: 'B2:B100', rules: [
    { type: 'cellIs', operator: 'greaterThan', formulae: [1000], style: { font: { bold: true, color: { rgb: '9C0006' } }, fill: { fgColor: { rgb: 'FFC7CE' } } } },
    { type: 'expression', formula: '$C2="late"', style: { font: { italic: true } } },
  ]},
  { ref: 'C2:C100', rules: [{ type: 'dataBar', color: { rgb: '638EC6' } }] },
  { ref: 'D2:D100', rules: [{ type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], colors: [{ rgb: 'F8696B' }, { rgb: '63BE7B' }] }] },
  { ref: 'E2:E100', rules: [{ type: 'iconSet', iconSet: '3TrafficLights1' }] },
];
```

Rule types: `cellIs`, `expression`, `containsText`, `notContainsText`, `beginsWith`, `endsWith`, `top10`, `aboveAverage`, `duplicateValues`, `uniqueValues`, `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors`, `timePeriod`, `colorScale`, `dataBar`, `iconSet`.

## Data validation

```ts
ws['!dataValidation'] = [
  { ref: 'C2:C100', type: 'list', values: ['Open', 'Closed'], prompt: { title: 'Status', message: 'Pick one' } },
  { ref: 'D2:D100', type: 'list', formula1: "'Lists'!$A$1:$A$20" },
  { ref: 'E2:E100', type: 'whole', operator: 'between', formula1: 1, formula2: 100, error: { message: '1 to 100', style: 'stop' } },
  { ref: 'F2:F100', type: 'date', operator: 'greaterThan', formula1: new Date(2026, 0, 1) },
];
```

Types: `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`, `any`. Inline lists are limited to 255 characters (an Excel rule). For longer lists, use a range.

## Excel tables

```ts
ws['!tables'] = [{
  name: 'Sales', ref: 'A1:D20', style: 'TableStyleMedium2', totalsRow: true,
  columns: [{ totalsRowLabel: 'Total' }, {}, { totalsRowFunction: 'sum' }, { totalsRowFunction: 'average' }],
}];
```

Header cells are written as the column names, and totals cells get `SUBTOTAL()` formulas. Table names must be unique and must not look like cell references.

## Protection

```ts
ws['!protect'] = { password: 'secret', formatCells: false, sort: false };   // flags follow OOXML: true = locked
ws.B2.s = { protection: { locked: false } };                               // editable cell
wb.Workbook = { Protection: { password: 'secret' } };                      // lock sheet structure
```

Passwords are hashed with SHA-512, as current Excel versions do. This is an editing restriction, not encryption.

## Images

```ts
ws['!images'] = [
  { data: fs.readFileSync('logo.png'), anchor: { from: 'B2' } },                    // natural size
  { data: jpegBuffer, anchor: { from: 'E2', width: 200, height: 120 } },           // pixels
  { data: 'data:image/gif;base64,R0lGOD...', anchor: { from: 'B10', to: 'D20' } }, // stretched over cells
];
```

PNG, JPEG and GIF are supported. Identical images are stored once.

## Charts

```ts
ws['!charts'] = [{
  type: 'column',                     // column bar line area pie doughnut scatter
  grouping: 'stacked',                // clustered (default) stacked percentStacked
  title: 'Sales by month',
  series: [
    { name: "'Data'!$B$1", categories: "'Data'!$A$2:$A$13", values: "'Data'!$B$2:$B$13" },
    { name: 'Target', categories: 'Data!A2:A13', values: 'Data!C2:C13', color: { rgb: 'ED7D31' } },
  ],
  legend: 'b', dataLabels: false,
  yAxis: { title: 'Units', min: 0, numFmt: '#,##0' },
  anchor: { from: 'E2', width: 480, height: 288 },
}];
```

Series point at worksheet ranges, and their current values are embedded so the chart displays before any recalculation.

## Pivot tables

```ts
utils.add_pivot_table(wb, {
  name: 'RegionalSales',
  sourceSheet: 'Data',                 // first row = field names
  targetSheet: 'Pivot',                // created if missing
  targetCell: 'A3',
  title: 'Sales by region',            // optional caption above the pivot
  rows: ['Region', 'Category'],        // outer -> inner, with subtotals
  cols: ['Year'],
  values: [
    { field: 'Sales', name: 'Revenue', numFmt: '$#,##0' },
    { field: 'Units', name: 'Avg Units', subtotal: 'average' },
  ],
  filters: [{ field: 'Channel', selected: 'Online' }, 'Country'],   // report filters
  showRowTotals: true, showColTotals: true, showSubtotals: true,
  styleName: 'PivotStyleLight16',
});
```

What this produces:
- **A real Excel pivot table** (pivot cache and definition) in tabular form. It is marked refresh-on-open, so Excel rebuilds it from the source data.
- **The same summary as ordinary cells**, so viewers that don't process pivots (and anything reading the file programmatically) also see the result.

Summary functions: `sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevp`, `var`, `varp`.

For the numbers without a workbook, use `utils.compute_pivot_table(rows, { rows, cols, values, filters })`.

## Other formats and utilities

All SheetJS CE `utils` are available: `aoa_to_sheet`, `json_to_sheet`, `sheet_add_aoa` / `sheet_add_json` (`origin: -1` appends), `sheet_to_json` (`header: 1 | 'A' | string[]`, `range`, `defval`, `raw`, `rawNumbers`, `blankrows`, `skipHidden`), `sheet_to_csv`, `sheet_to_txt`, `sheet_to_html`, `sheet_to_formulae`, `table_to_sheet`, `table_to_book`, `format_cell`, `encode_*`/`decode_*`, `split_cell`, `book_new`, `book_append_sheet`, `book_set_sheet_visibility`, `sheet_get_cell`, `cell_set_number_format`, `cell_set_hyperlink`, `cell_set_internal_link`, `cell_add_comment`, `sheet_set_array_formula`. `XLSX.SSF` is the SheetJS number-format engine.

`sheet_to_html` writes cell styles as inline CSS and merges as `colspan`/`rowspan`. `table_to_sheet` reads them back.

## Differences from SheetJS CE

- **Output formats:** only `.xlsx` (plus CSV/TSV/HTML/JSON) is written. Legacy formats (`.xls`, `.xlsb`, `.ods`, …) are rejected instead of being written with the wrong content.
- **`t: 'd'` cells:** these hold real local `Date` objects. SheetJS stores UTC-shifted dates internally; `sheet_to_json` output matches either way.
- **Sheet names:** duplicates are rejected case-insensitively, as Excel does.
- **Defaults:** styles, formulas and number formats are read by default, and every cell gets a formatted `w`.
- **Not supported yet:** dense sheet mode, streaming, file encryption, VBA, formula calculation, threaded comments, sparklines, combo charts, and round-trip preservation of parts node-xls does not model (SheetJS Pro "Edit").

---

## Verifying output

```bash
npm run showcase       # one workbook per feature area -> out/showcase/
npm run render         # LibreOffice -> PDF -> PNG in out/render/ (also a smoke test)
npm run inspect -- out/showcase/02-styles.xlsx   # list parts, pretty-print XML to out/inspect/
npm run excel-check    # focused files + CHECKLIST.md for manual testing in Microsoft Excel
```

## Development

```bash
npm install
npm run schemas        # download the ECMA-376 schemas used by the tests (once)
npm test               # vitest; every written file is schema-validated and read by SheetJS CE + ExcelJS
npm run test:tz        # the suite under several timezones
npm run build
```

See [CLAUDE.md](CLAUDE.md) for the architecture and contributor rules.

## License

MIT
