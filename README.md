# node-xls

> Complete SheetJS (`xlsx`) compatible spreadsheet engine with first-class **Cell Styling** and **Pivot Tables**.

SheetJS (`xlsx`) on npm is standard for reading and manipulating Excel spreadsheets in JavaScript, but the open-source version lacks two critical features:
1. **No styling** (fonts, background colors, borders, alignments, and custom number formatting are stripped or ignored).
2. **No pivot tables** (creating, reading, and rendering pivot tables is not supported).

`node-xls` solves both problems while maintaining 100% API compatibility with `xlsx`.

---

## Features

- **SheetJS (`xlsx`) Compatible API**: Drop-in replacement for `XLSX.read`, `XLSX.readFile`, `XLSX.write`, `XLSX.writeFile`, and `XLSX.utils`.
- **First-Class Cell Styling**:
  - Fonts: name, size, bold, italic, underline, strikethrough, RGB/theme colors.
  - Fills: solid, pattern types, foreground and background colors.
  - Borders: top, bottom, left, right, diagonal (thin, medium, thick, double, dashed, dotted, hair, etc.).
  - Alignment: horizontal (left, center, right, justify), vertical (top, center, bottom), text wrapping, rotation, indent.
  - Number Formats: built-in and custom format strings (currency, percentages, dates, decimals).
- **Comprehensive Pivot Table Support**:
  - Interactive OpenXML Pivot Tables (`pivotTableDefinition`, `pivotCacheDefinition`, `pivotCacheRecords`).
  - Pre-computed rendered summary cells with professional formatting for instant display in all spreadsheet viewers.
  - Multi-level row and column groupings.
  - Multiple aggregations: `sum`, `count`, `average`, `min`, `max`, `product`, `countNums`, `stdDev`, `stdDevp`, `var`, `varp`.
  - Filters / Page fields.
  - Row and column grand totals and subtotals.
- **Synchronous & Asynchronous Support**:
  - Synchronous functions for exact SheetJS compatibility: `read`, `readFile`, `write`, `writeFile`.
  - Asynchronous alternatives: `readAsync`, `readFileAsync`, `writeAsync`, `writeFileAsync`.
- **TypeScript**: Written in strict TypeScript with complete type definitions.

---

## Installation

```bash
npm install node-xls
```

---

## Quick Start

```typescript
import XLSX, { utils } from 'node-xls';

// 1. Create a new workbook
const wb = utils.book_new();

// 2. Create sheet from JSON data
const data = [
  { Region: 'East', Category: 'Furniture', Sales: 1200, Profit: 300 },
  { Region: 'East', Category: 'Technology', Sales: 2500, Profit: 800 },
  { Region: 'West', Category: 'Furniture', Sales: 1800, Profit: 450 },
  { Region: 'West', Category: 'Technology', Sales: 3200, Profit: 950 },
  { Region: 'West', Category: 'Supplies', Sales: 600, Profit: 120 },
];
const ws = utils.json_to_sheet(data);

// 3. Apply cell styling (Header styling)
utils.set_range_style(ws, 'A1:D1', {
  font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: 'FFFFFF' } },
  fill: { patternType: 'solid', fgColor: { rgb: '1F497D' } },
  alignment: { horizontal: 'center', vertical: 'center' },
  border: {
    bottom: { style: 'medium', color: { rgb: '000000' } },
  },
});

// Format numeric columns
utils.set_range_style(ws, 'C2:D6', {
  numFmt: '$#,##0.00',
  alignment: { horizontal: 'right' },
});

utils.book_append_sheet(wb, ws, 'Sales Data');

// 4. Add a Pivot Table
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
  showRowTotals: true,
  showColTotals: true,
});

// 5. Write to file
XLSX.writeFile(wb, 'Sales_Report.xlsx');
```

---

## Cell Styling Guide

Styles are defined on `cell.s` or applied using helper functions `utils.set_cell_style` and `utils.set_range_style`.

### Style Properties

```typescript
interface CellStyle {
  font?: {
    name?: string;               // e.g. 'Arial', 'Calibri', 'Times New Roman'
    sz?: number;                 // font size in pt, e.g. 11, 14, 16
    bold?: boolean;
    italic?: boolean;
    underline?: boolean | 'single' | 'double';
    strike?: boolean;
    color?: { rgb: string };     // hex RGB string: 'FF0000' or '#FF0000'
  };
  fill?: {
    patternType?: 'solid' | 'none' | 'gray125' | 'darkGray' | 'lightGray';
    fgColor?: { rgb: string };   // background fill color in hex
    bgColor?: { rgb: string };
  };
  border?: {
    top?: { style: BorderStyle; color?: { rgb: string } };
    bottom?: { style: BorderStyle; color?: { rgb: string } };
    left?: { style: BorderStyle; color?: { rgb: string } };
    right?: { style: BorderStyle; color?: { rgb: string } };
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'justify';
    vertical?: 'top' | 'center' | 'bottom' | 'justify';
    wrapText?: boolean;
    textRotation?: number;       // -90 to 90 degrees
    indent?: number;
    shrinkToFit?: boolean;
  };
  numFmt?: string | number;       // e.g. '$#,##0.00', '0.00%', 'yyyy-mm-dd'
}
```

### Direct Cell Assignment

```typescript
ws['A1'].s = {
  font: { bold: true, color: { rgb: 'E74C3C' }, sz: 14 },
  fill: { fgColor: { rgb: 'F9EBEA' } },
  alignment: { horizontal: 'center' },
};
```

---

## Pivot Table Guide

`node-xls` provides both OpenXML pivot cache integration and pre-computed rendered summary grids.

### `utils.add_pivot_table(workbook, options)`

```typescript
utils.add_pivot_table(wb, {
  name: 'RegionalAnalysis',         // Pivot table name
  sourceSheet: 'RawData',           // Sheet containing source table
  sourceRange: 'A1:F500',           // Optional: default is sourceSheet['!ref']
  targetSheet: 'PivotReport',       // Sheet where pivot table will be placed
  targetCell: 'A3',                 // Top-left cell of the pivot report
  rows: ['Region', 'SubRegion'],    // Fields grouped on rows
  cols: ['Quarter'],                // Fields grouped on columns
  values: [                         // Aggregated fields
    { field: 'Revenue', subtotal: 'sum', name: 'Total Revenue' },
    { field: 'Units', subtotal: 'average', name: 'Avg Units' },
  ],
  filters: ['Department'],          // Report filter fields
  showRowTotals: true,              // Show row grand total
  showColTotals: true,              // Show column grand total
  styleName: 'PivotStyleLight16',   // Excel pivot style theme
  renderCells: true,                // Pre-compute formatted cells (default: true)
});
```

### Standalone Pivot Calculation

If you need the calculated pivot matrix without modifying a workbook:

```typescript
const result = utils.compute_pivot_table(dataArray, {
  rows: ['Department'],
  cols: ['Year'],
  values: [{ field: 'Expenses', subtotal: 'sum' }],
});

console.log(result.matrix);
console.log(result.rowTotals);
console.log(result.grandTotal);
```

---

## API Reference

### Core Functions

- `XLSX.readFile(filename, opts)`: Synchronously reads a file from disk.
- `XLSX.read(data, opts)`: Synchronously parses Buffer, Uint8Array, base64, or string.
- `XLSX.writeFile(wb, filename, opts)`: Synchronously writes workbook to disk.
- `XLSX.write(wb, opts)`: Synchronously generates XLSX buffer, binary, or base64.
- `XLSX.readFileAsync(filename, opts)`: Async file reading.
- `XLSX.readAsync(data, opts)`: Async data parsing.
- `XLSX.writeFileAsync(wb, filename, opts)`: Async file writing.
- `XLSX.writeAsync(wb, opts)`: Async data generation.

### `XLSX.utils`

- `book_new()`: Creates a new empty workbook.
- `book_append_sheet(wb, ws, name)`: Adds a sheet to workbook.
- `aoa_to_sheet(aoa, opts)`: Converts array of arrays to a worksheet.
- `json_to_sheet(json, opts)`: Converts array of objects to a worksheet.
- `sheet_to_json(ws, opts)`: Converts worksheet to JSON objects or 2D array (`header: 1`).
- `sheet_to_csv(ws, opts)`: Converts worksheet to CSV string.
- `sheet_to_txt(ws, opts)`: Converts worksheet to TSV string.
- `sheet_to_html(ws, opts)`: Converts worksheet to HTML table string with inline CSS styles.
- `sheet_to_formulae(ws)`: Generates list of formulas.
- `encode_cell(cellAddress)`: `{ c: 0, r: 0 }` -> `'A1'`.
- `decode_cell(cellRef)`: `'A1'` -> `{ c: 0, r: 0 }`.
- `encode_range(range)`: `{ s: { c: 0, r: 0 }, e: { c: 1, r: 1 } }` -> `'A1:B2'`.
- `decode_range(rangeStr)`: `'A1:B2'` -> range object.
- `set_cell_style(ws, cellRef, style)`: Applies styling to a cell.
- `set_range_style(ws, range, style)`: Applies styling to a range.
- `add_pivot_table(wb, options)`: Creates, pre-renders, and attaches a pivot table.
- `compute_pivot_table(data, options)`: Computes pivot table matrix and totals.

---

## License

MIT © [andrew.yohanes123@gmail.com](mailto:andrew.yohanes123@gmail.com)
