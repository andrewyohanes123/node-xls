import { describe, it, expect } from 'vitest';
import XLSX, { utils } from '../src/index.js';

describe('Pivot Table Engine & OpenXML Integration', () => {
  const salesData = [
    { Region: 'East', Category: 'Furniture', Sales: 100, Quantity: 2 },
    { Region: 'East', Category: 'Technology', Sales: 250, Quantity: 1 },
    { Region: 'West', Category: 'Furniture', Sales: 150, Quantity: 3 },
    { Region: 'West', Category: 'Technology', Sales: 300, Quantity: 2 },
    { Region: 'West', Category: 'Supplies', Sales: 50, Quantity: 5 },
  ];

  it('should compute pivot aggregations correctly', () => {
    const comp = utils.compute_pivot_table(salesData, {
      rows: ['Region'],
      cols: ['Category'],
      values: [
        { field: 'Sales', subtotal: 'sum' },
        { field: 'Quantity', subtotal: 'sum' },
      ],
    });

    expect(comp.rowKeys).toEqual([['East'], ['West']]);
    expect(comp.colKeys).toEqual([['Furniture'], ['Supplies'], ['Technology']]);

    // East / Furniture / Sum of Sales
    expect(comp.matrix['East']['Furniture']['Sum of Sales']).toBe(100);
    // West / Technology / Sum of Sales
    expect(comp.matrix['West']['Technology']['Sum of Sales']).toBe(300);
    // Grand total sales
    expect(comp.grandTotal['Sum of Sales']).toBe(850);
    // Grand total quantity
    expect(comp.grandTotal['Sum of Quantity']).toBe(13);
  });

  it('should create, render and write a workbook with pivot table', () => {
    const wb = utils.book_new();
    const dataSheet = utils.json_to_sheet(salesData);
    utils.book_append_sheet(wb, dataSheet, 'Data');

    const pivotDef = utils.add_pivot_table(wb, {
      name: 'SalesSummaryPivot',
      sourceSheet: 'Data',
      targetSheet: 'Summary',
      rows: ['Region'],
      cols: ['Category'],
      values: [{ field: 'Sales', subtotal: 'sum', name: 'Total Sales' }],
      showRowTotals: true,
      showColTotals: true,
    });

    expect(pivotDef.name).toBe('SalesSummaryPivot');
    expect(wb.SheetNames).toContain('Summary');
    expect(wb.PivotTables?.length).toBe(1);

    const summarySheet = wb.Sheets['Summary'];
    expect(summarySheet['!ref']).toBeDefined();

    // Verify rendered cells in summary sheet
    // Title
    expect(summarySheet['A3'].v).toBe('SalesSummaryPivot');
    // Header
    expect(summarySheet['A5'].v).toBe('Region');

    // Write to XLSX buffer
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);

    // Read back workbook
    const readWb = XLSX.read(buffer, { type: 'buffer' });
    expect(readWb.SheetNames).toContain('Data');
    expect(readWb.SheetNames).toContain('Summary');

    // Verify pivot table definition was preserved & parsed
    expect(readWb.PivotTables).toBeDefined();
    expect(readWb.PivotTables?.length).toBe(1);

    const parsedPivot = readWb.PivotTables![0];
    expect(parsedPivot.name).toBe('SalesSummaryPivot');
    expect(parsedPivot.sourceSheet).toBe('Data');
    expect(parsedPivot.rows).toEqual(['Region']);
    expect(parsedPivot.cols).toEqual(['Category']);
    expect(parsedPivot.values[0].field).toBe('Sales');
    expect(parsedPivot.values[0].subtotal).toBe('sum');

    // Verify summary sheet cells still have data
    const readSummary = readWb.Sheets['Summary'];
    expect(readSummary['A3'].v).toBe('SalesSummaryPivot');
  });

  it('should support multiple aggregations: average, min, max, count', () => {
    const comp = utils.compute_pivot_table(salesData, {
      rows: ['Region'],
      values: [
        { field: 'Sales', subtotal: 'average', name: 'Avg Sales' },
        { field: 'Sales', subtotal: 'min', name: 'Min Sales' },
        { field: 'Sales', subtotal: 'max', name: 'Max Sales' },
        { field: 'Sales', subtotal: 'count', name: 'Count Sales' },
      ],
    });

    // East: 100, 250 -> Avg = 175, Min = 100, Max = 250, Count = 2
    expect(comp.matrix['East']['']['Avg Sales']).toBe(175);
    expect(comp.matrix['East']['']['Min Sales']).toBe(100);
    expect(comp.matrix['East']['']['Max Sales']).toBe(250);
    expect(comp.matrix['East']['']['Count Sales']).toBe(2);
  });
});
