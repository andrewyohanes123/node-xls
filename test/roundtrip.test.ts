import { describe, it, expect, afterAll } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import XLSX, { utils } from '../src/index.js';

describe('File System Operations & Multi-format Roundtrip', () => {
  const tmpDir = path.join(__dirname, '../tmp-test');

  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  afterAll(() => {
    if (fs.existsSync(tmpDir)) {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it('should write and read files on disk with styles and pivots', () => {
    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet([
      ['Department', 'Quarter', 'Budget'],
      ['Engineering', 'Q1', 50000],
      ['Engineering', 'Q2', 60000],
      ['Marketing', 'Q1', 30000],
      ['Marketing', 'Q2', 35000],
    ]);

    utils.set_cell_style(ws, 'A1', {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '002060' } },
    });

    utils.book_append_sheet(wb, ws, 'Budget');

    utils.add_pivot_table(wb, {
      name: 'BudgetPivot',
      sourceSheet: 'Budget',
      rows: ['Department'],
      cols: ['Quarter'],
      values: [{ field: 'Budget', subtotal: 'sum' }],
    });

    const filePath = path.join(tmpDir, 'budget_test.xlsx');
    XLSX.writeFile(wb, filePath);
    expect(fs.existsSync(filePath)).toBe(true);

    const readWb = XLSX.readFile(filePath);
    expect(readWb.SheetNames).toContain('Budget');
    expect(readWb.SheetNames).toContain('Budget Pivot');

    const readWs = readWb.Sheets['Budget'];
    expect(readWs['A1'].s?.font?.bold).toBe(true);
    expect(readWs['A1'].s?.font?.color).toEqual({ rgb: 'FFFFFF' });
    expect(readWs['A1'].s?.fill?.fgColor).toEqual({ rgb: '002060' });

    expect(readWb.PivotTables?.length).toBe(1);
    expect(readWb.PivotTables![0].name).toBe('BudgetPivot');
  });

  it('should read and write CSV files', () => {
    const csvContent = 'ID,Name,Active\n1,Alice,true\n2,Bob,false\n';
    const csvPath = path.join(tmpDir, 'test.csv');
    fs.writeFileSync(csvPath, csvContent, 'utf-8');

    const wb = XLSX.readFile(csvPath);
    expect(wb.SheetNames[0]).toBe('Sheet1');
    const ws = wb.Sheets['Sheet1'];
    expect(ws['A1'].v).toBe('ID');
    expect(ws['B2'].v).toBe('Alice');
    expect(ws['C2'].v).toBe(true);

    const outCsvPath = path.join(tmpDir, 'out.csv');
    XLSX.writeFile(wb, outCsvPath, { bookType: 'csv' });
    const readOut = fs.readFileSync(outCsvPath, 'utf-8');
    expect(readOut).toContain('ID,Name,Active');
    expect(readOut).toContain('1,Alice,true');
  });

  it('should support async read and write', async () => {
    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet([['Async', 'Test'], [1, 2]]);
    utils.book_append_sheet(wb, ws, 'AsyncSheet');

    const buf = await XLSX.writeAsync(wb, { type: 'buffer' });
    expect(buf).toBeInstanceOf(Buffer);

    const readWb = await XLSX.readAsync(buf);
    expect(readWb.SheetNames).toContain('AsyncSheet');
    expect(readWb.Sheets['AsyncSheet']['A1'].v).toBe('Async');
  });
});
