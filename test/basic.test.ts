import { describe, it, expect } from 'vitest';
import XLSX, { utils } from '../src/index.js';

describe('Basic SheetJS Compatibility', () => {
  it('should encode and decode cell addresses', () => {
    expect(utils.encode_cell({ c: 0, r: 0 })).toBe('A1');
    expect(utils.encode_cell({ c: 25, r: 0 })).toBe('Z1');
    expect(utils.encode_cell({ c: 26, r: 9 })).toBe('AA10');

    expect(utils.decode_cell('A1')).toEqual({ c: 0, r: 0 });
    expect(utils.decode_cell('Z1')).toEqual({ c: 25, r: 0 });
    expect(utils.decode_cell('AA10')).toEqual({ c: 26, r: 9 });
    expect(utils.decode_cell('$B$5')).toEqual({ c: 1, r: 4 });
  });

  it('should encode and decode ranges', () => {
    const range = { s: { c: 0, r: 0 }, e: { c: 3, r: 10 } };
    expect(utils.encode_range(range)).toBe('A1:D11');
    expect(utils.decode_range('A1:D11')).toEqual(range);
  });

  it('should convert aoa to sheet and sheet to aoa/json', () => {
    const data = [
      ['Name', 'Age', 'Score'],
      ['Alice', 28, 95.5],
      ['Bob', 34, 88.0],
    ];

    const ws = utils.aoa_to_sheet(data);
    expect(ws['!ref']).toBe('A1:C3');
    expect(ws['A1'].v).toBe('Name');
    expect(ws['B2'].v).toBe(28);
    expect(ws['C2'].v).toBe(95.5);

    const json = utils.sheet_to_json(ws);
    expect(json).toEqual([
      { Name: 'Alice', Age: 28, Score: 95.5 },
      { Name: 'Bob', Age: 34, Score: 88.0 },
    ]);

    const aoaBack = utils.sheet_to_json(ws, { header: 1 });
    expect(aoaBack).toEqual(data);
  });

  it('should convert json to sheet', () => {
    const data = [
      { Product: 'Widget', Price: 19.99, Active: true },
      { Product: 'Gadget', Price: 29.99, Active: false },
    ];

    const ws = utils.json_to_sheet(data);
    expect(ws['!ref']).toBe('A1:C3');
    expect(ws['A1'].v).toBe('Product');
    expect(ws['B2'].v).toBe(19.99);
    expect(ws['C2'].v).toBe(true);

    const json = utils.sheet_to_json(ws);
    expect(json).toEqual(data);
  });

  it('should convert sheet to csv and txt', () => {
    const data = [
      ['City', 'Country', 'Population'],
      ['Tokyo', 'Japan', 14000000],
      ['Paris, France', 'France', 2161000],
    ];

    const ws = utils.aoa_to_sheet(data);
    const csv = utils.sheet_to_csv(ws);
    expect(csv).toContain('City,Country,Population');
    expect(csv).toContain('"Paris, France",France,2161000');

    const txt = utils.sheet_to_txt(ws);
    expect(txt).toContain('City\tCountry\tPopulation');
  });

  it('should convert sheet to html', () => {
    const data = [
      ['A', 'B'],
      [1, 2],
    ];
    const ws = utils.aoa_to_sheet(data);
    const html = utils.sheet_to_html(ws);
    expect(html).toContain('<table>');
    expect(html).toContain('<td id="A1">A</td>');
    expect(html).toContain('<td id="B2">2</td>');
  });

  it('should support formulas in sheet', () => {
    const ws = utils.aoa_to_sheet([
      [10, 20],
      [null, null],
    ]);
    ws['A2'] = { t: 'n', f: 'SUM(A1:B1)', v: 30 };
    const formulae = utils.sheet_to_formulae(ws);
    expect(formulae).toContain('A2=SUM(A1:B1)');
  });
});
