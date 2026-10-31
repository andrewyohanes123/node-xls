/**
 * Sheet layout and workbook structure: merges, column/row sizing, outlines,
 * freeze panes, views, tab colors, filters, print settings, sheet visibility,
 * defined names and document properties.
 */
import { describe, it, expect } from 'vitest';
import XLSX, { utils } from '../src/index.js';
import { book, roundtrip, roundtripSheet } from './helpers/util.js';

const grid = () =>
  utils.aoa_to_sheet([
    ['h1', 'h2', 'h3', 'h4'],
    [1, 2, 3, 4],
    [5, 6, 7, 8],
  ]);

describe('merges', () => {
  it('round-trip and are seen by SheetJS and ExcelJS', async () => {
    const ws = grid();
    ws['!merges'] = [utils.decode_range('A1:B1'), utils.decode_range('C2:D3')];
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back['!merges']).toEqual(ws['!merges']);
    expect(rep.sheetjs.Sheets.S['!merges']).toEqual(ws['!merges']);
    expect((rep.exceljs.getWorksheet('S') as any).model.merges.sort()).toEqual(['A1:B1', 'C2:D3']);
  });

  it('overlapping merges are rejected', () => {
    const ws = grid();
    ws['!merges'] = [utils.decode_range('A1:B2'), utils.decode_range('B2:C3')];
    expect(() => XLSX.write(book(['S', ws]), { type: 'array' })).toThrow(/Overlapping/);
  });
});

describe('columns and rows', () => {
  it('widths in wch / wpx / width, hidden and outline levels', async () => {
    const ws = grid();
    ws['!cols'] = [{ wch: 20 }, { wpx: 100 }, { width: 15.5, level: 1 }, { hidden: true, level: 2 }];
    const { ws: back, rep } = await roundtripSheet(ws);
    const ex = rep.exceljs.getWorksheet('S')!;
    expect(back['!cols']![0].wch).toBe(20);
    expect(back['!cols']![1].wpx).toBe(100);
    expect(back['!cols']![2]).toMatchObject({ width: 15.5, level: 1 });
    expect(back['!cols']![3]).toMatchObject({ hidden: true, level: 2 });
    expect(ex.getColumn(4).hidden).toBe(true);
    expect(ex.getColumn(3).outlineLevel).toBe(1);
    expect(rep.xml('xl/worksheets/sheet1.xml')).toContain('outlineLevelCol="2"');
    const sj = rep.sheetjs.Sheets.S['!cols']!;
    expect(sj[0].wch).toBe(20);
    expect(sj[1].wpx).toBe(100);
  });

  it('row heights (pt and px), hidden rows, outline levels; empty rows with properties are kept', async () => {
    const ws = grid();
    ws['!rows'] = [{ hpt: 30 }, { hpx: 40 }, { hidden: true, level: 1 }];
    ws['!rows'][9] = { hpt: 50 };
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back['!rows']![0].hpt).toBe(30);
    expect(back['!rows']![1].hpx).toBe(40);
    expect(back['!rows']![2]).toMatchObject({ hidden: true, level: 1 });
    expect(back['!rows']![9].hpt).toBe(50);
    const ex = rep.exceljs.getWorksheet('S')!;
    expect(ex.getRow(1).height).toBe(30);
    expect(ex.getRow(3).hidden).toBe(true);
    expect(rep.sheetjs.Sheets.S['!rows']![0].hpt).toBe(30);
  });
});

describe('sheet views', () => {
  const cases: Array<[any, any]> = [
    ['B2', { xSplit: 1, ySplit: 1, topLeftCell: 'B2' }],
    ['A2', { xSplit: 0, ySplit: 1, topLeftCell: 'A2' }],
    ['C1', { xSplit: 2, ySplit: 0, topLeftCell: 'C1' }],
    [{ ySplit: 3 }, { xSplit: 0, ySplit: 3, topLeftCell: 'A4' }],
  ];
  for (const [input, expected] of cases) {
    it(`freeze panes ${JSON.stringify(input)}`, async () => {
      const ws = grid();
      ws['!freeze'] = input;
      const { ws: back, rep } = await roundtripSheet(ws);
      expect(back['!freeze']).toEqual(expected);
      const view = rep.exceljs.getWorksheet('S')!.views[0] as any;
      expect(view).toMatchObject({ state: 'frozen', xSplit: expected.xSplit, ySplit: expected.ySplit });
    });
  }

  it('gridlines, zoom, right-to-left, tab color', async () => {
    const ws = grid();
    ws['!views'] = { showGridLines: false, zoomScale: 150, rightToLeft: true };
    ws['!tabColor'] = { rgb: '00B050' };
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back['!views']).toEqual({ showGridLines: false, zoomScale: 150, rightToLeft: true });
    expect(back['!tabColor']).toEqual({ rgb: '00B050' });
    const ex = rep.exceljs.getWorksheet('S')!;
    expect(ex.views[0]).toMatchObject({ showGridLines: false, zoomScale: 150, rightToLeft: true });
    expect(ex.properties.tabColor).toEqual({ argb: 'FF00B050' });
  });
});

describe('autofilter', () => {
  it('writes the filter and the hidden _FilterDatabase name Excel expects', async () => {
    const ws = grid();
    ws['!autofilter'] = { ref: 'A1:D3' };
    const { ws: back, rep, back: wb } = await roundtripSheet(ws);
    expect(back['!autofilter']).toEqual({ ref: 'A1:D3' });
    expect(rep.xml('xl/workbook.xml')).toContain(
      '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">S!$A$1:$D$3</definedName>'
    );
    expect(wb.Workbook?.Names).toContainEqual({ Name: '_xlnm._FilterDatabase', Ref: 'S!$A$1:$D$3', Sheet: 0, Hidden: true });
    expect(rep.exceljs.getWorksheet('S')!.autoFilter).toBe('A1:D3');
  });
});

describe('print settings', () => {
  it('margins and page setup', async () => {
    const ws = grid();
    ws['!margins'] = { left: 0.25, right: 0.25, top: 1, bottom: 1 };
    ws['!pageSetup'] = { orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0 };
    const { ws: back, rep } = await roundtripSheet(ws);
    expect(back['!margins']).toEqual({ left: 0.25, right: 0.25, top: 1, bottom: 1, header: 0.3, footer: 0.3 });
    expect(back['!pageSetup']).toEqual({ orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0 });
    expect(rep.exceljs.getWorksheet('S')!.pageSetup).toMatchObject({ orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1 });
    expect(rep.sheetjs.Sheets.S['!margins']).toMatchObject({ left: 0.25, top: 1 });
  });
});

describe('workbook structure', () => {
  it('hidden and very hidden sheets; active tab moves to a visible sheet', async () => {
    const wb = book(['A', grid()], ['B', grid()], ['C', grid()]);
    utils.book_set_sheet_visibility(wb, 'A', 1);
    utils.book_set_sheet_visibility(wb, 'C', 2);
    const { back, rep } = await roundtrip(wb);
    expect(back.Workbook?.Sheets?.map((s) => s.Hidden)).toEqual([1, 0, 2]);
    expect(rep.exceljs.worksheets.map((w) => w.state)).toEqual(['hidden', 'visible', 'veryHidden']);
    expect(rep.xml('xl/worksheets/sheet2.xml')).toContain('tabSelected="1"');
    expect(back.Workbook?.Views?.[0].activeTab).toBe(1);
  });

  it('all sheets hidden is rejected', () => {
    const wb = book(['A', grid()]);
    utils.book_set_sheet_visibility(wb, 'A', 1);
    expect(() => XLSX.write(wb, { type: 'array' })).toThrow(/visible/);
  });

  it('explicit active tab', async () => {
    const wb = book(['A', grid()], ['B', grid()]);
    wb.Workbook = { ...wb.Workbook, Views: [{ activeTab: 1 }] };
    const { rep } = await roundtrip(wb);
    expect(rep.xml('xl/workbook.xml')).toContain('activeTab="1"');
    expect(rep.xml('xl/worksheets/sheet1.xml')).not.toContain('tabSelected');
  });

  it('defined names (global and sheet-scoped)', async () => {
    const wb = book(['Data', grid()], ['Other', grid()]);
    wb.Workbook = {
      Names: [
        { Name: 'AllData', Ref: 'Data!$A$1:$D$3' },
        { Name: 'Local', Ref: "'Other'!$B$2", Sheet: 1, Comment: 'scoped' },
        { Name: 'Rate', Ref: '=0.07' },
      ],
    };
    const { back, rep } = await roundtrip(wb);
    expect(back.Workbook?.Names).toEqual([
      { Name: 'AllData', Ref: 'Data!$A$1:$D$3' },
      { Name: 'Local', Ref: "'Other'!$B$2", Sheet: 1, Comment: 'scoped' },
      { Name: 'Rate', Ref: '0.07' },
    ]);
    expect(rep.sheetjs.Workbook?.Names?.map((n: any) => n.Name)).toEqual(['AllData', 'Local', 'Rate']);
  });

  it('document properties', async () => {
    const wb = book(['S', grid()]);
    wb.Props = {
      Title: 'Quarterly <Report>',
      Subject: 'Sales',
      Author: 'Ann',
      Keywords: 'q1 sales',
      Comments: 'notes',
      Category: 'finance',
      Company: 'ACME & Co',
      Manager: 'Bo',
      CreatedDate: new Date('2026-01-02T03:04:05Z'),
    };
    const { back, rep } = await roundtrip(wb);
    expect(back.Props).toMatchObject({
      Title: 'Quarterly <Report>',
      Subject: 'Sales',
      Author: 'Ann',
      Keywords: 'q1 sales',
      Comments: 'notes',
      Category: 'finance',
      Company: 'ACME & Co',
      Manager: 'Bo',
      LastAuthor: 'Ann',
    });
    expect((back.Props!.CreatedDate as Date).toISOString()).toBe('2026-01-02T03:04:05.000Z');
    expect(rep.sheetjs.Props).toMatchObject({ Title: 'Quarterly <Report>', Author: 'Ann', Company: 'ACME & Co' });
    expect(rep.exceljs.creator).toBe('Ann');
  });

  it('reads only selected sheets and limits rows', async () => {
    const big = utils.aoa_to_sheet(Array.from({ length: 50 }, (_, i) => [i]));
    const buf = XLSX.write(book(['A', big], ['B', grid()]), { type: 'array' });
    const onlyB = XLSX.read(buf, { sheets: 'B' });
    expect(Object.keys(onlyB.Sheets.A)).toEqual([]);
    expect(onlyB.Sheets.B.A1.v).toBe('h1');
    const limited = XLSX.read(buf, { sheetRows: 10 });
    expect(limited.Sheets.A['!ref']).toBe('A1:A10');
    expect(limited.Sheets.A.A11).toBeUndefined();
  });
});
