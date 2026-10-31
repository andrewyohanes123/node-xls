/**
 * utils.* behaviour, differentially tested against SheetJS CE where the APIs
 * overlap. Cell objects are compared on t/v/f (and z/w where SheetJS sets them).
 */
import { describe, it, expect } from 'vitest';
import * as SheetJS from 'xlsx';
import { utils } from '../src/index.js';

const pick = (ws: any) =>
  Object.fromEntries(
    Object.keys(ws)
      .filter((k) => !k.startsWith('!'))
      .map((k) => [k, { t: ws[k].t, v: ws[k].v instanceof Date ? ws[k].v.getTime() : ws[k].v, f: ws[k].f }])
  );

describe('address utils (vs SheetJS)', () => {
  it('encode/decode columns, rows, cells and ranges', () => {
    for (const c of [0, 1, 25, 26, 27, 51, 52, 701, 702, 16383]) {
      expect(utils.encode_col(c)).toBe(SheetJS.utils.encode_col(c));
      expect(utils.decode_col(SheetJS.utils.encode_col(c))).toBe(c);
    }
    for (const ref of ['A1', 'Z99', 'AA10', 'XFD1048576', '$B$5']) {
      expect(utils.decode_cell(ref)).toEqual(SheetJS.utils.decode_cell(ref.replace(/\$/g, '')));
    }
    for (const r of ['A1:D11', 'B2', 'A1:XFD1048576']) {
      expect(utils.decode_range(r)).toEqual(SheetJS.utils.decode_range(r));
      expect(utils.encode_range(utils.decode_range(r))).toBe(SheetJS.utils.encode_range(SheetJS.utils.decode_range(r)));
    }
    expect(utils.split_cell('$AB$12')).toEqual(SheetJS.utils.split_cell('$AB$12'));
  });

  it('rejects malformed cell references', () => {
    expect(() => utils.decode_cell('1A')).toThrow();
    expect(() => utils.encode_col(-1)).toThrow();
  });
});

describe('aoa_to_sheet / sheet_add_aoa (vs SheetJS)', () => {
  const cases: Array<[string, any[][], any?]> = [
    ['mixed types', [['a', 1, true, false, 0, -2.5, '']]],
    ['null and undefined', [[1, null, undefined, 2], [null]]],
    ['null with sheetStubs', [[1, null]], { sheetStubs: true }],
    ['null with nullError', [[1, null]], { nullError: true }],
    ['NaN / Infinity become errors', [[NaN, Infinity, -Infinity]]],
    ['ragged rows and holes', [['a'], [], ['b', 'c', 'd'], undefined as any]],
    ['origin string', [[1, 2]], { origin: 'C3' }],
    ['origin number', [[1, 2]], { origin: 4 }],
    ['value+formula pair', [[[3, 'A2+1']]]],
  ];
  for (const [name, aoa, opts] of cases) {
    it(name, () => {
      const ours = utils.aoa_to_sheet(aoa, opts);
      const ref = SheetJS.utils.aoa_to_sheet(aoa, opts);
      expect(ours['!ref']).toBe(ref['!ref']);
      expect(pick(ours)).toEqual(pick(ref));
    });
  }

  it('dates match SheetJS serials, z and w in any timezone', () => {
    const aoa = [[new Date(2026, 0, 15), new Date(2026, 5, 30, 13, 45, 10), new Date(1900, 0, 1), new Date(1900, 2, 1)]];
    const ours = utils.aoa_to_sheet(aoa);
    const ref = SheetJS.utils.aoa_to_sheet(aoa);
    for (const k of ['A1', 'B1', 'D1']) {
      expect(ours[k].v).toBeCloseTo(ref[k].v, 9);
      expect(ours[k].z).toBe(ref[k].z);
      expect(ours[k].w).toBe(ref[k].w);
    }
    // 1900-01-01 is serial 1 in Excel (SheetJS is off by one for Jan/Feb 1900)
    expect(ours.C1.v).toBe(1);
    const nf = utils.aoa_to_sheet(aoa, { dateNF: 'yyyy-mm-dd' });
    expect(nf.A1.w).toBe('2026-01-15');
  });

  it('cellDates keeps JS Dates with t:"d"', () => {
    const d = new Date(2026, 0, 15, 8, 30);
    const ws = utils.aoa_to_sheet([[d]], { cellDates: true });
    expect(ws.A1.t).toBe('d');
    expect(ws.A1.v).toEqual(d);
  });

  it('accepts cell objects inside the array (styles inline)', () => {
    const ws = utils.aoa_to_sheet([[{ t: 'n', v: 5, s: { font: { bold: true } } }]]);
    expect(ws.A1).toEqual({ t: 'n', v: 5, s: { font: { bold: true } } });
  });

  it('sheet_add_aoa appends with origin -1 and extends the range', () => {
    const ours = utils.aoa_to_sheet([['a'], ['b']]);
    const ref = SheetJS.utils.aoa_to_sheet([['a'], ['b']]);
    utils.sheet_add_aoa(ours, [['c', 'd']], { origin: -1 });
    SheetJS.utils.sheet_add_aoa(ref, [['c', 'd']], { origin: -1 });
    expect(ours['!ref']).toBe(ref['!ref']);
    expect(pick(ours)).toEqual(pick(ref));
  });

  it('throws on non-array rows', () => {
    expect(() => utils.aoa_to_sheet([{ a: 1 } as any])).toThrow(/array of arrays/);
  });
});

describe('json_to_sheet / sheet_add_json (vs SheetJS)', () => {
  const data = [
    { name: 'Al', age: 3, ok: true, note: null },
    { name: 'Bo', extra: 'x', age: 4 },
  ];
  const cases: Array<[string, any?]> = [
    ['default', undefined],
    ['explicit header order', { header: ['age', 'name'] }],
    ['skipHeader', { skipHeader: true }],
    ['origin', { origin: 'B3' }],
  ];
  for (const [name, opts] of cases) {
    it(name, () => {
      const ours = utils.json_to_sheet(data, opts);
      const ref = SheetJS.utils.json_to_sheet(data, opts);
      expect(ours['!ref']).toBe(ref['!ref']);
      expect(pick(ours)).toEqual(pick(ref));
    });
  }

  it('dates and nested cell objects', () => {
    const rows = [{ d: new Date(2026, 1, 3), c: { t: 's', v: 'raw', s: { font: { italic: true } } } }];
    const ours = utils.json_to_sheet(rows);
    const ref = SheetJS.utils.json_to_sheet(rows);
    expect(ours.A2.v).toBeCloseTo(ref.A2.v, 9);
    expect(ours.A2.z).toBe(ref.A2.z);
    expect(ours.B2.s).toEqual({ font: { italic: true } });
  });

  it('sheet_add_json updates existing cells but keeps their style', () => {
    const ws = utils.json_to_sheet([{ a: 1 }]);
    ws.A2.s = { font: { bold: true } };
    utils.sheet_add_json(ws, [{ a: 9 }], { origin: 'A2', skipHeader: true });
    expect(ws.A2).toEqual({ t: 'n', v: 9, s: { font: { bold: true } } });
  });

  it('empty input', () => {
    expect(utils.json_to_sheet([])['!ref']).toBe(SheetJS.utils.json_to_sheet([])['!ref']);
  });
});

describe('sheet_to_json (vs SheetJS)', () => {
  const aoa = [
    ['Name', 'Age', 'Name', '', 'When'],
    ['Al', 3, 'x', 1, new Date(2026, 0, 2)],
    [],
    ['Bo', null, 'y', 2, null],
    ['Cy', 5.5, true, undefined, new Date(2026, 0, 3)],
  ];
  const variants: Array<[string, any?]> = [
    ['default objects', undefined],
    ['header: 1', { header: 1 }],
    ['header: "A"', { header: 'A' }],
    ['explicit header names', { header: ['n', 'a', 'n2', 'e', 'w'] }],
    ['defval', { defval: '' }],
    ['blankrows true', { blankrows: true }],
    ['blankrows false with header 1', { header: 1, blankrows: false }],
    ['raw false (formatted text)', { raw: false }],
    ['rawNumbers', { raw: false, rawNumbers: true }],
    ['range number', { range: 1, header: 1 }],
    ['range string', { range: 'B1:C4' }],
    ['dateNF', { raw: false, dateNF: 'yyyy-mm-dd' }],
  ];
  for (const [name, opts] of variants) {
    it(name, () => {
      const ours = utils.sheet_to_json(utils.aoa_to_sheet(aoa), opts);
      const ref = SheetJS.utils.sheet_to_json(SheetJS.utils.aoa_to_sheet(aoa), opts);
      expect(ours).toEqual(ref);
    });
  }

  it('sets a non-enumerable __rowNum__', () => {
    const rows: any[] = utils.sheet_to_json(utils.aoa_to_sheet([['h'], [1], [], [2]]));
    expect(rows.map((r) => r.__rowNum__)).toEqual([1, 3]);
    expect(Object.keys(rows[0])).toEqual(['h']);
  });

  it('skipHidden drops hidden rows and columns', () => {
    const ws = utils.aoa_to_sheet([
      ['a', 'b'],
      [1, 2],
      [3, 4],
    ]);
    ws['!rows'] = [undefined as any, { hidden: true }];
    ws['!cols'] = [{ hidden: true }];
    expect(utils.sheet_to_json(ws, { skipHidden: true })).toEqual([{ b: 4 }]);
  });

  it('error cells: #NULL! -> null, others skipped', () => {
    const ws = utils.aoa_to_sheet([['a', 'b'], [{ t: 'e', v: 0 }, { t: 'e', v: 7 }]]);
    const ref = SheetJS.utils.aoa_to_sheet([['a', 'b'], [{ t: 'e', v: 0 }, { t: 'e', v: 7 }]]);
    expect(utils.sheet_to_json(ws)).toEqual(SheetJS.utils.sheet_to_json(ref));
  });

  it('returns [] for an empty sheet', () => {
    expect(utils.sheet_to_json({})).toEqual([]);
  });
});

describe('sheet_to_csv / sheet_to_txt / sheet_to_formulae (vs SheetJS)', () => {
  const aoa = [
    ['ID', 'Text', 'Num', 'Bool'],
    [1, 'has,comma', 1.5, true],
    [2, 'has "quote"', -3, false],
    [3, 'multi\nline', 0, null],
    [],
    [4, '', 1e21, true],
  ];
  const variants: Array<[string, any?]> = [
    ['default', undefined],
    ['FS ;', { FS: ';' }],
    ['RS CRLF', { RS: '\r\n' }],
    ['blankrows false', { blankrows: false }],
    ['strip', { strip: true }],
    ['forceQuotes', { forceQuotes: true }],
    ['rawNumbers', { rawNumbers: true }],
  ];
  for (const [name, opts] of variants) {
    it(`csv ${name}`, () => {
      expect(utils.sheet_to_csv(utils.aoa_to_sheet(aoa), opts)).toBe(SheetJS.utils.sheet_to_csv(SheetJS.utils.aoa_to_sheet(aoa), opts));
    });
  }

  it('csv uses number formats and dates', () => {
    const make = (lib: any) => {
      const ws = lib.utils.aoa_to_sheet([[0.256, new Date(2026, 0, 15), 1234.5]]);
      ws.A1.z = '0.0%';
      ws.C1.z = '#,##0.00';
      return ws;
    };
    expect(utils.sheet_to_csv(make({ utils }))).toBe(SheetJS.utils.sheet_to_csv(make(SheetJS)));
  });

  it('csv writes formula-only cells', () => {
    const ws = utils.aoa_to_sheet([[1, 2]]);
    ws.C1 = { t: 'n', f: 'SUM(A1,B1)' };
    ws['!ref'] = 'A1:C1';
    const ref = SheetJS.utils.aoa_to_sheet([[1, 2]]);
    ref.C1 = { t: 'n', f: 'SUM(A1,B1)' };
    ref['!ref'] = 'A1:C1';
    expect(utils.sheet_to_csv(ws)).toBe(SheetJS.utils.sheet_to_csv(ref));
  });

  it('txt uses tabs', () => {
    expect(utils.sheet_to_txt(utils.aoa_to_sheet(aoa))).toBe(SheetJS.utils.sheet_to_csv(SheetJS.utils.aoa_to_sheet(aoa), { FS: '\t' }));
  });

  it('formulae', () => {
    const build = (lib: any) => {
      const ws = lib.utils.aoa_to_sheet([
        [1, 'two', true, null],
        [3.5, '', false],
      ]);
      ws.D1 = { t: 'n', f: 'A1*2', v: 2 };
      ws['!ref'] = 'A1:D2';
      return ws;
    };
    expect(utils.sheet_to_formulae(build({ utils }))).toEqual(SheetJS.utils.sheet_to_formulae(build(SheetJS)));
  });

  it('array formulae', () => {
    const ws = utils.aoa_to_sheet([[1], [2]]);
    utils.sheet_set_array_formula(ws, 'B1:B2', 'A1:A2*2');
    const ref = SheetJS.utils.aoa_to_sheet([[1], [2]]);
    SheetJS.utils.sheet_set_array_formula(ref, 'B1:B2', 'A1:A2*2');
    expect(utils.sheet_to_formulae(ws)).toEqual(SheetJS.utils.sheet_to_formulae(ref));
    expect(ws['!ref']).toBe(ref['!ref']);
  });
});

describe('format_cell', () => {
  it('matches SheetJS for common formats', () => {
    const cells = [
      { t: 'n', v: 1234.5678 },
      { t: 'n', v: 1234.5678, z: '#,##0.00' },
      { t: 'n', v: 0.125, z: '0.0%' },
      { t: 'n', v: 46037.5, z: 'yyyy-mm-dd hh:mm' },
      { t: 'n', v: -5, z: '[Red]0.00;[Blue](0.00)' },
      { t: 'b', v: true },
      { t: 'e', v: 0x07 },
      { t: 's', v: 'text' },
      { t: 'z' },
    ];
    for (const c of cells) expect(utils.format_cell(c as any)).toBe(SheetJS.utils.format_cell(c as any));
  });

  it('prefers cell.w', () => {
    expect(utils.format_cell({ t: 'n', v: 1, w: 'one' })).toBe('one');
  });
});

describe('workbook utils', () => {
  it('book_new with a sheet, book_append_sheet naming and roll', () => {
    const wb = utils.book_new(utils.aoa_to_sheet([[1]]));
    expect(wb.SheetNames).toEqual(['Sheet1']);
    expect(utils.book_append_sheet(wb, utils.aoa_to_sheet([[2]]))).toBe('Sheet2');
    expect(utils.book_append_sheet(wb, utils.aoa_to_sheet([[3]]), 'Sheet1', true)).toBe('Sheet3');
    expect(utils.book_append_sheet(wb, utils.aoa_to_sheet([[3]]), 'Data')).toBe('Data');
    expect(utils.book_append_sheet(wb, utils.aoa_to_sheet([[3]]), 'Data', true)).toBe('Data1');
    expect(() => utils.book_append_sheet(wb, utils.aoa_to_sheet([[3]]), 'Data')).toThrow(/already exists/);
  });

  it('validates sheet names like SheetJS', () => {
    for (const bad of ['x'.repeat(32), 'a:b', 'a\\b', 'a/b', 'a?b', 'a*b', 'a[b', 'a]b', "'a", "a'", 'History']) {
      expect(() => utils.book_append_sheet(utils.book_new(), {}, bad), bad).toThrow();
      expect(() => SheetJS.utils.book_append_sheet(SheetJS.utils.book_new(), {}, bad), bad).toThrow();
    }
    expect(() => utils.book_append_sheet(utils.book_new(), {}, 'x'.repeat(31))).not.toThrow();
  });

  it('book_set_sheet_visibility', () => {
    const wb = utils.book_new(utils.aoa_to_sheet([[1]]), 'A');
    utils.book_append_sheet(wb, utils.aoa_to_sheet([[1]]), 'B');
    utils.book_set_sheet_visibility(wb, 'B', utils.consts.SHEET_VERY_HIDDEN);
    expect(wb.Workbook?.Sheets?.[1].Hidden).toBe(2);
    expect(() => utils.book_set_sheet_visibility(wb, 'nope', 1)).toThrow();
  });

  it('cell helpers', () => {
    const ws = utils.aoa_to_sheet([[1]]);
    const c = utils.sheet_get_cell(ws, 'C3');
    expect(c).toEqual({ t: 'z' });
    utils.cell_set_number_format(ws.A1, '0.00');
    utils.cell_set_hyperlink(ws.A1, 'https://example.com', 'tip');
    utils.cell_add_comment(ws.A1, 'note', 'Ann');
    expect(ws.A1).toMatchObject({ z: '0.00', l: { Target: 'https://example.com', Tooltip: 'tip' }, c: [{ t: 'note', a: 'Ann' }] });
    utils.cell_set_internal_link(ws.A1, 'Sheet2!A1');
    expect(ws.A1.l).toEqual({ Target: '#Sheet2!A1' });
    utils.cell_set_hyperlink(ws.A1);
    expect(ws.A1.l).toBeUndefined();
  });

  it('set_cell_style / set_range_style merge nested objects and extend !ref', () => {
    const ws = utils.aoa_to_sheet([['a']]);
    utils.set_range_style(ws, 'A1:B2', { font: { bold: true }, border: { top: { style: 'thin' } } });
    utils.set_cell_style(ws, 'B2', { font: { italic: true }, border: { top: { color: { rgb: 'FF0000' } }, left: { style: 'thick' } } });
    expect(ws['!ref']).toBe('A1:B2');
    expect(ws.B2.s).toEqual({
      font: { bold: true, italic: true },
      border: { top: { style: 'thin', color: { rgb: 'FF0000' } }, left: { style: 'thick' } },
    });
    expect(ws.A1.s).toEqual({ font: { bold: true }, border: { top: { style: 'thin' } } });
  });
});
