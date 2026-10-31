/**
 * Input/output: data types, files (sync/async), CSV/TSV/HTML/JSON formats,
 * HTML round trips with styles and merges.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import XLSX, { utils } from '../src/index.js';
import { book } from './helpers/util.js';
import { validateXlsx } from './helpers/validate.js';

let dir: string;
beforeAll(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'node-xls-io-'));
});
afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

const sample = () => {
  const ws = utils.aoa_to_sheet([
    ['Name', 'Qty', 'Price'],
    ['Pen, blue', 3, 1.5],
    ['Ink "black"', 10, 0.25],
  ]);
  ws.C2.z = '$0.00';
  return book(['Items', ws], ['Second', utils.aoa_to_sheet([['only on sheet 2']])]);
};

describe('write/read output types', () => {
  for (const type of ['buffer', 'array', 'base64', 'binary'] as const) {
    it(type, async () => {
      const out = XLSX.write(sample(), { type });
      if (type === 'buffer') expect(Buffer.isBuffer(out)).toBe(true);
      if (type === 'array') expect(out).toBeInstanceOf(Uint8Array);
      if (type === 'base64' || type === 'binary') expect(typeof out).toBe('string');
      const back = XLSX.read(out, { type });
      expect(back.Sheets.Items.A2.v).toBe('Pen, blue');
    });
  }

  it('reads ArrayBuffer, Buffer slices and base64 without a type', async () => {
    const buf: Buffer = XLSX.write(sample(), { type: 'buffer' });
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    expect(XLSX.read(ab).Sheets.Items.B3.v).toBe(10);
    const padded = Buffer.concat([Buffer.from('xxxx'), buf]).subarray(4);
    expect(XLSX.read(padded).Sheets.Items.B3.v).toBe(10);
    expect(XLSX.read(buf.toString('base64')).Sheets.Items.B3.v).toBe(10);
  });

  it('uncompressed output is still valid', async () => {
    const out = XLSX.write(sample(), { type: 'array', compression: false });
    await validateXlsx(out);
    expect(out.length).toBeGreaterThan(XLSX.write(sample(), { type: 'array' }).length);
  });

  it('async APIs', async () => {
    const buf = await XLSX.writeAsync(sample(), { type: 'buffer' });
    const back = await XLSX.readAsync(buf);
    expect(back.SheetNames).toEqual(['Items', 'Second']);
  });
});

describe('files', () => {
  it('writeFile / readFile (sync and async) pick the format from the extension', async () => {
    const wb = sample();
    for (const ext of ['xlsx', 'csv', 'txt', 'html', 'json']) {
      const file = path.join(dir, `out.${ext}`);
      XLSX.writeFile(wb, file);
      expect(fs.statSync(file).size).toBeGreaterThan(0);
    }
    await XLSX.writeFileAsync(wb, path.join(dir, 'async.xlsx'));
    await validateXlsx(fs.readFileSync(path.join(dir, 'out.xlsx')));
    expect(XLSX.readFile(path.join(dir, 'out.xlsx')).Sheets.Items.C2.w).toBe('$1.50');
    expect((await XLSX.readFileAsync(path.join(dir, 'async.xlsx'))).SheetNames).toEqual(['Items', 'Second']);
    expect(XLSX.readFile(path.join(dir, 'out.csv')).Sheets.Sheet1.A2.v).toBe('Pen, blue');
    expect(XLSX.readFile(path.join(dir, 'out.html')).Sheets.Sheet1.C2.v).toBe(1.5);
    expect(JSON.parse(fs.readFileSync(path.join(dir, 'out.json'), 'utf8')).Items[0]).toEqual({ Name: 'Pen, blue', Qty: 3, Price: 1.5 });
    // byte output of CSV/TXT starts with a UTF-8 BOM so Excel detects the encoding (as SheetJS)
    expect(fs.readFileSync(path.join(dir, 'out.txt'), 'utf8').split('\n')[0]).toBe('\ufeffName\tQty\tPrice');
  });

  it('rejects unsupported extensions instead of writing mislabeled files', () => {
    expect(() => XLSX.writeFile(sample(), path.join(dir, 'legacy.xls'))).toThrow(/Unsupported file extension/);
    expect(() => XLSX.writeFile(sample(), path.join(dir, 'doc.ods'))).toThrow(/Unsupported/);
  });
});

describe('CSV', () => {
  it('formatted output, sheet selection, separators', () => {
    const wb = sample();
    expect(XLSX.write(wb, { bookType: 'csv', type: 'string' })).toBe('Name,Qty,Price\n"Pen, blue",3,$1.50\n"Ink ""black""",10,0.25');
    expect(XLSX.write(wb, { bookType: 'csv', type: 'string', sheet: 'Second' })).toBe('only on sheet 2');
    expect(XLSX.write(wb, { bookType: 'csv', type: 'string', sheet: 1 })).toBe('only on sheet 2');
    expect(() => XLSX.write(wb, { bookType: 'csv', sheet: 'Nope' })).toThrow(/Cannot find sheet/);
  });

  it('reading: BOM, CRLF, quotes, separators, type detection, raw', () => {
    const csv = '﻿id;name;ok;amount\r\n1;"semi;colon";TRUE;007\r\n2;"multi\nline";false;3.5\r\n';
    const ws = XLSX.read(csv, { type: 'string' }).Sheets.Sheet1;
    expect(ws.A1.v).toBe('id');
    expect(ws.B2.v).toBe('semi;colon');
    expect(ws.C2.v).toBe(true);
    expect(ws.D2.v).toBe('007');
    expect(ws.B3.v).toBe('multi\nline');
    expect(ws.D3.v).toBe(3.5);
    const raw = XLSX.read(csv, { type: 'string', raw: true }).Sheets.Sheet1;
    expect(raw.C2.v).toBe('TRUE');
    const tsv = XLSX.read('a\tb\n1\t2', { type: 'string' }).Sheets.Sheet1;
    expect(tsv.B2.v).toBe(2);
  });
});

describe('HTML', () => {
  it('sheet_to_html -> table_to_sheet keeps values, types, styles, links and merges', () => {
    const ws = utils.aoa_to_sheet([
      ['Title', null, null],
      ['a & b', 42, true],
      ['line1\nline2', 0, false],
    ]);
    ws['!merges'] = [utils.decode_range('A1:C1')];
    ws.A1.s = { font: { bold: true, sz: 14, color: { rgb: 'C00000' } }, fill: { fgColor: { rgb: 'FFF2CC' } }, alignment: { horizontal: 'center' } };
    ws.B2.s = { border: { bottom: { style: 'medium', color: { rgb: '0000FF' } } } };
    ws.B2.z = '0.0';
    ws.A2.l = { Target: 'https://example.com', Tooltip: 'go' };
    const html = utils.sheet_to_html(ws, { id: 'tbl' });
    expect(html).toContain('colspan="3"');
    expect(html).toContain('<br/>');
    expect(html).toContain('>42.0<');
    expect(html).toContain('>FALSE<');
    expect(html).toContain('a &amp; b');
    const back = utils.table_to_sheet(html);
    expect(back['!merges']).toEqual([utils.decode_range('A1:C1')]);
    expect(back.A1.s).toEqual({ font: { bold: true, sz: 14, color: { rgb: 'C00000' } }, fill: { patternType: 'solid', fgColor: { rgb: 'FFF2CC' } }, alignment: { horizontal: 'center' } });
    expect(back.B2).toMatchObject({ t: 'n', v: 42, z: '0.0', s: { border: { bottom: { style: 'medium', color: { rgb: '0000FF' } } } } });
    expect(back.C2).toMatchObject({ t: 'b', v: true });
    expect(back.A3.v).toBe('line1\nline2');
    expect(back.A2.l).toEqual({ Target: 'https://example.com', Tooltip: 'go' });
  });

  it('neutralises javascript: links and escapes attribute values', () => {
    const ws = utils.aoa_to_sheet([['x']]);
    ws.A1.l = { Target: 'javascript:alert(1)' };
    ws.A1.s = { font: { name: 'Evil"; background:url(x)' } };
    const html = utils.sheet_to_html(ws);
    expect(html).toContain('href="#"');
    expect(html).not.toContain('url(x)"');
    expect(html).not.toMatch(/style="[^"]*"[^>]*background:url/);
  });

  it('table_to_sheet handles rowspan, th, entities and inferred numbers', () => {
    const html = `<!DOCTYPE html><html><body><table>
      <tr><th rowspan="2">R</th><th>A&nbsp;B</th><th>&#x263A;</th></tr>
      <tr><td>1,5</td><td> 12 </td></tr></table></body></html>`;
    const wb = XLSX.read(html, { type: 'string' });
    const ws = wb.Sheets.Sheet1;
    expect(ws['!merges']).toEqual([utils.decode_range('A1:A2')]);
    expect(ws.B1.v).toBe('A B');
    expect(ws.C1.v).toBe('☺');
    expect(ws.B2.v).toBe('1,5');
    expect(ws.C2.v).toBe(12);
  });

  it('write html bookType produces a full document', () => {
    const html = XLSX.write(sample(), { bookType: 'html', type: 'string' });
    expect(html).toMatch(/^<!DOCTYPE html>/);
    expect(html).toContain('<meta charset="utf-8"/>');
  });
});
