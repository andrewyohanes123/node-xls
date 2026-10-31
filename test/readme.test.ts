/**
 * Keeps README examples honest: the Quick start code block is extracted from
 * README.md and executed (writeFile redirected to memory).
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import XLSX, { utils } from '../src/index.js';
import { validateXlsx } from './helpers/validate.js';

describe('README', () => {
  it('quick start runs and produces a valid workbook', async () => {
    const readme = fs.readFileSync(path.resolve(__dirname, '../README.md'), 'utf8');
    const section = readme.slice(readme.indexOf('## Quick start'));
    const code = /```ts\n([\s\S]*?)```/.exec(section)![1].replace(/^import .*$/m, '');
    let written: any;
    const fakeXLSX = { ...XLSX, writeFile: (wb: any) => (written = wb) };
    new Function('XLSX', 'utils', code)(fakeXLSX, utils);
    expect(written.SheetNames).toEqual(['Sales', 'Summary']);
    const rep = await validateXlsx(XLSX.write(written, { type: 'array' }));
    expect(rep.files['xl/pivotTables/pivotTable1.xml']).toBeDefined();
    expect(written.Sheets.Sales['!freeze']).toBe('A2');
  });
});
