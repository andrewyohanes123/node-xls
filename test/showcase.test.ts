/**
 * Validates every workbook produced by `npm run showcase` (skipped when the
 * showcase has not been generated). These are the files meant for manual
 * inspection in Excel, so they must pass the same checks as unit-test output.
 */
import { describe, it } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { validateXlsx } from './helpers/validate.js';

const dir = path.resolve(__dirname, '../out/showcase');
const files = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => f.endsWith('.xlsx')) : [];

describe.skipIf(files.length === 0)('showcase workbooks', () => {
  for (const f of files) {
    it(f, async () => {
      await validateXlsx(fs.readFileSync(path.join(dir, f)));
    });
  }
});
