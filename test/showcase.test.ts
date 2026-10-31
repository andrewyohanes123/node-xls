/**
 * Validates every workbook produced by `npm run showcase` / `npm run
 * excel-check` (skipped when they have not been generated). These are the files meant for manual
 * inspection in Excel, so they must pass the same checks as unit-test output.
 */
import { describe, it } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { validateXlsx } from './helpers/validate.js';

for (const sub of ['showcase', 'excel-check']) {
  const dir = path.resolve(__dirname, '../out', sub);
  const files = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => f.endsWith('.xlsx')) : [];
  describe.skipIf(files.length === 0)(`${sub} workbooks`, () => {
    for (const f of files) {
      it(f, async () => {
        await validateXlsx(fs.readFileSync(path.join(dir, f)));
      });
    }
  });
}
