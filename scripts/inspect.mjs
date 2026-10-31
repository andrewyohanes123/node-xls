#!/usr/bin/env node
// Usage: node scripts/inspect.mjs <file.xlsx> [outDir]
// Prints a package summary and writes every part (XML pretty-printed) to
// out/inspect/<file-name>/ for manual review.
import { readFileSync, writeFileSync, mkdirSync, rmSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { unzipSync, strFromU8 } from 'fflate';

const file = process.argv[2];
if (!file) {
  console.error('usage: node scripts/inspect.mjs <file.xlsx> [outDir]');
  process.exit(1);
}
const outDir = process.argv[3] ?? join('out', 'inspect', basename(file).replace(/\.xlsx$/i, ''));
const files = unzipSync(readFileSync(file));
const hasXmllint = spawnSync('xmllint', ['--version']).status === 0;

rmSync(outDir, { recursive: true, force: true });
const ct = strFromU8(files['[Content_Types].xml'] ?? new Uint8Array());
const typeOf = (part) => {
  const o = new RegExp(`PartName="/${part.replace(/[.[\]]/g, '\\$&')}" ContentType="([^"]+)"`).exec(ct);
  if (o) return o[1];
  const ext = part.slice(part.lastIndexOf('.') + 1);
  return new RegExp(`Extension="${ext}" ContentType="([^"]+)"`, 'i').exec(ct)?.[1] ?? '(none)';
};

console.log(`\n${file}\n${'='.repeat(file.length)}`);
console.log('parts:');
for (const part of Object.keys(files).sort()) {
  const bytes = files[part];
  const dest = join(outDir, part);
  mkdirSync(dirname(dest), { recursive: true });
  if (/\.(xml|rels|vml)$/i.test(part)) {
    const text = strFromU8(bytes);
    const pretty = hasXmllint ? spawnSync('xmllint', ['--format', '-'], { input: text, encoding: 'utf8' }) : null;
    writeFileSync(dest, pretty && pretty.status === 0 ? pretty.stdout : text);
  } else {
    writeFileSync(dest, bytes);
  }
  console.log(`  ${part.padEnd(52)} ${String(bytes.length).padStart(8)} B  ${typeOf(part).replace('application/vnd.openxmlformats-', '')}`);
}

const wb = strFromU8(files['xl/workbook.xml'] ?? new Uint8Array());
const sheets = [...wb.matchAll(/<sheet [^>]*name="([^"]*)"/g)].map((m) => m[1]);
const count = (re, s) => (s.match(re) || []).length;
console.log('\nsummary:');
console.log(`  sheets:          ${sheets.join(', ')}`);
sheets.forEach((name, i) => {
  const xml = strFromU8(files[`xl/worksheets/sheet${i + 1}.xml`] ?? new Uint8Array());
  const dim = /<dimension ref="([^"]+)"/.exec(xml)?.[1];
  console.log(`  [${name}] dimension=${dim ?? '?'} cells=${count(/<c /g, xml)} merges=${count(/<mergeCell /g, xml)}`);
});
const styles = strFromU8(files['xl/styles.xml'] ?? new Uint8Array());
const cellXfs = /<cellXfs count="(\d+)"/.exec(styles)?.[1] ?? '0';
const sst = strFromU8(files['xl/sharedStrings.xml'] ?? new Uint8Array());
console.log(`  cell styles:     ${cellXfs}`);
console.log(`  shared strings:  ${count(/<si>/g, sst)}`);
for (const [label, re] of [
  ['pivot tables', /^xl\/pivotTables\/pivotTable\d+\.xml$/],
  ['tables', /^xl\/tables\/table\d+\.xml$/],
  ['charts', /^xl\/charts\/chart\d+\.xml$/],
  ['images', /^xl\/media\//],
  ['comments', /^xl\/comments\d+\.xml$/],
]) {
  const n = Object.keys(files).filter((p) => re.test(p)).length;
  if (n) console.log(`  ${(label + ':').padEnd(17)}${n}`);
}
console.log(`\nparts written to ${outDir}/`);
