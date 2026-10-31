/**
 * Independent validation of generated .xlsx packages.
 *
 * A file passing this helper has: complete content types, resolvable
 * relationships, well-formed XML, consistent shared-string / style indexes,
 * ordered rows & cells, schema-valid parts (ECMA-376 Transitional XSD via
 * xmllint), and loads in both SheetJS CE and ExcelJS.
 */
import { unzipSync, strFromU8 } from 'fflate';
import { XMLParser, XMLValidator } from 'fast-xml-parser';
import { execFileSync, spawnSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as SheetJS from 'xlsx';
import ExcelJS from 'exceljs';

export const SCHEMA_DIR = path.resolve(__dirname, '../../.cache/ecma/xsd');

const HAS_XMLLINT = spawnSync('xmllint', ['--version']).status === 0;
const HAS_SCHEMAS = fs.existsSync(path.join(SCHEMA_DIR, 'sml.xsd'));
export const SCHEMA_VALIDATION = HAS_XMLLINT && HAS_SCHEMAS;
if (!SCHEMA_VALIDATION) {
  console.warn('[validate] XSD validation disabled: run scripts/fetch-schemas.sh and install xmllint');
}

const SCHEMA_BY_CONTENT_TYPE: Array<[RegExp, string]> = [
  [
    /spreadsheetml\.(sheet\.main|worksheet|styles|sharedStrings|pivotTable|pivotCacheDefinition|pivotCacheRecords|table|comments)\+xml$/,
    'sml.xsd',
  ],
  [/drawingml\.chart\+xml$/, 'dml-chart.xsd'],
  [/officedocument\.drawing\+xml$/, 'dml-spreadsheetDrawing.xsd'],
  [/extended-properties\+xml$/, 'shared-documentPropertiesExtended.xsd'],
];

const INVALID_SHEET_NAME = /[\\/?*[\]:]/;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  trimValues: false,
  isArray: (name) =>
    ['Default', 'Override', 'Relationship', 'sheet', 'row', 'c', 'si', 'xf', 'font', 'fill', 'border', 'numFmt', 'dxf'].includes(name),
});

export interface XlsxReport {
  files: Record<string, Uint8Array>;
  contentTypes: Record<string, string>;
  /** Returns the decoded text of a part (throws if missing). */
  xml(part: string): string;
  /** Parsed (attribute-preserving) object for a part. */
  parse(part: string): any;
  sheetjs: SheetJS.WorkBook;
  exceljs: ExcelJS.Workbook;
}

function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const base = path.posix.dirname(sourcePart);
  return path.posix.normalize(path.posix.join(base === '.' ? '' : base, target));
}

function relsSourcePart(relsPath: string): string {
  // xl/worksheets/_rels/sheet1.xml.rels -> xl/worksheets/sheet1.xml ; _rels/.rels -> ''
  const dir = path.posix.dirname(path.posix.dirname(relsPath));
  const file = path.posix.basename(relsPath).replace(/\.rels$/, '');
  return dir === '.' ? file : `${dir}/${file}`;
}

/**
 * Minimal Markup Compatibility processing (ECMA-376 Part 3) so files written by
 * Excel/ExcelJS validate against the base schemas: keeps mc:Fallback content of
 * mc:AlternateContent and drops elements/attributes in mc:Ignorable namespaces.
 * node-xls itself never emits MCE markup.
 */
function stripMarkupCompatibility(xml: string): string {
  let out = xml.replace(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g, (block) => {
    const fallback = /<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/.exec(block);
    return fallback ? fallback[1] : '';
  });
  const ignorable = /\smc:Ignorable="([^"]*)"/.exec(out);
  if (!ignorable) return out;
  out = out.replace(/\smc:Ignorable="[^"]*"/g, '');
  for (const p of ignorable[1].trim().split(/\s+/).filter(Boolean)) {
    out = out.replace(new RegExp(`<${p}:[\\w.-]+\\b[^>]*?/>`, 'g'), '');
    out = out.replace(new RegExp(`<${p}:([\\w.-]+)\\b[^>]*>[\\s\\S]*?</${p}:\\1>`, 'g'), '');
    out = out.replace(new RegExp(`\\s${p}:[\\w.-]+="[^"]*"`, 'g'), '');
  }
  return out;
}

function schemaValidate(parts: Record<string, string>, contentTypes: Record<string, string>, errors: string[]) {
  if (!SCHEMA_VALIDATION) return;
  const groups = new Map<string, string[]>();
  for (const part of Object.keys(parts)) {
    const ct = contentTypes[part];
    if (!ct) continue;
    const match = SCHEMA_BY_CONTENT_TYPE.find(([re]) => re.test(ct));
    if (!match) continue;
    const list = groups.get(match[1]) ?? [];
    list.push(part);
    groups.set(match[1], list);
  }
  if (groups.size === 0) return;

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'node-xls-xsd-'));
  try {
    for (const [schema, list] of groups) {
      const filePaths = list.map((part) => {
        const p = path.join(dir, part.replace(/\//g, '__'));
        // The ECMA XSDs don't import xml.xsd, so xml:space (which Excel itself
        // writes on <t> elements) would be reported as undeclared.
        fs.writeFileSync(p, stripMarkupCompatibility(parts[part]).replace(/ xml:space="preserve"/g, ''));
        return p;
      });
      try {
        execFileSync('xmllint', ['--noout', '--schema', path.join(SCHEMA_DIR, schema), ...filePaths], {
          stdio: ['ignore', 'pipe', 'pipe'],
        });
      } catch (e: any) {
        const lines = String(e.stderr || '')
          .split('\n')
          .filter((l) => l.includes('error') || l.includes('fails to validate'))
          .map((l) => l.replace(dir + '/', '').replace(/__/g, '/'));
        errors.push(`XSD (${schema}):\n    ${lines.join('\n    ')}`);
      }
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

export async function validateXlsx(data: Uint8Array | ArrayBuffer): Promise<XlsxReport> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const errors: string[] = [];
  const files = unzipSync(bytes);
  const partNames = Object.keys(files).filter((n) => !n.endsWith('/'));
  const text: Record<string, string> = {};
  for (const n of partNames) {
    if (/\.(xml|rels|vml)$/i.test(n)) text[n] = strFromU8(files[n]);
  }

  // 1. Well-formed XML
  for (const [n, s] of Object.entries(text)) {
    const res = XMLValidator.validate(s);
    if (res !== true) errors.push(`${n}: malformed XML: ${res.err.msg} (line ${res.err.line})`);
  }
  const parsed = (part: string) => parser.parse(text[part]);

  // 2. Content types
  const contentTypes: Record<string, string> = {};
  if (!text['[Content_Types].xml']) {
    errors.push('missing [Content_Types].xml');
  } else {
    const ct = parsed('[Content_Types].xml').Types || {};
    const defaults = new Map<string, string>();
    for (const d of ct.Default || []) defaults.set(String(d['@_Extension']).toLowerCase(), d['@_ContentType']);
    const overrides = new Map<string, string>();
    for (const o of ct.Override || []) {
      const name = String(o['@_PartName']).replace(/^\//, '');
      overrides.set(name, o['@_ContentType']);
      if (!files[name]) errors.push(`content type override for missing part: ${name}`);
    }
    for (const n of partNames) {
      if (n === '[Content_Types].xml') continue;
      const base = path.posix.basename(n);
      const ext = base.includes('.') ? base.slice(base.lastIndexOf('.') + 1).toLowerCase() : '';
      const type = overrides.get(n) ?? defaults.get(ext);
      if (!type) errors.push(`no content type for part: ${n}`);
      else contentTypes[n] = type;
    }
  }

  // 3. Relationships resolve
  for (const n of partNames.filter((p) => p.endsWith('.rels'))) {
    const source = relsSourcePart(n);
    if (source && !files[source]) errors.push(`${n}: rels for missing source part ${source}`);
    const rels = parsed(n).Relationships?.Relationship || [];
    const ids = new Set<string>();
    for (const r of rels) {
      if (ids.has(r['@_Id'])) errors.push(`${n}: duplicate relationship id ${r['@_Id']}`);
      ids.add(r['@_Id']);
      if (r['@_TargetMode'] === 'External') continue;
      const target = resolveTarget(source, r['@_Target']);
      if (!files[target]) errors.push(`${n}: ${r['@_Id']} -> missing part ${target}`);
    }
  }

  // 4. Workbook sheet names / ids
  const wbPart = 'xl/workbook.xml';
  if (text[wbPart]) {
    const sheets = parsed(wbPart).workbook?.sheets?.sheet || [];
    const names = new Set<string>();
    const ids = new Set<string>();
    for (const s of sheets) {
      const name = String(s['@_name']);
      if (name.length === 0 || name.length > 31) errors.push(`sheet name length invalid: "${name}"`);
      if (INVALID_SHEET_NAME.test(name)) errors.push(`sheet name has invalid characters: "${name}"`);
      if (names.has(name.toLowerCase())) errors.push(`duplicate sheet name: "${name}"`);
      names.add(name.toLowerCase());
      if (ids.has(s['@_sheetId'])) errors.push(`duplicate sheetId ${s['@_sheetId']}`);
      ids.add(s['@_sheetId']);
    }
  }

  // 5. Shared strings + styles + worksheet cell ordering
  const sstCount = text['xl/sharedStrings.xml'] ? (parsed('xl/sharedStrings.xml').sst?.si || []).length : 0;
  let xfCount = 1;
  if (text['xl/styles.xml']) {
    const ss = parsed('xl/styles.xml').styleSheet || {};
    const counts: Record<string, number> = {};
    for (const [coll, item] of [
      ['numFmts', 'numFmt'],
      ['fonts', 'font'],
      ['fills', 'fill'],
      ['borders', 'border'],
      ['cellXfs', 'xf'],
      ['dxfs', 'dxf'],
    ] as const) {
      const node = ss[coll];
      if (!node) continue;
      const n = (node[item] || []).length;
      counts[coll] = n;
      if (node['@_count'] !== undefined && Number(node['@_count']) !== n) {
        errors.push(`styles.xml: <${coll} count="${node['@_count']}"> but has ${n} items`);
      }
    }
    xfCount = counts.cellXfs ?? 0;
    const customFmts = new Set<number>((ss.numFmts?.numFmt || []).map((f: any) => Number(f['@_numFmtId'])));
    (ss.cellXfs?.xf || []).forEach((xf: any, i: number) => {
      if (Number(xf['@_fontId'] ?? 0) >= (counts.fonts ?? 0)) errors.push(`cellXfs[${i}] fontId out of range`);
      if (Number(xf['@_fillId'] ?? 0) >= (counts.fills ?? 0)) errors.push(`cellXfs[${i}] fillId out of range`);
      if (Number(xf['@_borderId'] ?? 0) >= (counts.borders ?? 0)) errors.push(`cellXfs[${i}] borderId out of range`);
      const nf = Number(xf['@_numFmtId'] ?? 0);
      if (nf >= 164 && !customFmts.has(nf)) errors.push(`cellXfs[${i}] numFmtId ${nf} not defined`);
    });
  }

  for (const n of partNames.filter((p) => contentTypes[p]?.endsWith('spreadsheetml.worksheet+xml'))) {
    const ws = parsed(n).worksheet || {};
    let prevRow = 0;
    for (const row of ws.sheetData?.row || []) {
      const r = Number(row['@_r']);
      if (!(r > prevRow)) errors.push(`${n}: row ${r} not ascending`);
      prevRow = r;
      let prevCol = -1;
      for (const c of row.c || []) {
        const m = /^([A-Z]+)(\d+)$/.exec(c['@_r']);
        if (!m) {
          errors.push(`${n}: bad cell ref ${c['@_r']}`);
          continue;
        }
        if (Number(m[2]) !== r) errors.push(`${n}: cell ${c['@_r']} inside row ${r}`);
        const col = m[1].split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
        if (!(col > prevCol)) errors.push(`${n}: cell ${c['@_r']} not ascending in row ${r}`);
        prevCol = col;
        if (c['@_s'] !== undefined && Number(c['@_s']) >= xfCount) errors.push(`${n}: ${c['@_r']} style ${c['@_s']} >= ${xfCount}`);
        if (c['@_t'] === 's') {
          const idx = Number(typeof c.v === 'object' ? c.v['#text'] : c.v);
          if (!(idx >= 0 && idx < sstCount)) errors.push(`${n}: ${c['@_r']} shared string ${idx} out of range`);
        }
        const v = typeof c.v === 'object' ? c.v?.['#text'] : c.v;
        if ((c['@_t'] === undefined || c['@_t'] === 'n') && v !== undefined && !Number.isFinite(Number(v))) {
          errors.push(`${n}: ${c['@_r']} non-finite numeric value "${v}"`);
        }
        if (c.f !== undefined) {
          const f = typeof c.f === 'object' ? c.f['#text'] : c.f;
          if (typeof f === 'string' && f.startsWith('=')) errors.push(`${n}: ${c['@_r']} formula starts with "="`);
        }
      }
    }
  }

  // 6. Schema validation
  schemaValidate(text, contentTypes, errors);

  // 7. Independent readers
  let sheetjs: SheetJS.WorkBook | undefined;
  try {
    sheetjs = SheetJS.read(bytes, { type: 'array', cellStyles: true, cellFormula: true, cellNF: true });
  } catch (e: any) {
    errors.push(`SheetJS CE failed to read: ${e.message}`);
  }
  const exceljs = new ExcelJS.Workbook();
  try {
    await exceljs.xlsx.load(Buffer.from(bytes) as any);
  } catch (e: any) {
    errors.push(`ExcelJS failed to read: ${e.message}`);
  }

  if (errors.length > 0) {
    throw new Error(`Invalid xlsx package:\n  - ${errors.join('\n  - ')}`);
  }

  return {
    files,
    contentTypes,
    xml(part: string) {
      if (!files[part]) throw new Error(`missing part ${part}`);
      return strFromU8(files[part]);
    },
    parse(part: string) {
      if (!files[part]) throw new Error(`missing part ${part}`);
      return parser.parse(strFromU8(files[part]));
    },
    sheetjs: sheetjs!,
    exceljs,
  };
}
