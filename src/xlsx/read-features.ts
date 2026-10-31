import { XML_READ_OPTIONS } from './xml.js';
/**
 * Parsers for worksheet features: conditional formatting, data validation,
 * tables and comments. Their output matches the shapes accepted by the writer
 * so read -> write round-trips.
 */
import { XMLParser } from 'fast-xml-parser';
import {
  CFStyle,
  CFValue,
  Comment,
  ConditionalFormat,
  ConditionalFormatRule,
  DataValidation,
  TableDefinition,
  TableColumn,
} from '../types.js';
import { parseColorNode } from '../styles/parser.js';
import { decodeXstring } from './xml.js';

export const featureParser = new XMLParser({
  ...XML_READ_OPTIONS,
  isArray: (name) => ['tableColumn', 'author', 'comment', 'r'].includes(name),
});

/** Text content of a parsed node ('' when absent). */
export function textOf(node: any): string {
  if (node === undefined || node === null) return '';
  if (typeof node === 'object') return node['#text'] !== undefined ? String(node['#text']) : '';
  return String(node);
}

const bool = (v: any) => v === '1' || v === 'true';
const isFalse = (v: any) => v === '0' || v === 'false';
/** Plain-JSON copy so rules/cells never share mutable objects. */
const clone = <T>(o: T): T => (o && typeof o === 'object' ? (JSON.parse(JSON.stringify(o)) as T) : o);
const numOrString = (s: string): string | number => (s.trim() !== '' && Number.isFinite(Number(s)) ? Number(s) : s);
const asArray = <T>(v: T | T[] | undefined): T[] => (v === undefined ? [] : Array.isArray(v) ? v : [v]);

function parseCfvo(n: any): CFValue {
  const v: CFValue = { type: n['@_type'] };
  if (n['@_val'] !== undefined) v.value = numOrString(n['@_val']);
  return v;
}

export function parseConditionalFormatting(nodes: any[], dxfs: CFStyle[]): ConditionalFormat[] {
  const out: ConditionalFormat[] = [];
  for (const cf of nodes) {
    const rules: ConditionalFormatRule[] = [];
    const ruleNodes = asArray<any>(cf.cfRule).sort((a, b) => Number(a['@_priority'] ?? 0) - Number(b['@_priority'] ?? 0));
    for (const r of ruleNodes) {
      const type = r['@_type'];
      const style: CFStyle = r['@_dxfId'] !== undefined ? clone(dxfs[Number(r['@_dxfId'])] ?? {}) : {};
      const formulas = asArray(r.formula).map(textOf);
      const stop = {
        ...(bool(r['@_stopIfTrue']) ? { stopIfTrue: true } : {}),
        ...(r['@_priority'] !== undefined ? { priority: Number(r['@_priority']) } : {}),
      };
      const prio = r['@_priority'] !== undefined ? { priority: Number(r['@_priority']) } : {};
      switch (type) {
        case 'cellIs':
          rules.push({ type, operator: r['@_operator'], formulae: formulas.map(numOrString), style, ...stop });
          break;
        case 'expression':
          rules.push({ type, formula: formulas[0] ?? '', style, ...stop });
          break;
        case 'containsText':
        case 'notContainsText':
        case 'beginsWith':
        case 'endsWith':
          // without a text attribute the rule is defined by its formula alone
          if (r['@_text'] === undefined) rules.push({ type: 'expression', formula: formulas[0] ?? 'FALSE', style, ...stop });
          else rules.push({ type, text: r['@_text'], style, ...stop });
          break;
        case 'top10':
          rules.push({
            type,
            rank: Number(r['@_rank'] ?? 10),
            ...(bool(r['@_bottom']) ? { bottom: true } : {}),
            ...(bool(r['@_percent']) ? { percent: true } : {}),
            style,
            ...stop,
          });
          break;
        case 'aboveAverage':
          rules.push({
            type,
            ...(isFalse(r['@_aboveAverage']) ? { below: true } : {}),
            ...(bool(r['@_equalAverage']) ? { equalAverage: true } : {}),
            ...(r['@_stdDev'] !== undefined ? { stdDev: Number(r['@_stdDev']) } : {}),
            style,
            ...stop,
          });
          break;
        case 'duplicateValues':
        case 'uniqueValues':
        case 'containsBlanks':
        case 'notContainsBlanks':
        case 'containsErrors':
        case 'notContainsErrors':
          rules.push({ type, style, ...stop });
          break;
        case 'timePeriod':
          rules.push({ type, timePeriod: r['@_timePeriod'], style, ...stop });
          break;
        case 'colorScale': {
          const cs = r.colorScale ?? {};
          rules.push({ type, cfvo: asArray(cs.cfvo).map(parseCfvo), colors: asArray(cs.color).map((c) => parseColorNode(c)!).filter(Boolean), ...prio });
          break;
        }
        case 'dataBar': {
          const db = r.dataBar ?? {};
          const cfvo = asArray(db.cfvo).map(parseCfvo);
          rules.push({
            type,
            color: parseColorNode(asArray(db.color)[0]) ?? { rgb: '638EC6' },
            min: cfvo[0],
            max: cfvo[1],
            ...(isFalse(db['@_showValue']) ? { showValue: false } : {}),
            ...prio,
          });
          break;
        }
        case 'iconSet': {
          const is = r.iconSet ?? {};
          rules.push({
            type,
            iconSet: is['@_iconSet'] ?? '3TrafficLights1',
            cfvo: asArray(is.cfvo).map(parseCfvo),
            ...(bool(is['@_reverse']) ? { reverse: true } : {}),
            ...(isFalse(is['@_showValue']) ? { showValue: false } : {}),
            ...prio,
          });
          break;
        }
        default:
          break; // unsupported rule types are dropped
      }
    }
    if (rules.length) out.push({ ref: cf['@_sqref'], rules });
  }
  return out;
}

function splitList(f: string): string[] | undefined {
  if (!(f.startsWith('"') && f.endsWith('"'))) return undefined;
  return f
    .slice(1, -1)
    .replace(/""/g, '"')
    .split(',');
}

export function parseDataValidations(nodes: any[]): DataValidation[] {
  return nodes.map((n) => {
    const type = n['@_type'] === undefined || n['@_type'] === 'none' ? 'any' : n['@_type'];
    const dv: DataValidation = { ref: n['@_sqref'], type };
    const f1 = n.formula1 !== undefined ? textOf(n.formula1) : undefined;
    const f2 = n.formula2 !== undefined ? textOf(n.formula2) : undefined;
    const list = type === 'list' && f1 !== undefined ? splitList(f1) : undefined;
    if (list) dv.values = list;
    else if (f1 !== undefined) dv.formula1 = f1;
    if (f2 !== undefined) dv.formula2 = f2;
    if (n['@_operator'] && type !== 'list') dv.operator = n['@_operator'];
    // the schema default for allowBlank is false; node-xls defaults it to true
    if (!bool(n['@_allowBlank'])) dv.allowBlank = false;
    if (bool(n['@_showDropDown'])) dv.hideDropDown = true;
    // both default to false in the schema
    if (!bool(n['@_showErrorMessage'])) dv.showErrorMessage = false;
    if (!bool(n['@_showInputMessage'])) dv.showInputMessage = false;
    if (n['@_prompt'] !== undefined || n['@_promptTitle'] !== undefined) {
      dv.prompt = { message: n['@_prompt'] ?? '', ...(n['@_promptTitle'] !== undefined ? { title: n['@_promptTitle'] } : {}) };
    }
    if (n['@_error'] !== undefined || n['@_errorTitle'] !== undefined || n['@_errorStyle'] !== undefined) {
      dv.error = {
        message: n['@_error'] ?? '',
        ...(n['@_errorTitle'] !== undefined ? { title: n['@_errorTitle'] } : {}),
        ...(n['@_errorStyle'] ? { style: n['@_errorStyle'] } : {}),
      };
    }
    return dv;
  });
}

export function parseTableXml(xml: string): TableDefinition {
  const t = featureParser.parse(xml).table ?? {};
  const columns: TableColumn[] = asArray<any>(t.tableColumns?.tableColumn).map((c) => {
    const col: TableColumn = { name: decodeXstring(c['@_name'] ?? '') };
    if (c['@_totalsRowFunction']) col.totalsRowFunction = c['@_totalsRowFunction'];
    if (c.totalsRowFormula !== undefined) col.totalsRowFormula = textOf(c.totalsRowFormula);
    if (c['@_totalsRowLabel'] !== undefined) col.totalsRowLabel = c['@_totalsRowLabel'];
    if (c.calculatedColumnFormula !== undefined) col.calculatedColumnFormula = textOf(c.calculatedColumnFormula);
    return col;
  });
  const si = t.tableStyleInfo ?? {};
  const def: TableDefinition = { name: t['@_name'], ref: t['@_ref'], columns };
  if (t['@_displayName'] && t['@_displayName'] !== t['@_name']) def.displayName = t['@_displayName'];
  if (t['@_headerRowCount'] === '0') def.headerRow = false;
  if (Number(t['@_totalsRowCount'] ?? 0) > 0) def.totalsRow = true;
  if (si['@_name']) def.style = si['@_name'];
  else if (t.tableStyleInfo) def.style = null; // explicitly unstyled
  if (si['@_showRowStripes'] !== undefined && !bool(si['@_showRowStripes'])) def.showRowStripes = false;
  if (bool(si['@_showColumnStripes'])) def.showColumnStripes = true;
  if (bool(si['@_showFirstColumn'])) def.showFirstColumn = true;
  if (bool(si['@_showLastColumn'])) def.showLastColumn = true;
  if (!t.autoFilter && def.headerRow !== false) def.autoFilter = false;
  return def;
}

function richText(node: any): string {
  if (!node) return '';
  if (node.t !== undefined) return decodeXstring(textOf(node.t));
  return asArray<any>(node.r)
    .map((r) => decodeXstring(textOf(r.t)))
    .join('');
}

/** Comments by cell ref (a cell may carry several <comment> elements). `visible` marks always-shown notes (from VML). */
export function parseCommentsXml(xml: string, vml?: string): Map<string, { comments: Comment[]; visible: boolean }> {
  const root = featureParser.parse(xml).comments ?? {};
  const authors: string[] = asArray<any>(root.authors?.author).map(textOf);
  const visible = new Set<string>();
  if (vml) {
    for (const m of vml.matchAll(/<x:ClientData[^>]*>([\s\S]*?)<\/x:ClientData>/g)) {
      if (!/<x:Visible\s*\/?>/.test(m[1])) continue;
      const r = /<x:Row>(\d+)<\/x:Row>/.exec(m[1]);
      const c = /<x:Column>(\d+)<\/x:Column>/.exec(m[1]);
      if (r && c) visible.add(`${r[1]},${c[1]}`);
    }
  }
  const out = new Map<string, { comments: Comment[]; visible: boolean }>();
  for (const c of asArray<any>(root.commentList?.comment)) {
    const ref = c['@_ref'];
    const author = authors[Number(c['@_authorId'] ?? 0)];
    const comment: Comment = { t: richText(c.text) };
    if (author !== undefined && author !== '') comment.a = author;
    const entry = out.get(ref);
    if (entry) entry.comments.push(comment);
    else out.set(ref, { comments: [comment], visible: false });
  }
  if (visible.size) {
    for (const [ref, entry] of out) {
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!m) continue;
      const col = m[1].split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      if (visible.has(`${Number(m[2]) - 1},${col}`)) entry.visible = true;
    }
  }
  return out;
}
