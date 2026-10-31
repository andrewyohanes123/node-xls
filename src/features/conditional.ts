/**
 * Conditional formatting (<conditionalFormatting>/<cfRule>).
 *
 * Highlight rules reference a differential format (dxf) in styles.xml.
 * Text, blank, error and time-period rules carry the same helper formula
 * Excel writes, relative to the top-left cell of the first target range.
 */
import { CFValue, Color, ConditionalFormat, ConditionalFormatRule, CFTimePeriod } from '../types.js';
import { StyleBuilder, colorXml } from '../styles/index.js';
import { attrs, AttrValue, escapeAttr, escapeXml } from '../xlsx/xml.js';
import { decode_range, encode_cell } from '../utils/address.js';

const TIME_PERIOD_FORMULAS: Record<CFTimePeriod, (c: string) => string> = {
  today: (c) => `FLOOR(${c},1)=TODAY()`,
  yesterday: (c) => `FLOOR(${c},1)=TODAY()-1`,
  tomorrow: (c) => `FLOOR(${c},1)=TODAY()+1`,
  last7Days: (c) => `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`,
  thisMonth: (c) => `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`,
  lastMonth: (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`,
  nextMonth: (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`,
  thisWeek: (c) => `AND(TODAY()-ROUNDDOWN(${c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${c},0)-TODAY()<=7-WEEKDAY(TODAY()))`,
  lastWeek: (c) => `AND(TODAY()-ROUNDDOWN(${c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${c},0)<(WEEKDAY(TODAY())+7))`,
  nextWeek: (c) => `AND(ROUNDDOWN(${c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${c},0)-TODAY()<(15-WEEKDAY(TODAY())))`,
};

const ICON_COUNTS: Record<string, number> = { '3': 3, '4': 4, '5': 5 };

/** Formula text for a literal string inside a formula ("a""b"). */
export function formulaString(s: string): string {
  return `"${s.replace(/"/g, '""')}"`;
}

function formulaValue(v: string | number): string {
  return typeof v === 'number' ? String(v) : v.replace(/^=/, '');
}

function cfvoXml(v: CFValue): string {
  return `<cfvo${attrs([
    ['type', v.type],
    ['val', v.value !== undefined ? formulaValue(v.value) : undefined],
  ])}/>`;
}

function colorsXml(colors: Color[]): string {
  return colors.map((c) => colorXml('color', c)).join('');
}

function topLeft(sqref: string): string {
  const first = sqref.trim().split(/\s+/)[0];
  return encode_cell(decode_range(first.replace(/\$/g, '')).s);
}

function ruleXml(rule: ConditionalFormatRule, sqref: string, priority: number, styles: StyleBuilder): string {
  const base: Array<[string, AttrValue]> = [['type', rule.type]];
  let children = '';
  const cell = topLeft(sqref);
  const f = (x: string) => `<formula>${escapeXml(x)}</formula>`;

  switch (rule.type) {
    case 'colorScale': {
      if (rule.cfvo.length !== rule.colors.length || rule.cfvo.length < 2 || rule.cfvo.length > 3) {
        throw new Error('colorScale needs 2 or 3 cfvo entries and the same number of colors');
      }
      children = `<colorScale>${rule.cfvo.map(cfvoXml).join('')}${colorsXml(rule.colors)}</colorScale>`;
      return `<cfRule${attrs([...base, ['priority', priority]])}>${children}</cfRule>`;
    }
    case 'dataBar': {
      const min = rule.min ?? { type: 'min' };
      const max = rule.max ?? { type: 'max' };
      children = `<dataBar${attrs([['showValue', rule.showValue === false ? '0' : undefined]])}>${cfvoXml(min)}${cfvoXml(max)}${colorXml(
        'color',
        rule.color
      )}</dataBar>`;
      return `<cfRule${attrs([...base, ['priority', priority]])}>${children}</cfRule>`;
    }
    case 'iconSet': {
      const n = ICON_COUNTS[rule.iconSet.charAt(0)] ?? 3;
      const cfvo = rule.cfvo ?? Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((100 * i) / n) }));
      if (cfvo.length !== n) throw new Error(`iconSet ${rule.iconSet} needs ${n} cfvo entries`);
      children = `<iconSet${attrs([
        ['iconSet', rule.iconSet],
        ['showValue', rule.showValue === false ? '0' : undefined],
        ['reverse', rule.reverse],
      ])}>${cfvo.map(cfvoXml).join('')}</iconSet>`;
      return `<cfRule${attrs([...base, ['priority', priority]])}>${children}</cfRule>`;
    }
  }

  const dxfId = styles.registerDxf(rule.style);
  const a: Array<[string, AttrValue]> = [...base, ['dxfId', dxfId], ['priority', priority], ['stopIfTrue', rule.stopIfTrue]];
  switch (rule.type) {
    case 'cellIs':
      a.push(['operator', rule.operator]);
      children = rule.formulae.map((x) => f(formulaValue(x))).join('');
      if (rule.formulae.length !== (rule.operator === 'between' || rule.operator === 'notBetween' ? 2 : 1)) {
        throw new Error(`cellIs ${rule.operator} expects ${rule.operator.includes('etween') ? 2 : 1} formulae`);
      }
      break;
    case 'expression':
      children = f(formulaValue(rule.formula));
      break;
    case 'containsText': {
      const t = formulaString(rule.text);
      a.push(['operator', 'containsText'], ['text', rule.text]);
      children = f(`NOT(ISERROR(SEARCH(${t},${cell})))`);
      break;
    }
    case 'notContainsText': {
      const t = formulaString(rule.text);
      a.push(['operator', 'notContains'], ['text', rule.text]);
      children = f(`ISERROR(SEARCH(${t},${cell}))`);
      break;
    }
    case 'beginsWith': {
      const t = formulaString(rule.text);
      a.push(['operator', 'beginsWith'], ['text', rule.text]);
      children = f(`LEFT(${cell},LEN(${t}))=${t}`);
      break;
    }
    case 'endsWith': {
      const t = formulaString(rule.text);
      a.push(['operator', 'endsWith'], ['text', rule.text]);
      children = f(`RIGHT(${cell},LEN(${t}))=${t}`);
      break;
    }
    case 'top10':
      a.push(['percent', rule.percent], ['bottom', rule.bottom], ['rank', rule.rank]);
      break;
    case 'aboveAverage':
      a.push(['aboveAverage', rule.below ? '0' : undefined], ['equalAverage', rule.equalAverage], ['stdDev', rule.stdDev]);
      break;
    case 'containsBlanks':
      children = f(`LEN(TRIM(${cell}))=0`);
      break;
    case 'notContainsBlanks':
      children = f(`LEN(TRIM(${cell}))>0`);
      break;
    case 'containsErrors':
      children = f(`ISERROR(${cell})`);
      break;
    case 'notContainsErrors':
      children = f(`NOT(ISERROR(${cell}))`);
      break;
    case 'timePeriod':
      a.push(['timePeriod', rule.timePeriod]);
      children = f(TIME_PERIOD_FORMULAS[rule.timePeriod](cell));
      break;
    case 'duplicateValues':
    case 'uniqueValues':
      break;
  }
  return `<cfRule${attrs(a)}${children ? `>${children}</cfRule>` : '/>'}`;
}

export function conditionalFormattingXml(list: ConditionalFormat[], styles: StyleBuilder): string {
  const blocks = list.filter((cf) => cf.rules && cf.rules.length > 0);
  // Priorities are sheet-wide: rules with an explicit priority come first (in that order),
  // the rest follow in document order. The result is renumbered 1..n without gaps.
  const order = blocks
    .flatMap((cf, b) => cf.rules.map((r, i) => ({ key: `${b}:${i}`, p: r.priority ?? Infinity, seq: b * 1e6 + i })))
    .sort((a, b) => a.p - b.p || a.seq - b.seq);
  const priority = new Map(order.map((o, i) => [o.key, i + 1]));
  return blocks
    .map((cf, b) => {
      const sqref = cf.ref.replace(/\$/g, '').trim();
      const rules = cf.rules.map((r, i) => ruleXml(r, sqref, priority.get(`${b}:${i}`)!, styles)).join('');
      return `<conditionalFormatting sqref="${escapeAttr(sqref)}">${rules}</conditionalFormatting>`;
    })
    .join('');
}
