import { XMLParser } from 'fast-xml-parser';
import { PivotTableDefinition, PivotDataField, PivotSubtotal, PivotFilter } from '../types.js';
import { decode_range } from '../utils/address.js';
import { unescapeXml } from '../xlsx/xml.js';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  trimValues: false,
  isArray: (name) => ['field', 'cacheField', 'dataField', 'pivotField', 'pageField', 'item'].includes(name),
});

type Shared = Array<string | number | boolean | null>;

/** Shared items of every cache field, in document order (types are interleaved, so parse by regex). */
function sharedItems(cacheXml: string): Shared[] {
  const out: Shared[] = [];
  for (const field of cacheXml.matchAll(/<cacheField\b[^>]*?(?:\/>|>([\s\S]*?)<\/cacheField>)/g)) {
    const items: Shared = [];
    const body = /<sharedItems\b[^>]*>([\s\S]*?)<\/sharedItems>/.exec(field[1] ?? '')?.[1] ?? '';
    for (const m of body.matchAll(/<(s|n|b|d|m|e)\b(?:\s+v="([^"]*)")?[^>]*\/?>/g)) {
      const v = m[2] === undefined ? undefined : unescapeXml(m[2]);
      if (m[1] === 'm') items.push(null);
      else if (m[1] === 'n') items.push(Number(v));
      else if (m[1] === 'b') items.push(v === '1' || v === 'true');
      else items.push(v ?? '');
    }
    out.push(items);
  }
  return out;
}

export function parsePivotTableXml(pivotXml: string, cacheDefXml: string, targetSheetName: string): PivotTableDefinition {
  const pivot = parser.parse(pivotXml)?.pivotTableDefinition || {};
  const cache = parser.parse(cacheDefXml)?.pivotCacheDefinition || {};

  const ws = cache.cacheSource?.worksheetSource || {};
  const fieldNames: string[] = (cache.cacheFields?.cacheField || []).map((cf: any) => cf['@_name'] || '');
  const shared = sharedItems(cacheDefXml);
  const pivotFields: any[] = pivot.pivotFields?.pivotField || [];
  const name = (idx: number) => (idx >= 0 && fieldNames[idx] !== undefined ? fieldNames[idx] : undefined);

  const location = pivot.location || {};
  const targetRange = decode_range(location['@_ref'] || 'A1');
  const pageRows = Number(location['@_rowPageCount'] || 0);
  const targetCell = { r: pageRows ? Math.max(0, targetRange.s.r - pageRows - 1) : targetRange.s.r, c: targetRange.s.c };

  const fieldsOf = (node: any): string[] =>
    (node?.field || [])
      .map((f: any) => Number(f['@_x']))
      .filter((x: number) => x >= 0)
      .map(name)
      .filter((n: string | undefined): n is string => n !== undefined);

  const filters: PivotFilter[] = (pivot.pageFields?.pageField || [])
    .map((pf: any) => {
      const fld = Number(pf['@_fld']);
      const field = name(fld);
      if (field === undefined) return undefined;
      const filter: PivotFilter = { field };
      if (pf['@_item'] !== undefined) {
        const item = (pivotFields[fld]?.items?.item || [])[Number(pf['@_item'])];
        const x = item?.['@_x'];
        const value = x !== undefined ? shared[fld]?.[Number(x)] : undefined;
        if (value !== undefined && value !== null) filter.selected = value;
      }
      return filter;
    })
    .filter((f: PivotFilter | undefined): f is PivotFilter => f !== undefined);

  const values: PivotDataField[] = (pivot.dataFields?.dataField || []).map((df: any) => {
    const fld = Number(df['@_fld']);
    const field = name(fld) ?? `Field${fld}`;
    const subtotal = (df['@_subtotal'] || 'sum') as PivotSubtotal;
    return { field, name: df['@_name'] ?? `${subtotal} of ${field}`, subtotal };
  });

  const axisFields = pivotFields.filter((f) => f['@_axis'] === 'axisRow' || f['@_axis'] === 'axisCol');
  const def: PivotTableDefinition = {
    name: pivot['@_name'] || 'PivotTable',
    sourceSheet: ws['@_sheet'] || '',
    sourceRange: decode_range(ws['@_ref'] || 'A1'),
    targetSheet: targetSheetName,
    targetCell,
    targetRange,
    rows: fieldsOf(pivot.rowFields),
    cols: fieldsOf(pivot.colFields),
    values,
    filters,
    showRowTotals: pivot['@_rowGrandTotals'] !== '0',
    showColTotals: pivot['@_colGrandTotals'] !== '0',
    showSubtotals: !axisFields.length || axisFields.some((f) => f['@_defaultSubtotal'] !== '0'),
  };
  const style = pivot.pivotTableStyleInfo?.['@_name'];
  if (style) def.styleName = style;
  return def;
}
