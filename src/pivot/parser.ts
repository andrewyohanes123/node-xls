import { XML_READ_OPTIONS } from '../xlsx/xml.js';
import { XMLParser } from 'fast-xml-parser';
import { PivotTableDefinition, PivotDataField, PivotSubtotal, PivotFilter } from '../types.js';
import { decode_range } from '../utils/address.js';
import { unescapeXml } from '../xlsx/xml.js';
import { isoToSerial } from '../utils/dates.js';

const parser = new XMLParser({
  ...XML_READ_OPTIONS,
  isArray: (name) => ['field', 'cacheField', 'dataField', 'pivotField', 'pageField', 'item'].includes(name),
});

type Shared = Array<string | number | boolean | null | { error: string }>;

/** Shared items of every cache field, in document order (types are interleaved, so parse by regex). */
function sharedItems(cacheXml: string, date1904: boolean): Shared[] {
  const out: Shared[] = [];
  for (const field of cacheXml.matchAll(/<(?:[\w.-]+:)?cacheField\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?cacheField>)/g)) {
    const items: Shared = [];
    const body = /<(?:[\w.-]+:)?sharedItems\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?sharedItems>/.exec(field[1] ?? '')?.[1] ?? '';
    for (const m of body.matchAll(/<(?:[\w.-]+:)?(s|n|b|d|m|e)\b(?:\s+v="([^"]*)")?[^>]*\/?>/g)) {
      const v = m[2] === undefined ? undefined : unescapeXml(m[2]);
      if (m[1] === 'm') items.push(null);
      else if (m[1] === 'n') items.push(Number(v));
      else if (m[1] === 'b') items.push(v === '1' || v === 'true');
      else if (m[1] === 'd') items.push(isoToSerial(v ?? '', { date1904 }) ?? v ?? ''); // dates are serials in node-xls
      else if (m[1] === 'e') items.push({ error: v ?? '#N/A' });
      else items.push(v ?? '');
    }
    out.push(items);
  }
  return out;
}

const isFalse = (v: unknown) => v === '0' || v === 'false';

export function parsePivotTableXml(
  pivotXml: string,
  cacheDefXml: string,
  targetSheetName: string,
  numFmtCode: (id: number) => string | undefined = () => undefined,
  date1904 = false
): PivotTableDefinition {
  const pivot = parser.parse(pivotXml)?.pivotTableDefinition || {};
  const cache = parser.parse(cacheDefXml)?.pivotCacheDefinition || {};

  const sourceType = cache.cacheSource?.['@_type'] ?? 'worksheet';
  const ws = (sourceType === 'worksheet' && cache.cacheSource?.worksheetSource) || {};
  const fieldNames: string[] = (cache.cacheFields?.cacheField || []).map((cf: any) => cf['@_name'] || '');
  const shared = sharedItems(cacheDefXml, date1904);
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
      const items: any[] = (pivotFields[fld]?.items?.item || []).filter((it: any) => it['@_t'] === undefined || it['@_t'] === 'data');
      let item = pf['@_item'] !== undefined ? items[Number(pf['@_item'])] : undefined;
      if (!item) {
        // LibreOffice expresses a single selection by hiding every other item
        const visible = items.filter((it) => !(it['@_h'] === '1' || it['@_h'] === 'true'));
        if (visible.length === 1 && visible.length < items.length) item = visible[0];
      }
      const x = item?.['@_x'];
      const value = x !== undefined ? shared[fld]?.[Number(x)] : undefined;
      if (value !== undefined && !(typeof value === 'object' && value !== null)) filter.selected = value;
      return filter;
    })
    .filter((f: PivotFilter | undefined): f is PivotFilter => f !== undefined);

  const values: PivotDataField[] = (pivot.dataFields?.dataField || []).map((df: any) => {
    const fld = Number(df['@_fld']);
    const field = name(fld) ?? `Field${fld}`;
    const subtotal = (df['@_subtotal'] || 'sum') as PivotSubtotal;
    const out: PivotDataField = { field, name: df['@_name'] ?? `${subtotal} of ${field}`, subtotal };
    const fmt = df['@_numFmtId'] !== undefined ? numFmtCode(Number(df['@_numFmtId'])) : undefined;
    if (fmt && fmt !== 'General') out.numFmt = fmt;
    return out;
  });

  const axisFields = pivotFields.filter((f) => f['@_axis'] === 'axisRow' || f['@_axis'] === 'axisCol');
  const def: PivotTableDefinition = {
    name: pivot['@_name'] || 'PivotTable',
    // a ref wins over a name; a name alone (table / defined name) is resolved by the reader
    sourceSheet: ws['@_ref'] ? ws['@_sheet'] || '' : '',
    sourceRange: decode_range(ws['@_ref'] || 'A1'),
    ...(ws['@_name'] && !ws['@_ref'] ? { sourceName: ws['@_name'] } : {}),
    targetSheet: targetSheetName,
    targetCell,
    targetRange,
    rows: fieldsOf(pivot.rowFields),
    cols: fieldsOf(pivot.colFields),
    values,
    filters,
    showRowTotals: !isFalse(pivot['@_rowGrandTotals']),
    showColTotals: !isFalse(pivot['@_colGrandTotals']),
    showSubtotals: !axisFields.length || axisFields.some((f) => !isFalse(f['@_defaultSubtotal'])),
  };
  const style = pivot.pivotTableStyleInfo?.['@_name'];
  if (style) def.styleName = style;
  return def;
}
