/**
 * Pivot cache definition, cache records and pivot table definition parts.
 *
 * - Axis and page fields list their unique values as shared items (source
 *   order) and records reference them with <x v>; other fields are stored
 *   inline. Type flags describe each field's values as Excel expects.
 * - pivotField <items> are in display (sorted) order; rowItems/colItems index
 *   into them and mirror the layout used for the pre-rendered cells.
 * - refreshOnLoad lets Excel rebuild the cache from the source range.
 */
import { PivotTableDefinition, WorkSheet } from '../types.js';
import { escapeAttr, XML_HEADER, NS } from '../xlsx/xml.js';
import { encode_range } from '../utils/address.js';
import { StyleBuilder } from '../styles/index.js';
import { readSource, PivotValue, SourceData } from './source.js';
import { AxisEntry, PivotLayout, buildLayout, itemKey } from './layout.js';

export interface OpenXmlPivotParts {
  cacheDefinitionXml: string;
  cacheRecordsXml: string;
  pivotTableXml: string;
  layout: PivotLayout;
}

const DAY_MS = 86_400_000;

function isoDate(serial: number): string {
  const s = serial < 61 ? serial + 1 : serial;
  return new Date(Math.round(s * DAY_MS) + Date.UTC(1899, 11, 30)).toISOString().slice(0, 19);
}

function valueXml(v: PivotValue, isDate: boolean): string {
  if (v === null) return '<m/>';
  if (typeof v === 'number') return isDate ? `<d v="${isoDate(v)}"/>` : `<n v="${v}"/>`;
  if (typeof v === 'boolean') return `<b v="${v ? 1 : 0}"/>`;
  return `<s v="${escapeAttr(v)}"/>`;
}

function isDateField(source: SourceData, field: number): boolean {
  const values = source.records.map((r) => r[field]);
  return (
    source.dateFields.has(field) &&
    values.some((v) => typeof v === 'number') &&
    !values.some((v) => typeof v === 'string' || typeof v === 'boolean')
  );
}

function sharedItemsXml(source: SourceData, field: number, listed: PivotValue[] | undefined): string {
  const values = source.records.map((r) => r[field]);
  const strings = values.filter((v): v is string => typeof v === 'string');
  const numbers = values.filter((v): v is number => typeof v === 'number');
  const hasBool = values.some((v) => typeof v === 'boolean');
  const hasBlank = values.some((v) => v === null);
  const isDate = isDateField(source, field);
  const kinds = [strings.length > 0, numbers.length > 0, hasBool].filter(Boolean).length;
  const min = numbers.length ? Math.min(...numbers) : 0;
  const max = numbers.length ? Math.max(...numbers) : 0;
  const a: Array<[string, string | number | undefined]> = [
    ['containsSemiMixedTypes', strings.length === 0 && !hasBlank ? '0' : undefined],
    ['containsNonDate', isDate ? '0' : undefined],
    ['containsDate', isDate ? '1' : undefined],
    ['containsString', strings.length === 0 ? '0' : undefined],
    ['containsBlank', hasBlank ? '1' : undefined],
    ['containsMixedTypes', kinds > 1 ? '1' : undefined],
    ['containsNumber', numbers.length && !isDate ? '1' : undefined],
    ['containsInteger', numbers.length && !isDate && numbers.every(Number.isInteger) ? '1' : undefined],
    ['minValue', numbers.length && !isDate ? min : undefined],
    ['maxValue', numbers.length && !isDate ? max : undefined],
    ['minDate', isDate ? isoDate(min) : undefined],
    ['maxDate', isDate ? isoDate(max + 1) : undefined],
    ['count', listed ? listed.length : undefined],
    ['longText', strings.some((s) => s.length > 255) ? '1' : undefined],
  ];
  const attrs = a
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => ` ${k}="${v}"`)
    .join('');
  if (!listed) return `<sharedItems${attrs}/>`;
  return `<sharedItems${attrs}>${listed.map((v) => valueXml(v, isDate)).join('')}</sharedItems>`;
}

/** <i> elements for an axis: r = members repeated from the previous item, i = data field index. */
function axisItemsXml(tag: 'rowItems' | 'colItems', entries: AxisEntry[]): string {
  let prev: number[] = [];
  const out = entries.map((e) => {
    if (e.type === 'grand') {
      prev = [];
      return `<i t="grand"${e.data ? ` i="${e.data}"` : ''}><x/></i>`;
    }
    let r = 0;
    while (r < e.path.length - 1 && r < prev.length && e.path[r] === prev[r]) r++;
    const xs = e.path
      .slice(r)
      .map((v) => (v ? `<x v="${v}"/>` : '<x/>'))
      .join('');
    const attrs = `${e.type === 'subtotal' ? ' t="default"' : ''}${r ? ` r="${r}"` : ''}${e.data ? ` i="${e.data}"` : ''}`;
    prev = e.type === 'item' ? e.path : [];
    return xs ? `<i${attrs}>${xs}</i>` : `<i${attrs}/>`;
  });
  return `<${tag} count="${out.length}">${out.join('')}</${tag}>`;
}

export function generateOpenXmlPivotParts(
  sourceSheet: WorkSheet,
  pivot: PivotTableDefinition,
  id: number,
  recordsRid: string,
  styles?: StyleBuilder
): OpenXmlPivotParts {
  const source = readSource(sourceSheet, pivot.sourceRange);
  const layout = buildLayout(pivot, source);
  const { fields, records } = source;
  const axisSet = new Set([...layout.rowFields, ...layout.colFields, ...layout.pageFields.map((p) => p.field)]);

  // cache definition + records
  const cacheFields = fields
    .map((name, f) => {
      const numFmtId = isDateField(source, f) && styles ? styles.registerNumFmt(source.formats[f] ?? 14) : 0;
      return `<cacheField name="${escapeAttr(name)}" numFmtId="${numFmtId}">${sharedItemsXml(source, f, axisSet.has(f) ? layout.shared.get(f) : undefined)}</cacheField>`;
    })
    .join('');
  const cacheDefinitionXml =
    `${XML_HEADER}<pivotCacheDefinition xmlns="${NS.main}" xmlns:r="${NS.r}" r:id="${recordsRid}" refreshOnLoad="1" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="${records.length}">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${encode_range(source.range)}" sheet="${escapeAttr(pivot.sourceSheet)}"/></cacheSource>` +
    `<cacheFields count="${fields.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;

  const dateFields = fields.map((_, f) => isDateField(source, f));
  const recordXml = records
    .map(
      (rec) =>
        `<r>${rec.map((v, f) => (axisSet.has(f) ? `<x v="${layout.sharedIndex.get(f)!.get(itemKey(v))}"/>` : valueXml(v, dateFields[f]))).join('')}</r>`
    )
    .join('');
  const cacheRecordsXml = `${XML_HEADER}<pivotCacheRecords xmlns="${NS.main}" xmlns:r="${NS.r}" count="${records.length}">${recordXml}</pivotCacheRecords>`;

  // pivot table definition
  const dataFieldSet = new Set(layout.dataFields.map((d) => d.field));
  const subtotalsOff = !layout.showSubtotals;
  const pivotFields = fields
    .map((_, f) => {
      const axis = layout.rowFields.includes(f)
        ? 'axisRow'
        : layout.colFields.includes(f)
          ? 'axisCol'
          : layout.pageFields.some((p) => p.field === f)
            ? 'axisPage'
            : undefined;
      const common = `${axis ? ` axis="${axis}"` : ''}${dataFieldSet.has(f) ? ' dataField="1"' : ''} compact="0" outline="0" showAll="0"`;
      if (!axis) return `<pivotField${common}/>`;
      const display = layout.items.get(f)!;
      const shared = layout.sharedIndex.get(f)!;
      const noDefault = subtotalsOff && axis !== 'axisPage';
      const itemsXml = display.map((v) => `<item x="${shared.get(itemKey(v))}"/>`).join('') + (noDefault ? '' : '<item t="default"/>');
      return `<pivotField${common}${noDefault ? ' defaultSubtotal="0"' : ''}><items count="${display.length + (noDefault ? 0 : 1)}">${itemsXml}</items></pivotField>`;
    })
    .join('');

  const g = layout.geometry;
  const location = `<location ref="${encode_range(g.location)}" firstHeaderRow="${g.firstHeaderRow}" firstDataRow="${g.firstDataRow}" firstDataCol="${g.firstDataCol}"${
    layout.pageFields.length ? ` rowPageCount="${layout.pageFields.length}" colPageCount="1"` : ''
  }/>`;
  const rowFieldsXml = `<rowFields count="${layout.rowFields.length}">${layout.rowFields.map((f) => `<field x="${f}"/>`).join('')}</rowFields>`;
  const colXs = layout.valuesOnCols ? [...layout.colFields, -2] : layout.colFields;
  const colFieldsXml = colXs.length ? `<colFields count="${colXs.length}">${colXs.map((f) => `<field x="${f}"/>`).join('')}</colFields>` : '';
  const pageFieldsXml = layout.pageFields.length
    ? `<pageFields count="${layout.pageFields.length}">${layout.pageFields
        .map((p) => `<pageField fld="${p.field}"${p.selected !== undefined ? ` item="${p.selected}"` : ''} hier="-1"/>`)
        .join('')}</pageFields>`
    : '';
  const dataFieldsXml = `<dataFields count="${layout.dataFields.length}">${layout.dataFields
    .map((d) => {
      const numFmtId = d.numFmt && styles ? styles.registerNumFmt(d.numFmt) : undefined;
      return `<dataField name="${escapeAttr(d.name)}" fld="${d.field}"${d.subtotal !== 'sum' ? ` subtotal="${d.subtotal}"` : ''} baseField="0" baseItem="0"${
        numFmtId !== undefined ? ` numFmtId="${numFmtId}"` : ''
      }/>`;
    })
    .join('')}</dataFields>`;

  const pivotTableXml =
    `${XML_HEADER}<pivotTableDefinition xmlns="${NS.main}" name="${escapeAttr(pivot.name)}" cacheId="${id}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="3" minRefreshableVersion="3" useAutoFormatting="1"${
      pivot.showColTotals === false ? ' colGrandTotals="0"' : ''
    }${pivot.showRowTotals === false ? ' rowGrandTotals="0"' : ''} itemPrintTitles="1" createdVersion="3" indent="0" compact="0" compactData="0" multipleFieldFilters="0">` +
    location +
    `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>` +
    rowFieldsXml +
    axisItemsXml('rowItems', layout.rowAxis) +
    colFieldsXml +
    axisItemsXml('colItems', layout.colAxis) +
    pageFieldsXml +
    dataFieldsXml +
    `<pivotTableStyleInfo name="${escapeAttr(pivot.styleName || 'PivotStyleLight16')}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>` +
    '</pivotTableDefinition>';

  return { cacheDefinitionXml, cacheRecordsXml, pivotTableXml, layout };
}
