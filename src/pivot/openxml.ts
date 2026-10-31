import { PivotTableDefinition, WorkSheet, Range } from '../types.js';
import { encode_cell, encode_range } from '../utils/address.js';
import { extractSheetData } from './compute.js';
import { escapeXml } from '../styles/index.js';

export interface OpenXmlPivotParts {
  cacheDefinitionXml: string;
  cacheRecordsXml: string;
  cacheDefRelsXml: string;
  pivotTableXml: string;
  pivotTableRelsXml: string;
}

export function generateOpenXmlPivotParts(
  sourceSheet: WorkSheet,
  pivot: PivotTableDefinition,
  id: number
): OpenXmlPivotParts {
  const { headers, rows } = extractSheetData(sourceSheet, pivot.sourceRange);
  const recordCount = rows.length;

  // Map header name to source index
  const headerIndexMap = new Map<string, number>();
  headers.forEach((h, idx) => headerIndexMap.set(h, idx));

  // Determine unique values for each field
  const fieldItemsMap = new Map<string, any[]>();
  for (const h of headers) {
    const uniqueVals = Array.from(new Set(rows.map((r) => r[h]).filter((v) => v !== undefined && v !== null)));
    fieldItemsMap.set(h, uniqueVals);
  }

  // 1. Pivot Cache Definition XML
  const cacheFieldsXml = headers
    .map((name) => {
      const vals = fieldItemsMap.get(name) || [];
      const isNum = vals.every((v) => typeof v === 'number');
      let itemsXml = '';
      if (vals.length > 0) {
        itemsXml = `<sharedItems count="${vals.length}">${vals
          .map((v) => {
            if (typeof v === 'number') {
              return `<n v="${v}"/>`;
            } else if (typeof v === 'boolean') {
              return `<b v="${v ? '1' : '0'}"/>`;
            } else {
              return `<s v="${escapeXml(String(v))}"/>`;
            }
          })
          .join('')}</sharedItems>`;
      } else {
        itemsXml = '<sharedItems/>';
      }
      return `<cacheField name="${escapeXml(name)}" numFmtId="0">${itemsXml}</cacheField>`;
    })
    .join('');

  const sourceRangeStr = encode_range(pivot.sourceRange);
  const cacheDefinitionXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshOnLoad="1" recordCount="${recordCount}"><cacheSource type="worksheet"><worksheetSource sheet="${escapeXml(
    pivot.sourceSheet
  )}" ref="${sourceRangeStr}"/></cacheSource><cacheFields count="${headers.length}">${cacheFieldsXml}</cacheFields></pivotCacheDefinition>`;

  // 2. Pivot Cache Records XML
  const recordsXml = rows
    .map((r) => {
      const cellsXml = headers
        .map((h) => {
          const v = r[h];
          if (v === undefined || v === null || v === '') {
            return '<m/>';
          }
          if (typeof v === 'number') {
            return `<n v="${v}"/>`;
          }
          if (typeof v === 'boolean') {
            return `<b v="${v ? '1' : '0'}"/>`;
          }
          return `<s v="${escapeXml(String(v))}"/>`;
        })
        .join('');
      return `<r>${cellsXml}</r>`;
    })
    .join('');

  const cacheRecordsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${recordCount}">${recordsXml}</pivotCacheRecords>`;

  // Cache def rels
  const cacheDefRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords${id}.xml"/></Relationships>`;

  // 3. Pivot Table Definition XML
  const rowFieldIndices = pivot.rows
    .map((rf) => headerIndexMap.get(rf))
    .filter((idx): idx is number => idx !== undefined);

  const colFieldIndices = pivot.cols
    .map((cf) => headerIndexMap.get(cf))
    .filter((idx): idx is number => idx !== undefined);

  const filterFieldIndices = pivot.filters
    .map((ff) => headerIndexMap.get(ff))
    .filter((idx): idx is number => idx !== undefined);

  const dataFieldsConfig = pivot.values.map((v) => {
    const fldIdx = headerIndexMap.get(v.field) ?? 0;
    const sub = v.subtotal || 'sum';
    const subLabel = sub.charAt(0).toUpperCase() + sub.slice(1);
    return {
      name: v.name || `${subLabel} of ${v.field}`,
      fld: fldIdx,
      subtotal: sub,
    };
  });

  const pivotFieldsXml = headers
    .map((h, idx) => {
      const isRow = rowFieldIndices.includes(idx);
      const isCol = colFieldIndices.includes(idx);
      const isFilter = filterFieldIndices.includes(idx);
      const isData = dataFieldsConfig.some((df) => df.fld === idx);

      const items = fieldItemsMap.get(h) || [];
      const itemElements = items.map((_, i) => `<item x="${i}"/>`).join('') + '<item t="default"/>';

      if (isRow) {
        return `<pivotField axis="axisRow" showAll="0"><items count="${items.length + 1}">${itemElements}</items></pivotField>`;
      }
      if (isCol) {
        return `<pivotField axis="axisCol" showAll="0"><items count="${items.length + 1}">${itemElements}</items></pivotField>`;
      }
      if (isFilter) {
        return `<pivotField axis="axisPage" showAll="0"/>`;
      }
      if (isData) {
        return `<pivotField data="1"/>`;
      }
      return '<pivotField/>';
    })
    .join('');

  const rowFieldsXml =
    rowFieldIndices.length > 0
      ? `<rowFields count="${rowFieldIndices.length}">${rowFieldIndices
          .map((idx) => `<field x="${idx}"/>`)
          .join('')}</rowFields>`
      : '';

  const colFieldsXml =
    colFieldIndices.length > 0
      ? `<colFields count="${colFieldIndices.length}">${colFieldIndices
          .map((idx) => `<field x="${idx}"/>`)
          .join('')}</colFields>`
      : '';

  const pageFieldsXml =
    filterFieldIndices.length > 0
      ? `<pageFields count="${filterFieldIndices.length}">${filterFieldIndices
          .map((idx) => `<pageField fld="${idx}"/>`)
          .join('')}</pageFields>`
      : '';

  const dataFieldsXml =
    dataFieldsConfig.length > 0
      ? `<dataFields count="${dataFieldsConfig.length}">${dataFieldsConfig
          .map((df) => `<dataField name="${escapeXml(df.name)}" fld="${df.fld}" subtotal="${df.subtotal}" baseField="0" baseItem="0"/>`)
          .join('')}</dataFields>`
      : '';

  const targetRangeStr = pivot.targetRange
    ? encode_range(pivot.targetRange)
    : `${encode_cell(pivot.targetCell)}:${encode_cell({
        c: pivot.targetCell.c + Math.max(colFieldIndices.length + dataFieldsConfig.length, 2),
        r: pivot.targetCell.r + Math.max(recordCount, 5),
      })}`;

  const styleName = pivot.styleName || 'PivotStyleLight16';

  const pivotTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${escapeXml(
    pivot.name
  )}" cacheId="${id}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Data"><location ref="${targetRangeStr}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="${
    headers.length
  }">${pivotFieldsXml}</pivotFields>${rowFieldsXml}${colFieldsXml}${pageFieldsXml}${dataFieldsXml}<pivotTableStyleInfo name="${styleName}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;

  const pivotTableRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${id}.xml"/></Relationships>`;

  return {
    cacheDefinitionXml,
    cacheRecordsXml,
    cacheDefRelsXml,
    pivotTableXml,
    pivotTableRelsXml,
  };
}
