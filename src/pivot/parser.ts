import { XMLParser } from 'fast-xml-parser';
import { PivotTableDefinition, CellAddress, PivotDataField, PivotSubtotal, PivotFilter } from '../types.js';
import { decode_range } from '../utils/address.js';

export function parsePivotTableXml(
  pivotXml: string,
  cacheDefXml: string,
  targetSheetName: string
): PivotTableDefinition {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseTagValue: false,
    trimValues: false,
    isArray: (name) => {
      return ['field', 'cacheField', 'dataField', 'pivotField', 'pageField'].includes(name);
    },
  });

  const parsedPivot = parser.parse(pivotXml)?.pivotTableDefinition || {};
  const parsedCache = parser.parse(cacheDefXml)?.pivotCacheDefinition || {};

  // Extract source
  const wsSource = parsedCache.cacheSource?.worksheetSource || {};
  const sourceSheet = wsSource['@_sheet'] || '';
  const sourceRefStr = wsSource['@_ref'] || 'A1';
  const sourceRange = decode_range(sourceRefStr);

  // Extract cache field names
  const cacheFields = parsedCache.cacheFields?.cacheField || [];
  const fieldNames: string[] = cacheFields.map((cf: any) => cf['@_name'] || '');

  // Extract layout
  const location = parsedPivot.location || {};
  const targetRefStr = location['@_ref'] || 'A1:B2';
  const targetRange = decode_range(targetRefStr);
  const targetCell: CellAddress = targetRange.s;

  const rowFieldNodes = parsedPivot.rowFields?.field || [];
  const rows: string[] = rowFieldNodes
    .map((rf: any) => {
      const idx = Number(rf['@_x']);
      return !isNaN(idx) && fieldNames[idx] ? fieldNames[idx] : undefined;
    })
    .filter((f: any): f is string => Boolean(f));

  const colFieldNodes = parsedPivot.colFields?.field || [];
  const cols: string[] = colFieldNodes
    .map((cf: any) => {
      const idx = Number(cf['@_x']);
      return !isNaN(idx) && fieldNames[idx] ? fieldNames[idx] : undefined;
    })
    .filter((f: any): f is string => Boolean(f));

  const pageFieldNodes = parsedPivot.pageFields?.pageField || [];
  const filters: PivotFilter[] = pageFieldNodes
    .map((pf: any) => {
      const idx = Number(pf['@_fld']);
      return !isNaN(idx) && fieldNames[idx] ? { field: fieldNames[idx] } : undefined;
    })
    .filter((f: any): f is PivotFilter => Boolean(f));

  const dataFieldNodes = parsedPivot.dataFields?.dataField || [];
  const values: PivotDataField[] = dataFieldNodes.map((df: any) => {
    const fldIdx = Number(df['@_fld']);
    const fieldName = !isNaN(fldIdx) && fieldNames[fldIdx] ? fieldNames[fldIdx] : `Field${fldIdx}`;
    const subtotal = (df['@_subtotal'] || 'sum') as PivotSubtotal;
    const subLabel = subtotal.charAt(0).toUpperCase() + subtotal.slice(1);
    const name = df['@_name'] || `${subLabel} of ${fieldName}`;
    return {
      field: fieldName,
      name,
      subtotal,
    };
  });

  const styleInfo = parsedPivot.pivotTableStyleInfo || {};
  const styleName = styleInfo['@_name'];

  return {
    name: parsedPivot['@_name'] || 'PivotTable',
    sourceSheet,
    sourceRange,
    targetSheet: targetSheetName,
    targetCell,
    targetRange,
    rows,
    cols,
    values,
    filters,
    styleName,
  };
}
