/**
 * XML primitives shared by every part writer.
 *
 * Two escaping rules matter for OOXML:
 *  - XML 1.0 forbids most C0 control characters, even as character references.
 *  - SpreadsheetML text (ST_Xstring) encodes such characters as `_xHHHH_`; a
 *    literal `_xHHHH_` in user text must itself be escaped as `_x005F_xHHHH_`.
 */

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export const NS = {
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
};

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const REL_TYPES = {
  officeDocument: `${REL}/officeDocument`,
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProps: `${REL}/extended-properties`,
  worksheet: `${REL}/worksheet`,
  styles: `${REL}/styles`,
  sharedStrings: `${REL}/sharedStrings`,
  theme: `${REL}/theme`,
  pivotTable: `${REL}/pivotTable`,
  pivotCacheDefinition: `${REL}/pivotCacheDefinition`,
  pivotCacheRecords: `${REL}/pivotCacheRecords`,
  table: `${REL}/table`,
  comments: `${REL}/comments`,
  vmlDrawing: `${REL}/vmlDrawing`,
  drawing: `${REL}/drawing`,
  image: `${REL}/image`,
  chart: `${REL}/chart`,
  hyperlink: `${REL}/hyperlink`,
};

const CT = 'application/vnd.openxmlformats-officedocument';
export const CONTENT_TYPES = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  workbook: `${CT}.spreadsheetml.sheet.main+xml`,
  worksheet: `${CT}.spreadsheetml.worksheet+xml`,
  styles: `${CT}.spreadsheetml.styles+xml`,
  sharedStrings: `${CT}.spreadsheetml.sharedStrings+xml`,
  theme: `${CT}.theme+xml`,
  coreProps: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProps: `${CT}.extended-properties+xml`,
  pivotTable: `${CT}.spreadsheetml.pivotTable+xml`,
  pivotCacheDefinition: `${CT}.spreadsheetml.pivotCacheDefinition+xml`,
  pivotCacheRecords: `${CT}.spreadsheetml.pivotCacheRecords+xml`,
  table: `${CT}.spreadsheetml.table+xml`,
  comments: `${CT}.spreadsheetml.comments+xml`,
  vml: `${CT}.vmlDrawing`,
  drawing: `${CT}.drawing+xml`,
  chart: `${CT}.drawingml.chart+xml`,
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
};

// Characters not allowed anywhere in an XML 1.0 document (lone surrogates included).
const INVALID_XML_CHARS =
  // eslint-disable-next-line no-control-regex
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/** Escapes text content. Characters illegal in XML are dropped. */
export function escapeXml(value: unknown): string {
  return String(value)
    .replace(INVALID_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Escapes an attribute value (also preserves tabs/newlines). */
export function escapeAttr(value: unknown): string {
  return escapeXml(value)
    .replace(/"/g, '&quot;')
    .replace(/\t/g, '&#9;')
    .replace(/\n/g, '&#10;')
    .replace(/\r/g, '&#13;');
}

/** Encodes SpreadsheetML text (ST_Xstring): control chars become `_xHHHH_`. */
export function encodeXstring(value: unknown): string {
  const s = String(value)
    .replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
    .replace(INVALID_XML_CHARS, (ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`)
    // XML parsers normalise CR/CRLF to LF, so Excel stores CR escaped
    .replace(/\r/g, '_x000D_');
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Decodes the predefined XML entities and numeric character references. */
export function unescapeXml(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e: string) => {
    switch (e) {
      case 'amp':
        return '&';
      case 'lt':
        return '<';
      case 'gt':
        return '>';
      case 'quot':
        return '"';
      case 'apos':
        return "'";
    }
    const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return Number.isFinite(code) ? String.fromCodePoint(code) : m;
  });
}

/** Decodes `_xHHHH_` escapes produced by Excel (single pass, like Excel). */
export function decodeXstring(value: string): string {
  if (value.indexOf('_x') === -1) return value;
  return value.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

/** True when the string needs xml:space="preserve" to survive parsing. */
export function needsPreserve(s: string): boolean {
  return /^\s|\s$|[\t\n\r]/.test(s);
}

/** `<t>` element with correct whitespace preservation. */
export function textElement(value: string, tag = 't'): string {
  return `<${tag}${needsPreserve(value) ? ' xml:space="preserve"' : ''}>${encodeXstring(value)}</${tag}>`;
}

export type AttrValue = string | number | boolean | undefined | null;

/**
 * Serialises attributes in the given order. `undefined`/`null`/`false` values
 * are omitted, `true` becomes "1".
 */
export function attrs(values: Record<string, AttrValue> | Array<[string, AttrValue]>): string {
  const entries = Array.isArray(values) ? values : Object.entries(values);
  let out = '';
  for (const [k, v] of entries) {
    if (v === undefined || v === null || v === false) continue;
    out += ` ${k}="${v === true ? '1' : escapeAttr(v)}"`;
  }
  return out;
}

/** Builds an element; self-closing when there are no children. */
export function el(name: string, attributes?: Record<string, AttrValue> | Array<[string, AttrValue]>, children?: string): string {
  const a = attributes ? attrs(attributes) : '';
  return children ? `<${name}${a}>${children}</${name}>` : `<${name}${a}/>`;
}

/** Number serialisation for <v> / attribute values. Throws on non-finite input. */
export function num(n: number): string {
  if (!Number.isFinite(n)) throw new Error(`Cannot serialise non-finite number ${n}`);
  return String(n);
}
