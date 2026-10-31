import { XMLParser } from 'fast-xml-parser';
import { CellStyle, Font, Fill, Border, BorderSide, Alignment, Color, CFStyle, CellProtection } from '../types.js';
import { denormalizeColor } from './index.js';
import { BUILTIN_FORMATS } from '../utils/format.js';

export interface ParsedStyles {
  /** cellXfs as CellStyle objects (index = `s` attribute). */
  cellStyles: CellStyle[];
  /** Number format code per cellXf (undefined for General). */
  cellFormats: Array<string | undefined>;
  /** Differential formats used by conditional formatting. */
  dxfs: CFStyle[];
}

export const styleParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  trimValues: false,
  isArray: (name) => ['font', 'fill', 'border', 'xf', 'numFmt', 'dxf'].includes(name),
});

const isTrue = (node: any): boolean => {
  if (node === undefined || node === null) return false;
  const v = typeof node === 'object' ? node['@_val'] : undefined;
  return v === undefined || (v !== '0' && v !== 'false');
};

export function parseColorNode(node: any): Color | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const tint = node['@_tint'] !== undefined ? Number(node['@_tint']) : undefined;
  if (node['@_rgb']) return { rgb: denormalizeColor(node['@_rgb']) ?? node['@_rgb'], ...(tint ? { tint } : {}) };
  if (node['@_theme'] !== undefined) return { theme: Number(node['@_theme']), ...(tint ? { tint } : {}) };
  if (node['@_indexed'] !== undefined) return { indexed: Number(node['@_indexed']), ...(tint ? { tint } : {}) };
  if (node['@_auto'] === '1' || node['@_auto'] === 'true') return { auto: true };
  return undefined;
}

/** Parses <font>/<rPr> children into a Font. */
export function parseFontNode(fn: any): Font {
  const font: Font = {};
  if (!fn || typeof fn !== 'object') return font;
  const name = fn.name?.['@_val'] ?? fn.rFont?.['@_val'];
  if (name !== undefined) font.name = name;
  if (fn.sz?.['@_val'] !== undefined) font.sz = Number(fn.sz['@_val']);
  if (fn.b !== undefined && isTrue(fn.b)) font.bold = true;
  if (fn.i !== undefined && isTrue(fn.i)) font.italic = true;
  if (fn.strike !== undefined && isTrue(fn.strike)) font.strike = true;
  if (fn.outline !== undefined && isTrue(fn.outline)) font.outline = true;
  if (fn.shadow !== undefined && isTrue(fn.shadow)) font.shadow = true;
  if (fn.u !== undefined) {
    const val = typeof fn.u === 'object' ? fn.u['@_val'] : undefined;
    if (val !== 'none') font.underline = !val || val === 'single' ? true : val;
  }
  if (fn.vertAlign?.['@_val'] && fn.vertAlign['@_val'] !== 'baseline') font.vertAlign = fn.vertAlign['@_val'];
  const color = parseColorNode(fn.color);
  if (color) font.color = color;
  if (fn.family?.['@_val'] !== undefined) font.family = Number(fn.family['@_val']);
  if (fn.charset?.['@_val'] !== undefined) font.charset = Number(fn.charset['@_val']);
  if (fn.scheme?.['@_val']) font.scheme = fn.scheme['@_val'];
  return font;
}

function parseFillNode(fn: any, differential = false): Fill {
  const fill: Fill = {};
  const pat = fn?.patternFill;
  if (pat && typeof pat === 'object') {
    if (pat['@_patternType']) fill.patternType = pat['@_patternType'];
    const fg = parseColorNode(pat.fgColor);
    const bg = parseColorNode(pat.bgColor);
    if (differential && !fill.patternType) {
      // dxf solid fills store the cell color in bgColor
      if (bg) fill.fgColor = bg;
      fill.patternType = 'solid';
    } else {
      if (fg) fill.fgColor = fg;
      // indexed 64 is the "system background" placeholder Excel writes for solid fills
      if (bg && (bg as any).indexed !== 64) fill.bgColor = bg;
    }
  }
  return fill;
}

function parseBorderSide(node: any): BorderSide | undefined {
  if (!node || typeof node !== 'object' || !node['@_style']) return undefined;
  const side: BorderSide = { style: node['@_style'] };
  const color = parseColorNode(node.color);
  if (color && !(color as any).auto) side.color = color;
  return side;
}

function parseBorderNode(bn: any): Border {
  const border: Border = {};
  for (const k of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) {
    const side = parseBorderSide(bn?.[k]);
    if (side) border[k] = side;
  }
  if (bn?.['@_diagonalUp'] === '1' || bn?.['@_diagonalUp'] === 'true') border.diagonalUp = true;
  if (bn?.['@_diagonalDown'] === '1' || bn?.['@_diagonalDown'] === 'true') border.diagonalDown = true;
  return border;
}

function decodeRotation(r: number): number {
  return r > 90 && r <= 180 ? 90 - r : r;
}

function parseAlignment(al: any): Alignment | undefined {
  if (!al || typeof al !== 'object') return undefined;
  const a: Alignment = {};
  if (al['@_horizontal'] && al['@_horizontal'] !== 'general') a.horizontal = al['@_horizontal'];
  if (al['@_vertical'] && al['@_vertical'] !== 'bottom') a.vertical = al['@_vertical'];
  if (al['@_wrapText'] === '1' || al['@_wrapText'] === 'true') a.wrapText = true;
  if (al['@_textRotation'] !== undefined && Number(al['@_textRotation']) !== 0) a.textRotation = decodeRotation(Number(al['@_textRotation']));
  if (al['@_indent'] !== undefined && Number(al['@_indent']) !== 0) a.indent = Number(al['@_indent']);
  if (al['@_shrinkToFit'] === '1' || al['@_shrinkToFit'] === 'true') a.shrinkToFit = true;
  if (al['@_readingOrder'] !== undefined && Number(al['@_readingOrder']) !== 0) a.readingOrder = Number(al['@_readingOrder']);
  return Object.keys(a).length ? a : undefined;
}

function parseProtection(p: any): CellProtection | undefined {
  if (!p || typeof p !== 'object') return undefined;
  const out: CellProtection = {};
  if (p['@_locked'] !== undefined) out.locked = !(p['@_locked'] === '0' || p['@_locked'] === 'false');
  if (p['@_hidden'] !== undefined) out.hidden = p['@_hidden'] === '1' || p['@_hidden'] === 'true';
  return Object.keys(out).length ? out : undefined;
}

const isEmpty = (o: object) => Object.keys(o).length === 0;

export function parseStyles(xml: string): ParsedStyles {
  const styleSheet = styleParser.parse(xml).styleSheet || {};

  const numFmts: Record<number, string> = { ...BUILTIN_FORMATS };
  for (const nf of styleSheet.numFmts?.numFmt || []) {
    const id = Number(nf['@_numFmtId']);
    if (!isNaN(id) && nf['@_formatCode'] !== undefined) numFmts[id] = nf['@_formatCode'];
  }
  const fonts: Font[] = (styleSheet.fonts?.font || []).map((f: any) => parseFontNode(f));
  const fills: Fill[] = (styleSheet.fills?.fill || []).map((f: any) => parseFillNode(f));
  const borders: Border[] = (styleSheet.borders?.border || []).map((b: any) => parseBorderNode(b));

  const cellStyles: CellStyle[] = [];
  const cellFormats: Array<string | undefined> = [];
  for (const xf of styleSheet.cellXfs?.xf || []) {
    const style: CellStyle = {};
    const fontId = Number(xf['@_fontId'] ?? 0);
    const fillId = Number(xf['@_fillId'] ?? 0);
    const borderId = Number(xf['@_borderId'] ?? 0);
    const numFmtId = Number(xf['@_numFmtId'] ?? 0);
    if (fontId > 0 && fonts[fontId] && !isEmpty(fonts[fontId])) style.font = fonts[fontId];
    if (fillId > 1 && fills[fillId] && fills[fillId].patternType && fills[fillId].patternType !== 'none') style.fill = fills[fillId];
    if (borderId > 0 && borders[borderId] && !isEmpty(borders[borderId])) style.border = borders[borderId];
    const fmt = numFmtId > 0 ? numFmts[numFmtId] : undefined;
    if (numFmtId > 0) style.numFmt = fmt ?? numFmtId;
    const alignment = parseAlignment(xf.alignment);
    if (alignment) style.alignment = alignment;
    const protection = parseProtection(xf.protection);
    if (protection) style.protection = protection;
    cellStyles.push(style);
    cellFormats.push(fmt);
  }

  const dxfs: CFStyle[] = (styleSheet.dxfs?.dxf || []).map((d: any) => {
    const s: CFStyle = {};
    if (d.font?.[0]) s.font = parseFontNode(d.font[0]);
    if (d.numFmt?.[0]?.['@_formatCode'] !== undefined) s.numFmt = d.numFmt[0]['@_formatCode'];
    if (d.fill?.[0]) s.fill = parseFillNode(d.fill[0], true);
    if (d.border?.[0]) s.border = parseBorderNode(d.border[0]);
    return s;
  });

  return { cellStyles, cellFormats, dxfs };
}

/** @deprecated use parseStyles */
export function parseStylesXml(xml: string): CellStyle[] {
  return parseStyles(xml).cellStyles;
}
