import { XMLParser } from 'fast-xml-parser';
import { CellStyle, Font, Fill, Border, BorderSide, Alignment, Color } from '../types.js';
import { BUILTIN_NUM_FMTS, denormalizeColor } from './index.js';

function parseColorNode(node: any): Color | undefined {
  if (!node) return undefined;
  if (typeof node === 'string') return node;
  if (node['@_rgb']) {
    return { rgb: denormalizeColor(node['@_rgb']) || node['@_rgb'] };
  }
  if (node['@_theme'] !== undefined) {
    return {
      theme: Number(node['@_theme']),
      tint: node['@_tint'] ? Number(node['@_tint']) : undefined,
    };
  }
  if (node['@_auto']) {
    return { auto: true };
  }
  return undefined;
}

function parseBorderSide(node: any): BorderSide | undefined {
  if (!node) return undefined;
  const style = node['@_style'];
  if (!style) return undefined;
  const color = parseColorNode(node.color);
  return { style, color };
}

export function parseStylesXml(xml: string): CellStyle[] {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name) => {
      return ['font', 'fill', 'border', 'xf', 'numFmt'].includes(name);
    },
  });

  const parsed = parser.parse(xml);
  const styleSheet = parsed.styleSheet || {};

  // 1. Number formats
  const numFmtsMap: Record<number, string> = { ...BUILTIN_NUM_FMTS };
  const numFmtNodes = styleSheet.numFmts?.numFmt || [];
  for (const nf of numFmtNodes) {
    const id = Number(nf['@_numFmtId']);
    const code = nf['@_formatCode'];
    if (!isNaN(id) && code) {
      numFmtsMap[id] = code;
    }
  }

  // 2. Fonts
  const fontNodes = styleSheet.fonts?.font || [];
  const fonts: Font[] = fontNodes.map((fn: any) => {
    const font: Font = {};
    if (fn.name?.['@_val']) font.name = fn.name['@_val'];
    if (fn.sz?.['@_val']) font.sz = Number(fn.sz['@_val']);
    if (fn.b !== undefined) font.bold = fn.b?.['@_val'] !== '0' && fn.b?.['@_val'] !== 'false';
    if (fn.i !== undefined) font.italic = fn.i?.['@_val'] !== '0' && fn.i?.['@_val'] !== 'false';
    if (fn.strike !== undefined) font.strike = fn.strike?.['@_val'] !== '0' && fn.strike?.['@_val'] !== 'false';
    if (fn.u !== undefined) {
      const val = fn.u?.['@_val'];
      font.underline = val ? val : true;
    }
    if (fn.color) {
      font.color = parseColorNode(fn.color);
    }
    if (fn.family?.['@_val']) font.family = Number(fn.family['@_val']);
    if (fn.scheme?.['@_val']) font.scheme = fn.scheme['@_val'];
    return font;
  });

  // 3. Fills
  const fillNodes = styleSheet.fills?.fill || [];
  const fills: Fill[] = fillNodes.map((fn: any) => {
    const fill: Fill = {};
    const pat = fn.patternFill;
    if (pat) {
      fill.patternType = pat['@_patternType'];
      if (pat.fgColor) fill.fgColor = parseColorNode(pat.fgColor);
      if (pat.bgColor) fill.bgColor = parseColorNode(pat.bgColor);
    }
    return fill;
  });

  // 4. Borders
  const borderNodes = styleSheet.borders?.border || [];
  const borders: Border[] = borderNodes.map((bn: any) => {
    const border: Border = {};
    if (bn.left) border.left = parseBorderSide(bn.left);
    if (bn.right) border.right = parseBorderSide(bn.right);
    if (bn.top) border.top = parseBorderSide(bn.top);
    if (bn.bottom) border.bottom = parseBorderSide(bn.bottom);
    if (bn.diagonal) border.diagonal = parseBorderSide(bn.diagonal);
    if (bn['@_diagonalUp'] === '1' || bn['@_diagonalUp'] === 'true') border.diagonalUp = true;
    if (bn['@_diagonalDown'] === '1' || bn['@_diagonalDown'] === 'true') border.diagonalDown = true;
    return border;
  });

  // 5. CellXfs
  const xfNodes = styleSheet.cellXfs?.xf || [];
  const cellStyles: CellStyle[] = xfNodes.map((xf: any) => {
    const style: CellStyle = {};
    const fontId = Number(xf['@_fontId'] || 0);
    const fillId = Number(xf['@_fillId'] || 0);
    const borderId = Number(xf['@_borderId'] || 0);
    const numFmtId = Number(xf['@_numFmtId'] || 0);

    if (fontId > 0 && fonts[fontId]) {
      style.font = fonts[fontId];
    }
    if (fillId > 0 && fills[fillId]) {
      style.fill = fills[fillId];
    }
    if (borderId > 0 && borders[borderId]) {
      style.border = borders[borderId];
    }
    if (numFmtId > 0) {
      style.numFmt = numFmtsMap[numFmtId] || numFmtId;
    }

    const al = xf.alignment;
    if (al) {
      const alignment: Alignment = {};
      if (al['@_horizontal']) alignment.horizontal = al['@_horizontal'];
      if (al['@_vertical']) alignment.vertical = al['@_vertical'];
      if (al['@_wrapText'] === '1' || al['@_wrapText'] === 'true') alignment.wrapText = true;
      if (al['@_textRotation'] !== undefined) alignment.textRotation = Number(al['@_textRotation']);
      if (al['@_indent'] !== undefined) alignment.indent = Number(al['@_indent']);
      if (al['@_shrinkToFit'] === '1' || al['@_shrinkToFit'] === 'true') alignment.shrinkToFit = true;
      if (al['@_readingOrder'] !== undefined) alignment.readingOrder = Number(al['@_readingOrder']);
      style.alignment = alignment;
    }

    return style;
  });

  return cellStyles;
}
