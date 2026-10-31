import { CellStyle, Font, Fill, Border, BorderSide, Alignment, Color } from '../types.js';

export const BUILTIN_NUM_FMTS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

const REVERSE_BUILTIN_NUM_FMTS: Record<string, number> = {};
for (const [id, fmt] of Object.entries(BUILTIN_NUM_FMTS)) {
  REVERSE_BUILTIN_NUM_FMTS[fmt.toLowerCase()] = Number(id);
}

export function normalizeColor(color?: Color): string | undefined {
  if (!color) return undefined;
  let hex: string | undefined;
  if (typeof color === 'string') {
    hex = color.replace(/^#/, '');
  } else if (color.rgb) {
    hex = color.rgb.replace(/^#/, '');
  }
  if (!hex) return undefined;
  if (hex.length === 6) {
    return 'FF' + hex.toUpperCase();
  }
  if (hex.length === 8) {
    return hex.toUpperCase();
  }
  if (hex.length === 3) {
    return 'FF' + hex.split('').map((c) => c + c).join('').toUpperCase();
  }
  return hex.toUpperCase();
}

export function denormalizeColor(argb?: string): string | undefined {
  if (!argb) return undefined;
  if (argb.length === 8 && argb.startsWith('FF')) {
    return argb.substring(2);
  }
  return argb;
}

export function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export class StyleBuilder {
  private fonts: string[] = [];
  private fontMap = new Map<string, number>();

  private fills: string[] = [];
  private fillMap = new Map<string, number>();

  private borders: string[] = [];
  private borderMap = new Map<string, number>();

  private numFmts: Array<{ id: number; code: string }> = [];
  private numFmtMap = new Map<string, number>();
  private nextCustomNumFmtId = 164;

  private cellXfs: string[] = [];
  private xfMap = new Map<string, number>();

  constructor() {
    // Default Font (index 0)
    this.addFontXml('<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>');

    // Standard OpenXML fills: index 0 = none, index 1 = gray125
    this.addFillXml('<fill><patternFill patternType="none"/></fill>');
    this.addFillXml('<fill><patternFill patternType="gray125"/></fill>');

    // Default Border (index 0)
    this.addBorderXml('<border><left/><right/><top/><bottom/><diagonal/></border>');

    // Default normal cellXf (index 0)
    this.addXfXml('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
  }

  private addFontXml(xml: string): number {
    if (this.fontMap.has(xml)) {
      return this.fontMap.get(xml)!;
    }
    const idx = this.fonts.length;
    this.fonts.push(xml);
    this.fontMap.set(xml, idx);
    return idx;
  }

  private addFillXml(xml: string): number {
    if (this.fillMap.has(xml)) {
      return this.fillMap.get(xml)!;
    }
    const idx = this.fills.length;
    this.fills.push(xml);
    this.fillMap.set(xml, idx);
    return idx;
  }

  private addBorderXml(xml: string): number {
    if (this.borderMap.has(xml)) {
      return this.borderMap.get(xml)!;
    }
    const idx = this.borders.length;
    this.borders.push(xml);
    this.borderMap.set(xml, idx);
    return idx;
  }

  private addXfXml(xml: string): number {
    if (this.xfMap.has(xml)) {
      return this.xfMap.get(xml)!;
    }
    const idx = this.cellXfs.length;
    this.cellXfs.push(xml);
    this.xfMap.set(xml, idx);
    return idx;
  }

  public registerNumFmt(formatCode: string): number {
    const lower = formatCode.toLowerCase();
    if (REVERSE_BUILTIN_NUM_FMTS[lower] !== undefined) {
      return REVERSE_BUILTIN_NUM_FMTS[lower];
    }
    if (this.numFmtMap.has(formatCode)) {
      return this.numFmtMap.get(formatCode)!;
    }
    const id = this.nextCustomNumFmtId++;
    this.numFmtMap.set(formatCode, id);
    this.numFmts.push({ id, code: formatCode });
    return id;
  }

  public registerFont(font?: Font): number {
    if (!font) return 0;
    const parts: string[] = [];
    if (font.bold) parts.push('<b/>');
    if (font.italic) parts.push('<i/>');
    if (font.strike) parts.push('<strike/>');
    if (font.underline) {
      if (typeof font.underline === 'string') {
        parts.push(`<u val="${font.underline}"/>`);
      } else {
        parts.push('<u/>');
      }
    }
    const sz = font.sz || 11;
    parts.push(`<sz val="${sz}"/>`);

    const col = normalizeColor(font.color);
    if (col) {
      parts.push(`<color rgb="${col}"/>`);
    } else {
      parts.push('<color theme="1"/>');
    }

    const name = font.name || 'Calibri';
    parts.push(`<name val="${escapeXml(name)}"/>`);
    if (font.family !== undefined) {
      parts.push(`<family val="${font.family}"/>`);
    }
    if (font.scheme) {
      parts.push(`<scheme val="${font.scheme}"/>`);
    }

    const xml = `<font>${parts.join('')}</font>`;
    return this.addFontXml(xml);
  }

  public registerFill(fill?: Fill): number {
    if (!fill) return 0;
    const patternType = fill.patternType || (fill.fgColor ? 'solid' : 'none');
    if (patternType === 'none' && !fill.fgColor && !fill.bgColor) return 0;

    const parts: string[] = [];
    const fg = normalizeColor(fill.fgColor);
    const bg = normalizeColor(fill.bgColor);

    if (fg) {
      parts.push(`<fgColor rgb="${fg}"/>`);
    }
    if (bg) {
      parts.push(`<bgColor rgb="${bg}"/>`);
    } else if (fg && patternType === 'solid') {
      parts.push('<bgColor indexed="64"/>');
    }

    const xml = `<fill><patternFill patternType="${patternType}">${parts.join('')}</patternFill></fill>`;
    return this.addFillXml(xml);
  }

  private renderBorderSide(name: string, side?: BorderSide): string {
    if (!side || !side.style || side.style === 'none') {
      return `<${name}/>`;
    }
    const col = normalizeColor(side.color);
    if (col) {
      return `<${name} style="${side.style}"><color rgb="${col}"/></${name}>`;
    }
    return `<${name} style="${side.style}"><color auto="1"/></${name}>`;
  }

  public registerBorder(border?: Border): number {
    if (!border) return 0;
    const left = this.renderBorderSide('left', border.left);
    const right = this.renderBorderSide('right', border.right);
    const top = this.renderBorderSide('top', border.top);
    const bottom = this.renderBorderSide('bottom', border.bottom);
    const diagonal = this.renderBorderSide('diagonal', border.diagonal);

    let attrs = '';
    if (border.diagonalUp) attrs += ' diagonalUp="1"';
    if (border.diagonalDown) attrs += ' diagonalDown="1"';

    const xml = `<border${attrs}>${left}${right}${top}${bottom}${diagonal}</border>`;
    return this.addBorderXml(xml);
  }

  public getStyleIndex(style?: CellStyle): number {
    if (!style) return 0;

    const fontId = this.registerFont(style.font);
    const fillId = this.registerFill(style.fill);
    const borderId = this.registerBorder(style.border);

    let numFmtId = 0;
    if (style.numFmt !== undefined) {
      if (typeof style.numFmt === 'number') {
        numFmtId = style.numFmt;
      } else {
        numFmtId = this.registerNumFmt(style.numFmt);
      }
    }

    const align = style.alignment;
    let alignXml = '';
    let hasAlign = false;
    if (align) {
      const attrs: string[] = [];
      if (align.horizontal) attrs.push(`horizontal="${align.horizontal}"`);
      if (align.vertical) attrs.push(`vertical="${align.vertical}"`);
      if (align.wrapText) attrs.push('wrapText="1"');
      if (align.textRotation !== undefined) attrs.push(`textRotation="${align.textRotation}"`);
      if (align.indent !== undefined) attrs.push(`indent="${align.indent}"`);
      if (align.shrinkToFit) attrs.push('shrinkToFit="1"');
      if (align.readingOrder !== undefined) attrs.push(`readingOrder="${align.readingOrder}"`);
      if (attrs.length > 0) {
        alignXml = `<alignment ${attrs.join(' ')}/>`;
        hasAlign = true;
      }
    }

    const applyAttrs: string[] = [];
    if (fontId > 0) applyAttrs.push('applyFont="1"');
    if (fillId > 0) applyAttrs.push('applyFill="1"');
    if (borderId > 0) applyAttrs.push('applyBorder="1"');
    if (numFmtId > 0) applyAttrs.push('applyNumberFormat="1"');
    if (hasAlign) applyAttrs.push('applyAlignment="1"');

    const applyStr = applyAttrs.length > 0 ? ` ${applyAttrs.join(' ')}` : '';
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${applyStr}>${alignXml}</xf>`;

    return this.addXfXml(xml);
  }

  public buildXml(): string {
    const numFmtsXml =
      this.numFmts.length > 0
        ? `<numFmts count="${this.numFmts.length}">${this.numFmts
            .map((nf) => `<numFmt numFmtId="${nf.id}" formatCode="${escapeXml(nf.code)}"/>`)
            .join('')}</numFmts>`
        : '';

    const fontsXml = `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`;
    const fillsXml = `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`;
    const bordersXml = `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`;
    const cellStyleXfsXml = '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    const cellXfsXml = `<cellXfs count="${this.cellXfs.length}">${this.cellXfs.join('')}</cellXfs>`;
    const cellStylesXml = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmtsXml}${fontsXml}${fillsXml}${bordersXml}${cellStyleXfsXml}${cellXfsXml}${cellStylesXml}</styleSheet>`;
  }
}
