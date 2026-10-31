/**
 * styles.xml generation.
 *
 * Cell styles (`cell.s`) are deduplicated into fonts/fills/borders/numFmts and
 * cellXfs; conditional-formatting styles become differential formats (dxfs).
 * Index 0 of every collection is Excel's default; fills 0 and 1 are reserved
 * (none, gray125) as Excel requires.
 */
import {
  CellStyle,
  Font,
  Fill,
  Border,
  BorderSide,
  Alignment,
  Color,
  CellProtection,
  CFStyle,
} from '../types.js';
import { XML_HEADER, NS, attrs, AttrValue, escapeAttr } from '../xlsx/xml.js';
import { builtinFormatId, BUILTIN_FORMATS } from '../utils/format.js';

export { escapeXml } from '../xlsx/xml.js';
export { BUILTIN_FORMATS as BUILTIN_NUM_FMTS } from '../utils/format.js';

/** Normalises '#F00', 'FF0000', '80FF0000' to uppercase ARGB. */
export function normalizeRgb(hex: string): string {
  const h = hex.replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{3}$/.test(h)) return 'FF' + h.split('').map((c) => c + c).join('');
  if (/^[0-9A-F]{6}$/.test(h)) return 'FF' + h;
  if (/^[0-9A-F]{8}$/.test(h)) return h;
  throw new Error(`Invalid color "${hex}": expected RGB or ARGB hex`);
}

/** @deprecated kept for compatibility; returns ARGB. */
export function normalizeColor(color?: Color): string | undefined {
  if (!color) return undefined;
  const rgb = typeof color === 'string' ? color : color.rgb;
  return rgb ? normalizeRgb(rgb) : undefined;
}

/** ARGB -> 6-digit RGB when fully opaque (reader convenience). */
export function denormalizeColor(argb?: string): string | undefined {
  if (!argb) return undefined;
  const up = argb.toUpperCase();
  return up.length === 8 && up.startsWith('FF') ? up.substring(2) : up;
}

/** CT_Color attributes for a color, or undefined when the color is empty. */
export function colorAttrs(color?: Color): Array<[string, AttrValue]> | undefined {
  if (color === undefined || color === null) return undefined;
  if (typeof color === 'string') return [['rgb', normalizeRgb(color)]];
  const out: Array<[string, AttrValue]> = [];
  if (color.auto) out.push(['auto', true]);
  else if (color.rgb) out.push(['rgb', normalizeRgb(color.rgb)]);
  else if (color.theme !== undefined) out.push(['theme', color.theme]);
  else if (color.indexed !== undefined) out.push(['indexed', color.indexed]);
  else return undefined;
  if (color.tint !== undefined && color.tint !== 0) out.push(['tint', color.tint]);
  return out;
}

export function colorXml(tag: string, color?: Color): string {
  const a = colorAttrs(color);
  return a ? `<${tag}${attrs(a)}/>` : '';
}

interface FontXmlOptions {
  /** 'name' for cell fonts, 'rFont' for rich-text runs. */
  nameTag?: 'name' | 'rFont';
  /** Fill in Excel's default size/color/name (cell fonts). */
  withDefaults?: boolean;
}

/** Inner XML of a <font>/<rPr> element in Excel's canonical child order. */
export function fontChildren(font: Font, opts: FontXmlOptions = {}): string {
  const { nameTag = 'name', withDefaults = true } = opts;
  let x = '';
  if (font.bold) x += '<b/>';
  if (font.italic) x += '<i/>';
  if (font.strike) x += '<strike/>';
  if (font.outline) x += '<outline/>';
  if (font.shadow) x += '<shadow/>';
  if (font.underline) {
    x += font.underline === true || font.underline === 'single' ? '<u/>' : `<u val="${escapeAttr(font.underline)}"/>`;
  }
  if (font.vertAlign) x += `<vertAlign val="${font.vertAlign}"/>`;
  const sz = font.sz ?? (withDefaults ? 11 : undefined);
  if (sz !== undefined) x += `<sz val="${sz}"/>`;
  const color = colorXml('color', font.color);
  x += color || (withDefaults ? '<color theme="1"/>' : '');
  const name = font.name ?? (withDefaults ? 'Calibri' : undefined);
  if (name !== undefined) x += `<${nameTag} val="${escapeAttr(name)}"/>`;
  const family = font.family ?? (withDefaults && font.name === undefined ? 2 : undefined);
  if (family !== undefined) x += `<family val="${family}"/>`;
  if (font.charset !== undefined) x += `<charset val="${font.charset}"/>`;
  const scheme = font.scheme ?? (withDefaults && font.name === undefined ? 'minor' : undefined);
  if (scheme !== undefined && scheme !== 'none') x += `<scheme val="${scheme}"/>`;
  return x;
}

function borderSideXml(name: string, side?: BorderSide): string {
  if (!side || !side.style || side.style === 'none') return `<${name}/>`;
  return `<${name} style="${side.style}">${colorXml('color', side.color) || '<color auto="1"/>'}</${name}>`;
}

function borderXml(border: Border): string {
  return `<border${attrs([
    ['diagonalUp', border.diagonalUp],
    ['diagonalDown', border.diagonalDown],
  ])}>${borderSideXml('left', border.left)}${borderSideXml('right', border.right)}${borderSideXml('top', border.top)}${borderSideXml(
    'bottom',
    border.bottom
  )}${borderSideXml('diagonal', border.diagonal)}</border>`;
}

/** Converts SheetJS-style -90..90 rotation to Excel's 0..180 encoding. */
function encodeRotation(r: number): number {
  if (r === 255) return 255;
  if (r < 0) return Math.min(180, 90 - r);
  return Math.min(r, 180);
}

function alignmentXml(a?: Alignment): string {
  if (!a) return '';
  const a2 = attrs([
    ['horizontal', a.horizontal],
    ['vertical', a.vertical],
    ['textRotation', a.textRotation !== undefined ? encodeRotation(a.textRotation) : undefined],
    ['wrapText', a.wrapText],
    ['indent', a.indent],
    ['shrinkToFit', a.shrinkToFit],
    ['readingOrder', a.readingOrder],
  ]);
  return a2 ? `<alignment${a2}/>` : '';
}

function protectionXml(p?: CellProtection): string {
  if (!p) return '';
  const a = attrs([
    ['locked', p.locked === undefined ? undefined : p.locked ? '1' : '0'],
    ['hidden', p.hidden === undefined ? undefined : p.hidden ? '1' : '0'],
  ]);
  return a ? `<protection${a}/>` : '';
}

const MAX_CELL_FORMATS = 65430;

/** Office value ranges for style properties. */
function checkStyle(s: CellStyle): void {
  const bad = (what: string, v: unknown) => {
    throw new Error(`Invalid style: ${what} (got ${v})`);
  };
  if (s.font?.sz !== undefined && !(s.font.sz >= 1 && s.font.sz <= 409)) bad('font size must be 1-409', s.font.sz);
  if (s.font?.family !== undefined && !(Number.isInteger(s.font.family) && s.font.family >= 0 && s.font.family <= 14)) bad('font family must be 0-14', s.font.family);
  if (s.alignment?.indent !== undefined && !(Number.isInteger(s.alignment.indent) && s.alignment.indent >= 0 && s.alignment.indent <= 250)) bad('indent must be 0-250', s.alignment.indent);
  const rot = s.alignment?.textRotation;
  if (rot !== undefined && !(rot === 255 || (rot >= -90 && rot <= 180))) bad('textRotation must be -90..180 or 255', rot);
  if (typeof s.numFmt === 'string' && s.numFmt.length > 255) bad('number format longer than 255 characters', s.numFmt.length);
}

class IndexedSet {
  readonly items: string[] = [];
  private map = new Map<string, number>();
  add(xml: string): number {
    const hit = this.map.get(xml);
    if (hit !== undefined) return hit;
    const idx = this.items.length;
    this.items.push(xml);
    this.map.set(xml, idx);
    return idx;
  }
  get size(): number {
    return this.items.length;
  }
}

export class StyleBuilder {
  private fonts = new IndexedSet();
  private fills = new IndexedSet();
  private borders = new IndexedSet();
  private cellXfs = new IndexedSet();
  private dxfs = new IndexedSet();
  private numFmts: Array<{ id: number; code: string }> = [];
  private numFmtIds = new Map<string, number>();
  private nextNumFmtId = 164;
  private cache = new Map<string, number>();

  constructor() {
    this.fonts.add(`<font>${fontChildren({})}</font>`);
    this.fills.add('<fill><patternFill patternType="none"/></fill>');
    this.fills.add('<fill><patternFill patternType="gray125"/></fill>');
    this.borders.add('<border><left/><right/><top/><bottom/><diagonal/></border>');
    this.cellXfs.add('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
  }

  get cellXfCount(): number {
    return this.cellXfs.size;
  }

  registerNumFmt(fmt: string | number): number {
    if (typeof fmt === 'number') {
      // ids < 164 are Excel built-ins; custom ids only exist inside the workbook that defined them
      if (Number.isInteger(fmt) && fmt >= 0 && fmt < 164) return fmt;
      throw new Error(`numFmt ${fmt} is not a built-in format id (0-163); pass the format code string instead`);
    }
    const builtin = builtinFormatId(fmt);
    if (builtin !== undefined) return builtin;
    const hit = this.numFmtIds.get(fmt);
    if (hit !== undefined) return hit;
    const id = this.nextNumFmtId++;
    this.numFmtIds.set(fmt, id);
    this.numFmts.push({ id, code: fmt });
    return id;
  }

  registerFont(font?: Font): number {
    if (!font) return 0;
    return this.fonts.add(`<font>${fontChildren(font)}</font>`);
  }

  registerFill(fill?: Fill): number {
    if (!fill) return 0;
    const fg = colorXml('fgColor', fill.fgColor);
    const bg = colorXml('bgColor', fill.bgColor);
    const pattern = fill.patternType ?? (fg || bg ? 'solid' : 'none');
    if (pattern === 'none' && !fg && !bg) return 0;
    if (pattern === 'gray125' && !fg && !bg) return 1;
    const children = fg + (bg || (fg && pattern === 'solid' ? '<bgColor indexed="64"/>' : ''));
    return this.fills.add(`<fill><patternFill patternType="${pattern}"${children ? `>${children}</patternFill>` : '/>'}</fill>`);
  }

  registerBorder(border?: Border): number {
    if (!border) return 0;
    return this.borders.add(borderXml(border));
  }

  /** cellXfs index for a style (0 = default). */
  getStyleIndex(style?: CellStyle): number {
    if (!style) return 0;
    const key = JSON.stringify(style);
    const cached = this.cache.get(key);
    if (cached !== undefined) return cached;
    checkStyle(style);

    const fontId = this.registerFont(style.font);
    const fillId = this.registerFill(style.fill);
    const borderId = this.registerBorder(style.border);
    const numFmtId = style.numFmt !== undefined ? this.registerNumFmt(style.numFmt) : 0;
    const align = alignmentXml(style.alignment);
    const prot = protectionXml(style.protection);

    const xml = `<xf${attrs([
      ['numFmtId', numFmtId],
      ['fontId', fontId],
      ['fillId', fillId],
      ['borderId', borderId],
      ['xfId', 0],
      ['applyNumberFormat', numFmtId > 0],
      ['applyFont', fontId > 0],
      ['applyFill', fillId > 0],
      ['applyBorder', borderId > 0],
      ['applyAlignment', align !== ''],
      ['applyProtection', prot !== ''],
    ])}${align || prot ? `>${align}${prot}</xf>` : '/>'}`;
    const idx = this.cellXfs.add(xml);
    // Excel's limit on distinct cell formats in a workbook
    if (this.cellXfs.size > MAX_CELL_FORMATS) throw new Error(`More than ${MAX_CELL_FORMATS} distinct cell styles; Excel cannot open the workbook`);
    this.cache.set(key, idx);
    return idx;
  }

  /** dxfs index for a conditional-formatting style. */
  registerDxf(style: CFStyle): number {
    let x = '';
    if (style.font) x += `<font>${fontChildren(style.font, { withDefaults: false })}</font>`;
    if (style.numFmt !== undefined) {
      const id = this.registerNumFmt(style.numFmt);
      const code = typeof style.numFmt === 'string' ? style.numFmt : BUILTIN_FORMATS[id] ?? 'General';
      x += `<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`;
    }
    if (style.fill) {
      const f = style.fill;
      const pattern = f.patternType ?? 'solid';
      if (pattern === 'solid') {
        // In differential formats the solid fill color is stored as bgColor.
        x += `<fill><patternFill>${colorXml('bgColor', f.fgColor ?? f.bgColor)}</patternFill></fill>`;
      } else {
        x += `<fill><patternFill patternType="${pattern}">${colorXml('fgColor', f.fgColor)}${colorXml('bgColor', f.bgColor)}</patternFill></fill>`;
      }
    }
    if (style.border) x += borderXml(style.border);
    return this.dxfs.add(`<dxf>${x}</dxf>`);
  }

  buildXml(): string {
    const numFmts = this.numFmts.length
      ? `<numFmts count="${this.numFmts.length}">${this.numFmts
          .map((nf) => `<numFmt numFmtId="${nf.id}" formatCode="${escapeAttr(nf.code)}"/>`)
          .join('')}</numFmts>`
      : '';
    const list = (tag: string, set: IndexedSet) => `<${tag} count="${set.size}">${set.items.join('')}</${tag}>`;
    return (
      `${XML_HEADER}<styleSheet xmlns="${NS.main}">` +
      numFmts +
      list('fonts', this.fonts) +
      list('fills', this.fills) +
      list('borders', this.borders) +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      list('cellXfs', this.cellXfs) +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      (this.dxfs.size ? list('dxfs', this.dxfs) : '') +
      '</styleSheet>'
    );
  }
}
