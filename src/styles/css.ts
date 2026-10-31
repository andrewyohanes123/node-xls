/**
 * CellStyle <-> CSS conversion used by sheet_to_html / table_to_sheet.
 */
import { BorderSide, BorderStyle, CellStyle, Color } from '../types.js';
import { THEME_INDEX_RGB } from './theme.js';

/** Excel's legacy indexed palette (0-63). */
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function applyTint(hex: string, tint: number): string {
  const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = ch.map((c) => Math.round(tint < 0 ? c * (1 + tint) : c + (255 - c) * tint));
  return out.map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Resolves a Color to 6-digit hex (no '#'), or undefined. */
export function colorToHex(color?: Color): string | undefined {
  if (!color) return undefined;
  if (typeof color === 'string') return color.replace(/^#/, '').slice(-6).toUpperCase();
  let hex: string | undefined;
  if (color.rgb) hex = color.rgb.replace(/^#/, '').slice(-6).toUpperCase();
  else if (color.theme !== undefined) hex = THEME_INDEX_RGB[color.theme];
  else if (color.indexed !== undefined) hex = INDEXED[color.indexed];
  if (!hex || !/^[0-9A-F]{6}$/i.test(hex)) return undefined;
  return color.tint ? applyTint(hex, color.tint) : hex;
}

const BORDER_CSS: Record<BorderStyle, string> = {
  none: 'none',
  thin: '1px solid',
  hair: '1px solid',
  medium: '2px solid',
  thick: '3px solid',
  double: '3px double',
  dashed: '1px dashed',
  dotted: '1px dotted',
  mediumDashed: '2px dashed',
  dashDot: '1px dashed',
  mediumDashDot: '2px dashed',
  dashDotDot: '1px dotted',
  mediumDashDotDot: '2px dotted',
  slantDashDot: '2px dashed',
};

const cssSafe = (s: string) => s.replace(/[;"'<>{}\\]/g, '');

function borderCss(side?: BorderSide): string | undefined {
  if (!side?.style || side.style === 'none') return undefined;
  return `${BORDER_CSS[side.style] ?? '1px solid'} #${colorToHex(side.color) ?? '000000'}`;
}

/** CSS declarations for a cell style (values sanitised; caller escapes the attribute). */
export function styleToCss(s?: CellStyle): string {
  if (!s) return '';
  const css: string[] = [];
  const f = s.font;
  if (f?.bold) css.push('font-weight: bold');
  if (f?.italic) css.push('font-style: italic');
  const deco = [f?.underline ? 'underline' : '', f?.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) css.push(`text-decoration: ${deco}`);
  if (f?.sz) css.push(`font-size: ${Number(f.sz)}pt`);
  if (f?.name) css.push(`font-family: '${cssSafe(f.name)}'`);
  const color = colorToHex(f?.color);
  if (color) css.push(`color: #${color}`);
  const fill = s.fill;
  if (fill && fill.patternType !== 'none') {
    const bg = colorToHex(fill.fgColor ?? fill.bgColor);
    if (bg) css.push(`background-color: #${bg}`);
  }
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const b = borderCss(s.border?.[side]);
    if (b) css.push(`border-${side}: ${b}`);
  }
  const a = s.alignment;
  if (a?.horizontal && ['left', 'center', 'right', 'justify'].includes(a.horizontal)) css.push(`text-align: ${a.horizontal}`);
  if (a?.vertical) css.push(`vertical-align: ${a.vertical === 'center' ? 'middle' : a.vertical === 'top' ? 'top' : 'bottom'}`);
  if (a?.wrapText) css.push('white-space: pre-wrap');
  if (a?.indent) css.push(`padding-left: ${Number(a.indent) * 9}px`);
  return css.join('; ');
}

function parseCssColor(v: string): string | undefined {
  const s = v.trim().toLowerCase();
  let m = /^#([0-9a-f]{6})$/.exec(s);
  if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f]{3})$/.exec(s);
  if (m) return m[1].split('').map((c) => c + c).join('').toUpperCase();
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (m) return [m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase();
  const named: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080' };
  return named[s];
}

function parseBorderCss(v: string): BorderSide | undefined {
  const parts = v.trim().toLowerCase().split(/\s+/);
  if (parts.includes('none') || parts.includes('hidden')) return undefined;
  const width = parseFloat(parts.find((p) => /^[\d.]+px$/.test(p)) ?? '1');
  const kind = parts.find((p) => ['solid', 'dashed', 'dotted', 'double'].includes(p)) ?? 'solid';
  const colorToken = /#[0-9a-f]{6}\b|#[0-9a-f]{3}\b|rgba?\([^)]*\)/i.exec(v)?.[0] ?? parts.find((p) => parseCssColor(p));
  const color = colorToken ? parseCssColor(colorToken) : undefined;
  let style: BorderStyle = 'thin';
  if (kind === 'double') style = 'double';
  else if (kind === 'dashed') style = width >= 2 ? 'mediumDashed' : 'dashed';
  else if (kind === 'dotted') style = 'dotted';
  else style = width >= 3 ? 'thick' : width >= 2 ? 'medium' : 'thin';
  return { style, ...(color ? { color: { rgb: color } } : {}) };
}

/** Parses an inline CSS declaration list back into a CellStyle. */
export function cssToStyle(css: string): CellStyle | undefined {
  const s: CellStyle = {};
  for (const decl of css.split(';')) {
    const i = decl.indexOf(':');
    if (i === -1) continue;
    const prop = decl.slice(0, i).trim().toLowerCase();
    const val = decl.slice(i + 1).trim();
    const lower = val.toLowerCase();
    switch (prop) {
      case 'font-weight':
        if (lower === 'bold' || Number(lower) >= 600) (s.font ??= {}).bold = true;
        break;
      case 'font-style':
        if (lower === 'italic') (s.font ??= {}).italic = true;
        break;
      case 'text-decoration':
      case 'text-decoration-line':
        if (lower.includes('underline')) (s.font ??= {}).underline = true;
        if (lower.includes('line-through')) (s.font ??= {}).strike = true;
        break;
      case 'font-size': {
        const n = parseFloat(lower);
        if (n) (s.font ??= {}).sz = lower.endsWith('px') ? Math.round(n * 0.75 * 2) / 2 : n;
        break;
      }
      case 'font-family':
        (s.font ??= {}).name = val.split(',')[0].replace(/["']/g, '').trim();
        break;
      case 'color': {
        const c = parseCssColor(val);
        if (c) (s.font ??= {}).color = { rgb: c };
        break;
      }
      case 'background':
      case 'background-color': {
        const c = parseCssColor(val);
        if (c) s.fill = { patternType: 'solid', fgColor: { rgb: c } };
        break;
      }
      case 'text-align':
        if (['left', 'center', 'right', 'justify'].includes(lower)) (s.alignment ??= {}).horizontal = lower as any;
        break;
      case 'vertical-align':
        if (lower === 'middle') (s.alignment ??= {}).vertical = 'center';
        else if (lower === 'top' || lower === 'bottom') (s.alignment ??= {}).vertical = lower;
        break;
      case 'white-space':
        if (lower.startsWith('pre') || lower === 'normal') (s.alignment ??= {}).wrapText = true;
        break;
      default:
        if (/^border-(top|right|bottom|left)$/.test(prop)) {
          const side = parseBorderCss(val);
          if (side) (s.border ??= {})[prop.slice(7) as 'top'] = side;
        } else if (prop === 'border') {
          const side = parseBorderCss(val);
          if (side) s.border = { top: side, right: side, bottom: side, left: side };
        }
    }
  }
  return Object.keys(s).length ? s : undefined;
}
