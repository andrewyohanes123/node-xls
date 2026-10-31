/**
 * Default Office theme (2013-2022 palette, Calibri). Excel resolves
 * `theme="n"` colors and the `minor` font scheme against this part.
 *
 * SpreadsheetML theme indexes: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 accent1-6,
 * 10 hlink, 11 folHlink.
 */
import { XML_HEADER, NS } from '../xlsx/xml.js';

export const THEME_COLORS = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

/** Theme color index -> RGB (for renderers such as sheet_to_html). */
export const THEME_INDEX_RGB = [
  THEME_COLORS.lt1,
  THEME_COLORS.dk1,
  THEME_COLORS.lt2,
  THEME_COLORS.dk2,
  THEME_COLORS.accent1,
  THEME_COLORS.accent2,
  THEME_COLORS.accent3,
  THEME_COLORS.accent4,
  THEME_COLORS.accent5,
  THEME_COLORS.accent6,
  THEME_COLORS.hlink,
  THEME_COLORS.folHlink,
];

const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';

export function buildThemeXml(): string {
  const c = THEME_COLORS;
  const clr = (name: string, rgb: string) =>
    name === 'dk1'
      ? `<a:dk1><a:sysClr val="windowText" lastClr="${rgb}"/></a:dk1>`
      : name === 'lt1'
        ? `<a:lt1><a:sysClr val="window" lastClr="${rgb}"/></a:lt1>`
        : `<a:${name}><a:srgbClr val="${rgb}"/></a:${name}>`;
  const font = (latin: string) => `<a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  const ln = (w: number) =>
    `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  return (
    `${XML_HEADER}<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements>` +
    `<a:clrScheme name="Office">${Object.entries(c)
      .map(([k, v]) => clr(k, v))
      .join('')}</a:clrScheme>` +
    `<a:fontScheme name="Office"><a:majorFont>${font('Calibri Light')}</a:majorFont><a:minorFont>${font('Calibri')}</a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office">` +
    `<a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst>` +
    `<a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst>` +
    `</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`
  );
}
