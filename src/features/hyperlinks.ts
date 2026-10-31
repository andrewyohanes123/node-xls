import { Hyperlink } from '../types.js';
import { PackageWriter } from '../xlsx/package.js';
import { REL_TYPES, attrs } from '../xlsx/xml.js';

/**
 * Relationship targets must be valid URIs: percent-encode spaces, non-ASCII
 * and other disallowed characters, leaving existing %XX escapes intact.
 */
export function encodeUriTarget(target: string): string {
  let out = '';
  for (const ch of target) {
    const code = ch.codePointAt(0)!;
    if (ch === '%') out += '%';
    else if (code > 0x20 && code < 0x7f && !'"<>^`{|}'.includes(ch)) out += ch;
    else out += Array.from(new TextEncoder().encode(ch), (b) => `%${b.toString(16).toUpperCase().padStart(2, '0')}`).join('');
  }
  // a '%' not starting an escape must itself be escaped
  return out.replace(/%(?![0-9A-Fa-f]{2})/g, '%25');
}

/**
 * <hyperlinks> for a worksheet. Internal links ('#Sheet2!A1') use `location`;
 * everything else becomes an external relationship.
 */
export function hyperlinksXml(links: Array<{ ref: string; link: Hyperlink }>, pkg: PackageWriter, sheetPart: string): string {
  if (links.length === 0) return '';
  // Office limits (MS-OI29500): 65,530 links per sheet, 255-character tooltips, 2,084-character targets
  if (links.length > 65530) throw new Error(`Sheet has ${links.length} hyperlinks; Excel allows 65530`);
  const items = links.map(({ ref, link }) => {
    const target = String(link.Target ?? '');
    const internal = target.startsWith('#');
    if (link.Tooltip !== undefined && String(link.Tooltip).length > 255) throw new Error(`Hyperlink ${ref}: tooltip longer than 255 characters`);
    if (target.length > 2084) throw new Error(`Hyperlink ${ref}: target longer than 2084 characters`);
    const rid = internal ? undefined : pkg.addRel(sheetPart, REL_TYPES.hyperlink, encodeUriTarget(target), true);
    return `<hyperlink${attrs([
      ['ref', ref],
      ['r:id', rid],
      ['location', internal ? target.slice(1) : undefined],
      ['tooltip', link.Tooltip],
    ])}/>`;
  });
  return `<hyperlinks>${items.join('')}</hyperlinks>`;
}
