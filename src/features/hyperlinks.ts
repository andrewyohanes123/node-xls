import { Hyperlink } from '../types.js';
import { PackageWriter } from '../xlsx/package.js';
import { REL_TYPES, attrs } from '../xlsx/xml.js';

/**
 * <hyperlinks> for a worksheet. Internal links ('#Sheet2!A1') use `location`;
 * everything else becomes an external relationship.
 */
export function hyperlinksXml(links: Array<{ ref: string; link: Hyperlink }>, pkg: PackageWriter, sheetPart: string): string {
  if (links.length === 0) return '';
  const items = links.map(({ ref, link }) => {
    const target = String(link.Target ?? '');
    const internal = target.startsWith('#');
    const rid = internal ? undefined : pkg.addRel(sheetPart, REL_TYPES.hyperlink, target, true);
    return `<hyperlink${attrs([
      ['ref', ref],
      ['r:id', rid],
      ['location', internal ? target.slice(1) : undefined],
      ['tooltip', link.Tooltip],
    ])}/>`;
  });
  return `<hyperlinks>${items.join('')}</hyperlinks>`;
}
