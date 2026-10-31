import { Font, RichTextRun } from '../types.js';
import { XML_HEADER, NS, textElement } from './xml.js';
import { fontChildren } from '../styles/index.js';

/** Shared string table: deduplicated plain and rich strings. */
export class SharedStringsBuilder {
  private items: string[] = [];
  private index = new Map<string, number>();
  private refs = 0;

  private add(key: string, xml: string): number {
    this.refs++;
    const hit = this.index.get(key);
    if (hit !== undefined) return hit;
    const idx = this.items.length;
    this.items.push(xml);
    this.index.set(key, idx);
    return idx;
  }

  getIndex(str: string): number {
    return this.add(`s:${str}`, `<si>${textElement(str)}</si>`);
  }

  /**
   * Rich string. Runs with a font get complete run properties (the cell's
   * name/size/colour, else Excel's default, overlaid with the run font) like
   * Excel writes; partial run properties render with an application default
   * font. On/off flags (bold, italic, ...) are not inherited: an rPr is the
   * run's whole formatting, so a plain run in a bold cell stays plain.
   */
  getRichIndex(runs: RichTextRun[], baseFont?: Font): number {
    const { name, sz, family, charset, scheme, color } = baseFont ?? {};
    const base: Font = Object.fromEntries(Object.entries({ name, sz, family, charset, scheme, color }).filter(([, v]) => v !== undefined));
    const xml = runs
      .map((run) => {
        const rPr = run.font ? fontChildren({ ...base, ...run.font }, { nameTag: 'rFont' }) : '';
        return `<r>${rPr ? `<rPr>${rPr}</rPr>` : ''}${textElement(String(run.text ?? ''))}</r>`;
      })
      .join('');
    return this.add(`r:${xml}`, `<si>${xml}</si>`);
  }

  get count(): number {
    return this.items.length;
  }

  buildXml(): string {
    return `${XML_HEADER}<sst xmlns="${NS.main}" count="${this.refs}" uniqueCount="${this.items.length}">${this.items.join('')}</sst>`;
  }
}
