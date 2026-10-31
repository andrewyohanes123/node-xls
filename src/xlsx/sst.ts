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
   * Rich string. Runs with a font get complete run properties (the cell font,
   * else Excel's default, overlaid with the run font) like Excel writes;
   * partial run properties render with an application default font.
   */
  getRichIndex(runs: RichTextRun[], baseFont?: Font): number {
    const xml = runs
      .map((run) => {
        const rPr = run.font ? fontChildren({ ...(baseFont ?? {}), ...run.font }, { nameTag: 'rFont' }) : '';
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
