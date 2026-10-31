import { strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { resolveTarget, relsPathFor } from './package.js';

export interface Rel {
  type: string;
  /** Package path (internal) or URI (external). */
  target: string;
  external: boolean;
}

const relsParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  trimValues: false,
  isArray: (n) => n === 'Relationship',
});

/** Relationships of a part ('' = package root), keyed by id, with resolved targets. */
export function readRels(files: Record<string, Uint8Array>, source: string): Map<string, Rel> {
  const out = new Map<string, Rel>();
  const bytes = files[relsPathFor(source)];
  if (!bytes) return out;
  for (const r of relsParser.parse(strFromU8(bytes)).Relationships?.Relationship ?? []) {
    const external = r['@_TargetMode'] === 'External';
    out.set(r['@_Id'], {
      type: r['@_Type'] ?? '',
      target: external ? r['@_Target'] : resolveTarget(source, r['@_Target'] ?? ''),
      external,
    });
  }
  return out;
}

export function findRel(rels: Map<string, Rel>, type: string): Rel | undefined {
  const tail = type.split('/').pop();
  for (const r of rels.values()) if (r.type === type || r.type.endsWith('/' + tail)) return r;
  return undefined;
}
