/**
 * OPC package assembly: parts, content types and relationships.
 *
 * Every relationship is registered from its source part with an absolute
 * target path; relative targets and rIds are computed here so that part
 * writers never hand-build `.rels` files or `[Content_Types].xml`.
 */
import { strToU8 } from 'fflate';
import { CONTENT_TYPES, NS, XML_HEADER, attrs } from './xml.js';

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function dirOf(part: string): string {
  const i = part.lastIndexOf('/');
  return i === -1 ? '' : part.slice(0, i);
}

/** Relative path from directory `fromDir` to part `to` (both package-absolute, no leading slash). */
export function relativePath(fromDir: string, to: string): string {
  const from = fromDir ? fromDir.split('/') : [];
  const target = to.split('/');
  let i = 0;
  while (i < from.length && i < target.length - 1 && from[i] === target[i]) i++;
  return [...Array(from.length - i).fill('..'), ...target.slice(i)].join('/');
}

/** Resolves a relationship target against its source part (reader side). */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const out: string[] = dirOf(sourcePart) ? dirOf(sourcePart).split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') out.pop();
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return out.join('/');
}

/** Path of the .rels part for a source part ('' = package root). */
export function relsPathFor(sourcePart: string): string {
  if (!sourcePart) return '_rels/.rels';
  const dir = dirOf(sourcePart);
  const file = sourcePart.slice(dir.length ? dir.length + 1 : 0);
  return `${dir ? dir + '/' : ''}_rels/${file}.rels`;
}

export class PackageWriter {
  private parts = new Map<string, Uint8Array>();
  private overrides = new Map<string, string>();
  private defaults = new Map<string, string>([
    ['rels', CONTENT_TYPES.rels],
    ['xml', CONTENT_TYPES.xml],
  ]);
  private rels = new Map<string, Relationship[]>();

  /** Adds a part. A content type registers an Override unless `asDefault` is set. */
  addPart(path: string, content: string | Uint8Array, contentType?: string): void {
    this.parts.set(path, typeof content === 'string' ? strToU8(content) : content);
    if (contentType) this.overrides.set(path, contentType);
  }

  /** Registers a Default content type for a file extension (e.g. png). */
  addDefault(extension: string, contentType: string): void {
    this.defaults.set(extension.toLowerCase(), contentType);
  }

  hasPart(path: string): boolean {
    return this.parts.has(path);
  }

  /**
   * Adds a relationship from `source` ('' for the package root) and returns its
   * id. `target` is a package-absolute part path, or a URI when `external`.
   */
  addRel(source: string, type: string, target: string, external = false): string {
    const list = this.rels.get(source) ?? [];
    const id = `rId${list.length + 1}`;
    list.push({ id, type, target, external });
    this.rels.set(source, list);
    return id;
  }

  build(): Record<string, Uint8Array> {
    const files: Record<string, Uint8Array> = {};

    for (const [source, list] of this.rels) {
      const base = dirOf(source);
      const items = list
        .map(
          (r) =>
            `<Relationship${attrs([
              ['Id', r.id],
              ['Type', r.type],
              ['Target', r.external ? r.target : relativePath(base, r.target)],
              ['TargetMode', r.external ? 'External' : undefined],
            ])}/>`
        )
        .join('');
      files[relsPathFor(source)] = strToU8(`${XML_HEADER}<Relationships xmlns="${NS.rels}">${items}</Relationships>`);
    }

    const defaults = [...this.defaults]
      .map(([ext, ct]) => `<Default${attrs([['Extension', ext], ['ContentType', ct]])}/>`)
      .join('');
    const overrides = [...this.overrides]
      .map(([part, ct]) => `<Override${attrs([['PartName', `/${part}`], ['ContentType', ct]])}/>`)
      .join('');
    files['[Content_Types].xml'] = strToU8(`${XML_HEADER}<Types xmlns="${NS.contentTypes}">${defaults}${overrides}</Types>`);

    for (const [path, bytes] of this.parts) files[path] = bytes;
    return files;
  }
}
