import { XML_READ_OPTIONS } from '../xlsx/xml.js';
/**
 * Reads images (and chart references) from a sheet drawing part.
 */
import { strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { DrawingAnchor, SheetImage, WorkSheet } from '../types.js';
import { encode_cell } from '../utils/address.js';
import { EMU_PER_PIXEL } from '../utils/units.js';
import { readRels } from '../xlsx/rels.js';
import { detectImageType, imageTypeFromExtension } from './image.js';
import { readChartPart } from '../charts/read.js';

const parser = new XMLParser({
  ...XML_READ_OPTIONS,
  isArray: (n) => ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor', 'pic', 'graphicFrame'].includes(n),
});

const px = (emu: unknown) => Math.round(Number(emu ?? 0) / EMU_PER_PIXEL);

function marker(m: any) {
  return { c: Number(m?.col ?? 0), r: Number(m?.row ?? 0), dx: px(m?.colOff), dy: px(m?.rowOff) };
}

export function readDrawing(files: Record<string, Uint8Array>, part: string, ws: WorkSheet): void {
  const bytes = files[part];
  if (!bytes) return;
  const xml = strFromU8(bytes);
  const root = parser.parse(xml).wsDr ?? {};
  const rels = readRels(files, part);
  // keep document (stacking) order across anchor kinds
  const twos = [...(root.twoCellAnchor ?? [])];
  const ones = [...(root.oneCellAnchor ?? [])];
  const anchors: any[] = [];
  for (const m of xml.matchAll(/<(?:[\w.-]+:)?(twoCellAnchor|oneCellAnchor)\b/g)) {
    const next = m[1] === 'twoCellAnchor' ? twos.shift() : ones.shift();
    if (next) anchors.push(next);
  }

  for (const anchor of anchors) {
    const from = marker(anchor.from);
    const to = anchor.to ? marker(anchor.to) : undefined;
    const base = {
      from: encode_cell({ c: from.c, r: from.r }),
      ...(from.dx ? { offsetX: from.dx } : {}),
      ...(from.dy ? { offsetY: from.dy } : {}),
    };
    // size: the object's own extent if set (some writers leave it 0), else the anchor's
    // "to" marker (with offsets), else a one-cell anchor's extent
    const sizeOf = (ext: any): Partial<DrawingAnchor> => {
      if (ext && Number(ext['@_cx']) > 0 && Number(ext['@_cy']) > 0) return { width: px(ext['@_cx']), height: px(ext['@_cy']) };
      if (to) return { to: encode_cell({ c: to.c, r: to.r }), ...(to.dx ? { toOffsetX: to.dx } : {}), ...(to.dy ? { toOffsetY: to.dy } : {}) };
      if (anchor.ext) return { width: px(anchor.ext['@_cx']), height: px(anchor.ext['@_cy']) };
      return { to: encode_cell({ c: from.c + 1, r: from.r + 1 }) };
    };

    for (const pic of anchor.pic ?? []) {
      const rid = pic.blipFill?.blip?.['@_embed'];
      const rel = rid ? rels.get(rid) : undefined;
      const data = rel && !rel.external ? files[rel.target] : undefined;
      if (!data) continue;
      const type = detectImageType(data) ?? imageTypeFromExtension(rel!.target);
      if (!type) continue; // e.g. SVG-only pictures: not representable, skipped rather than breaking a re-write
      const img: SheetImage = { data, type, anchor: { ...base, ...sizeOf(pic.spPr?.xfrm?.ext) } };
      const nv = pic.nvPicPr?.cNvPr;
      if (nv?.['@_name']) img.name = nv['@_name'];
      if (nv?.['@_descr']) img.description = nv['@_descr'];
      (ws['!images'] ??= []).push(img);
    }

    for (const frame of anchor.graphicFrame ?? []) {
      const rid = frame.graphic?.graphicData?.chart?.['@_id'];
      const rel = rid ? rels.get(rid) : undefined;
      if (!rel || !rel.type.endsWith('/chart')) continue;
      const chart = readChartPart(files, rel.target, { ...base, ...sizeOf(frame.xfrm?.ext) });
      if (chart) {
        const name = frame.nvGraphicFramePr?.cNvPr?.['@_name'];
        if (name) chart.name = name;
        (ws['!charts'] ??= []).push(chart);
      }
    }
  }
}
