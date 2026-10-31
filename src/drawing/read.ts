/**
 * Reads images (and chart references) from a sheet drawing part.
 */
import { strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { SheetImage, WorkSheet } from '../types.js';
import { encode_cell } from '../utils/address.js';
import { EMU_PER_PIXEL } from '../utils/units.js';
import { readRels } from '../xlsx/rels.js';
import { REL_TYPES } from '../xlsx/xml.js';
import { detectImageType } from './image.js';
import { readChartPart } from '../charts/read.js';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  trimValues: false,
  removeNSPrefix: true,
  isArray: (n) => ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor', 'pic', 'graphicFrame'].includes(n),
});

const px = (emu: unknown) => Math.round(Number(emu ?? 0) / EMU_PER_PIXEL);

function marker(m: any) {
  return { c: Number(m?.col ?? 0), r: Number(m?.row ?? 0), dx: px(m?.colOff), dy: px(m?.rowOff) };
}

export function readDrawing(files: Record<string, Uint8Array>, part: string, ws: WorkSheet): void {
  const bytes = files[part];
  if (!bytes) return;
  const root = parser.parse(strFromU8(bytes)).wsDr ?? {};
  const rels = readRels(files, part);
  const anchors = [...(root.twoCellAnchor ?? []), ...(root.oneCellAnchor ?? [])];

  for (const anchor of anchors) {
    const from = marker(anchor.from);
    const to = anchor.to ? marker(anchor.to) : undefined;
    const base = {
      from: encode_cell({ c: from.c, r: from.r }),
      ...(from.dx ? { offsetX: from.dx } : {}),
      ...(from.dy ? { offsetY: from.dy } : {}),
    };

    for (const pic of anchor.pic ?? []) {
      const rid = pic.blipFill?.blip?.['@_embed'];
      const rel = rid ? rels.get(rid) : undefined;
      const data = rel && !rel.external ? files[rel.target] : undefined;
      if (!data) continue;
      const ext = pic.spPr?.xfrm?.ext;
      const size = ext ? { width: px(ext['@_cx']), height: px(ext['@_cy']) } : anchor.ext ? { width: px(anchor.ext['@_cx']), height: px(anchor.ext['@_cy']) } : undefined;
      const img: SheetImage = {
        data,
        type: detectImageType(data),
        anchor: size ? { ...base, ...size } : { ...base, to: encode_cell({ c: to?.c ?? from.c + 1, r: to?.r ?? from.r + 1 }) },
      };
      const nv = pic.nvPicPr?.cNvPr;
      if (nv?.['@_name']) img.name = nv['@_name'];
      if (nv?.['@_descr']) img.description = nv['@_descr'];
      (ws['!images'] ??= []).push(img);
    }

    for (const frame of anchor.graphicFrame ?? []) {
      const rid = frame.graphic?.graphicData?.chart?.['@_id'];
      const rel = rid ? rels.get(rid) : undefined;
      if (!rel || rel.type !== REL_TYPES.chart) continue;
      const xfrm = frame.xfrm?.ext;
      const size = to
        ? { to: encode_cell({ c: to.c, r: to.r }) }
        : xfrm
          ? { width: px(xfrm['@_cx']), height: px(xfrm['@_cy']) }
          : {};
      const chart = readChartPart(files, rel.target, { ...base, ...size });
      if (chart) {
        const name = frame.nvGraphicFramePr?.cNvPr?.['@_name'];
        if (name) chart.name = name;
        (ws['!charts'] ??= []).push(chart);
      }
    }
  }
}
