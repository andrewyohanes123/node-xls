/**
 * Sheet drawing part (xl/drawings/drawingN.xml) holding images and charts.
 */
import { createHash } from 'node:crypto';
import { SheetImage, WorkSheet } from '../types.js';
import { WriteContext } from '../xlsx/context.js';
import { CONTENT_TYPES, NS, REL_TYPES, XML_HEADER, attrs } from '../xlsx/xml.js';
import { resolveAnchor, twoCellAnchorXml, emu } from './anchor.js';
import { IMAGE_CONTENT_TYPES, IMAGE_EXTENSIONS, detectImageType, imageSize, toBytes } from './image.js';
import { writeChart } from '../charts/chart.js';

function mediaPart(ctx: WriteContext, bytes: Uint8Array, ext: string, contentType: string): string {
  const key = createHash('sha1').update(bytes).digest('hex');
  let part = ctx.media.get(key);
  if (!part) {
    part = `xl/media/image${++ctx.seq.images}.${ext}`;
    ctx.pkg.addPart(part, bytes);
    ctx.pkg.addDefault(ext, contentType);
    ctx.media.set(key, part);
  }
  return part;
}

function pictureXml(ctx: WriteContext, ws: WorkSheet, img: SheetImage, drawingPart: string, id: number): string {
  const bytes = toBytes(img.data);
  const type = img.type ?? detectImageType(bytes);
  if (!type) throw new Error('Unsupported image format (use PNG, JPEG or GIF)');
  const detected = detectImageType(bytes);
  if (detected && detected !== type) throw new Error(`Image data is ${detected} but type "${type}" was given`);
  const part = mediaPart(ctx, bytes, IMAGE_EXTENSIONS[type], IMAGE_CONTENT_TYPES[type]);
  const rid = ctx.pkg.addRel(drawingPart, REL_TYPES.image, part);
  const anchor = resolveAnchor(ws, img.anchor, imageSize(bytes, type));
  const pic =
    '<xdr:pic>' +
    `<xdr:nvPicPr><xdr:cNvPr${attrs([
      ['id', id],
      ['name', img.name ?? `Picture ${id - 1}`],
      ['descr', img.description],
    ])}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(anchor.width)}" cy="${emu(anchor.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
    '</xdr:pic>';
  return twoCellAnchorXml(anchor, pic, 'oneCell');
}

/** Writes the sheet drawing (images + charts); returns the <drawing> element. */
export function writeDrawing(ctx: WriteContext, ws: WorkSheet, sheetName: string, sheetPart: string): string {
  const images = ws['!images'] ?? [];
  const charts = ws['!charts'] ?? [];
  if (images.length === 0 && charts.length === 0) return '';
  const drawingPart = `xl/drawings/drawing${++ctx.seq.drawings}.xml`;
  let id = 1;
  const anchors: string[] = [];
  for (const img of images) anchors.push(pictureXml(ctx, ws, img, drawingPart, ++id));
  for (const chart of charts) anchors.push(writeChart(ctx, ws, sheetName, chart, drawingPart, ++id));
  ctx.pkg.addPart(
    drawingPart,
    `${XML_HEADER}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:c="${NS.c}">${anchors.join('')}</xdr:wsDr>`,
    CONTENT_TYPES.drawing
  );
  return `<drawing r:id="${ctx.pkg.addRel(sheetPart, REL_TYPES.drawing, drawingPart)}"/>`;
}
