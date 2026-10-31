/**
 * Images (SheetJS "Pro Image" tier): embedding, anchoring, media sharing and
 * read-back, verified by the drawing schema and ExcelJS.
 */
import { describe, it, expect } from 'vitest';
import XLSX, { utils } from '../src/index.js';
import { detectImageType, imageSize } from '../src/drawing/image.js';
import { book, roundtrip, roundtripSheet } from './helpers/util.js';
import { makePng, JPEG_3x2, GIF_4x5 } from './helpers/png.js';

const png = makePng(40, 30, (x, y) => [x * 6, y * 8, 160, 255]);
const sheet = () => utils.aoa_to_sheet([['image sheet']]);

describe('image utilities', () => {
  it('detects type and pixel size of PNG, JPEG and GIF', () => {
    const jpg = Buffer.from(JPEG_3x2, 'base64');
    const gif = Buffer.from(GIF_4x5, 'base64');
    expect([detectImageType(png), detectImageType(jpg), detectImageType(gif)]).toEqual(['png', 'jpeg', 'gif']);
    expect(imageSize(png, 'png')).toEqual({ width: 40, height: 30 });
    expect(imageSize(jpg, 'jpeg')).toEqual({ width: 3, height: 2 });
    expect(imageSize(gif, 'gif')).toEqual({ width: 4, height: 5 });
    expect(detectImageType(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]))).toBeUndefined();
  });
});

describe('embedding images', () => {
  it('PNG at its natural size: schema-valid, read by ExcelJS and node-xls', async () => {
    const ws = sheet();
    ws['!images'] = [{ data: png, anchor: { from: 'B2' }, name: 'Logo', description: 'Company logo' }];
    const { rep, ws: back } = await roundtripSheet(ws);
    expect(rep.files['xl/media/image1.png']).toEqual(png);
    expect(rep.contentTypes['xl/media/image1.png']).toBe('image/png');
    const ex = rep.exceljs.getWorksheet('S')!;
    const images = ex.getImages();
    expect(images).toHaveLength(1);
    expect(images[0].range.tl).toMatchObject({ nativeCol: 1, nativeRow: 1 });
    const media = rep.exceljs.getImage(Number(images[0].imageId)) as any;
    expect(media.extension).toBe('png');
    expect(new Uint8Array(media.buffer)).toEqual(png);
    expect(back['!images']).toEqual([{ data: png, type: 'png', anchor: { from: 'B2', width: 40, height: 30 }, name: 'Logo', description: 'Company logo' }]);
  });

  it('JPEG and GIF from base64 / data URIs, explicit size and offsets', async () => {
    const ws = sheet();
    ws['!images'] = [
      { data: JPEG_3x2, anchor: { from: 'A3', width: 90, height: 60, offsetX: 5, offsetY: 4 } },
      { data: `data:image/gif;base64,${GIF_4x5}`, anchor: { from: 'D3', width: 40, height: 50 } },
    ];
    const { rep, ws: back } = await roundtripSheet(ws);
    expect(Object.keys(rep.files).filter((f) => f.startsWith('xl/media/')).sort()).toEqual(['xl/media/image1.jpeg', 'xl/media/image2.gif']);
    expect(rep.contentTypes['xl/media/image1.jpeg']).toBe('image/jpeg');
    expect(back['!images']!.map((i) => [i.type, i.anchor])).toEqual([
      ['jpeg', { from: 'A3', offsetX: 5, offsetY: 4, width: 90, height: 60 }],
      ['gif', { from: 'D3', width: 40, height: 50 }],
    ]);
  });

  it('pixel sizes become two-cell anchors using the sheet column widths and row heights', async () => {
    const ws = sheet();
    ws['!cols'] = [{ wpx: 100 }, { wpx: 50 }, { hidden: true }];
    ws['!rows'] = [{ hpx: 40 }];
    ws['!images'] = [{ data: png, anchor: { from: 'A1', width: 170, height: 50 } }];
    const { rep } = await roundtripSheet(ws);
    const drawing = rep.xml('xl/drawings/drawing1.xml');
    // 100 + 50 + hidden(0) -> 20px into column D; 40 -> 10px into row 2
    expect(drawing).toContain('<xdr:to><xdr:col>3</xdr:col><xdr:colOff>190500</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:to>');
  });

  it('explicit from/to cell anchors', async () => {
    const ws = sheet();
    ws['!images'] = [{ data: png, anchor: { from: 'B2', to: 'E9' } }];
    const { rep } = await roundtripSheet(ws);
    const ex = rep.exceljs.getWorksheet('S')!.getImages()[0];
    expect(ex.range.br).toMatchObject({ nativeCol: 4, nativeRow: 8 });
  });

  it('identical images are stored once across sheets', async () => {
    const a = sheet();
    const b = sheet();
    a['!images'] = [
      { data: png, anchor: { from: 'A1' } },
      { data: Buffer.from(png), anchor: { from: 'D1' } },
    ];
    b['!images'] = [{ data: png.buffer.slice(0), anchor: { from: 'A1' } }];
    const { rep, back } = await roundtrip(book(['A', a], ['B', b]));
    expect(Object.keys(rep.files).filter((f) => f.startsWith('xl/media/'))).toEqual(['xl/media/image1.png']);
    expect(back.Sheets.A['!images']).toHaveLength(2);
    expect(back.Sheets.B['!images']![0].data).toEqual(png);
  });

  it('coexists with comments, hyperlinks and tables on the same sheet', async () => {
    const ws = utils.aoa_to_sheet([
      ['k', 'v'],
      ['a', 1],
    ]);
    utils.cell_add_comment(ws.A1, 'note', 'x');
    ws.B2.l = { Target: 'https://example.com' };
    ws['!tables'] = [{ name: 'T_img', ref: 'A1:B2' }];
    ws['!images'] = [{ data: png, anchor: { from: 'D2' } }];
    const { rep } = await roundtripSheet(ws);
    const xml = rep.xml('xl/worksheets/sheet1.xml');
    expect(xml.indexOf('<hyperlinks>')).toBeLessThan(xml.indexOf('<drawing '));
    expect(xml.indexOf('<drawing ')).toBeLessThan(xml.indexOf('<legacyDrawing '));
    expect(xml.indexOf('<legacyDrawing ')).toBeLessThan(xml.indexOf('<tableParts'));
  });

  it('read -> write keeps images', async () => {
    const ws = sheet();
    ws['!images'] = [{ data: png, anchor: { from: 'C3' } }];
    const first = await roundtripSheet(ws);
    const second = await roundtripSheet(first.ws);
    expect(second.ws['!images']![0]).toMatchObject({ type: 'png', anchor: { from: 'C3', width: 40, height: 30 } });
  });

  it('rejects unknown data and mismatched types', () => {
    const bad = (img: any) => () => {
      const ws = sheet();
      ws['!images'] = [img];
      XLSX.write(book(['S', ws]), { type: 'array' });
    };
    expect(bad({ data: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]), anchor: { from: 'A1' } })).toThrow(/Unsupported image/);
    expect(bad({ data: png, type: 'jpeg', anchor: { from: 'A1' } })).toThrow(/is png/);
  });
});
