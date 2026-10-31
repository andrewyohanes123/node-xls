/**
 * Image input handling: bytes from Buffer/ArrayBuffer/base64/data-URI, type
 * sniffing and pixel size from the image header.
 */
import { ImageType } from '../types.js';

export const IMAGE_CONTENT_TYPES: Record<ImageType, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export const IMAGE_EXTENSIONS: Record<ImageType, string> = { png: 'png', jpeg: 'jpeg', gif: 'gif', bmp: 'bmp', tiff: 'tiff', emf: 'emf', wmf: 'wmf' };

/** Image type for a media part's file extension (reader fallback). */
export function imageTypeFromExtension(path: string): ImageType | undefined {
  const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
  const map: Record<string, ImageType> = { png: 'png', jpg: 'jpeg', jpeg: 'jpeg', jpe: 'jpeg', gif: 'gif', bmp: 'bmp', dib: 'bmp', tif: 'tiff', tiff: 'tiff', emf: 'emf', wmf: 'wmf' };
  return map[ext];
}

export function toBytes(data: Uint8Array | ArrayBuffer | string): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (typeof data === 'string') {
    const b64 = data.replace(/^data:[^;,]*;base64,/, '');
    const buf = Buffer.from(b64, 'base64');
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  }
  throw new Error('Image data must be a Buffer, Uint8Array, ArrayBuffer or base64 string');
}

export function detectImageType(b: Uint8Array): ImageType | undefined {
  if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpeg';
  if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return 'gif';
  if (b.length > 26 && b[0] === 0x42 && b[1] === 0x4d) return 'bmp';
  if (b.length > 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 0x2a))) return 'tiff';
  if (b.length > 44 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46) return 'emf';
  if (b.length > 4 && ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || ((b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0))) return 'wmf';
  return undefined;
}

/** Pixel dimensions read from the image header. */
export function imageSize(b: Uint8Array, type: ImageType): { width: number; height: number } | undefined {
  const be16 = (i: number) => (b[i] << 8) | b[i + 1];
  const be32 = (i: number) => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
  switch (type) {
    case 'png':
      return b.length >= 24 ? { width: be32(16), height: be32(20) } : undefined;
    case 'gif':
      return b.length >= 10 ? { width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) } : undefined;
    case 'bmp': {
      if (b.length < 26) return undefined;
      const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
      // OS/2 BITMAPCOREHEADER (12 bytes) stores 16-bit sizes; later headers 32-bit
      if (v.getUint32(14, true) === 12) return { width: v.getUint16(18, true), height: v.getUint16(20, true) };
      return { width: Math.abs(v.getInt32(18, true)), height: Math.abs(v.getInt32(22, true)) };
    }
    case 'jpeg': {
      let i = 2;
      while (i + 9 < b.length) {
        if (b[i] !== 0xff) return undefined;
        const marker = b[i + 1];
        if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
          i += 2;
          continue;
        }
        const len = be16(i + 2);
        // SOF0..SOF15 except DHT (C4), JPG (C8) and DAC (CC)
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { height: be16(i + 5), width: be16(i + 7) };
        }
        i += 2 + len;
      }
      return undefined;
    }
    default:
      return undefined;
  }
}
