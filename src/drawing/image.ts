/**
 * Image input handling: bytes from Buffer/ArrayBuffer/base64/data-URI, type
 * sniffing and pixel size from the image header.
 */
import { ImageType } from '../types.js';

export const IMAGE_CONTENT_TYPES: Record<ImageType, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
};

export const IMAGE_EXTENSIONS: Record<ImageType, string> = { png: 'png', jpeg: 'jpeg', gif: 'gif' };

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
  }
}
