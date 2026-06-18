import { existsSync, readFileSync, statSync } from 'node:fs';

export function screenshotFileLooksLikeImage(path: string): boolean {
  if (!existsSync(path) || !statSync(path).isFile()) return false;
  const bytes = readFileSync(path);
  return screenshotBytesLookLikeImage(bytes);
}

export function screenshotBytesLookLikeImage(bytes: Buffer): boolean {
  if (bytes.length < 32) return false;
  const header = bytes.subarray(0, 12);
  const isPng = header.length >= 8 && header[0] === 0x89 && header[1] === 0x50 && header[2] === 0x4e && header[3] === 0x47 && header[4] === 0x0d && header[5] === 0x0a && header[6] === 0x1a && header[7] === 0x0a;
  const isJpeg = header.length >= 3 && header[0] === 0xff && header[1] === 0xd8 && header[2] === 0xff;
  const isWebp = header.length >= 12 && header.subarray(0, 4).toString('ascii') === 'RIFF' && header.subarray(8, 12).toString('ascii') === 'WEBP';
  const isBmp = header.length >= 2 && header[0] === 0x42 && header[1] === 0x4d;
  if (isPng) return bytes.includes(Buffer.from('IEND', 'ascii'));
  if (isJpeg) return bytes.at(-2) === 0xff && bytes.at(-1) === 0xd9;
  if (isWebp) return bytes.length >= 64;
  if (isBmp) return bytes.length >= 54;
  return false;
}

export function tinyPng(): Buffer {
  return Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/luznWQAAAABJRU5ErkJggg==',
    'base64'
  );
}
