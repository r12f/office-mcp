import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import zlib from 'node:zlib';

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const COMMON_ROOT = dirname(SCRIPT_DIR);
const ASSET_ROOT = join(COMMON_ROOT, 'assets');

const SIZES = [16, 20, 24, 32, 48, 64, 80, 128, 256];
const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

mkdirSync(ASSET_ROOT, { recursive: true });

for (const size of SIZES) {
  const png = renderPng(size);
  writeFileSync(join(ASSET_ROOT, `icon-${size}.png`), png);
}

function renderPng(size) {
  const rows = [];
  for (let y = 0; y < size; y += 1) {
    const row = Buffer.alloc(1 + size * 4);
    row[0] = 0;
    for (let x = 0; x < size; x += 1) {
      const [r, g, b, a] = pixel(size, x, y);
      const offset = 1 + x * 4;
      row[offset] = r;
      row[offset + 1] = g;
      row[offset + 2] = b;
      row[offset + 3] = a;
    }
    rows.push(row);
  }

  const raw = Buffer.concat(rows);
  return Buffer.concat([
    PNG_SIGNATURE,
    chunk('IHDR', ihdr(size, size)),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0))
  ]);
}

function pixel(size, x, y) {
  const scale = size / 256;
  const px = (x + 0.5) / scale;
  const py = (y + 0.5) / scale;
  const radius = 56;
  if (!roundedRect(px, py, 0, 0, 256, 256, radius)) return [0, 0, 0, 0];

  let color = [15, 23, 42, 255];
  if (roundedRect(px, py, 56, 66, 118, 124, 10)) color = [234, 242, 255, 255];
  if (roundedRect(px, py, 112, 54, 110, 148, 10)) color = [72, 214, 164, 255];
  if (line(px, py, 82, 92, 154, 92, 10) || line(px, py, 82, 118, 140, 118, 10) || line(px, py, 82, 144, 124, 144, 10)) {
    color = [36, 87, 214, 255];
  }
  if (line(px, py, 134, 92, 192, 92, 10) || line(px, py, 134, 120, 178, 120, 10) || line(px, py, 134, 148, 192, 148, 10)) {
    color = [15, 23, 42, 184];
  }
  const controlBorder = distance(px, py, 190, 188);
  if (controlBorder <= 40) color = [15, 23, 42, 255];
  if (controlBorder <= 28) color = [248, 216, 74, 255];
  if (line(px, py, 178, 188, 202, 188, 10) || line(px, py, 190, 176, 190, 200, 10)) {
    color = [15, 23, 42, 255];
  }
  return color;
}

function roundedRect(px, py, x, y, width, height, radius) {
  const rx = Math.max(x + radius, Math.min(px, x + width - radius));
  const ry = Math.max(y + radius, Math.min(py, y + height - radius));
  return distance(px, py, rx, ry) <= radius || (px >= x && px <= x + width && py >= y && py <= y + height);
}

function line(px, py, x1, y1, x2, y2, width) {
  const lengthSq = (x2 - x1) ** 2 + (y2 - y1) ** 2;
  const t = Math.max(0, Math.min(1, ((px - x1) * (x2 - x1) + (py - y1) * (y2 - y1)) / lengthSq));
  const ix = x1 + t * (x2 - x1);
  const iy = y1 + t * (y2 - y1);
  return distance(px, py, ix, iy) <= width / 2;
}

function distance(x1, y1, x2, y2) {
  return Math.hypot(x1 - x2, y1 - y2);
}

function ihdr(width, height) {
  const data = Buffer.alloc(13);
  data.writeUInt32BE(width, 0);
  data.writeUInt32BE(height, 4);
  data[8] = 8;
  data[9] = 6;
  return data;
}

function chunk(type, data) {
  const typeBuffer = Buffer.from(type, 'ascii');
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])), 0);
  return Buffer.concat([length, typeBuffer, data, crc]);
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}
