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
  const radius = 48;
  if (!roundedRect(px, py, 0, 0, 256, 256, radius)) return [0, 0, 0, 0];

  let color = [15, 23, 42, 255];
  if (leftPane(px, py)) color = [234, 242, 255, 255];
  if (middlePane(px, py)) color = [75, 215, 168, 255];
  if (rightPane(px, py)) color = [36, 59, 122, 255];
  if (line(px, py, 64, 76, 130, 76, 10) || line(px, py, 64, 104, 118, 104, 10) || line(px, py, 64, 132, 128, 132, 10)) {
    color = [36, 87, 214, 255];
  }
  if (line(px, py, 94, 104, 162, 104, 10) || line(px, py, 94, 132, 144, 132, 10) || line(px, py, 94, 160, 156, 160, 10)) {
    color = [15, 23, 42, 184];
  }
  if (line(px, py, 124, 132, 178, 132, 10) || line(px, py, 124, 160, 160, 160, 10)) {
    color = [234, 242, 255, 235];
  }
  if (line(px, py, 48, 199, 92, 199, 16) || line(px, py, 92, 199, 147, 173, 16) || line(px, py, 147, 173, 169, 173, 16) || line(px, py, 169, 173, 211, 152, 16) || line(px, py, 211, 152, 218, 152, 16)) {
    color = [36, 87, 214, 255];
  }
  const firstNode = distance(px, py, 88, 199);
  if (firstNode <= 21) color = [15, 23, 42, 255];
  if (firstNode <= 13) color = [75, 215, 168, 255];
  const secondNode = distance(px, py, 147, 173);
  if (secondNode <= 21) color = [15, 23, 42, 255];
  if (secondNode <= 13) color = [234, 242, 255, 255];
  const controlBorder = distance(px, py, 211, 152);
  if (controlBorder <= 40) color = [15, 23, 42, 255];
  if (controlBorder <= 28) color = [248, 216, 74, 255];
  if (line(px, py, 192, 152, 230, 152, 10) || line(px, py, 211, 133, 211, 171, 10) || controlBorder <= 7) {
    color = [15, 23, 42, 255];
  }
  return color;
}

function leftPane(px, py) {
  return roundedRect(px, py, 46, 54, 122, 138, 22);
}

function middlePane(px, py) {
  return roundedRect(px, py, 74, 78, 126, 132, 22);
}

function rightPane(px, py) {
  return roundedRect(px, py, 102, 102, 108, 108, 22);
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
