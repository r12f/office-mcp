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
  const radius = 52;
  if (!roundedRect(px, py, 0, 0, 256, 256, radius)) return [0, 0, 0, 0];

  let color = [15, 23, 42, 255];
  if (leftPane(px, py)) color = [234, 242, 255, 255];
  if (middlePane(px, py)) color = [75, 215, 168, 255];
  if (rightPane(px, py)) color = [36, 59, 122, 255];
  if (line(px, py, 70, 80, 119, 80, 9) || line(px, py, 70, 106, 106, 106, 9) || line(px, py, 70, 132, 118, 132, 9)) {
    color = [36, 87, 214, 255];
  }
  if (line(px, py, 106, 102, 154, 102, 9) || line(px, py, 106, 128, 140, 128, 9) || line(px, py, 106, 154, 152, 154, 9)) {
    color = [15, 23, 42, 184];
  }
  if (line(px, py, 146, 124, 182, 124, 9) || line(px, py, 146, 150, 172, 150, 9) || line(px, py, 146, 176, 180, 176, 9)) {
    color = [234, 242, 255, 235];
  }
  if (line(px, py, 54, 198, 97, 198, 13) || line(px, py, 97, 198, 139, 180, 13) || line(px, py, 139, 180, 167, 180, 13) || line(px, py, 167, 180, 210, 152, 13) || line(px, py, 210, 152, 217, 152, 13)) {
    color = [36, 87, 214, 255];
  }
  const firstNode = distance(px, py, 72, 198);
  if (firstNode <= 16.5) color = [15, 23, 42, 255];
  if (firstNode <= 9.5) color = [75, 215, 168, 255];
  const secondNode = distance(px, py, 139, 180);
  if (secondNode <= 16.5) color = [15, 23, 42, 255];
  if (secondNode <= 9.5) color = [234, 242, 255, 255];
  const controlBorder = distance(px, py, 210, 152);
  if (controlBorder <= 35.5) color = [15, 23, 42, 255];
  if (controlBorder <= 24.5) color = [248, 216, 74, 255];
  if (line(px, py, 194, 152, 226, 152, 9) || line(px, py, 210, 136, 210, 168, 9) || controlBorder <= 6) {
    color = [15, 23, 42, 255];
  }
  return color;
}

function leftPane(px, py) {
  return roundedRect(px, py, 50, 48, 93, 116, 20);
}

function middlePane(px, py) {
  return roundedRect(px, py, 86, 70, 95, 118, 18);
}

function rightPane(px, py) {
  return roundedRect(px, py, 122, 92, 80, 112, 18);
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
