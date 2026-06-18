import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import zlib from 'node:zlib';

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const COMMON_ROOT = dirname(SCRIPT_DIR);
const REPO_ROOT = resolve(COMMON_ROOT, '..', '..', '..');
const ASSET_ROOT = join(COMMON_ROOT, 'assets');
const outputPath = resolve(readOption('--output') ?? join(REPO_ROOT, 'artifacts/logo-rendered-size-review.json'));
const sheetPath = resolve(readOption('--sheet') ?? join(REPO_ROOT, 'artifacts/logo-rendered-size-review.png'));

const SURFACES = [
  { key: 'logo_tray_size', label: 'Tray 16 px', size: 16, asset: 'icon-16.png' },
  { key: 'logo_ribbon_size', label: 'Ribbon 32 px', size: 32, asset: 'icon-32.png' },
  { key: 'logo_catalog_thumbnail', label: 'Catalog 80 px', size: 80, asset: 'icon-80.png' },
  { key: 'logo_daemon_titlebar', label: 'Title bar 20 px', size: 20, asset: 'icon-20.png' },
  { key: 'logo_installer_metadata', label: 'Installer 256 px', size: 256, asset: 'icon-256.png' }
];

const CELL_SIZE = 320;
const sheet = blankRgba(CELL_SIZE * SURFACES.length, CELL_SIZE, [247, 248, 250, 255]);
const surfaceReports = [];

SURFACES.forEach((surface, index) => {
  const iconPath = join(ASSET_ROOT, surface.asset);
  const icon = parsePngRgba(readFileSync(iconPath));
  const colors = uniqueOpaqueColors(icon.rgba);
  const bounds = opaqueBounds(icon);
  const cellX = index * CELL_SIZE;
  drawCell(sheet, cellX, surface.size);
  blit(sheet, icon, cellX + Math.floor((CELL_SIZE - icon.width) / 2), Math.floor((CELL_SIZE - icon.height) / 2));
  surfaceReports.push({
    key: surface.key,
    label: surface.label,
    asset_path: iconPath,
    rendered_size_px: surface.size,
    width: icon.width,
    height: icon.height,
    opaque_color_count: colors.size,
    opaque_bounds: bounds,
    non_empty: bounds !== null,
    palette_ready: colors.size >= 4,
    expected_size_ready: icon.width === surface.size && icon.height === surface.size,
    screenshot_path: sheetPath
  });
});

mkdirSync(dirname(sheetPath), { recursive: true });
writeFileSync(sheetPath, encodePng(sheet.width, sheet.height, sheet.rgba));

const ready = surfaceReports.every((surface) => surface.non_empty && surface.palette_ready && surface.expected_size_ready);
const report = {
  schema_version: 1,
  kind: 'rendered_logo_review',
  recorded_at: new Date().toISOString(),
  product_name: 'Office MCP Control',
  sheet_path: sheetPath,
  surfaces: surfaceReports,
  ready
};

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, JSON.stringify(report, null, 2));
console.log(JSON.stringify(report, null, 2));
if (!ready) process.exit(1);

function drawCell(sheet, cellX, iconSize) {
  drawRect(sheet, cellX + 8, 8, CELL_SIZE - 16, CELL_SIZE - 16, [255, 255, 255, 255]);
  drawBorder(sheet, cellX + 8, 8, CELL_SIZE - 16, CELL_SIZE - 16, [216, 222, 230, 255]);
  drawRect(sheet, cellX + 16, 16, Math.max(24, iconSize), 4, [36, 87, 214, 255]);
  drawRect(sheet, cellX + 16, 26, Math.max(24, iconSize), 4, [72, 214, 164, 255]);
}

function blankRgba(width, height, color) {
  const rgba = Buffer.alloc(width * height * 4);
  for (let offset = 0; offset < rgba.length; offset += 4) {
    rgba[offset] = color[0];
    rgba[offset + 1] = color[1];
    rgba[offset + 2] = color[2];
    rgba[offset + 3] = color[3];
  }
  return { width, height, rgba };
}

function drawRect(image, x, y, width, height, color) {
  for (let yy = y; yy < y + height; yy += 1) {
    for (let xx = x; xx < x + width; xx += 1) setPixel(image, xx, yy, color);
  }
}

function drawBorder(image, x, y, width, height, color) {
  drawRect(image, x, y, width, 1, color);
  drawRect(image, x, y + height - 1, width, 1, color);
  drawRect(image, x, y, 1, height, color);
  drawRect(image, x + width - 1, y, 1, height, color);
}

function setPixel(image, x, y, color) {
  if (x < 0 || y < 0 || x >= image.width || y >= image.height) return;
  const offset = (y * image.width + x) * 4;
  image.rgba[offset] = color[0];
  image.rgba[offset + 1] = color[1];
  image.rgba[offset + 2] = color[2];
  image.rgba[offset + 3] = color[3];
}

function blit(target, source, x, y) {
  for (let yy = 0; yy < source.height; yy += 1) {
    for (let xx = 0; xx < source.width; xx += 1) {
      const sourceOffset = (yy * source.width + xx) * 4;
      const alpha = source.rgba[sourceOffset + 3] / 255;
      if (alpha === 0) continue;
      const targetX = x + xx;
      const targetY = y + yy;
      const targetOffset = (targetY * target.width + targetX) * 4;
      for (let channel = 0; channel < 3; channel += 1) {
        target.rgba[targetOffset + channel] = Math.round(source.rgba[sourceOffset + channel] * alpha + target.rgba[targetOffset + channel] * (1 - alpha));
      }
      target.rgba[targetOffset + 3] = 255;
    }
  }
}

function parsePngRgba(png) {
  if (!png.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) throw new Error('Not a PNG file.');
  let offset = 8;
  let width = 0;
  let height = 0;
  const idat = [];
  while (offset < png.length) {
    const length = png.readUInt32BE(offset);
    const type = png.subarray(offset + 4, offset + 8).toString('ascii');
    const data = png.subarray(offset + 8, offset + 8 + length);
    if (type === 'IHDR') {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      if (data[8] !== 8 || data[9] !== 6) throw new Error('Expected 8-bit RGBA PNG.');
    } else if (type === 'IDAT') {
      idat.push(data);
    } else if (type === 'IEND') {
      break;
    }
    offset += 12 + length;
  }
  const raw = zlib.inflateSync(Buffer.concat(idat));
  const stride = width * 4;
  const rgba = Buffer.alloc(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    if (raw[y * (stride + 1)] !== 0) throw new Error('Expected unfiltered PNG scanline.');
    raw.copy(rgba, y * stride, y * (stride + 1) + 1, y * (stride + 1) + 1 + stride);
  }
  return { width, height, rgba };
}

function opaqueBounds(image) {
  let minX = image.width;
  let minY = image.height;
  let maxX = -1;
  let maxY = -1;
  for (let y = 0; y < image.height; y += 1) {
    for (let x = 0; x < image.width; x += 1) {
      const alpha = image.rgba[(y * image.width + x) * 4 + 3];
      if (alpha > 0) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }
  return maxX === -1 ? null : { min_x: minX, min_y: minY, max_x: maxX, max_y: maxY };
}

function uniqueOpaqueColors(rgba) {
  const colors = new Set();
  for (let offset = 0; offset < rgba.length; offset += 4) {
    if (rgba[offset + 3] > 0) colors.add(`${rgba[offset]},${rgba[offset + 1]},${rgba[offset + 2]},${rgba[offset + 3]}`);
  }
  return colors;
}

function encodePng(width, height, rgba) {
  const rows = [];
  const stride = width * 4;
  for (let y = 0; y < height; y += 1) {
    const row = Buffer.alloc(1 + stride);
    row[0] = 0;
    rgba.copy(row, 1, y * stride, y * stride + stride);
    rows.push(row);
  }
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', ihdr(width, height)),
    chunk('IDAT', zlib.deflateSync(Buffer.concat(rows), { level: 9 })),
    chunk('IEND', Buffer.alloc(0))
  ]);
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
    for (let index = 0; index < 8; index += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function readOption(name) {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}
