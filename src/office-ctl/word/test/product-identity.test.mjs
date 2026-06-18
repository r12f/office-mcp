import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { inflateSync } from 'node:zlib';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

const REPO_ROOT = join(process.cwd(), '..', '..', '..');
const ASSET_ROOT = join(REPO_ROOT, 'src', 'office-ctl', 'common', 'assets');
const WORD_ROOT = join(REPO_ROOT, 'src', 'office-ctl', 'word');
const EXCEL_ROOT = join(REPO_ROOT, 'src', 'office-ctl', 'excel');
const ICON_SIZES = [16, 20, 24, 32, 48, 64, 80, 128, 256];
const PRODUCT_NAME = 'Office MCP Control';
const RENDERED_REVIEW_SURFACES = [
  ['logo_tray_size', 16],
  ['logo_ribbon_size', 32],
  ['logo_catalog_thumbnail', 80],
  ['logo_daemon_titlebar', 20],
  ['logo_installer_metadata', 256]
];
const BANNED_IDENTITY_PATTERNS = [
  /DefaultValue="office-mcp/i,
  /DefaultValue="Open"/,
  /Office MCP Project/,
  /placeholder/i,
  /debug/i,
  /prototype/i,
  /experiment/i,
  /gear[-_ ]?icon/i,
  /office[-_ ]?logo/i,
  /word[-_ ]?logo/i,
  /excel[-_ ]?logo/i,
  /fabric[-_ ]?icon/i,
  /fluent[-_ ]?emoji/i
];

test('Office add-ins use mature product identity metadata', () => {
  for (const [host, root, context] of [
    ['Word', WORD_ROOT, 'document'],
    ['Excel', EXCEL_ROOT, 'workbook']
  ]) {
    const manifest = readFileSync(join(root, 'manifest.xml'), 'utf8');
    const taskpane = readFileSync(join(root, 'public', 'taskpane.html'), 'utf8');

    assert.match(manifest, new RegExp(`<ProviderName>${PRODUCT_NAME}</ProviderName>`), `${host} provider is product branded`);
    assert.match(manifest, new RegExp(`<DisplayName DefaultValue="${PRODUCT_NAME}" \/>`), `${host} display name is product branded`);
    assert.match(manifest, /DefaultValue="Office MCP"/, `${host} ribbon group has a stable product label`);
    assert.match(manifest, /DefaultValue="Open Control Panel"/, `${host} command label is action-oriented`);
    assert.match(manifest, /local productivity automation control utility/, `${host} description states product type`);
    assert.match(manifest, new RegExp(`Open ${PRODUCT_NAME} for this ${context}`), `${host} tooltip names the product and host context`);
    assert.match(taskpane, new RegExp(`<title>${PRODUCT_NAME}</title>`), `${host} task pane title is product branded`);
    assert.match(taskpane, /<img class="product-mark" src="\/assets\/icon-32\.png" width="32" height="32" alt="" aria-hidden="true" \/>/, `${host} task pane chrome uses the product icon`);
    assert.match(taskpane, new RegExp(`<h1>${PRODUCT_NAME}</h1>`), `${host} task pane heading is product branded`);
    assert.doesNotMatch(taskpane, /<h1>\s*(Add-in|Task Pane|office-mcp)/i, `${host} task pane heading must not read as a scaffold`);
    for (const pattern of BANNED_IDENTITY_PATTERNS) {
      assert.doesNotMatch(manifest, pattern, `${host} manifest must not contain ${pattern}`);
    }
  }
});

test('brand design note documents the office-control visual brief', () => {
  const note = readFileSync(join(ASSET_ROOT, 'brand-design.md'), 'utf8');
  assert.match(note, /local office automation under user\s+control/i);
  assert.match(note, /abstract document or app panes/i);
  assert.match(note, /command spine/i);
  assert.match(note, /operator nodes/i);
  assert.match(note, /control dial/i);
  assert.match(note, /not a Microsoft Office app palette/i);
  assert.match(note, /16 px/);
  assert.match(note, /32 px/);
  assert.match(note, /80 px/);
  assert.match(note, /must not copy, trace, remix, or visually impersonate/i);
  assert.match(note, /Office ribbon tiles/i);
  assert.match(note, /Microsoft 365\s+multi-color gradients/i);
  assert.match(note, /Word document silhouettes/i);
  assert.match(note, /Excel grid marks/i);
  assert.match(note, /standalone gear icon/i);
});

test('generated brand icons are original non-placeholder assets', () => {
  const svg = readFileSync(join(ASSET_ROOT, 'brand-mark.svg'), 'utf8');
  assert.match(svg, /aria-label="Office MCP control mark"/);
  assert.match(svg, /#0F172A/);
  assert.match(svg, /#2457D6/);
  assert.match(svg, /#48D6A4/);
  assert.match(svg, /#F8D84A/);
  assert.match(svg, /M70 196h62c24 0 36-34 60-34h15/);
  assert.match(svg, /cx="203" cy="162" r="31"/);
  assert.doesNotMatch(svg, /office[-_ ]?logo|word[-_ ]?logo|excel[-_ ]?logo|gear/i);

  for (const size of ICON_SIZES) {
    const png = readFileSync(join(ASSET_ROOT, `icon-${size}.png`));
    const image = parsePngRgba(png);
    assert.equal(image.width, size);
    assert.equal(image.height, size);
    const colors = uniqueOpaqueColors(image.rgba);
    assert.ok(colors.size >= 4, `icon-${size}.png should contain the product palette, not a single-color placeholder`);
  }
});

test('rendered-size logo review artifact covers first-contact product surfaces', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-rendered-logo-review-'));
  try {
    const output = join(dir, 'logo-rendered-size-review.json');
    const sheet = join(dir, 'logo-rendered-size-review.png');
    const result = spawnSync(process.execPath, [
      join(ASSET_ROOT, '..', 'scripts', 'record-rendered-logo-review.mjs'),
      '--output', output,
      '--sheet', sheet
    ], { encoding: 'utf8' });
    assert.equal(result.status, 0, result.stdout + result.stderr);

    const report = JSON.parse(readFileSync(output, 'utf8'));
    assert.equal(report.kind, 'rendered_logo_review');
    assert.equal(report.product_name, PRODUCT_NAME);
    assert.equal(report.ready, true);
    assert.equal(report.sheet_path, sheet);
    assert.equal(report.surfaces.length, RENDERED_REVIEW_SURFACES.length);
    const sheetImage = parsePngRgba(readFileSync(sheet));
    assert.equal(sheetImage.width, 320 * RENDERED_REVIEW_SURFACES.length);
    assert.equal(sheetImage.height, 320);

    for (const [key, size] of RENDERED_REVIEW_SURFACES) {
      const surface = report.surfaces.find((item) => item.key === key);
      assert.ok(surface, `${key} is present in rendered logo review`);
      assert.equal(surface.rendered_size_px, size);
      assert.equal(surface.width, size);
      assert.equal(surface.height, size);
      assert.equal(surface.non_empty, true);
      assert.equal(surface.palette_ready, true);
      assert.equal(surface.expected_size_ready, true);
      assert.ok(surface.opaque_color_count >= 4, `${key} keeps product palette at rendered size`);
      assert.equal(surface.screenshot_path, sheet);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

function parsePngRgba(png) {
  assert.deepEqual([...png.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
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
      assert.equal(data[8], 8, 'brand icons must be 8-bit PNGs');
      assert.equal(data[9], 6, 'brand icons must be RGBA PNGs');
    } else if (type === 'IDAT') {
      idat.push(data);
    } else if (type === 'IEND') {
      break;
    }
    offset += 12 + length;
  }

  const raw = inflateSync(Buffer.concat(idat));
  const stride = width * 4;
  const rgba = Buffer.alloc(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    const filter = raw[y * (stride + 1)];
    assert.equal(filter, 0, 'brand icon generator uses unfiltered scanlines');
    raw.copy(rgba, y * stride, y * (stride + 1) + 1, y * (stride + 1) + 1 + stride);
  }
  return { width, height, rgba };
}

function uniqueOpaqueColors(rgba) {
  const colors = new Set();
  for (let offset = 0; offset < rgba.length; offset += 4) {
    if (rgba[offset + 3] > 0) {
      colors.add(`${rgba[offset]},${rgba[offset + 1]},${rgba[offset + 2]},${rgba[offset + 3]}`);
    }
  }
  return colors;
}

