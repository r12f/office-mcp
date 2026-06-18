import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { screenshotBytesLookLikeImage, screenshotFileLooksLikeImage, tinyPng } from './image-evidence.js';

test('image evidence accepts complete screenshots only', () => {
  assert.equal(screenshotBytesLookLikeImage(tinyPng()), true);
  assert.equal(screenshotBytesLookLikeImage(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00])), false);
  assert.equal(screenshotBytesLookLikeImage(Buffer.from('not an image')), false);
});

test('image evidence checks files before accepting screenshot paths', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-image-evidence-'));
  try {
    const valid = join(dir, 'valid.png');
    const truncated = join(dir, 'truncated.png');
    writeFileSync(valid, tinyPng());
    writeFileSync(truncated, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));

    assert.equal(screenshotFileLooksLikeImage(valid), true);
    assert.equal(screenshotFileLooksLikeImage(truncated), false);
    assert.equal(screenshotFileLooksLikeImage(join(dir, 'missing.png')), false);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
