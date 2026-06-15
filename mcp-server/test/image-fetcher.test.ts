import assert from 'node:assert/strict';
import test from 'node:test';
import { fetchImageAsBase64, isForbiddenAddress } from '../src/image-fetcher.js';

const PNG_1X1 = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]);

test('rejects private and reserved IP addresses', () => {
  assert.equal(isForbiddenAddress('127.0.0.1'), true);
  assert.equal(isForbiddenAddress('10.1.2.3'), true);
  assert.equal(isForbiddenAddress('172.16.0.1'), true);
  assert.equal(isForbiddenAddress('192.168.1.1'), true);
  assert.equal(isForbiddenAddress('169.254.169.254'), true);
  assert.equal(isForbiddenAddress('8.8.8.8'), false);
});

test('rejects non-https image URLs', async () => {
  await assert.rejects(fetchImageAsBase64('http://example.com/a.png', mockFetch(PNG_1X1), publicResolver), /https/);
});

test('fetches PNG images as base64 through the policy fetcher', async () => {
  const image = await fetchImageAsBase64('https://example.com/a.png', mockFetch(PNG_1X1), publicResolver);
  assert.equal(image.mimeType, 'image/png');
  assert.equal(image.base64, Buffer.from(PNG_1X1).toString('base64'));
});

test('rejects non-image payloads', async () => {
  await assert.rejects(fetchImageAsBase64('https://example.com/a.txt', mockFetch(Uint8Array.from([1, 2, 3])), publicResolver), /PNG and JPEG/);
});

function mockFetch(bytes: Uint8Array): typeof fetch {
  return (async () => new Response(bytes, { status: 200 })) as typeof fetch;
}

async function publicResolver() {
  return [{ address: '93.184.216.34' }];
}
