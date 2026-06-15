import { lookup } from 'node:dns/promises';
import { isIP } from 'node:net';

const MAX_IMAGE_BYTES = 10 * 1024 * 1024;
const MAX_REDIRECTS = 3;

export type FetchedImage = {
  base64: string;
  mimeType: 'image/png' | 'image/jpeg';
  byteLength: number;
};

type LookupAddress = { address: string };
type ResolveHost = (hostname: string) => Promise<LookupAddress[]>;

export class ImageFetchError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ImageFetchError';
  }
}

export async function fetchImageAsBase64(url: string, fetchImpl: typeof fetch = fetch, resolveHost: ResolveHost = defaultResolveHost): Promise<FetchedImage> {
  const response = await fetchWithPolicy(new URL(url), fetchImpl, resolveHost, 0);
  if (!response.ok) throw new ImageFetchError(`Image fetch failed with HTTP ${response.status}.`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (bytes.byteLength > MAX_IMAGE_BYTES) throw new ImageFetchError('Image exceeds 10 MiB limit.');
  const mimeType = detectImageMime(bytes);
  if (!mimeType) throw new ImageFetchError('Only PNG and JPEG images are supported.');
  return { base64: Buffer.from(bytes).toString('base64'), mimeType, byteLength: bytes.byteLength };
}

async function fetchWithPolicy(url: URL, fetchImpl: typeof fetch, resolveHost: ResolveHost, redirectCount: number): Promise<Response> {
  await assertPublicHttpsUrl(url, resolveHost);
  const response = await fetchImpl(url, {
    method: 'GET',
    redirect: 'manual',
    headers: { accept: 'image/png,image/jpeg' }
  });
  if (isRedirect(response.status)) {
    if (redirectCount >= MAX_REDIRECTS) throw new ImageFetchError('Too many redirects while fetching image.');
    const location = response.headers.get('location');
    if (!location) throw new ImageFetchError('Redirect response is missing Location header.');
    return fetchWithPolicy(new URL(location, url), fetchImpl, resolveHost, redirectCount + 1);
  }
  return response;
}

async function assertPublicHttpsUrl(url: URL, resolveHost: ResolveHost): Promise<void> {
  if (url.protocol !== 'https:') throw new ImageFetchError('Image URL must use https.');
  if (url.username || url.password) throw new ImageFetchError('Image URL credentials are not allowed.');
  const records = await resolveHost(url.hostname);
  if (records.length === 0) throw new ImageFetchError('Image URL hostname did not resolve.');
  for (const record of records) {
    if (isForbiddenAddress(record.address)) {
      throw new ImageFetchError(`Image URL resolves to a private or reserved address: ${record.address}.`);
    }
  }
}

async function defaultResolveHost(hostname: string): Promise<LookupAddress[]> {
  return lookup(hostname, { all: true, verbatim: true });
}

export function isForbiddenAddress(address: string): boolean {
  const family = isIP(address);
  if (family === 4) return isForbiddenIPv4(address);
  if (family === 6) return isForbiddenIPv6(address);
  return true;
}

function isForbiddenIPv4(address: string): boolean {
  const parts = address.split('.').map((part) => Number.parseInt(part, 10));
  const [a, b] = parts;
  return (
    a === 0 ||
    a === 10 ||
    a === 127 ||
    (a === 100 && b >= 64 && b <= 127) ||
    (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 168) ||
    (a === 198 && (b === 18 || b === 19)) ||
    a >= 224
  );
}

function isForbiddenIPv6(address: string): boolean {
  const normalized = address.toLowerCase();
  return (
    normalized === '::' ||
    normalized === '::1' ||
    normalized.startsWith('fc') ||
    normalized.startsWith('fd') ||
    normalized.startsWith('fe80') ||
    normalized.startsWith('ff')
  );
}

function isRedirect(status: number): boolean {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

function detectImageMime(bytes: Uint8Array): FetchedImage['mimeType'] | null {
  if (bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47 && bytes[4] === 0x0d && bytes[5] === 0x0a && bytes[6] === 0x1a && bytes[7] === 0x0a) {
    return 'image/png';
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg';
  return null;
}
