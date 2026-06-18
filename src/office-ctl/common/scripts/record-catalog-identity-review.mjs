import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';

const PRODUCT_NAME = 'Office MCP Control';
const PRODUCT_TYPE = 'Local productivity automation control utility';
const PRODUCT_DESCRIPTION = /local productivity automation control utility/i;
const BANNED_FIRST_IMPRESSION = /(sample|debug|prototype|experimental|protocol bridge|mcp server|developer tool|office-mcp-(word|excel|powerpoint)|DefaultValue="Open"|Task Pane|Add-in)/i;

const catalogPath = resolve(readOption('--catalog-path') ?? 'addin-catalog');
const outputPath = resolve(readOption('--output') ?? 'artifacts/catalog-identity-review.json');
const catalogType = readOption('--catalog-type') ?? PRODUCT_TYPE;

const hosts = [
  { key: 'word', label: 'Word', file: 'office-mcp-word.xml', path: '/word/taskpane.html' },
  { key: 'excel', label: 'Excel', file: 'office-mcp-excel.xml', path: '/excel/taskpane.html' },
  { key: 'powerpoint', label: 'PowerPoint', file: 'office-mcp-powerpoint.xml', path: '/powerpoint/taskpane.html' }
];

const hostReports = hosts.map((host) => reviewHost(host));
const origins = new Set(hostReports.map((host) => host.origin).filter(Boolean));
const failures = [
  ...hostReports.flatMap((host) => host.failures.map((failure) => `${host.label}: ${failure}`)),
  ...(catalogTypeLooksReady(catalogType) ? [] : [`Catalog type is not product-ready: ${catalogType}`]),
  ...(origins.size === 1 ? [] : ['Catalog manifests do not share one daemon origin.'])
];

const report = {
  schema_version: 1,
  kind: 'catalog_identity_review',
  recorded_at: new Date().toISOString(),
  product_name: PRODUCT_NAME,
  catalog_path: catalogPath,
  catalog_type: catalogType,
  shared_origin: origins.size === 1 ? [...origins][0] : null,
  hosts: Object.fromEntries(hostReports.map((host) => [host.key, host])),
  ready: failures.length === 0,
  failures
};

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, JSON.stringify(report, null, 2));
console.log(JSON.stringify(report, null, 2));
if (!report.ready) process.exit(1);

function reviewHost(host) {
  const manifestPath = join(catalogPath, host.file);
  const xml = readFileSync(manifestPath, 'utf8');
  const identity = {
    key: host.key,
    label: host.label,
    manifest_path: manifestPath,
    display_name: extractDefaultValue(xml, 'DisplayName'),
    provider: extractElementText(xml, 'ProviderName'),
    description: extractDefaultValue(xml, 'Description'),
    icon_url: extractDefaultValue(xml, 'IconUrl'),
    high_resolution_icon_url: extractDefaultValue(xml, 'HighResolutionIconUrl'),
    command_label: extractStringDefaultValue(xml, 'OfficeMcp.OpenPane.Label'),
    tooltip: extractStringDefaultValue(xml, 'OfficeMcp.OpenPane.Tooltip'),
    taskpane_url: extractUrlDefaultValue(xml, 'Taskpane.Url'),
    origin: null,
    ready: false,
    failures: []
  };
  identity.origin = originOf(identity.taskpane_url) ?? originOf(identity.icon_url) ?? originOf(identity.high_resolution_icon_url);
  identity.failures = hostFailures(host, identity, xml);
  identity.ready = identity.failures.length === 0;
  return identity;
}

function hostFailures(host, identity, xml) {
  const failures = [];
  if (identity.display_name !== PRODUCT_NAME) failures.push(`display name must be ${PRODUCT_NAME}.`);
  if (identity.provider !== PRODUCT_NAME) failures.push(`provider must be ${PRODUCT_NAME}.`);
  if (!PRODUCT_DESCRIPTION.test(identity.description ?? '')) failures.push('description must describe a local productivity automation control utility.');
  if (identity.command_label !== 'Open Control Panel') failures.push('ribbon command must be Open Control Panel.');
  if (typeof identity.tooltip !== 'string' || !identity.tooltip.includes(PRODUCT_NAME)) failures.push('tooltip must include the product name.');
  if (typeof identity.taskpane_url !== 'string' || !identity.taskpane_url.includes(host.path)) failures.push(`task pane URL must target ${host.path}.`);
  if (!assetUrlLooksReady(identity.icon_url, '/assets/icon-32.png')) failures.push('catalog/ribbon icon URL must use generated icon-32.png.');
  if (!assetUrlLooksReady(identity.high_resolution_icon_url, '/assets/icon-80.png')) failures.push('high-resolution icon URL must use generated icon-80.png.');
  if (BANNED_FIRST_IMPRESSION.test(xml)) failures.push('manifest contains prototype, debug, raw package, generic add-in, or generic task-pane wording.');
  return failures;
}

function extractDefaultValue(xml, elementName) {
  return matchFirst(xml, new RegExp(`<${elementName}\\b[^>]*DefaultValue="([^"]+)"`, 'i'));
}

function extractStringDefaultValue(xml, id) {
  return matchFirst(xml, new RegExp(`<bt:String\\b[^>]*id="${escapeRegExp(id)}"[^>]*DefaultValue="([^"]+)"`, 'i'));
}

function extractUrlDefaultValue(xml, id) {
  return matchFirst(xml, new RegExp(`<bt:Url\\b[^>]*id="${escapeRegExp(id)}"[^>]*DefaultValue="([^"]+)"`, 'i'));
}

function extractElementText(xml, elementName) {
  return matchFirst(xml, new RegExp(`<${elementName}>([^<]+)</${elementName}>`, 'i'));
}

function matchFirst(value, pattern) {
  return value.match(pattern)?.[1] ?? null;
}

function assetUrlLooksReady(value, suffix) {
  return typeof value === 'string' && value.startsWith('https://localhost:') && value.endsWith(suffix);
}

function catalogTypeLooksReady(value) {
  return typeof value === 'string' && /local productivity automation control utility/i.test(value) && !BANNED_FIRST_IMPRESSION.test(value);
}

function originOf(value) {
  if (typeof value !== 'string') return null;
  try {
    const url = new URL(value);
    return url.origin;
  } catch {
    return null;
  }
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function readOption(name) {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}
