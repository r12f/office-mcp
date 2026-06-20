import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';

const DEFAULT_OFFICE_E2E_DRIVER = fileURLToPath(new URL('./office-e2e-driver.mjs', import.meta.url));
const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

export function advertisedTools(addinRoot) {
  const source = readFileSync(join(addinRoot, 'public', 'taskpane.js'), 'utf8');
  const match = source.match(/const AVAILABLE_TOOLS = \[([\s\S]*?)\];/);
  assert.ok(match, 'taskpane.js must declare AVAILABLE_TOOLS');
  return [...match[1].matchAll(/'([^']+)'/g)].map((tool) => tool[1]);
}

export function daemonCatalogTools(host, catalogPath = join(REPO_ROOT, 'src', 'office-mcp', 'daemon', 'src', 'mcp', 'catalog.rs')) {
  const prefix = hostToolPrefix(host);
  const source = readFileSync(catalogPath, 'utf8');
  const section = catalogSection(source, prefix);
  return [...section.matchAll(/"((?:word|excel|powerpoint)\.[^"]+)"/g)]
    .map((match) => match[1])
    .filter((tool, index, tools) => tool.startsWith(`${prefix}.`) && tools.indexOf(tool) === index)
    .sort();
}

export function assertE2eCaseCoverage({ addinRoot, host, cases, catalogPath }) {
  const tools = advertisedTools(addinRoot);
  const caseNames = Object.keys(cases).sort();
  assert.deepEqual(caseNames, [...tools].sort(), `${host} E2E cases must cover every advertised tool exactly`);
  const catalogTools = daemonCatalogTools(host, catalogPath);
  assert.deepEqual(caseNames, catalogTools, `${host} E2E cases must cover every daemon catalog tool exactly`);

  for (const tool of tools) {
    const toolCase = cases[tool];
    assert.equal(toolCase.tool, tool, `${tool} case must name the tool`);
    assert.ok(toolCase.setup, `${tool} case must define deterministic setup content`);
    assert.ok(toolCase.call, `${tool} case must define an MCP tool call`);
    assert.ok(toolCase.verify, `${tool} case must define a verifier`);
    assert.match(toolCase.verify.kind, /^(direct-result|readback)$/);
    if (toolCase.verify.kind === 'readback' && toolCase.verify.expect) {
      assert.ok(toolCase.verify.readbackTool || toolCase.verify.resource, `${tool} readback verifier must define readbackTool or resource`);
    }
  }
}

function hostToolPrefix(host) {
  const normalized = String(host || '').trim().toLowerCase();
  if (normalized === 'word' || normalized === 'excel' || normalized === 'powerpoint') return normalized;
  throw new Error(`Unsupported E2E host ${host}.`);
}

function catalogSection(source, prefix) {
  const anchors = {
    word: 'pub const WORD_V1_TOOLS',
    excel: 'const EXCEL_V1_TOOLS',
    powerpoint: 'const POWERPOINT_V1_TOOLS'
  };
  const start = source.indexOf(anchors[prefix]);
  assert.notEqual(start, -1, `daemon catalog must define ${anchors[prefix]}`);
  const end = source.indexOf('];', start);
  assert.notEqual(end, -1, `daemon catalog section ${anchors[prefix]} must close with ];`);
  return source.slice(start, end + 2);
}

export function e2eCase(tool, { setup = 'fixed baseline content', args = {}, verify = 'readback' } = {}) {
  return {
    tool,
    setup,
    call: { name: tool, arguments: args },
    verify: normalizeVerifier(verify)
  };
}

export function directResult(expect = {}) {
  return { kind: 'direct-result', expect };
}

export function readbackByTool(tool, { arguments: readbackArguments = {}, expect = {} } = {}) {
  return {
    kind: 'readback',
    readbackTool: tool,
    readbackArguments,
    expect
  };
}

export function readbackByResource(resource, { expect = {} } = {}) {
  return {
    kind: 'readback',
    resource,
    expect
  };
}

export const wordReadback = Object.freeze({
  documentText(expect, args = {}) {
    return readbackByTool('word.get_text', { arguments: { limit: 20, ...args }, expect });
  },
  outline(expect, args = {}) {
    return readbackByTool('word.get_outline', { arguments: args, expect });
  },
  paragraph(index, expect, args = {}) {
    return readbackByTool('word.get_paragraph', { arguments: { index, ...args }, expect });
  },
  table(tableIndex, expect, args = {}) {
    return readbackByTool('word.read_table', { arguments: { table_index: tableIndex, ...args }, expect });
  },
  contentControls(tag, expect, args = {}) {
    return readbackByTool('word.list_content_controls', { arguments: { tag, ...args }, expect });
  },
  comments(expect) {
    return readbackByResource('office://word/${session_id}/comments', { expect });
  },
  trackChanges(expect) {
    return readbackByResource('office://word/${session_id}/track_changes', { expect });
  }
});

export const excelReadback = Object.freeze({
  workbook(expect, args = {}) {
    return readbackByTool('excel.get_workbook_info', { arguments: args, expect });
  },
  sheets(expect, args = {}) {
    return readbackByTool('excel.list_sheets', { arguments: args, expect });
  },
  range(sheet, address, expect, args = {}) {
    return readbackByTool('excel.read_range', { arguments: { sheet, address, ...args }, expect });
  },
  table(table, expect, args = {}) {
    return readbackByTool('excel.update_table', { arguments: { table, action: 'metadata', ...args }, expect });
  },
  chart(sheet, chart, expect, args = {}) {
    return readbackByTool('excel.update_chart', { arguments: { sheet, chart, action: 'metadata', ...args }, expect });
  },
  pivotTable(pivotTable, expect, args = {}) {
    return readbackByTool('excel.update_pivot_table', { arguments: { pivot_table: pivotTable, action: 'metadata', ...args }, expect });
  }
});

export const powerpointReadback = Object.freeze({
  presentation(expect, args = {}) {
    return readbackByTool('powerpoint.get_presentation_info', { arguments: args, expect });
  },
  slides(expect, args = {}) {
    return readbackByTool('powerpoint.list_slides', { arguments: args, expect });
  },
  shapes(slideIndex, expect, args = {}) {
    return readbackByTool('powerpoint.list_shapes', { arguments: { slide_index: slideIndex, ...args }, expect });
  },
  text(slideIndex, expect, args = {}) {
    return readbackByTool('powerpoint.read_text', { arguments: { slide_index: slideIndex, ...args }, expect });
  },
  table(slideIndex, shapeId, expect, args = {}) {
    return readbackByTool('powerpoint.read_table', { arguments: { slide_index: slideIndex, shape_id: shapeId, ...args }, expect });
  },
  activeView(expect, args = {}) {
    return readbackByTool('powerpoint.get_active_view', { arguments: args, expect });
  },
  selection(expect, args = {}) {
    return readbackByTool('powerpoint.get_selection', { arguments: args, expect });
  },
  tags(expect, args = {}) {
    return readbackByTool('powerpoint.update_tags', { arguments: { action: 'list', ...args }, expect });
  },
  layouts(expect, args = {}) {
    return readbackByTool('powerpoint.list_layouts', { arguments: args, expect });
  }
});

function normalizeVerifier(verify) {
  if (typeof verify === 'string') return { kind: verify };
  assert.equal(typeof verify, 'object', 'E2E verifier must be a string or object');
  assert.ok(verify, 'E2E verifier object is required');
  return { ...verify };
}

export function officeE2eEnabled() {
  return process.env.OFFICE_MCP_RUN_E2E === '1';
}

export async function runOfficeToolE2e({ host, cases, driver }) {
  assert.ok(host, 'E2E host name is required');
  assert.ok(driver, `${host} E2E driver is required`);
  assertDriverMethod(driver, 'startDaemon', host);
  assertDriverMethod(driver, 'createDocument', host);
  assertDriverMethod(driver, 'waitForSession', host);
  assertDriverMethod(driver, 'resetContent', host);
  assertDriverMethod(driver, 'setupContent', host);
  assertDriverMethod(driver, 'callTool', host);
  assertDriverMethod(driver, 'verifyResult', host);

  let daemon;
  let document;
  let session;
  try {
    daemon = await driver.startDaemon({ host });
    document = await driver.createDocument({ host, daemon });
    session = await driver.waitForSession(document, { host, daemon });
    assertSessionCaseCoverage({ host, session, cases });

    for (const toolCase of orderedCases(cases, session.availableTools)) {
      const run = e2eRunMetadata(toolCase);
      const runSession = { ...session, bindings: { ...(session.bindings || {}) } };
      mergeBindings(runSession, await driver.resetContent(toolCase, runSession, { host, daemon, document, run }));
      mergeBindings(runSession, await driver.setupContent(toolCase, runSession, { host, daemon, document, run }));
      const result = await driver.callTool(toolCase, runSession, { host, daemon, document, run });
      await driver.verifyResult(toolCase, result, runSession, { host, daemon, document, run });
    }
  } finally {
    if (document && typeof driver.cleanupDocument === 'function') {
      await driver.cleanupDocument(document, { host, daemon, session });
    }
    if (typeof driver.stopDaemon === 'function') {
      await driver.stopDaemon(daemon, { host, document, session });
    }
  }
}

function mergeBindings(session, stepResult) {
  if (!stepResult?.bindings) return;
  session.bindings = { ...(session.bindings || {}), ...stepResult.bindings };
}

export function requireOfficeE2eDriver(host) {
  const script = process.env.OFFICE_MCP_E2E_DRIVER || DEFAULT_OFFICE_E2E_DRIVER;
  return createExternalOfficeE2eDriver(host, script);
}

function assertDriverMethod(driver, name, host) {
  assert.equal(typeof driver[name], 'function', `${host} E2E driver must implement ${name}()`);
}

function assertSessionCaseCoverage({ host, session, cases }) {
  assert.ok(session?.sessionId, `${host} E2E driver must return a sessionId`);
  assert.ok(Array.isArray(session.availableTools), `${host} E2E driver must return availableTools`);
  assert.deepEqual(
    [...session.availableTools].sort(),
    Object.keys(cases).sort(),
    `${host} E2E session tools must match the case table exactly`
  );
}

function orderedCases(cases, availableTools) {
  return availableTools.map((tool) => cases[tool]);
}

function e2eRunMetadata(toolCase) {
  return {
    id: toolCase.tool,
    tool: toolCase.tool,
    requestId: `e2e-${toolCase.tool.replace(/[^a-z0-9]+/gi, '-')}-${Date.now()}`,
    verify: toolCase.verify?.kind || 'readback'
  };
}

function createExternalOfficeE2eDriver(host, script) {
  return {
    async startDaemon(context) {
      return runExternalDriverStep(script, host, 'startDaemon', context);
    },
    async createDocument(context) {
      return runExternalDriverStep(script, host, 'createDocument', context);
    },
    async waitForSession(document, context) {
      return runExternalDriverStep(script, host, 'waitForSession', { ...context, document });
    },
    async resetContent(toolCase, session, context) {
      return runExternalDriverStep(script, host, 'resetContent', { ...context, toolCase, session });
    },
    async setupContent(toolCase, session, context) {
      return runExternalDriverStep(script, host, 'setupContent', { ...context, toolCase, session });
    },
    async callTool(toolCase, session, context) {
      return runExternalDriverStep(script, host, 'callTool', { ...context, toolCase, session });
    },
    async verifyResult(toolCase, result, session, context) {
      return runExternalDriverStep(script, host, 'verifyResult', { ...context, toolCase, result, session });
    },
    async cleanupDocument(document, context) {
      return runExternalDriverStep(script, host, 'cleanupDocument', { ...context, document });
    },
    async stopDaemon(daemon, context) {
      return runExternalDriverStep(script, host, 'stopDaemon', { ...context, daemon });
    }
  };
}

async function runExternalDriverStep(script, host, step, context = {}) {
  const request = { host, step, context };
  const { status, stdout, stderr } = await runDriverProcess(script, request);
  if (status !== 0) {
    const detail = stderr.trim() || stdout.trim() || 'no driver output';
    throw new Error(`${host} E2E driver step ${step} failed with exit code ${status}: ${detail}`);
  }
  const text = stdout.trim();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`${host} E2E driver step ${step} returned invalid JSON: ${error.message}`);
  }
}

function runDriverProcess(script, request) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [script], {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true
    });
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    child.on('error', reject);
    child.on('close', (status) => resolve({ status, stdout, stderr }));
    child.stdin.end(JSON.stringify(request));
  });
}
