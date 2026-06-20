import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

export function advertisedTools(addinRoot) {
  const source = readFileSync(join(addinRoot, 'public', 'taskpane.js'), 'utf8');
  const match = source.match(/const AVAILABLE_TOOLS = \[([\s\S]*?)\];/);
  assert.ok(match, 'taskpane.js must declare AVAILABLE_TOOLS');
  return [...match[1].matchAll(/'([^']+)'/g)].map((tool) => tool[1]);
}

export function assertE2eCaseCoverage({ addinRoot, host, cases }) {
  const tools = advertisedTools(addinRoot);
  const caseNames = Object.keys(cases).sort();
  assert.deepEqual(caseNames, [...tools].sort(), `${host} E2E cases must cover every advertised tool exactly`);

  for (const tool of tools) {
    const toolCase = cases[tool];
    assert.equal(toolCase.tool, tool, `${tool} case must name the tool`);
    assert.ok(toolCase.setup, `${tool} case must define deterministic setup content`);
    assert.ok(toolCase.call, `${tool} case must define an MCP tool call`);
    assert.ok(toolCase.verify, `${tool} case must define a verifier`);
    assert.match(toolCase.verify.kind, /^(direct-result|readback)$/);
  }
}

export function e2eCase(tool, { setup = 'fixed baseline content', args = {}, verify = 'readback' } = {}) {
  return {
    tool,
    setup,
    call: { name: tool, arguments: args },
    verify: { kind: verify }
  };
}

export function realOfficeE2eEnabled() {
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
      await driver.resetContent(toolCase, session, { host, daemon, document });
      await driver.setupContent(toolCase, session, { host, daemon, document });
      const result = await driver.callTool(toolCase, session, { host, daemon, document });
      await driver.verifyResult(toolCase, result, session, { host, daemon, document });
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

export function requireRealOfficeE2eDriver(host) {
  const script = process.env.OFFICE_MCP_E2E_DRIVER;
  assert.ok(
    script,
    `${host} real Office E2E driver requires OFFICE_MCP_E2E_DRIVER; unset OFFICE_MCP_RUN_E2E or provide a host automation driver.`
  );
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

async function runExternalDriverStep(_script, host, step) {
  assert.fail(`${host} real Office E2E external driver step ${step} is not wired yet.`);
}
