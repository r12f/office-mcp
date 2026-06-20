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

export function requireRealOfficeE2eDriver(host) {
  assert.fail(`${host} real Office E2E driver is not implemented yet; unset OFFICE_MCP_RUN_E2E or implement the host driver.`);
}
