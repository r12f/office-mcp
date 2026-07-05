import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
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

export function assertConcreteE2eCases({ host, cases }) {
  assert.ok(host, 'E2E host name is required');
  assert.equal(typeof cases, 'object', `${host} E2E cases must be an object`);
  assert.ok(cases, `${host} E2E cases are required`);

  for (const [tool, toolCase] of Object.entries(cases)) {
    assert.equal(toolCase?.tool, tool, `${tool} case must name the tool`);
    assertConcreteSetup(tool, toolCase);
    assertConcreteCall(tool, toolCase);
    assertConcreteVerifier(tool, toolCase);
    for (const [index, scenario] of (toolCase.scenarios || []).entries()) {
      const scenarioName = `${tool} scenario ${index}`;
      assert.equal(scenario.tool, tool, `${scenarioName} must call the canonical tool`);
      assertConcreteSetup(scenarioName, scenario);
      assertConcreteCall(tool, scenario);
      assertConcreteVerifier(scenarioName, scenario);
    }
  }
}

function assertConcreteSetup(tool, toolCase) {
  assert.ok(Array.isArray(toolCase?.setup?.actions), `${tool} case must define setup.actions`);
  assert.ok(toolCase.setup.actions.length > 0, `${tool} case must define at least one setup action`);
  for (const [index, action] of toolCase.setup.actions.entries()) {
    assert.equal(typeof action, 'object', `${tool} setup action ${index} must be an object`);
    assert.ok(action, `${tool} setup action ${index} is required`);
    const target = action.tool || action.resource || action.driver;
    assert.equal(typeof target, 'string', `${tool} setup action ${index} must define tool, resource, or driver`);
    assert.ok(target.length > 0, `${tool} setup action ${index} must define a non-empty tool, resource, or driver`);
    if (action.tool) {
      assert.equal(typeof action.arguments, 'object', `${tool} setup action ${index} must define arguments`);
      assert.ok(action.arguments, `${tool} setup action ${index} arguments are required`);
    }
  }
}

function assertConcreteCall(tool, toolCase) {
  assert.equal(toolCase?.call?.name, tool, `${tool} case must call the same tool it covers`);
  assert.equal(typeof toolCase.call.arguments, 'object', `${tool} case must define call arguments`);
  assert.ok(toolCase.call.arguments, `${tool} case call arguments are required`);
}

function assertConcreteVerifier(tool, toolCase) {
  const verifier = toolCase?.verify;
  assert.ok(verifier, `${tool} case must define a verifier`);
  assert.match(verifier.kind, /^(direct-result|readback)$/, `${tool} case must use a supported verifier kind`);
  assertConcreteExpectation(tool, verifier);
  if (verifier.kind === 'readback') {
    assert.ok(verifier.readbackTool || verifier.resource, `${tool} readback verifier must define readbackTool or resource`);
    if (verifier.readbackTool) {
      assert.equal(typeof verifier.readbackTool, 'string', `${tool} readbackTool must be a string`);
      assert.equal(typeof verifier.readbackArguments, 'object', `${tool} readback verifier must define readbackArguments`);
      assert.ok(verifier.readbackArguments, `${tool} readbackArguments are required`);
    }
    if (verifier.resource) {
      assert.equal(typeof verifier.resource, 'string', `${tool} readback resource must be a string`);
    }
  }
}

function assertConcreteExpectation(tool, verifier) {
  assert.equal(typeof verifier.expect, 'object', `${tool} verifier must define expect`);
  assert.ok(verifier.expect, `${tool} verifier expect is required`);
  const expectationKeys = Object.keys(verifier.expect);
  const allowErrorCodes = Array.isArray(verifier.allowErrorCodes) ? verifier.allowErrorCodes : [];
  assert.ok(
    expectationKeys.length > 0 || allowErrorCodes.length > 0,
    `${tool} verifier must define at least one expectation or allowed host capability error`
  );
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

export function e2eCase(tool, { setup = 'fixed baseline content', args = {}, verify = 'readback', scenarios = [] } = {}) {
  return {
    tool,
    setup,
    call: { name: tool, arguments: args },
    verify: normalizeVerifier(verify),
    scenarios: scenarios.map((scenario) => e2eCase(tool, scenario))
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
    return readbackByTool('word.get_text', { arguments: { offset: index, limit: 1, include_metadata: true, ...args }, expect });
  },
  table(tableIndex, expect, args = {}) {
    return readbackByTool('word.read_table', { arguments: { table_index: tableIndex, ...args }, expect });
  },
  contentControls(tag, expect, args = {}) {
    return readbackByTool('word.list_content_controls', { arguments: { tag, ...args }, expect });
  },
  hyperlinks(expect, args = {}) {
    return readbackByTool('word.list_hyperlinks', { arguments: args, expect });
  },
  bookmarks(expect, args = {}) {
    return readbackByTool('word.list_bookmarks', { arguments: args, expect });
  },
  notes(kind, expect, args = {}) {
    return readbackByTool('word.list_notes', { arguments: { kind, ...args }, expect });
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

export async function runOfficeToolE2e({ host, cases, driver, reportPath }) {
  assert.ok(host, 'E2E host name is required');
  assert.ok(driver, `${host} E2E driver is required`);
  assertDriverMethod(driver, 'startDaemon', host);
  assertDriverMethod(driver, 'listTools', host);
  assertDriverMethod(driver, 'createDocument', host);
  assertDriverMethod(driver, 'waitForSession', host);
  assertDriverMethod(driver, 'resetContent', host);
  assertDriverMethod(driver, 'setupContent', host);
  assertDriverMethod(driver, 'callTool', host);
  assertDriverMethod(driver, 'verifyResult', host);
  assertDriverMethod(driver, 'cleanupDocument', host);
  assertDriverMethod(driver, 'stopDaemon', host);

  const report = createE2eReport(host);
  let daemon;
  let document;
  let session;
  try {
    report.lifecycle_counts.start_daemon += 1;
    daemon = await driver.startDaemon({ host });
    report.daemon = summarizeDaemon(daemon);
    report.lifecycle_counts.list_tools += 1;
    const daemonTools = await driver.listTools({ host, daemon });
    assertDaemonToolsCaseCoverage({ host, tools: daemonTools, cases });
    report.advertised_tools = hostNamedTools(host, daemonTools);
    report.lifecycle_counts.create_document += 1;
    document = await driver.createDocument({ host, daemon });
    report.document = summarizeDocument(document);
    if (typeof driver.activateAddin === 'function') {
      report.lifecycle_counts.activate_addin += 1;
      const activation = await driver.activateAddin(document, { host, daemon });
      applyActivatedDocument(document, activation);
      applyActivationArtifacts(document, activation);
      report.addin_activation = summarizeActivation(activation);
      assertActivationProof(host, report.addin_activation);
      report.document = summarizeDocument(document);
    }
    report.lifecycle_counts.wait_for_session += 1;
    session = await driver.waitForSession(document, { host, daemon });
    assertSessionCaseCoverage({ host, session, cases });
    report.session = summarizeSession(session);
    report.session_available_tools = [...session.availableTools];

    for (const toolCase of orderedCases(cases, session.availableTools).flatMap(expandToolCaseScenarios)) {
      const run = e2eRunMetadata(toolCase);
      const toolRun = createToolRunReport(toolCase, run);
      report.tool_runs.push(toolRun);
      const runSession = { ...session, bindings: { ...(session.bindings || {}) } };
      try {
        mergeBindings(runSession, await driver.resetContent(toolCase, runSession, { host, daemon, document, run }));
        mergeBindings(runSession, await driver.setupContent(toolCase, runSession, { host, daemon, document, run }));
        if (runSession.bindings?.__accepted_error_code) {
          toolRun.result = { ok: false, error_code: runSession.bindings.__accepted_error_code };
          toolRun.accepted_error_code = runSession.bindings.__accepted_error_code;
          toolRun.passed = true;
          continue;
        }
        const result = await driver.callTool(toolCase, runSession, { host, daemon, document, run });
        toolRun.result = summarizeToolResult(result);
        await driver.verifyResult(toolCase, result, runSession, { host, daemon, document, run });
        toolRun.passed = true;
      } catch (error) {
        toolRun.passed = false;
        toolRun.error = serializeError(error);
        throw error;
      } finally {
        toolRun.finished_at = new Date().toISOString();
      }
    }
    report.passed = true;
  } finally {
    let cleanupError;
    try {
      if (document) {
        report.lifecycle_counts.cleanup_document += 1;
        report.cleanup = summarizeCleanup(await driver.cleanupDocument(document, { host, daemon, session }));
        assertCleanupProof(host, report.cleanup);
      }
    } catch (error) {
      cleanupError = error;
      report.passed = false;
      report.error = serializeError(error);
    }
    try {
      report.lifecycle_counts.stop_daemon += 1;
      await driver.stopDaemon(daemon, { host, document, session });
    } catch (error) {
      report.passed = false;
      report.error = serializeError(error);
      throw error;
    } finally {
      report.finished_at = new Date().toISOString();
      report.executed_tools = report.tool_runs.map((run) => run.tool);
      if (report.passed !== true) {
        report.passed = false;
        report.error ??= firstToolRunError(report);
      }
      writeE2eReport(reportPath, report);
    }
    if (cleanupError) throw cleanupError;
  }
}

function applyActivatedDocument(document, activation) {
  if (!document || !activation || typeof activation.document_path !== 'string') return;
  if (!activation.document_path.trim()) return;
  document.original_path ??= document.path;
  document.path = activation.document_path;
}

function applyActivationArtifacts(document, activation) {
  if (!document || !activation || typeof activation.log_path !== 'string') return;
  document.activationLogPath = activation.log_path;
}

function createE2eReport(host) {
  return {
    schema_version: 1,
    kind: 'office_tool_e2e_report',
    host,
    started_at: new Date().toISOString(),
    finished_at: undefined,
    passed: false,
    lifecycle_counts: {
      start_daemon: 0,
      list_tools: 0,
      create_document: 0,
      activate_addin: 0,
      wait_for_session: 0,
      cleanup_document: 0,
      stop_daemon: 0
    },
    advertised_tools: [],
    session_available_tools: [],
    executed_tools: [],
    tool_runs: []
  };
}

function createToolRunReport(toolCase, run) {
  return {
    id: run.requestId,
    tool: toolCase.tool,
    started_at: new Date().toISOString(),
    finished_at: undefined,
    setup_action_count: Array.isArray(toolCase.setup?.actions) ? toolCase.setup.actions.length : 0,
    verifier: summarizeVerifier(toolCase.verify),
    passed: false
  };
}

function summarizeVerifier(verifier) {
  return {
    kind: verifier?.kind,
    readback_tool: verifier?.readbackTool,
    resource: verifier?.resource,
    expectation_keys: verifier?.expect && typeof verifier.expect === 'object' ? Object.keys(verifier.expect).sort() : []
  };
}

function summarizeDaemon(daemon) {
  if (!daemon || typeof daemon !== 'object') return undefined;
  return typeof daemon.endpoint === 'string' ? { endpoint: daemon.endpoint } : {};
}

function summarizeDocument(document) {
  if (!document || typeof document !== 'object') return undefined;
  return typeof document.path === 'string' ? { path: document.path } : {};
}

function summarizeActivation(activation) {
  if (!activation || typeof activation !== 'object') return undefined;
  return {
    activated: activation.activated === true,
    skipped: typeof activation.skipped === 'string' ? activation.skipped : undefined,
    activator: typeof activation.activator === 'string' ? activation.activator : undefined,
    activation_path: typeof activation.activation_path === 'string' ? activation.activation_path : undefined,
    control_opened: typeof activation.control_opened === 'boolean' ? activation.control_opened : undefined
  };
}

function assertActivationProof(host, activation) {
  if (!activation || typeof activation !== 'object') throw new Error(`${host} E2E add-in activation proof is missing.`);
  if (activation.activated !== true) throw new Error(`${host} E2E add-in activation did not run.`);
  if (typeof activation.skipped === 'string' && activation.skipped.length > 0) throw new Error(`${host} E2E add-in activation skipped: ${activation.skipped}`);
  if (typeof activation.activator !== 'string' || activation.activator.trim().length === 0) throw new Error(`${host} E2E add-in activator identity is missing.`);
  if (typeof activation.activation_path !== 'string' || activation.activation_path.trim().length === 0) throw new Error(`${host} E2E add-in activation path proof is missing.`);
}

function summarizeSession(session) {
  if (!session || typeof session !== 'object') return undefined;
  return {
    session_id: session.sessionId,
    available_tool_count: Array.isArray(session.availableTools) ? session.availableTools.length : undefined
  };
}

function summarizeCleanup(cleanup) {
  if (!cleanup || typeof cleanup !== 'object') return undefined;
  const deletedPaths = Array.isArray(cleanup.deletedPaths)
    ? cleanup.deletedPaths.filter((path) => typeof path === 'string' && path.trim().length > 0)
    : [];
  return {
    closed_by_driver: cleanup.closedByDriver === true,
    deleted: cleanup.deleted === true,
    deleted_path_count: deletedPaths.length,
    deleted_paths: [...new Set(deletedPaths)],
    skipped: typeof cleanup.skipped === 'string' ? cleanup.skipped : undefined
  };
}

function assertCleanupProof(host, cleanup) {
  if (!cleanup || typeof cleanup !== 'object') throw new Error(`${host} E2E cleanup proof is missing.`);
  if (typeof cleanup.skipped === 'string' && cleanup.skipped.length > 0) throw new Error(`${host} E2E cleanup skipped: ${cleanup.skipped}`);
  if (cleanup.closed_by_driver !== true) throw new Error(`${host} E2E cleanup did not close the driver-owned document.`);
  if (cleanup.deleted !== true) throw new Error(`${host} E2E cleanup did not delete the driver-owned document.`);
  if (typeof cleanup.deleted_path_count !== 'number' || cleanup.deleted_path_count < 1) throw new Error(`${host} E2E cleanup missing deleted path proof.`);
  if (!Array.isArray(cleanup.deleted_paths) || cleanup.deleted_paths.length < 1) throw new Error(`${host} E2E cleanup missing deleted paths.`);
}

function summarizeToolResult(result) {
  if (!result || typeof result !== 'object') return undefined;
  return {
    ok: result.ok,
    has_structured_content: Object.hasOwn(result, 'structuredContent'),
    has_content: Object.hasOwn(result, 'content'),
    error_code: result.error?.code
  };
}

function serializeError(error) {
  if (error instanceof Error) {
    return { name: error.name, message: error.message };
  }
  return { name: 'Error', message: String(error) };
}

function firstToolRunError(report) {
  return report.tool_runs.find((run) => run.error)?.error;
}

function writeE2eReport(reportPath, report) {
  if (!reportPath) return;
  mkdirSync(dirname(reportPath), { recursive: true });
  writeFileSync(reportPath, JSON.stringify(report, null, 2));
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

function assertDaemonToolsCaseCoverage({ host, tools, cases }) {
  const hostTools = hostNamedTools(host, tools);
  assert.deepEqual(
    hostTools,
    Object.keys(cases).sort(),
    `${host} E2E daemon tools/list must match the case table exactly`
  );
}

function hostNamedTools(host, tools) {
  assert.ok(Array.isArray(tools), `${host} E2E driver must return daemon tools/list tools`);
  const prefix = `${hostToolPrefix(host)}.`;
  return tools
    .map((tool) => (typeof tool === 'string' ? tool : tool?.name))
    .filter((name) => typeof name === 'string' && name.startsWith(prefix))
    .sort();
}

function orderedCases(cases, availableTools) {
  return availableTools.map((tool) => cases[tool]);
}

function expandToolCaseScenarios(toolCase) {
  return [toolCase, ...(toolCase.scenarios || [])];
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
    async listTools(context) {
      return runExternalDriverStep(script, host, 'listTools', context);
    },
    async createDocument(context) {
      return runExternalDriverStep(script, host, 'createDocument', context);
    },
    async activateAddin(document, context) {
      return runExternalDriverStep(script, host, 'activateAddin', { ...context, document });
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
