import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const REPO_ROOT = join(process.cwd(), '..', '..', '..');

function loadMainUi(overrides = {}) {
  const source = readFileSync(join(REPO_ROOT, 'src', 'office-ctl', 'common', 'main-ui.js'), 'utf8');
  const context = vm.createContext({
    Object,
    String,
    Math,
    globalThis: {},
    ...overrides
  });
  context.globalThis = context;
  vm.runInContext(source, context, { filename: 'main-ui.js' });
  return context.OfficeCtlMainUi;
}

function buttonFixture({ value, targetId, targetText = '' } = {}) {
  const target = { textContent: targetText };
  const button = {
    dataset: { ...(value !== undefined ? { copyValue: value } : {}), ...(targetId ? { copyTarget: targetId } : {}) },
    getAttribute(name) { return name === 'aria-label' ? 'Copy session ID' : ''; }
  };
  return {
    button,
    document: {
      getElementById(id) { return id === targetId ? target : null; }
    },
    event: {
      target: {
        closest(selector) { return selector === '[data-copy-target], [data-copy-value]' ? button : null; }
      }
    }
  };
}

test('common main UI copies metadata from explicit values and target elements', async () => {
  const mainUi = loadMainUi();
  const copied = [];
  const announcer = { textContent: '' };
  const direct = buttonFixture({ value: 'session-123' });

  assert.equal(await mainUi.copyMetadataValue(direct.event, {
    document: direct.document,
    navigator: { clipboard: { writeText: async (value) => copied.push(value) } },
    announcer
  }), true);
  assert.deepEqual(copied, ['session-123']);
  assert.equal(announcer.textContent, 'Copied Copy session ID');

  const targeted = buttonFixture({ targetId: 'session', targetText: 'session-from-target' });
  assert.equal(await mainUi.copyMetadataValue(targeted.event, {
    document: targeted.document,
    navigator: { clipboard: { writeText: async (value) => copied.push(value) } },
    announcer
  }), true);
  assert.deepEqual(copied, ['session-123', 'session-from-target']);
});

test('common main UI falls back and reports copy failures', async () => {
  const mainUi = loadMainUi();
  const fallbackValues = [];
  const warnings = [];
  const announcer = { textContent: '' };
  const fixture = buttonFixture({ value: 'daemon-endpoint' });

  assert.equal(await mainUi.copyMetadataValue(fixture.event, {
    document: fixture.document,
    navigator: {},
    announcer,
    fallbackCopy: (value) => fallbackValues.push(value)
  }), true);
  assert.deepEqual(fallbackValues, ['daemon-endpoint']);

  assert.equal(await mainUi.copyMetadataValue(fixture.event, {
    document: fixture.document,
    navigator: { clipboard: { writeText: async () => { throw new Error('denied'); } } },
    announcer,
    logger: { warn: (event, error) => warnings.push([event, error.message]) }
  }), false);
  assert.deepEqual(warnings, [['metadata_copy.failed', 'denied']]);
  assert.equal(announcer.textContent, 'Copy failed');
});

test('common main UI owns textarea clipboard fallback', async () => {
  const appended = [];
  const removed = [];
  const commands = [];
  const area = {
    value: '',
    style: {},
    setAttribute(name, value) { this[name] = value; },
    select() { this.selected = true; },
    remove() { removed.push(this); }
  };
  const document = {
    createElement(tag) {
      assert.equal(tag, 'textarea');
      return area;
    },
    body: {
      appendChild(element) { appended.push(element); }
    },
    execCommand(command) {
      commands.push(command);
      return true;
    }
  };
  const mainUi = loadMainUi({ document });
  const fixture = buttonFixture({ value: 'fallback-value' });

  assert.equal(await mainUi.copyMetadataValue(fixture.event, {
    document: { ...document, getElementById: fixture.document.getElementById },
    navigator: {},
    announcer: { textContent: '' }
  }), true);

  assert.equal(area.value, 'fallback-value');
  assert.equal(area.readonly, '');
  assert.equal(area.style.position, 'fixed');
  assert.equal(area.style.opacity, '0');
  assert.equal(area.selected, true);
  assert.deepEqual(appended, [area]);
  assert.deepEqual(commands, ['copy']);
  assert.deepEqual(removed, [area]);
});

test('common main UI maps connection and task states to badge classes', () => {
  const mainUi = loadMainUi();

  assert.equal(mainUi.statusClass('connected'), 'status-success');
  assert.equal(mainUi.statusClass('success'), 'status-success');
  assert.equal(mainUi.statusClass('connecting'), 'status-warning');
  assert.equal(mainUi.statusClass('reconnecting'), 'status-warning');
  assert.equal(mainUi.statusClass('running'), 'status-warning');
  assert.equal(mainUi.statusClass('failed'), 'status-danger');
  assert.equal(mainUi.statusClass('failure'), 'status-danger');
  assert.equal(mainUi.statusClass('unsupported'), 'status-danger');
  assert.equal(mainUi.statusClass('cancelled'), 'status-neutral');
  assert.equal(mainUi.statusClass('idle'), 'status-neutral');
});

test('common main UI renders connection state consistently', () => {
  const mainUi = loadMainUi();
  const badge = { textContent: '', className: '' };
  const detail = { textContent: 'Connecting' };
  const announcer = { textContent: '' };

  mainUi.setConnectionState({ badge, detail, announcer }, 'connected', 'Connected');

  assert.equal(badge.textContent, 'Connected');
  assert.equal(badge.className, 'status-badge status-success');
  assert.equal(detail.textContent, 'None');
  assert.equal(announcer.textContent, 'Connected');

  mainUi.setConnectionState({ badge, detail, announcer }, 'failed', 'Failed');

  assert.equal(badge.textContent, 'Failed');
  assert.equal(badge.className, 'status-badge status-danger');
  assert.equal(detail.textContent, 'None');
  assert.equal(announcer.textContent, 'Failed');
});

test('common main UI renders document protection and editability labels', () => {
  const mainUi = loadMainUi();

  assert.equal(mainUi.protectionLabel({}), 'Not protected');
  assert.equal(mainUi.protectionLabel({ protection: { kind: 'IRM' } }), 'IRM');
  assert.equal(mainUi.protectionLabel({ protection: { label: 'presentation protection' } }), 'presentation protection');
  assert.equal(mainUi.protectionLabel({ is_protected: true }), 'Protected');
  assert.equal(mainUi.protectionLabel({ protection: { kind: 'none' } }), 'Not protected');

  assert.equal(mainUi.documentStateLabel({}), 'Editable');
  assert.equal(mainUi.documentStateLabel({ is_dirty: true }), 'Editable, unsaved changes');
  assert.equal(mainUi.documentStateLabel({ is_read_only: true }), 'Read-only');
  assert.equal(mainUi.documentStateLabel({ is_protected: true }), 'Protected');
  assert.equal(mainUi.documentStateLabel({ protection: { kind: 'IRM' } }), 'Protected: IRM');
  assert.equal(mainUi.documentStateLabel({ protection: { label: 'presentation protection' } }), 'Protected: presentation protection');
});

test('common main UI renders static task pane metadata', () => {
  const mainUi = loadMainUi({
    Office: {
      context: {
        diagnostics: { host: 'Excel', version: '16.0' },
        platform: 'PC'
      }
    }
  });
  const session = { id: 'session', textContent: '', closest: () => null };
  const daemon = { id: 'daemon', textContent: '', closest: () => null };
  const server = { textContent: '' };
  const protocol = { textContent: '' };
  const hostPlatform = { textContent: '' };

  mainUi.renderStaticMetadata({
    session,
    daemon,
    serverVersion: server,
    protocolVersion: protocol,
    hostPlatform
  }, {
    sessionId: 'session-123',
    endpoint: 'wss://localhost:8765/addin',
    serverInfo: { serverVersion: '0.1.0', protocolVersion: '1.0' },
    protocolVersion: '1.0',
    defaultHost: 'Office'
  });

  assert.equal(session.textContent, 'session-123');
  assert.equal(daemon.textContent, 'wss://localhost:8765/addin');
  assert.equal(server.textContent, 'Server 0.1.0');
  assert.equal(protocol.textContent, 'Protocol 1.0');
  assert.equal(hostPlatform.textContent, 'Excel 16.0 / PC');
});

test('common main UI renders tool capability mode selection', () => {
  const mainUi = loadMainUi();
  const states = new Map();
  const buttons = ['read', 'write', 'all'].map((mode) => ({
    dataset: { toolMode: mode },
    setAttribute(name, value) { states.set(`${mode}:${name}`, value); }
  }));
  const control = {
    querySelectorAll(selector) {
      assert.equal(selector, '[data-tool-mode]');
      return buttons;
    }
  };

  mainUi.renderToolModeControl(control, 'write');

  assert.equal(states.get('read:aria-checked'), 'false');
  assert.equal(states.get('write:aria-checked'), 'true');
  assert.equal(states.get('all:aria-checked'), 'false');
});

test('common main UI renders copyable command IDs with middle truncation', () => {
  const mainUi = loadMainUi();
  const requestId = '0123456789abcdefghijklmnopqrstuvwxyz';

  const markup = mainUi.commandIdMarkup(requestId, { escapeHtml: (value) => String(value) });

  assert.match(markup, /class="task-meta task-command-id"/);
  assert.match(markup, /data-copy-value="0123456789abcdefghijklmnopqrstuvwxyz"/);
  assert.match(markup, /aria-label="Copy command ID"/);
  assert.match(markup, /title="0123456789abcdefghijklmnopqrstuvwxyz"/);
  assert.match(markup, /<code>0123456789abcd\.\.\.nopqrstuvwxyz<\/code>/);
  assert.equal(mainUi.commandIdMarkup('', { escapeHtml: (value) => String(value) }), '');
});

test('common main UI renders redacted task metadata rows', () => {
  const mainUi = loadMainUi();
  const task = {
    userIntent: 'replace token=secret text',
    deadlineAt: 1700000000000,
    cancelRequested: true,
    error: {
      office_mcp_code: 'HOST_ERROR',
      message: 'failed <unsafe>',
      retriable: false,
      partial_effect: 'none'
    }
  };

  const markup = mainUi.taskMetadataMarkup(task, {
    escapeHtml: (value) => String(value).replace(/</g, '&lt;').replace(/>/g, '&gt;'),
    formatTime: (value) => `time-${value}`,
    redactText: (value) => String(value).replace('token=secret', 'token=[redacted]'),
    valueLabel: (value) => value === true ? 'yes' : value === false ? 'no' : 'unknown'
  });

  assert.match(markup, /HOST_ERROR: failed &lt;unsafe&gt;/);
  assert.match(markup, /Retriable: no/);
  assert.match(markup, /Partial effect: none/);
  assert.match(markup, /replace token=\[redacted\] text/);
  assert.match(markup, /Deadline time-1700000000000/);
  assert.match(markup, /Cancel requested/);
  assert.equal(mainUi.taskMetadataMarkup({}, { escapeHtml: String, formatTime: String, redactText: String, valueLabel: String }), '');
});
