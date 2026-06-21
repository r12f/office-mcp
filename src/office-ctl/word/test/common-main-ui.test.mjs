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
