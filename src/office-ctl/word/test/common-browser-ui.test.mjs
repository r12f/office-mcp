import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const REPO_ROOT = join(process.cwd(), '..', '..', '..');

function loadCommon() {
  const source = readFileSync(join(REPO_ROOT, 'src', 'office-ctl', 'common', 'browser-ui.js'), 'utf8');
  const context = vm.createContext({
    globalThis: {},
    Intl,
    Date,
    Number,
    String,
    Object,
    RegExp
  });
  context.globalThis = context;
  vm.runInContext(source, context, { filename: 'browser-ui.js' });
  return context.OfficeCtlCommon;
}

test('common browser UI helpers redact sensitive task text', () => {
  const common = loadCommon();

  const redacted = common.redactText('Bearer abc token=secret base64,QUJDREVGRw==');

  assert.equal(redacted, 'Bearer [redacted] token=[redacted] base64,[redacted]');
});

test('common browser UI helpers format labels and filenames', () => {
  const common = loadCommon();

  assert.equal(common.boolLabel(true), 'yes');
  assert.equal(common.boolLabel(null), 'unknown');
  assert.equal(common.escapeHtml('<b>&"'), '&lt;b&gt;&amp;&quot;');
  assert.equal(common.fileName('C:\\Docs\\Report.docx'), 'Report.docx');
  assert.equal(common.titleCase('timed_out'), 'Timed Out');
  assert.match(common.formatDuration(1250), /1\.3s|1\.25s|1\.2s/);
});
