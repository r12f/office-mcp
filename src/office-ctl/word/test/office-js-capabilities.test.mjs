import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();
const OFFICE_JS_TYPES = readFileSync(join(ADDIN_ROOT, 'node_modules', '@types', 'office-js', 'index.d.ts'), 'utf8');
const TASKPANE = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

test('Office.js Word typings expose tracked-change read and mutation primitives used by v1', () => {
  assert.match(OFFICE_JS_TYPES, /getTrackedChanges\(\): Word\.TrackedChangeCollection/);
  assert.match(OFFICE_JS_TYPES, /accept\(\): void/);
  assert.match(OFFICE_JS_TYPES, /reject\(\): void/);
});

test('Office.js Word typings expose save but not portable save-as or PDF export APIs', () => {
  assert.match(OFFICE_JS_TYPES, /save\(saveBehavior\?: Word\.SaveBehavior, fileName\?: string\): void/);
  assert.doesNotMatch(OFFICE_JS_TYPES, /saveAs\s*\(/i);
  assert.doesNotMatch(OFFICE_JS_TYPES, /exportPdf\s*\(/i);
  assert.doesNotMatch(OFFICE_JS_TYPES, /exportAsPdf\s*\(/i);
});

test('Word task pane does not implement reserved document-output tools', () => {
  assert.doesNotMatch(TASKPANE, /word\.save_as/);
  assert.doesNotMatch(TASKPANE, /word\.export_pdf/);
});
