import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

test('Word task pane opts the current document into Office auto-open after connect', () => {
  const source = readFileSync(join(process.cwd(), '..', 'addin', 'public', 'taskpane.js'), 'utf8');

  assert.match(source, /Office\.AutoShowTaskpaneWithDocument/);
  assert.match(source, /Office\.context\.document\.settings\.saveAsync/);
  assert.match(source, /await enableDocumentAutoOpen\(\)/);
});

test('Word add-in manifest and task pane asset versions stay aligned', () => {
  const manifest = readFileSync(join(process.cwd(), '..', 'addin', 'manifest.xml'), 'utf8');
  const html = readFileSync(join(process.cwd(), '..', 'addin', 'public', 'taskpane.html'), 'utf8');
  const js = readFileSync(join(process.cwd(), '..', 'addin', 'public', 'taskpane.js'), 'utf8');

  assert.match(manifest, /<Version>1\.0\.0\.5<\/Version>/);
  assert.match(manifest, /taskpane\.html\?v=0\.1\.5/);
  assert.match(html, /taskpane\.css\?v=0\.1\.5/);
  assert.match(html, /taskpane\.js\?v=0\.1\.5/);
  assert.match(js, /ADDIN_VERSION = '0\.1\.5'/);
});

test('Word task pane exposes product UI regions and accessible endpoint settings', () => {
  const html = readFileSync(join(process.cwd(), '..', 'addin', 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(process.cwd(), '..', 'addin', 'public', 'taskpane.css'), 'utf8');
  const js = readFileSync(join(process.cwd(), '..', 'addin', 'public', 'taskpane.js'), 'utf8');

  assert.match(html, /id="connectionBadge"/);
  assert.match(html, /id="currentTask"/);
  assert.match(html, /id="historyList"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /type="url" inputmode="url" autocomplete="off" spellcheck="false"/);
  assert.match(html, /aria-label="Open Settings"/);
  assert.match(css, /:focus-visible/);
  assert.match(css, /prefers-reduced-motion/);
  assert.match(css, /forced-colors: active/);
  assert.match(js, /taskHistory\.splice\(20\)/);
  assert.match(js, /localStorage\.setItem\('office-mcp\.addin-endpoint'/);
  assert.match(js, /redactText/);
});
