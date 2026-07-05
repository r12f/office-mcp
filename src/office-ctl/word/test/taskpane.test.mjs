import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();

test('Word task pane opts the current document into Office auto-open after connect', () => {
  const source = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(source, /function isWordHost\(info\)/);
  assert.match(source, /Office\.HostType\?\.Word/);
  assert.match(source, /Office\.AutoShowTaskpaneWithDocument/);
  assert.match(source, /Office\.context\.document\.settings\.saveAsync/);
  assert.match(source, /await enableDocumentAutoOpen\(\)/);
});

test('Word add-in manifest and task pane asset versions stay aligned', () => {
  const manifest = readFileSync(join(ADDIN_ROOT, 'manifest.xml'), 'utf8');
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(manifest, /<Version>1\.0\.0\.18<\/Version>/);
  assert.match(manifest, /word\/taskpane\.html\?v=0\.1\.18/);
  assert.match(html, /common\/taskpane\.css\?v=0\.1\.18/);
  assert.match(html, /common\/browser-ui\.js\?v=0\.1\.18/);
  assert.match(html, /common\/addin-channel\.js\?v=0\.1\.18/);
  assert.match(html, /common\/logger\.js\?v=0\.1\.18/);
  assert.match(html, /common\/task-history\.js\?v=0\.1\.18/);
  assert.match(html, /common\/main-ui\.js\?v=0\.1\.18/);
  assert.match(html, /word\/taskpane\.js\?v=0\.1\.18/);
  assert.match(html, /<script src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
  assert.doesNotMatch(html, /<script async src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
  assert.match(js, /ADDIN_VERSION = '0\.1\.18'/);
});

test('Word add-in uses product identity metadata and generated icons', () => {
  const manifest = readFileSync(join(ADDIN_ROOT, 'manifest.xml'), 'utf8');
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const assetRoot = join(ADDIN_ROOT, '..', 'common', 'assets');

  assert.match(manifest, /<ProviderName>Office MCP Control<\/ProviderName>/);
  assert.match(manifest, /<DisplayName DefaultValue="Office MCP Control" \/>/);
  assert.match(manifest, /Control live Word documents through a local productivity automation control utility\./);
  assert.match(manifest, /<bt:String id="OfficeMcp\.GroupLabel" DefaultValue="Office MCP Control" \/>/);
  assert.match(manifest, /<bt:String id="OfficeMcp\.OpenPane\.Label" DefaultValue="Open Control Panel" \/>/);
  assert.match(manifest, /Office MCP Control for this document/);
  assert.doesNotMatch(manifest, /DefaultValue="office-mcp"/);
  assert.doesNotMatch(manifest, /DefaultValue="Open"/);
  assert.match(html, /<title>MCP Control<\/title>/);
  assert.match(html, /<img class="product-mark" src="\/assets\/icon-32\.png" width="32" height="32" alt="" aria-hidden="true" \/>/);
  assert.match(html, /<h1>MCP Control<\/h1>/);
  for (const size of [16, 20, 24, 32, 48, 64, 80, 128, 256]) {
    const png = readFileSync(join(assetRoot, `icon-${size}.png`));
    assert.deepEqual([...png.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
    assert.deepEqual(pngDimensions(png), [size, size]);
  }
  assert.match(readFileSync(join(assetRoot, 'brand-mark.svg'), 'utf8'), /Office MCP control mark/);
});

test('Office catalog registration wrapper delegates to shared Office catalog script', () => {
  const script = readFileSync(join(ADDIN_ROOT, 'scripts', 'register-word-catalog.ps1'), 'utf8');
  const commonScript = readFileSync(join(ADDIN_ROOT, '..', 'common', 'scripts', 'register-office-catalog.ps1'), 'utf8');

  assert.match(script, /register-office-catalog\.ps1/);
  assert.match(commonScript, /src\\office-ctl\\word\\manifest\.xml/);
  assert.match(commonScript, /src\\office-ctl\\excel\\manifest\.xml/);
  assert.match(commonScript, /src\\office-ctl\\powerpoint\\manifest\.xml/);
  assert.match(commonScript, /Word manifest:/);
  assert.match(commonScript, /Excel manifest:/);
  assert.match(commonScript, /PowerPoint manifest:/);
  assert.match(commonScript, /TrustedCatalogId = "\{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57\}"/);
  assert.match(commonScript, /TrustedCatalogRegistryRoot = "HKCU:\\Software\\Microsoft\\Office\\16\.0\\WEF\\TrustedCatalogs"/);
  assert.match(commonScript, /\$TrustedCatalogRegistryRoot\\\$TrustedCatalogId/);
  assert.match(commonScript, /\$TrustedCatalogRegistryRoot\\office-mcp/);
  assert.match(commonScript, /ClearOfficeCache/);
  assert.match(commonScript, /SkipOfficeCache/);
  assert.match(commonScript, /Remove-CustomUiValidationCache/);
  assert.match(commonScript, /CustomUIValidationCache/);
  assert.match(commonScript, /\$shouldClearOfficeCache = \(-not \$SkipOfficeCache\) -and \(\(-not \$SkipRegistry\) -or \$ClearOfficeCache\)/);
  assert.match(commonScript, /Close Office host processes before clearing add-in cache/);
  assert.match(commonScript, /SkipRegistry/);
  assert.match(commonScript, /Remove-DeveloperDebugRegistration/);
  assert.match(commonScript, /WEF\\\\Developer|WEF\\Developer/);
});

test('Office catalog registration script stages Office host manifests without registry mutation', () => {
  const catalogPath = mkdtempSync(join(tmpdir(), 'office-mcp-catalog-'));
  const reviewPath = join(catalogPath, 'catalog-identity-review.json');
  mkdirSync(join(catalogPath, 'word'), { recursive: true });
  mkdirSync(join(catalogPath, 'excel'), { recursive: true });
  mkdirSync(join(catalogPath, 'powerpoint'), { recursive: true });
  writeFileSync(join(catalogPath, 'word', 'manifest.xml'), '<legacy />');
  writeFileSync(join(catalogPath, 'excel', 'manifest.xml'), '<legacy />');
  writeFileSync(join(catalogPath, 'powerpoint', 'manifest.xml'), '<legacy />');
  writeFileSync(join(catalogPath, 'office-mcp-word.xml'), '<ProviderName>office-mcp</ProviderName><DisplayName DefaultValue="office-mcp" />');
  writeFileSync(join(catalogPath, 'office-mcp-excel.xml'), '<ProviderName>office-mcp</ProviderName><DisplayName DefaultValue="office-mcp for Excel" />');

  try {
    const result = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        join(ADDIN_ROOT, '..', 'common', 'scripts', 'register-office-catalog.ps1'),
        '-CatalogPath',
        catalogPath,
        '-BaseUrl',
        'https://localhost:8766',
        '-SkipRegistry'
      ],
      { cwd: join(ADDIN_ROOT, '..', '..', '..'), encoding: 'utf8' }
    );

    assert.equal(result.status, 0, result.stderr || result.stdout);
    assert.match(result.stdout, /Registered Office trusted catalog:/);
    assert.match(result.stdout, /Catalog URL: \\\\localhost\\[A-Z]\$\\/);
    assert.match(result.stdout, /Manifest origin: https:\/\/localhost:8766/);
    assert.match(result.stdout, /Word manifest:/);
    assert.match(result.stdout, /Excel manifest:/);
    assert.match(result.stdout, /PowerPoint manifest:/);
    const wordManifest = readFileSync(join(catalogPath, 'office-mcp-word.xml'), 'utf8');
    const excelManifest = readFileSync(join(catalogPath, 'office-mcp-excel.xml'), 'utf8');
    const powerpointManifest = readFileSync(join(catalogPath, 'office-mcp-powerpoint.xml'), 'utf8');
    assert.match(wordManifest, /<OfficeApp/);
    assert.match(wordManifest, /<ProviderName>Office MCP Control<\/ProviderName>/);
    assert.match(wordManifest, /<DisplayName DefaultValue="Office MCP Control" \/>/);
    assert.match(wordManifest, /local productivity automation control utility/);
    assert.match(wordManifest, /<SupportUrl DefaultValue="https:\/\/github\.com\/office-mcp\/office-mcp" \/>/);
    assert.match(wordManifest, /DefaultValue="Open Control Panel"/);
    assert.match(wordManifest, /Open the control panel to connect this Word document/);
    assert.match(wordManifest, /https:\/\/localhost:8766\/word\/taskpane\.html\?v=0\.1\.18/);
    assert.match(wordManifest, /<IconUrl DefaultValue="https:\/\/localhost:8766\/assets\/icon-32\.png" \/>/);
    assert.match(wordManifest, /<HighResolutionIconUrl DefaultValue="https:\/\/localhost:8766\/assets\/icon-80\.png" \/>/);
    assert.match(wordManifest, /https:\/\/localhost:8766\/assets\/icon-32\.png/);
    assert.match(wordManifest, /https:\/\/localhost:8766\/assets\/icon-80\.png/);
    assert.doesNotMatch(wordManifest, /<ProviderName>office-mcp<\/ProviderName>|DefaultValue="office-mcp"|DefaultValue="Open"|Open the office-mcp task pane/);
    assert.match(excelManifest, /<OfficeApp/);
    assert.match(excelManifest, /<ProviderName>Office MCP Control<\/ProviderName>/);
    assert.match(excelManifest, /<DisplayName DefaultValue="Office MCP Control" \/>/);
    assert.match(excelManifest, /local productivity automation control utility/);
    assert.match(excelManifest, /<SupportUrl DefaultValue="https:\/\/github\.com\/office-mcp\/office-mcp" \/>/);
    assert.match(excelManifest, /DefaultValue="Open Control Panel"/);
    assert.match(excelManifest, /Open the control panel to connect this Excel workbook/);
    assert.match(excelManifest, /https:\/\/localhost:8766\/excel\/taskpane\.html\?v=0\.1\.12/);
    assert.match(excelManifest, /<IconUrl DefaultValue="https:\/\/localhost:8766\/assets\/icon-32\.png" \/>/);
    assert.match(excelManifest, /<HighResolutionIconUrl DefaultValue="https:\/\/localhost:8766\/assets\/icon-80\.png" \/>/);
    assert.match(excelManifest, /https:\/\/localhost:8766\/assets\/icon-32\.png/);
    assert.match(excelManifest, /https:\/\/localhost:8766\/assets\/icon-80\.png/);
    assert.doesNotMatch(excelManifest, /<ProviderName>office-mcp<\/ProviderName>|DefaultValue="office-mcp for Excel"|DefaultValue="Open"|Open the office-mcp task pane/);
    assert.match(powerpointManifest, /<OfficeApp/);
    assert.match(powerpointManifest, /<ProviderName>Office MCP Control<\/ProviderName>/);
    assert.match(powerpointManifest, /<DisplayName DefaultValue="Office MCP Control" \/>/);
    assert.match(powerpointManifest, /local productivity automation control utility/);
    assert.match(powerpointManifest, /<SupportUrl DefaultValue="https:\/\/github\.com\/office-mcp\/office-mcp" \/>/);
    assert.match(powerpointManifest, /DefaultValue="Open Control Panel"/);
    assert.match(powerpointManifest, /Open the control panel to connect this PowerPoint presentation/);
    assert.match(powerpointManifest, /https:\/\/localhost:8766\/powerpoint\/taskpane\.html\?v=0\.1\.4/);
    assert.match(powerpointManifest, /<IconUrl DefaultValue="https:\/\/localhost:8766\/assets\/icon-32\.png" \/>/);
    assert.match(powerpointManifest, /<HighResolutionIconUrl DefaultValue="https:\/\/localhost:8766\/assets\/icon-80\.png" \/>/);
    assert.match(powerpointManifest, /https:\/\/localhost:8766\/assets\/icon-32\.png/);
    assert.match(powerpointManifest, /https:\/\/localhost:8766\/assets\/icon-80\.png/);
    assert.doesNotMatch(wordManifest, /https:\/\/localhost:8765/);
    assert.doesNotMatch(excelManifest, /https:\/\/localhost:8765/);
    assert.doesNotMatch(powerpointManifest, /https:\/\/localhost:8765/);
    assert.throws(() => readFileSync(join(catalogPath, 'word', 'manifest.xml'), 'utf8'));
    assert.throws(() => readFileSync(join(catalogPath, 'excel', 'manifest.xml'), 'utf8'));
    assert.throws(() => readFileSync(join(catalogPath, 'powerpoint', 'manifest.xml'), 'utf8'));

    const review = spawnSync(
      process.execPath,
      [
        join(ADDIN_ROOT, '..', 'common', 'scripts', 'record-catalog-identity-review.mjs'),
        '--catalog-path',
        catalogPath,
        '--output',
        reviewPath
      ],
      { cwd: join(ADDIN_ROOT, '..', '..', '..'), encoding: 'utf8' }
    );
    assert.equal(review.status, 0, review.stderr || review.stdout);
    const report = JSON.parse(readFileSync(reviewPath, 'utf8'));
    assert.equal(report.kind, 'catalog_identity_review');
    assert.equal(report.product_name, 'Office MCP Control');
    assert.equal(report.catalog_type, 'Local productivity automation control utility');
    assert.equal(report.shared_origin, 'https://localhost:8766');
    assert.equal(report.ready, true);
    assert.deepEqual(report.failures, []);
    for (const host of ['word', 'excel', 'powerpoint']) {
      assert.equal(report.hosts[host].display_name, 'Office MCP Control');
      assert.equal(report.hosts[host].provider, 'Office MCP Control');
      assert.equal(report.hosts[host].group_label, 'Office MCP Control');
      assert.equal(report.hosts[host].command_label, 'Open Control Panel');
      assert.equal(report.hosts[host].ready, true);
      assert.match(report.hosts[host].description, /local productivity automation control utility/i);
      assert.match(report.hosts[host].icon_url, /https:\/\/localhost:8766\/assets\/icon-32\.png/);
      assert.match(report.hosts[host].high_resolution_icon_url, /https:\/\/localhost:8766\/assets\/icon-80\.png/);
    }
  } finally {
    rmSync(catalogPath, { force: true, recursive: true });
  }
});

test('Office catalog identity review rejects prototype first-impression metadata', () => {
  const catalogPath = mkdtempSync(join(tmpdir(), 'office-mcp-catalog-review-bad-'));
  const reviewPath = join(catalogPath, 'catalog-identity-review.json');
  try {
    for (const [host, taskpane] of [
      ['word', '/word/taskpane.html'],
      ['excel', '/excel/taskpane.html'],
      ['powerpoint', '/powerpoint/taskpane.html']
    ]) {
      writeFileSync(join(catalogPath, `office-mcp-${host}.xml`), `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp>
  <ProviderName>${host === 'word' ? 'office-mcp' : 'Office MCP Control'}</ProviderName>
  <DisplayName DefaultValue="${host === 'word' ? 'office-mcp-word' : 'Office MCP Control'}" />
  <Description DefaultValue="${host === 'excel' ? 'Experimental protocol bridge debug panel.' : 'Control live documents through a local productivity automation control utility.'}" />
  <IconUrl DefaultValue="https://localhost:8766/assets/${host === 'powerpoint' ? 'blank.png' : 'icon-32.png'}" />
  <HighResolutionIconUrl DefaultValue="https://localhost:8766/assets/icon-80.png" />
  <bt:String id="OfficeMcp.GroupLabel" DefaultValue="${host === 'excel' ? 'Office MCP' : 'Office MCP Control'}" />
  <bt:String id="OfficeMcp.OpenPane.Label" DefaultValue="${host === 'word' ? 'Open' : 'Open Control Panel'}" />
  <bt:String id="OfficeMcp.OpenPane.Tooltip" DefaultValue="Open Office MCP Control." />
  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:8766${taskpane}?v=0.1.0" />
</OfficeApp>`);
    }

    const review = spawnSync(
      process.execPath,
      [
        join(ADDIN_ROOT, '..', 'common', 'scripts', 'record-catalog-identity-review.mjs'),
        '--catalog-path',
        catalogPath,
        '--output',
        reviewPath,
        '--catalog-type',
        'Task Pane Add-in protocol bridge'
      ],
      { cwd: join(ADDIN_ROOT, '..', '..', '..'), encoding: 'utf8' }
    );
    assert.notEqual(review.status, 0);
    const report = JSON.parse(readFileSync(reviewPath, 'utf8'));
    assert.equal(report.ready, false);
    assert.match(report.failures.join('\n'), /Word: display name must be Office MCP Control/);
    assert.match(report.failures.join('\n'), /Word: provider must be Office MCP Control/);
    assert.match(report.failures.join('\n'), /Excel: description must describe a local productivity automation control utility/);
    assert.match(report.failures.join('\n'), /Excel: ribbon group label must be Office MCP Control/);
    assert.match(report.failures.join('\n'), /PowerPoint: catalog\/ribbon icon URL must use generated icon-32\.png/);
    assert.match(report.failures.join('\n'), /Catalog type is not product-ready/);
  } finally {
    rmSync(catalogPath, { force: true, recursive: true });
  }
});

test('Office catalog registration can sync its origin from running daemon status', () => {
  const catalogPath = mkdtempSync(join(tmpdir(), 'office-mcp-catalog-status-'));
  const statusCommand = join(catalogPath, 'daemon-status.cmd');
  writeFileSync(statusCommand, '@echo off\r\necho {"running":true,"uiUrl":"https://localhost:8777/ui/"}\r\n');

  try {
    const result = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        join(ADDIN_ROOT, '..', 'common', 'scripts', 'register-office-catalog.ps1'),
        '-CatalogPath',
        catalogPath,
        '-DaemonStatusCommand',
        statusCommand,
        '-SkipRegistry'
      ],
      { cwd: join(ADDIN_ROOT, '..', '..', '..'), encoding: 'utf8' }
    );

    assert.equal(result.status, 0, result.stderr || result.stdout);
    assert.match(result.stdout, /Manifest origin: https:\/\/localhost:8777/);
    const wordManifest = readFileSync(join(catalogPath, 'office-mcp-word.xml'), 'utf8');
    const excelManifest = readFileSync(join(catalogPath, 'office-mcp-excel.xml'), 'utf8');
    const powerpointManifest = readFileSync(join(catalogPath, 'office-mcp-powerpoint.xml'), 'utf8');
    assert.match(wordManifest, /https:\/\/localhost:8777\/word\/taskpane\.html\?v=0\.1\.18/);
    assert.match(wordManifest, /https:\/\/localhost:8777\/assets\/icon-32\.png/);
    assert.match(wordManifest, /https:\/\/localhost:8777\/assets\/icon-80\.png/);
    assert.match(excelManifest, /https:\/\/localhost:8777\/excel\/taskpane\.html\?v=0\.1\.12/);
    assert.match(excelManifest, /https:\/\/localhost:8777\/assets\/icon-32\.png/);
    assert.match(excelManifest, /https:\/\/localhost:8777\/assets\/icon-80\.png/);
    assert.match(powerpointManifest, /https:\/\/localhost:8777\/powerpoint\/taskpane\.html\?v=0\.1\.4/);
    assert.match(powerpointManifest, /https:\/\/localhost:8777\/assets\/icon-32\.png/);
    assert.match(powerpointManifest, /https:\/\/localhost:8777\/assets\/icon-80\.png/);
    assert.doesNotMatch(wordManifest, /https:\/\/localhost:8765/);
    assert.doesNotMatch(excelManifest, /https:\/\/localhost:8765/);
    assert.doesNotMatch(powerpointManifest, /https:\/\/localhost:8765/);
  } finally {
    rmSync(catalogPath, { force: true, recursive: true });
  }
});

test('Office catalog registration writes a shared folder URL to the trusted catalog registry', () => {
  const catalogPath = mkdtempSync(join(tmpdir(), 'office-mcp-catalog-registry-'));
  const registryRoot = 'HKCU:\\Software\\office-mcp-tests\\TrustedCatalogs';
  const catalogId = '{01234567-89AB-CDEF-0123-456789ABCDEF}';
  const registryKey = `${registryRoot}\\${catalogId}`;
  const legacyRegistryKey = `${registryRoot}\\office-mcp`;

  try {
    const result = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-Command',
        [
          `New-Item '${legacyRegistryKey}' -Force | Out-Null`,
          `& '${join(ADDIN_ROOT, '..', 'common', 'scripts', 'register-office-catalog.ps1')}' -CatalogPath '${catalogPath}' -BaseUrl https://localhost:8778 -TrustedCatalogId '${catalogId}' -TrustedCatalogRegistryRoot '${registryRoot}' -SkipOfficeCache`,
          `$entry = Get-ItemProperty '${registryKey}'`,
          `$legacyExists = Test-Path '${legacyRegistryKey}'`,
          'Write-Output "REGISTRY_ID=$($entry.Id)"',
          'Write-Output "REGISTRY_URL=$($entry.Url)"',
          'Write-Output "LEGACY_EXISTS=$legacyExists"',
          `Remove-Item '${registryRoot}' -Recurse -Force`
        ].join('; ')
      ],
      { cwd: join(ADDIN_ROOT, '..', '..', '..'), encoding: 'utf8' }
    );

    assert.equal(result.status, 0, result.stderr || result.stdout);
    assert.match(result.stdout, /REGISTRY_ID=\{01234567-89AB-CDEF-0123-456789ABCDEF\}/);
    assert.match(result.stdout, /REGISTRY_URL=\\\\localhost\\[A-Z]\$\\/);
    assert.match(result.stdout, new RegExp(catalogPath.split(/[\\/]/).at(-1)));
    assert.match(result.stdout, /LEGACY_EXISTS=False/);
    assert.doesNotMatch(result.stdout, /REGISTRY_URL=[A-Z]:\\/i);
  } finally {
    rmSync(catalogPath, { force: true, recursive: true });
  }
});

test('word.insert_image handles paragraph anchors through clean image paragraphs', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const insertImageBody = functionBody(js, 'insertImage');

  assert.match(insertImageBody, /insertInlinePictureWithPlacement\(target, args\.anchor, base64, args\.placement\)/);
  assert.match(js, /function insertInlinePictureWithPlacement\(target, anchor, base64, placement\)/);
  assert.match(js, /function validateInsertImagePlacement\(anchor, placement\)/);
  assert.match(js, /const INSERT_IMAGE_PLACEMENTS = new Set\(\[/);
  assert.match(functionBody(js, 'validateInsertImagePlacement'), /if \(!placement\) return/);
  assert.match(functionBody(js, 'validateInsertImagePlacement'), /placement === 'selection' && anchor\.kind !== 'selection'/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /case 'new_paragraph_before':/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /case 'new_paragraph_after':/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /case 'replace_paragraph':/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /case 'before_paragraph':/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /case 'after_paragraph':/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /target\.insertParagraph\('', Word\.InsertLocation\.before\)/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /target\.insertParagraph\('', Word\.InsertLocation\.after\)/);
  assert.match(functionBody(js, 'insertInlinePictureWithPlacement'), /target\.getRange\(\)\.insertInlinePictureFromBase64\(base64, Word\.InsertLocation\.replace\)/);
  assert.match(functionBody(js, 'validateInsertImagePlacement'), /Unsupported word\.insert_image placement/);
  assert.match(functionBody(js, 'validateInsertImagePlacement'), /requires a paragraph-resolving anchor/);
  assert.match(functionBody(js, 'isParagraphAnchor'), /paragraph_index/);
  assert.match(functionBody(js, 'isParagraphAnchor'), /before_paragraph_index/);
  assert.match(functionBody(js, 'isParagraphAnchor'), /after_paragraph_index/);
  assert.match(functionBody(js, 'isParagraphAnchor'), /heading/);
});

test('Word mutating tools run preflight validation before Office mutation dispatch', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');

  assert.match(js, /const WORD_MUTATING_TOOLS = new Set\(\[/);
  assert.match(invokeBody, /preflightWordMutatingTool\(tool, args \|\| \{\}\);[\s\S]*switch \(tool\)/);
  assert.match(preflightBody, /case 'word\.insert_image':/);
  assert.match(preflightBody, /validateInsertImagePreflight\(args\)/);
  assert.match(preflightBody, /case 'word\.insert_table':/);
  assert.match(preflightBody, /requirePositiveInteger\(tool, 'rows', args\.rows\)/);
  assert.match(preflightBody, /case 'word\.insert_list':/);
  assert.match(preflightBody, /validateInsertListArgs\(args\)/);
  assert.match(preflightBody, /case 'word\.update_list':/);
  assert.match(preflightBody, /validateUpdateListArgs\(args\)/);
  assert.match(preflightBody, /case 'word\.update_header_footer':/);
  assert.match(preflightBody, /validateHeaderFooterArgs\(tool, args, true\)/);
  assert.match(preflightBody, /case 'word\.insert_hyperlink':/);
  assert.match(preflightBody, /validateHyperlinkArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.remove_hyperlink':/);
  assert.match(preflightBody, /validateRemoveHyperlinkArgs\(args\)/);
  assert.match(preflightBody, /case 'word\.replace_text':/);
  assert.match(preflightBody, /validateReplaceTextArgs\(args\)/);
  assert.match(preflightBody, /case 'word\.delete_range':/);
  assert.match(preflightBody, /validateExtentToolArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.update_table':/);
  assert.match(preflightBody, /validateUpdateTableArgs\(args\)/);
  assert.match(preflightBody, /case 'word\.update_tracked_change':/);
  assert.match(preflightBody, /validateTrackedChangeArgs\(args\)/);
  assert.match(preflightBody, /case 'word\.insert_note':/);
  assert.match(preflightBody, /validateInsertNoteArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.update_note':/);
  assert.match(preflightBody, /validateUpdateNoteArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.delete_note':/);
  assert.match(preflightBody, /validateDeleteNoteArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.insert_field':/);
  assert.match(preflightBody, /validateInsertFieldArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.update_field':/);
  assert.match(preflightBody, /validateUpdateFieldArgs\(tool, args\)/);
  assert.match(preflightBody, /case 'word\.delete_field':/);
  assert.match(preflightBody, /validateDeleteFieldArgs\(tool, args\)/);
});

test('Word mutating preflight helpers return specific no-effect validation errors', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(functionBody(js, 'invalidArgument'), /officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none'/);
  assert.match(functionBody(js, 'validateInsertImagePreflight'), /word\.insert_image requires image\.base64/);
  assert.match(functionBody(js, 'validateInsertImagePreflight'), /validateOptionalPositiveNumber\('word\.insert_image', 'width_pt', args\.width_pt\)/);
  assert.match(functionBody(js, 'validateInsertImagePreflight'), /validateOptionalPositiveNumber\('word\.insert_image', 'height_pt', args\.height_pt\)/);
  assert.match(functionBody(js, 'validateInsertImagePlacement'), /invalidArgumentWithSuggestion/);
  assert.match(functionBody(js, 'mapError'), /mapped\.suggestion = error\.suggestion/);
  assert.match(functionBody(js, 'validateInsertListArgs'), /word\.insert_list requires a non-empty items array/);
  assert.match(functionBody(js, 'validateInsertListArgs'), /word\.insert_list kind must be bulleted or numbered/);
  assert.match(functionBody(js, 'validateUpdateListArgs'), /Unsupported list action/);
  assert.match(functionBody(js, 'validateUpdateListArgs'), /word\.update_list add_item requires list_id/);
  assert.match(functionBody(js, 'validateUpdateListArgs'), /word\.update_list detach_paragraph requires paragraph_index/);
  assert.match(functionBody(js, 'validateHeaderFooterArgs'), /set_text[\s\S]*requires text/);
  assert.match(functionBody(js, 'normalizedHeaderFooterLocation'), /header\/footer location must be header or footer/);
  assert.match(functionBody(js, 'normalizedHeaderFooterAction'), /set_text, append_paragraph, or clear/);
  assert.match(functionBody(js, 'validateHyperlinkArgs'), /word\.insert_hyperlink requires a non-empty url/);
  assert.match(functionBody(js, 'validateHyperlinkUrl'), /https[\s\S]*http[\s\S]*mailto/);
  assert.match(functionBody(js, 'validateHyperlinkUrl'), /file: and javascript: URLs are not allowed/);
  assert.match(functionBody(js, 'validateRemoveHyperlinkArgs'), /keep_text must be a boolean/);
  assert.match(functionBody(js, 'validateReplaceTextArgs'), /word\.replace_text requires non-empty find text/);
  assert.match(functionBody(js, 'validateReplaceTextArgs'), /scope\.paragraph_range must be \[start, end\]/);
  assert.match(functionBody(js, 'validateExtentToolArgs'), /extent must be paragraph, sentence, or selection/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /Unsupported table action/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /delete_row/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /delete_column/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /merge_cells/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /set_column_width/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /distribute_columns/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /set_borders/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /set_header_row/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /validateTableIndexRange\('word\.update_table', 'row_range', args\.row_range\)/);
  assert.match(functionBody(js, 'validateUpdateTableArgs'), /validateTableBordersArg\(args\.borders\)/);
  assert.match(functionBody(js, 'validateContentControlTargetArgs'), /requires content_control_id, tag, or title/);
  assert.match(functionBody(js, 'validateDeleteContentControlMode'), /mode must be keep_content or delete_content/);
  assert.match(functionBody(js, 'validateNoteKind'), /kind must be footnote or endnote/);
  assert.match(functionBody(js, 'validateInsertNoteArgs'), /requires non-empty text/);
  assert.match(functionBody(js, 'validateUpdateNoteArgs'), /requires a non-negative integer index/);
  assert.match(functionBody(js, 'validateDeleteNoteArgs'), /requires a non-negative integer index/);
  assert.match(functionBody(js, 'validateChangeTrackingMode'), /mode must be off, track_all, or track_mine_only/);
  assert.match(functionBody(js, 'validateTrackedChangeArgs'), /requires expected_fingerprint/);
  assert.match(functionBody(js, 'validateTrackedChangeArgs'), /requires expected_count/);
});

test('Word change-tracking tools are advertised, grouped, gated, and dispatched through review owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');

  assert.match(js, /'word\.set_change_tracking'/);
  assert.match(js, /'word\.update_tracked_change'/);
  assert.match(js, /\{ label: 'Review', tools: \['word\.add_comment', 'word\.resolve_comment', 'word\.update_comment', 'word\.set_change_tracking', 'word\.update_tracked_change'\] \}/);
  assert.match(js, /\['word\.set_change_tracking', \{ category: 'Review', sideEffect: 'mutating', description: 'Set Track Changes mode\.' \}\]/);
  assert.match(js, /\['word\.update_tracked_change', \{ category: 'Review', sideEffect: 'destructive', description: 'Accept, reject, or bulk-finalize tracked changes\.' \}\]/);
  assert.match(invokeBody, /case 'word\.set_change_tracking':\s*data = await setChangeTracking\(args\);/);
  assert.match(invokeBody, /case 'word\.update_tracked_change':\s*data = await updateTrackedChange\(args\);/);
  assert.match(preflightBody, /case 'word\.set_change_tracking':\s*validateChangeTrackingMode\(args\.mode\);/);
  assert.match(preflightBody, /case 'word\.update_tracked_change':\s*validateTrackedChangeArgs\(args\);/);
  assert.match(js, /async function setChangeTracking\(args\)/);
  assert.match(js, /async function mutateAllTrackedChanges\(args, action\)/);
  assert.match(functionBody(js, 'updateTrackedChange'), /accept_all/);
  assert.match(functionBody(js, 'updateTrackedChange'), /reject_all/);
  assert.match(functionBody(js, 'mutateAllTrackedChanges'), /expected_count/);
  assert.match(functionBody(js, 'mutateAllTrackedChanges'), /acceptAll\(\)|rejectAll\(\)/);
  assert.match(functionBody(js, 'changeTrackingModeFromArg'), /Word\.ChangeTrackingMode\.trackAll/);
  assert.match(functionBody(js, 'changeTrackingModeToResult'), /track_mine_only/);
});

test('Word validation-only mode validates required mutating tools without writes', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const validateOnlyBody = functionBody(js, 'validateWordMutationOnly');

  assert.match(invokeBody, /if \(args\?\.validate_only\) data = await validateWordMutationOnly\(tool, args \|\| \{\}\)/);
  for (const tool of ['word.insert_image', 'word.update_image', 'word.insert_hyperlink', 'word.insert_note', 'word.insert_field', 'word.replace_text', 'word.update_paragraph', 'word.update_note', 'word.update_field', 'word.delete_range', 'word.delete_note', 'word.delete_field']) {
    assert.match(validateOnlyBody, new RegExp(`case '${tool.replace('.', '\\.')}'`));
  }
  assert.match(validateOnlyBody, /partial_effect: 'none'/);
  assert.match(validateOnlyBody, /valid: true/);
  assert.doesNotMatch(validateOnlyBody, /insertText\(/);
  assert.doesNotMatch(validateOnlyBody, /delete\(/);
  assert.doesNotMatch(validateOnlyBody, /\.hyperlink\s*=/);
  assert.doesNotMatch(validateOnlyBody, /insertInlinePictureFromBase64\(/);
  assert.doesNotMatch(validateOnlyBody, /insertFootnote\(/);
  assert.doesNotMatch(validateOnlyBody, /insertEndnote\(/);
  assert.doesNotMatch(validateOnlyBody, /insertField\(/);
});

test('Word image CRUD tools are advertised, grouped, gated, and dispatched through media owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');

  assert.match(js, /'word\.list_images'/);
  assert.match(js, /'word\.get_image'/);
  assert.match(js, /'word\.update_image'/);
  const availableToolsSource = js.match(/const AVAILABLE_TOOLS = \[([\s\S]*?)\];/)?.[1] || '';
  assert.doesNotMatch(availableToolsSource, /'word\.get_paragraph'/);
  assert.doesNotMatch(availableToolsSource, /'word\.resize_image'/);
  assert.doesNotMatch(availableToolsSource, /'word\.delete_image'/);
  assert.match(js, /'word\.list_shapes'/);
  assert.match(js, /'word\.insert_shape'/);
  assert.match(js, /'word\.update_shape'/);
  assert.match(js, /'word\.delete_shape'/);
  assert.match(js, /\{ label: 'Media', tools: \['word\.insert_image', 'word\.list_images', 'word\.get_image', 'word\.update_image', 'word\.list_shapes', 'word\.insert_shape', 'word\.update_shape', 'word\.delete_shape'\] \}/);
  assert.match(js, /\['word\.list_images', \{ category: 'Media', sideEffect: 'read', description: 'List inline images\.' \}\]/);
  assert.match(js, /\['word\.get_image', \{ category: 'Media', sideEffect: 'read', description: 'Export an inline image with metadata\.' \}\]/);
  assert.match(js, /\['word\.update_image', \{ category: 'Media', sideEffect: 'destructive', description: 'Resize, update, replace, or delete an inline image\.' \}\]/);
  assert.match(js, /\['word\.list_shapes', \{ category: 'Media', sideEffect: 'read', description: 'List desktop Word shapes and text boxes\.' \}\]/);
  assert.match(js, /\['word\.insert_shape', \{ category: 'Media', sideEffect: 'mutating', description: 'Insert a desktop Word shape or text box\.' \}\]/);
  assert.match(js, /\['word\.update_shape', \{ category: 'Media', sideEffect: 'mutating', description: 'Update desktop Word shape text, geometry, or visual settings\.' \}\]/);
  assert.match(js, /\['word\.delete_shape', \{ category: 'Media', sideEffect: 'destructive', description: 'Delete a desktop Word shape\.' \}\]/);
  assert.match(invokeBody, /case 'word\.list_images':\s*data = await listImages\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.get_image':\s*data = await getImage\(args\);/);
  assert.match(invokeBody, /case 'word\.update_image':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateImage\(args\);/);
  assert.match(invokeBody, /case 'word\.list_shapes':\s*data = await listShapes\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.insert_shape':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertShape\(args\);/);
  assert.match(invokeBody, /case 'word\.update_shape':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateShape\(args\);/);
  assert.match(invokeBody, /case 'word\.delete_shape':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await deleteShape\(args\);/);
  assert.match(preflightBody, /case 'word\.update_image':\s*validateUpdateImageArgs\(args\);/);
  assert.match(preflightBody, /case 'word\.insert_shape':\s*validateInsertShapeArgs\(tool, args\);/);
  assert.match(preflightBody, /case 'word\.update_shape':\s*validateUpdateShapeArgs\(tool, args\);/);
  assert.match(preflightBody, /case 'word\.delete_shape':\s*validateShapeId\(tool, args\?\.shape_id\);/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_3[\s\S]*word\.list_images[\s\S]*word\.update_image/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApiDesktop_1_2[\s\S]*word\.list_shapes[\s\S]*word\.delete_shape/);
  assert.match(js, /async function listImages\(args/);
  assert.match(js, /async function getImage\(args/);
  assert.match(js, /async function updateImage\(args/);
  assert.doesNotMatch(js, /async function deleteImage\(args/);
  assert.match(js, /async function listShapes\(args/);
  assert.match(js, /async function insertShape\(args/);
  assert.match(js, /async function updateShape\(args/);
  assert.match(js, /async function deleteShape\(args/);
  assert.match(js, /function validateImageLocator\(tool, image\)/);
  assert.match(js, /function validateInsertShapeArgs\(tool, args\)/);
  assert.match(js, /function validateUpdateShapeArgs\(tool, args\)/);
  assert.match(js, /function validateShapeId\(tool, shapeId\)/);
  assert.match(functionBody(js, 'validateUpdateImageArgs'), /requires action/);
  assert.match(functionBody(js, 'validateUpdateImageArgs'), /resize/);
  assert.match(functionBody(js, 'validateUpdateImageArgs'), /set_alt_text/);
  assert.match(functionBody(js, 'validateUpdateImageArgs'), /set_hyperlink/);
  assert.match(functionBody(js, 'validateUpdateImageArgs'), /replace/);
  assert.match(functionBody(js, 'validateUpdateImageArgs'), /delete/);
  assert.match(functionBody(js, 'updateImage'), /case 'resize'/);
  assert.match(functionBody(js, 'updateImage'), /case 'replace'/);
  assert.match(functionBody(js, 'updateImage'), /insertInlinePictureFromBase64\(args\.base64, Word\.InsertLocation\.replace\)/);
  assert.match(functionBody(js, 'updateImage'), /case 'delete'/);
  assert.match(functionBody(js, 'updateImage'), /picture\.delete\(\)/);
  assert.match(functionBody(js, 'insertShape'), /insertTextBox\(String\(args\.text \?\? ''\), shapeInsertOptions\(args\)\)/);
  assert.match(functionBody(js, 'insertShape'), /insertGeometricShape\(wordGeometricShapeType\(args\.shape_type\), shapeInsertOptions\(args\)\)/);
  assert.match(functionBody(js, 'insertShape'), /insertPictureFromBase64\(args\.image\.base64, shapeInsertOptions\(args\)\)/);
  assert.match(functionBody(js, 'updateShape'), /shape\.body\.insertText\(String\(args\.text\), Word\.InsertLocation\.replace\)/);
  assert.match(functionBody(js, 'deleteShape'), /shape\.delete\(\)/);
  assert.match(js, /function shapeMetadata\(shape\)/);
  assert.match(functionBody(js, 'shapeMetadata'), /text_preview: safeTextPreview\(/);
  assert.match(js, /function requireWordApiDesktop12\(tool\)/);
});

test('Word style tools are advertised, grouped, gated, and dispatched through document owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');

  assert.match(js, /'word\.list_styles'/);
  assert.match(js, /'word\.create_style'/);
  assert.match(js, /'word\.update_style'/);
  assert.match(js, /\{ label: 'Document & structure', tools: \[[\s\S]*'word\.list_styles'[\s\S]*'word\.create_style'[\s\S]*'word\.update_style'[\s\S]*\] \}/);
  assert.match(js, /\['word\.list_styles', \{ category: 'Document & structure', sideEffect: 'read', description: 'List built-in and custom document styles\.' \}\]/);
  assert.match(js, /\['word\.create_style', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Create a document style definition\.' \}\]/);
  assert.match(js, /\['word\.update_style', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Update a document style definition\.' \}\]/);
  assert.match(invokeBody, /case 'word\.list_styles':\s*data = await listStyles\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.create_style':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await createStyle\(args\);/);
  assert.match(invokeBody, /case 'word\.update_style':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateStyle\(args\);/);
  assert.match(preflightBody, /case 'word\.create_style':\s*validateCreateStyleArgs\(tool, args\);/);
  assert.match(preflightBody, /case 'word\.update_style':\s*validateUpdateStyleArgs\(tool, args\);/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_5[\s\S]*word\.list_styles/);
  assert.match(js, /async function listStyles\(args/);
  assert.match(js, /async function createStyle\(args/);
  assert.match(js, /async function updateStyle\(args/);
  assert.match(js, /function styleTypeToOffice\(type\)/);
  assert.match(js, /function applyStyleDefinitionFormatting\(style, args\)/);
});

test('Word document property tools are advertised, grouped, gated, and dispatched through document owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');

  assert.match(js, /'word\.get_document_properties'/);
  assert.match(js, /'word\.update_document_properties'/);
  assert.match(js, /\{ label: 'Document & structure', tools: \[[\s\S]*'word\.get_document_properties'[\s\S]*'word\.update_document_properties'[\s\S]*\] \}/);
  assert.match(js, /\['word\.get_document_properties', \{ category: 'Document & structure', sideEffect: 'read', description: 'Read document metadata and custom properties\.' \}\]/);
  assert.match(js, /\['word\.update_document_properties', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Update document metadata and custom properties\.' \}\]/);
  assert.match(invokeBody, /case 'word\.get_document_properties':\s*data = await getDocumentProperties\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.update_document_properties':\s*data = await updateDocumentProperties\(args \|\| \{\}\);/);
  assert.match(preflightBody, /case 'word\.update_document_properties':\s*validateUpdateDocumentPropertiesArgs\(tool, args\);/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_3[\s\S]*word\.get_document_properties[\s\S]*word\.update_document_properties/);
  assert.match(js, /async function getDocumentProperties\(args/);
  assert.match(js, /async function updateDocumentProperties\(args/);
  assert.match(js, /function validateUpdateDocumentPropertiesArgs\(tool, args\)/);
  assert.match(functionBody(js, 'validateUpdateDocumentPropertiesArgs'), /requires at least one writable property or custom operation/);
  assert.match(functionBody(js, 'updateDocumentProperties'), /custom_set/);
  assert.match(functionBody(js, 'updateDocumentProperties'), /custom_delete/);
});

test('Word field tools are advertised, grouped, gated, and dispatched through document owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');

  assert.match(js, /'word\.list_fields'/);
  assert.match(js, /'word\.insert_field'/);
  assert.match(js, /'word\.update_field'/);
  assert.match(js, /'word\.delete_field'/);
  assert.match(js, /\{ label: 'Document & structure', tools: \[[\s\S]*'word\.list_fields'[\s\S]*'word\.insert_field'[\s\S]*'word\.update_field'[\s\S]*'word\.delete_field'[\s\S]*\] \}/);
  assert.match(js, /\['word\.list_fields', \{ category: 'Document & structure', sideEffect: 'read', description: 'List document fields with bounded previews\.' \}\]/);
  assert.match(js, /\['word\.insert_field', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Insert a curated Word field at an anchored range\.' \}\]/);
  assert.match(js, /\['word\.update_field', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Refresh, lock, or unlock Word fields\.' \}\]/);
  assert.match(js, /\['word\.delete_field', \{ category: 'Document & structure', sideEffect: 'destructive', description: 'Delete a Word field by current index\.' \}\]/);
  assert.match(invokeBody, /case 'word\.list_fields':\s*data = await listFields\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.insert_field':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertField\(args\);/);
  assert.match(invokeBody, /case 'word\.update_field':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateField\(args\);/);
  assert.match(invokeBody, /case 'word\.delete_field':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await deleteField\(args\);/);
  assert.match(preflightBody, /case 'word\.insert_field':\s*validateInsertFieldArgs\(tool, args\);/);
  assert.match(preflightBody, /case 'word\.update_field':\s*validateUpdateFieldArgs\(tool, args\);/);
  assert.match(preflightBody, /case 'word\.delete_field':\s*validateDeleteFieldArgs\(tool, args\);/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_4[\s\S]*word\.list_fields/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_5[\s\S]*word\.insert_field/);
  assert.match(js, /async function listFields\(args/);
  assert.match(js, /async function insertField\(args/);
  assert.match(js, /async function updateField\(args/);
  assert.match(js, /async function deleteField\(args/);
  assert.match(functionBody(js, 'fieldTypeToOffice'), /toc[\s\S]*page[\s\S]*num_pages[\s\S]*styleref/);
  assert.match(functionBody(js, 'validateFieldType'), /INCLUDETEXT|IMPORT/);
});

test('Word note tools are advertised, grouped, gated, and dispatched through note owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');

  assert.match(js, /'word\.insert_note'/);
  assert.match(js, /'word\.list_notes'/);
  assert.match(js, /'word\.update_note'/);
  assert.match(js, /'word\.delete_note'/);
  assert.match(js, /\{ label: 'Notes', tools: \['word\.insert_note', 'word\.list_notes', 'word\.update_note', 'word\.delete_note'\] \}/);
  assert.match(js, /\['word\.insert_note', \{ category: 'Notes', sideEffect: 'mutating', description: 'Insert a footnote or endnote at an anchored range\.' \}\]/);
  assert.match(js, /\['word\.list_notes', \{ category: 'Notes', sideEffect: 'read', description: 'List footnotes or endnotes with reference locations\.' \}\]/);
  assert.match(js, /\['word\.update_note', \{ category: 'Notes', sideEffect: 'mutating', description: 'Replace a footnote or endnote body by index\.' \}\]/);
  assert.match(js, /\['word\.delete_note', \{ category: 'Notes', sideEffect: 'destructive', description: 'Delete a footnote or endnote by index\.' \}\]/);
  assert.match(invokeBody, /case 'word\.insert_note':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertNote\(args\);/);
  assert.match(invokeBody, /case 'word\.list_notes':\s*data = await listNotes\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.update_note':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateNote\(args\);/);
  assert.match(invokeBody, /case 'word\.delete_note':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await deleteNote\(args\);/);
  assert.match(js, /WordApi_1_5: requirements\.isSetSupported\('WordApi', '1\.5'\) \? '1\.5' : null/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_5[\s\S]*word\.insert_note/);
  assert.match(functionBody(js, 'insertNote'), /insertFootnote\(text\)|insertEndnote\(text\)/);
  assert.match(functionBody(js, 'noteCollectionForKind'), /footnotes[\s\S]*endnotes/);
});

test('Word hyperlink tools are advertised, grouped, and dispatched through range owners', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');

  assert.match(js, /'word\.insert_hyperlink'/);
  assert.match(js, /'word\.list_hyperlinks'/);
  assert.match(js, /'word\.remove_hyperlink'/);
  assert.match(js, /\{ label: 'Range & selection', tools: \[[\s\S]*'word\.insert_hyperlink'[\s\S]*'word\.list_hyperlinks'[\s\S]*'word\.remove_hyperlink'[\s\S]*\] \}/);
  assert.match(js, /\['word\.insert_hyperlink', \{ category: 'Range & selection', sideEffect: 'mutating', description: 'Insert or apply a hyperlink at an anchored range\.' \}\]/);
  assert.match(js, /\['word\.list_hyperlinks', \{ category: 'Range & selection', sideEffect: 'read', description: 'List document hyperlinks with paragraph-relative locations\.' \}\]/);
  assert.match(js, /\['word\.remove_hyperlink', \{ category: 'Range & selection', sideEffect: 'mutating', description: 'Remove a hyperlink while preserving text by default\.' \}\]/);
  assert.match(invokeBody, /case 'word\.insert_hyperlink':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertHyperlink\(args\);/);
  assert.match(invokeBody, /case 'word\.list_hyperlinks':\s*data = await listHyperlinks\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.remove_hyperlink':\s*data = await removeHyperlink\(args\);/);
  assert.match(functionBody(js, 'insertHyperlink'), /range\.hyperlink = url/);
  assert.match(js, /async function listHyperlinks[\s\S]*getHyperlinkRanges\(\)/);
  assert.match(functionBody(js, 'removeHyperlink'), /range\.hyperlink = ''/);
});

test('word.resolve_anchor returns safe anchor diagnostics without mutating', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const resolveBody = functionBody(js, 'resolveAnchorTool');
  const describeBody = functionBody(js, 'describeResolvedAnchor');

  assert.match(js, /'word\.resolve_anchor'/);
  assert.match(js, /'word\.insert_bookmark'/);
  assert.match(js, /'word\.list_bookmarks'/);
  assert.match(js, /'word\.delete_bookmark'/);
  assert.match(js, /\['word\.resolve_anchor', \{ category: 'Range & selection', sideEffect: 'read', description: 'Resolve an anchor to safe diagnostic metadata\.' \}\]/);
  assert.match(invokeBody, /case 'word\.resolve_anchor':\s*data = await resolveAnchorTool\(args\);/);
  assert.match(resolveBody, /requireAnchor\('word\.resolve_anchor', args\.anchor\)/);
  assert.match(resolveBody, /const resolved = await resolveAnchor\(context, args\.anchor\)/);
  assert.match(resolveBody, /describeResolvedAnchor\(context, args\.anchor, resolved, args\.include_text_preview !== false\)/);
  assert.match(describeBody, /resolved: true/);
  assert.match(describeBody, /object_type: resolvedAnchorObjectType\(anchor\)/);
  assert.match(describeBody, /supported_operations: supportedOperationsForAnchor\(anchor\)/);
  assert.match(describeBody, /unsupported_operations: unsupportedOperationsForAnchor\(anchor\)/);
  assert.match(describeBody, /tool_suitability: toolSuitabilityForAnchor\(anchor\)/);
  assert.match(describeBody, /untrusted_source: true/);
  assert.match(functionBody(js, 'toolSuitabilityForAnchor'), /image_insertion/);
  assert.match(functionBody(js, 'toolSuitabilityForAnchor'), /text_replacement/);
  assert.match(functionBody(js, 'toolSuitabilityForAnchor'), /deletion/);
  assert.match(functionBody(js, 'toolSuitabilityForAnchor'), /formatting/);
  assert.match(functionBody(js, 'safeTextPreview'), /text\.length > 80/);
  assert.doesNotMatch(functionBody(js, 'describeResolvedAnchor'), /document\.body\.text|body\.load\('text'/);
});

test('word.set_selection is advertised, grouped, and dispatches through anchor selection', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');
  const setSelectionBody = functionBody(js, 'setSelection');

  assert.match(js, /'word\.set_selection'/);
  assert.match(js, /\{ label: 'Range & selection', tools: \[[\s\S]*'word\.get_selection'[\s\S]*'word\.set_selection'[\s\S]*'word\.resolve_anchor'[\s\S]*\] \}/);
  assert.match(js, /\['word\.set_selection', \{ category: 'Range & selection', sideEffect: 'mutating', description: 'Set the current selection or cursor position from an anchor\.' \}\]/);
  assert.match(invokeBody, /case 'word\.set_selection':\s*data = await setSelection\(args\);/);
  assert.match(preflightBody, /case 'word\.set_selection':\s*validateSetSelectionArgs\(tool, args\);/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_3[\s\S]*word\.set_selection/);
  assert.match(setSelectionBody, /const target = await resolveRangeForExtent\(context, args\.anchor, args\.extent\);/);
  assert.match(setSelectionBody, /target\.select\(selectionModeForSetSelection\(args\.mode\)\);/);
  assert.match(setSelectionBody, /selected_text_preview: safeTextPreview\(target\.text\)/);
  assert.match(js, /function selectionModeForSetSelection\(mode\)/);
  assert.match(functionBody(js, 'selectionModeForSetSelection'), /cursor_start[\s\S]*Word\.SelectionMode\.start/);
  assert.match(functionBody(js, 'selectionModeForSetSelection'), /cursor_end[\s\S]*Word\.SelectionMode\.end/);
});

test('word HTML interchange tools are advertised, sanitized, and dispatched', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const invokeBody = functionBody(js, 'invokeTool');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');
  const getHtmlBody = functionBody(js, 'getHtml');
  const insertHtmlBody = functionBody(js, 'insertHtml');
  const sanitizeBody = functionBody(js, 'validateSafeHtmlForWord');

  assert.match(js, /'word\.get_html'/);
  assert.match(js, /'word\.insert_html'/);
  assert.match(js, /\{ label: 'Range & selection', tools: \[[\s\S]*'word\.get_selection'[\s\S]*'word\.set_selection'[\s\S]*'word\.get_html'[\s\S]*'word\.insert_html'[\s\S]*'word\.resolve_anchor'[\s\S]*\] \}/);
  assert.match(js, /\['word\.get_html', \{ category: 'Range & selection', sideEffect: 'read', description: 'Read document or anchored range HTML\.' \}\]/);
  assert.match(js, /\['word\.insert_html', \{ category: 'Range & selection', sideEffect: 'mutating', description: 'Insert sanitized HTML at an anchored range\.' \}\]/);
  assert.match(invokeBody, /case 'word\.get_html':\s*data = await getHtml\(args \|\| \{\}\);/);
  assert.match(invokeBody, /case 'word\.insert_html':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertHtml\(args\);/);
  assert.match(preflightBody, /case 'word\.insert_html':[\s\S]*validateInsertHtmlArgs\(tool, args\);/);
  assert.match(functionBody(js, 'availableToolsForRequirements'), /WordApi_1_3[\s\S]*word\.get_html[\s\S]*word\.insert_html/);
  assert.match(getHtmlBody, /args\.anchor \? await resolveRangeForExtent\(context, args\.anchor, args\.extent\) : context\.document\.body/);
  assert.match(getHtmlBody, /getHtml\(\)/);
  assert.match(getHtmlBody, /enforceResponseSizeLimit\(html/);
  assert.match(insertHtmlBody, /validateSafeHtmlForWord\(args\.html\)/);
  assert.match(insertHtmlBody, /const target = await resolveAnchor\(context, args\.anchor\);/);
  assert.match(insertHtmlBody, /target\.insertHtml\(args\.html, insertLocationForHtml\(args\.insert_location\)\);/);
  assert.match(js, /function validateInsertHtmlArgs\(tool, args\)/);
  assert.match(js, /function validateSafeHtmlForWord\(html\)/);
  assert.match(js, /function insertLocationForHtml\(location\)/);
  assert.match(sanitizeBody, /<script\b/i);
  assert.match(sanitizeBody, /\\bon\[a-z\]\+\\s\*=|\\bon\[a-z\]\+\\s\*=/i);
  assert.match(sanitizeBody, /javascript:/i);
  assert.match(sanitizeBody, /\\b\(\?:src\|srcset\|poster\|background\)\\s\*=/i);
  assert.match(sanitizeBody, /\\burl\\s\*\\\(/i);
});

test('word HTML sanitizer rejects unsafe payload classes before mutation', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const validateBody = functionBody(js, 'validateSafeHtmlForWord');
  const script = `function invalidArgument(message) { return Object.assign(new Error(message), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' }); }\nfunction validateSafeHtmlForWord(html) {${validateBody}}\nconst unsafe = [\n  '<script>alert(1)</script>',\n  '<p onclick="alert(1)">x</p>',\n  '<a href="javascript:alert(1)">x</a>',\n  '<img src="https://example.com/x.png">',\n  '<p style="background:url(https://example.com/x.png)">x</p>'\n];\nfor (const html of unsafe) {\n  let rejected = false;\n  try { validateSafeHtmlForWord(html); } catch (error) {\n    rejected = error.officeMcpCode === 'INVALID_ARGUMENT' && error.partialEffect === 'none';\n  }\n  if (!rejected) throw new Error('Expected unsafe HTML rejection for ' + html);\n}\nvalidateSafeHtmlForWord('<h1>Title</h1><p><strong>Bold</strong> <a href="https://example.com">link</a></p><table><tr><td>A</td></tr></table>');`;
  const tmpDir = mkdtempSync(join(tmpdir(), 'office-mcp-html-'));
  const scriptPath = join(tmpDir, 'html-sanitizer-test.cjs');
  try {
    writeFileSync(scriptPath, script);
    const result = spawnSync(process.execPath, [scriptPath], { encoding: 'utf8' });
    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
  } finally {
    rmSync(tmpDir, { recursive: true, force: true });
  }
});

test('word.apply_formatting supports paragraph layout formatting and readback metadata', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const preflightBody = functionBody(js, 'preflightWordMutatingTool');
  const applyBody = functionBody(js, 'applyFormatting');
  const getTextBody = functionBody(js, 'getText');

  assert.match(js, /function validateParagraphFormattingArg\(tool, paragraph/);
  assert.match(js, /function hasFormattingFields\(formatting\)/);
  assert.match(js, /function hasParagraphFormattingFields\(paragraph\)/);
  assert.match(preflightBody, /case 'word\.apply_formatting':[\s\S]*validateFormattingBlocks\(tool, args\.formatting, args\.paragraph\)/);
  assert.match(applyBody, /if \(args\.formatting\) applyRunFormatting\(range\.font, args\.formatting\);/);
  assert.match(applyBody, /if \(args\.paragraph\) await applyParagraphFormatting\(context, range, args\.paragraph\);/);
  assert.match(js, /async function applyParagraphFormatting\(context, range, paragraph\)/);
  assert.match(functionBody(js, 'applyParagraphFormatting'), /paragraphs\.load\('items'\)/);
  assert.match(functionBody(js, 'applyParagraphFormatting'), /applyParagraphFormattingToParagraph\(item, paragraph\)/);
  assert.match(js, /function applyParagraphFormattingToParagraph\(paragraph, formatting\)/);
  assert.match(functionBody(js, 'applyParagraphFormattingToParagraph'), /paragraph\.alignment = Word\.Alignment\.centered/);
  assert.match(functionBody(js, 'applyParagraphFormattingToParagraph'), /paragraph\.firstLineIndent = formatting\.first_line_indent_pt/);
  assert.match(getTextBody, /args\.include_formatting/);
  assert.match(js, /function paragraphFormattingMetadata\(paragraph\)/);
  assert.match(functionBody(js, 'paragraphFormattingMetadata'), /alignment: normalizedParagraphAlignment\(paragraph\.alignment\)/);
  assert.match(functionBody(js, 'paragraphFormattingMetadata'), /space_after_pt: paragraph\.spaceAfter/);
});

test('Word task pane preserves safe Office.js error debug context', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const mapBody = functionBody(js, 'mapError');
  const debugBody = functionBody(js, 'officeErrorDebug');

  assert.match(js, /mapError\(error, message\.params\?\.tool, message\.params\?\.args\)/);
  assert.match(js, /mapError\(error, tool, args\)/);
  assert.match(mapBody, /const debug = officeErrorDebug\(error, tool, args\)/);
  assert.match(mapBody, /if \(debug\) mapped\.debug = debug/);
  assert.match(functionBody(js, 'classifyOfficeError'), /InvalidArgument\|InvalidObjectPath\|InvalidSelection\|ItemNotFound/);
  assert.match(functionBody(js, 'errorMessage'), /Word\.js \$\{officeCode\} while running/);
  assert.match(debugBody, /office_error_code: officeCode/);
  assert.match(debugBody, /error_location: safeDebugString\(error\.debugInfo\?\.errorLocation/);
  assert.match(debugBody, /statement: safeDebugString\(error\.debugInfo\?\.statement\)/);
  assert.match(debugBody, /\.\.\.safeArgumentContext\(args\)/);
  assert.match(js, /context\.anchor_kind = String\(args\.anchor\.kind\)/);
  assert.match(js, /context\.placement = String\(args\.placement\)/);
  assert.match(js, /context\.image_mime_type = String\(args\.image\.mime_type\)/);
  assert.match(js, /context\.image_byte_length = args\.image\.byte_length/);
  assert.doesNotMatch(functionBody(js, 'officeErrorDebug'), /base64/);
  assert.doesNotMatch(functionBody(js, 'officeErrorDebug'), /text_preview|find|replace/);
  assert.match(functionBody(js, 'looksSensitive'), /base64\|data:image/);
});

test('Word task pane exposes product UI regions and accessible endpoint settings', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(html, /id="connectionBadge"/);
  assert.match(html, /id="runtimeVersions"/);
  assert.match(html, /id="serverVersion"/);
  assert.match(html, /id="protocolVersion"/);
  assert.match(html, /<dd id="runtimeVersions"><span id="serverVersion">Server Unknown<\/span> \/ <span id="protocolVersion">Protocol 1\.0<\/span><\/dd>/);
  assert.match(html, /id="hostPlatform"/);
  assert.match(html, /id="documentState"/);
  assert.match(html, /<dd id="protection">Not protected<\/dd>/);
  assert.match(html, /<dd id="documentState">Editable<\/dd>/);
  assert.match(html, /id="connectionDetail"/);
  assert.doesNotMatch(html, /id="endpointError"/);
  assert.doesNotMatch(html, /class="field-error"/);
  assert.match(html, /class="metadata-copy" data-copy-target="session" aria-label="Copy session ID" title="Copy session ID"/);
  assert.match(html, /class="daemon-endpoint-input" name="daemonEndpoint" type="url" inputmode="url" autocomplete="off" spellcheck="false" aria-label="Daemon endpoint"/);
  assert.match(html, /id="currentTask"/);
  assert.match(html, /id="historyList"/);
  assert.match(html, /class="panel summary-panel"/);
  assert.match(html, /class="summary-layout"/);
  assert.match(html, /class="metadata-section document-metadata"/);
  assert.match(html, /class="metadata-section daemon-metadata"/);
  assert.match(html, /<h2 id="daemonHeading">Daemon<\/h2>/);
  assert.match(html, /class="tools-panel"/);
  assert.match(html, /<span>Tools<\/span>/);
  assert.match(html, /id="toolList"/);
  assert.match(html, /0\/0/);
  assert.doesNotMatch(html, /Enabled \d+ of \d+/);
  assert.match(html, /id="toolModeControl" class="tool-mode-control" role="radiogroup" aria-label="Tool capability mode"/);
  assert.match(html, /data-tool-mode="read"/);
  assert.match(html, /data-tool-mode="write"/);
  assert.match(html, /data-tool-mode="all"/);
  assert.doesNotMatch(html, /Tool Permissions/);
  assert.doesNotMatch(html, /id="toolPermissionList"/);
  assert.doesNotMatch(html, /id="enabledToolCount"/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /class="daemon-endpoint-input" name="daemonEndpoint" type="url" inputmode="url" autocomplete="off" spellcheck="false" aria-label="Daemon endpoint"/);
  assert.match(html, /id="saveEndpoint" class="icon-button reconnect-button" type="submit" aria-label="Reconnect daemon" title="Reconnect daemon"/);
  assert.doesNotMatch(html, /id="settingsToggle"/);
  assert.doesNotMatch(html, /id="settingsPanel"/);
  assert.doesNotMatch(html, /Save Endpoint/);
  assert.ok(html.indexOf('class="daemon-endpoint-form"') < html.indexOf('id="runtimeVersions"'));
  assert.match(css, /:focus-visible/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8'), /\.summary-panel/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8'), /\.identity \{[\s\S]*grid-template-columns: 32px minmax\(0, 1fr\);/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8'), /\.product-mark \{[\s\S]*width: 32px;[\s\S]*height: 32px;/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8'), /\.control-glyph \{[\s\S]*width: 18px;[\s\S]*stroke: currentColor;/);
  assert.match(css, /\.tool-list/);
    assert.match(css, /\.tool-permission-row/);
  assert.match(css, /\.tool-toggle/);
  assert.doesNotMatch(css, /is-editing-tools/);
  assert.match(css, /\.metadata-copy \{[\s\S]*display: inline-flex;[\s\S]*min-height: 32px;[\s\S]*text-align: left;/);
  assert.match(css, /\.metadata-copy code \{[\s\S]*overflow-wrap: anywhere;[\s\S]*white-space: normal;/);
  assert.match(css, /\.task-command-id \{[\s\S]*display: flex;[\s\S]*gap: 4px;/);
  assert.match(css, /\.inline-copy \{[\s\S]*display: inline-flex;[\s\S]*min-height: 32px;[\s\S]*background: transparent;[\s\S]*cursor: pointer;/);
  assert.match(css, /\.metadata-copy:hover,[\s\S]*\.inline-copy:focus-visible \{[\s\S]*background: var\(--surface-raised\);/);
  assert.match(css, /prefers-reduced-motion/);
  assert.match(css, /forced-colors: active/);
  assert.match(js, /const TOOL_GROUPS = \[/);
  assert.match(js, /const TOOL_METADATA = new Map\(\[/);
  const availableToolsSource = js.match(/const AVAILABLE_TOOLS = \[([\s\S]*?)\];/)?.[1] || '';
  assert.match(js, /'word\.update_table'/);
  assert.match(js, /'word\.list_content_controls'/);
  assert.match(js, /'word\.insert_content_control'/);
  assert.match(js, /'word\.update_content_control'/);
  assert.match(js, /'word\.delete_content_control'/);
  assert.doesNotMatch(availableToolsSource, /'word\.resize_image'/);
  assert.doesNotMatch(availableToolsSource, /'word\.delete_image'/);
  assert.match(js, /'word\.update_tracked_change'/);
  assert.match(js, /'word\.update_comment'/);
  assert.match(js, /'word\.set_change_tracking'/);
  assert.match(js, /'word\.insert_note'/);
  assert.match(js, /'word\.list_notes'/);
  assert.match(js, /'word\.update_note'/);
  assert.match(js, /'word\.delete_note'/);
  assert.match(js, /'word\.resolve_anchor'/);
  assert.match(js, /'word\.get_header_footer'/);
  assert.match(js, /'word\.update_header_footer'/);
  assert.match(js, /'word\.insert_break'/);
  assert.match(js, /'word\.list_sections'/);
  assert.match(js, /'word\.update_page_setup'/);
  assert.match(js, /'word\.list_lists'/);
  assert.match(js, /'word\.update_list'/);
  assert.match(js, /\{ label: 'Document & structure', tools: \['word\.get_text', 'word\.get_outline', 'word\.get_header_footer', 'word\.update_header_footer', 'word\.get_document_properties', 'word\.update_document_properties', 'word\.insert_break', 'word\.list_sections', 'word\.update_page_setup', 'word\.list_fields', 'word\.insert_field', 'word\.update_field', 'word\.delete_field', 'word\.list_styles', 'word\.create_style', 'word\.update_style', 'word\.save'\] \}/);
  assert.match(js, /\{ label: 'Range & selection', tools: \['word\.get_selection', 'word\.set_selection', 'word\.get_html', 'word\.insert_html', 'word\.find_text', 'word\.resolve_anchor', 'word\.insert_bookmark', 'word\.list_bookmarks', 'word\.delete_bookmark', 'word\.insert_hyperlink', 'word\.list_hyperlinks', 'word\.remove_hyperlink', 'word\.replace_text', 'word\.delete_range', 'word\.apply_formatting', 'word\.apply_style'\] \}/);
  assert.match(js, /\{ label: 'Paragraphs & lists', tools: \['word\.insert_paragraph', 'word\.update_paragraph', 'word\.insert_list', 'word\.list_lists', 'word\.update_list'\] \}/);
  assert.match(js, /\{ label: 'Tables', tools: \['word\.read_table', 'word\.update_table'\] \}/);
  assert.match(js, /\{ label: 'Media', tools: \['word\.insert_image', 'word\.list_images', 'word\.get_image', 'word\.update_image', 'word\.list_shapes', 'word\.insert_shape', 'word\.update_shape', 'word\.delete_shape'\] \}/);
  assert.match(js, /\{ label: 'Content controls', tools: \['word\.list_content_controls', 'word\.insert_content_control', 'word\.update_content_control', 'word\.delete_content_control'\] \}/);
  assert.match(js, /\{ label: 'Notes', tools: \['word\.insert_note', 'word\.list_notes', 'word\.update_note', 'word\.delete_note'\] \}/);
  assert.match(js, /\{ label: 'Review', tools: \['word\.add_comment', 'word\.resolve_comment', 'word\.update_comment', 'word\.set_change_tracking', 'word\.update_tracked_change'\] \}/);
  assert.doesNotMatch(js, /'word\.insert_heading'/);
  assert.doesNotMatch(js, /'word\.set_heading_level'/);
  assert.doesNotMatch(js, /'word\.update_cell'/);
  assert.doesNotMatch(js, /'word\.add_row'/);
  assert.doesNotMatch(js, /'word\.add_column'/);
  assert.doesNotMatch(js, /'word\.format_cell'/);
  assert.doesNotMatch(js, /'word\.accept_change'/);
  assert.doesNotMatch(js, /'word\.reject_change'/);
  assert.doesNotMatch(availableToolsSource, /'word\.get_paragraph'/);
  assert.doesNotMatch(availableToolsSource, /'word\.insert_page_break'/);
  assert.match(js, /\['word\.update_table', \{ category: 'Tables', sideEffect: 'destructive', description: 'Update table cells, rows, columns, formatting, or lifecycle\.' \}\]/);
  assert.match(js, /\['word\.list_content_controls', \{ category: 'Content controls', sideEffect: 'read', description: 'List content-control metadata\.' \}\]/);
  assert.match(js, /\['word\.insert_content_control', \{ category: 'Content controls', sideEffect: 'mutating', description: 'Create a content control around an anchored range\.' \}\]/);
  assert.match(js, /\['word\.update_content_control', \{ category: 'Content controls', sideEffect: 'mutating', description: 'Update content-control metadata, locks, or text\.' \}\]/);
  assert.match(js, /\['word\.delete_content_control', \{ category: 'Content controls', sideEffect: 'destructive', description: 'Delete a content control with explicit content handling\.' \}\]/);
  assert.match(js, /\['word\.list_images', \{ category: 'Media', sideEffect: 'read', description: 'List inline images\.' \}\]/);
  assert.match(js, /\['word\.get_image', \{ category: 'Media', sideEffect: 'read', description: 'Export an inline image with metadata\.' \}\]/);
  assert.match(js, /\['word\.update_image', \{ category: 'Media', sideEffect: 'destructive', description: 'Resize, update, replace, or delete an inline image\.' \}\]/);
  assert.match(js, /\['word\.list_shapes', \{ category: 'Media', sideEffect: 'read', description: 'List desktop Word shapes and text boxes\.' \}\]/);
  assert.match(js, /\['word\.insert_shape', \{ category: 'Media', sideEffect: 'mutating', description: 'Insert a desktop Word shape or text box\.' \}\]/);
  assert.match(js, /\['word\.update_shape', \{ category: 'Media', sideEffect: 'mutating', description: 'Update desktop Word shape text, geometry, or visual settings\.' \}\]/);
  assert.match(js, /\['word\.delete_shape', \{ category: 'Media', sideEffect: 'destructive', description: 'Delete a desktop Word shape\.' \}\]/);
  assert.match(js, /\['word\.resolve_anchor', \{ category: 'Range & selection', sideEffect: 'read', description: 'Resolve an anchor to safe diagnostic metadata\.' \}\]/);
  assert.match(js, /\['word\.insert_bookmark', \{ category: 'Range & selection', sideEffect: 'mutating', description: 'Create a named bookmark around an anchored range\.' \}\]/);
  assert.match(js, /\['word\.list_bookmarks', \{ category: 'Range & selection', sideEffect: 'read', description: 'List bookmark names and locations\.' \}\]/);
  assert.match(js, /\['word\.delete_bookmark', \{ category: 'Range & selection', sideEffect: 'destructive', description: 'Delete a bookmark marker without deleting text\.' \}\]/);
  assert.match(js, /\['word\.get_header_footer', \{ category: 'Document & structure', sideEffect: 'read', description: 'Read section header or footer text\.' \}\]/);
  assert.match(js, /\['word\.update_header_footer', \{ category: 'Document & structure', sideEffect: 'destructive', description: 'Replace, append, or clear a section header or footer\.' \}\]/);
  assert.match(js, /\['word\.get_document_properties', \{ category: 'Document & structure', sideEffect: 'read', description: 'Read document metadata and custom properties\.' \}\]/);
  assert.match(js, /\['word\.update_document_properties', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Update document metadata and custom properties\.' \}\]/);
  assert.match(js, /\['word\.insert_break', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Insert a page, line, or section break\.' \}\]/);
  assert.match(js, /\['word\.list_sections', \{ category: 'Document & structure', sideEffect: 'read', description: 'List document sections\.' \}\]/);
  assert.match(js, /\['word\.update_page_setup', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Update document or section page setup\.' \}\]/);
  assert.match(js, /\['word\.list_lists', \{ category: 'Paragraphs & lists', sideEffect: 'read', description: 'List existing Word lists and their paragraph items\.' \}\]/);
  assert.match(js, /\['word\.update_list', \{ category: 'Paragraphs & lists', sideEffect: 'destructive', description: 'Mutate existing Word list membership, levels, or formatting\.' \}\]/);

  assert.match(js, /\['word\.list_fields', \{ category: 'Document & structure', sideEffect: 'read', description: 'List document fields with bounded previews\.' \}\]/);
  assert.match(js, /\['word\.insert_field', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Insert a curated Word field at an anchored range\.' \}\]/);
  assert.match(js, /\['word\.update_field', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Refresh, lock, or unlock Word fields\.' \}\]/);
  assert.match(js, /\['word\.delete_field', \{ category: 'Document & structure', sideEffect: 'destructive', description: 'Delete a Word field by current index\.' \}\]/);
  assert.match(js, /\['word\.list_styles', \{ category: 'Document & structure', sideEffect: 'read', description: 'List built-in and custom document styles\.' \}\]/);
  assert.match(js, /\['word\.create_style', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Create a document style definition\.' \}\]/);
  assert.match(js, /\['word\.update_style', \{ category: 'Document & structure', sideEffect: 'mutating', description: 'Update a document style definition\.' \}\]/);
  assert.match(js, /\['word\.insert_note', \{ category: 'Notes', sideEffect: 'mutating', description: 'Insert a footnote or endnote at an anchored range\.' \}\]/);
  assert.match(js, /\['word\.list_notes', \{ category: 'Notes', sideEffect: 'read', description: 'List footnotes or endnotes with reference locations\.' \}\]/);
  assert.match(js, /\['word\.update_note', \{ category: 'Notes', sideEffect: 'mutating', description: 'Replace a footnote or endnote body by index\.' \}\]/);
  assert.match(js, /\['word\.delete_note', \{ category: 'Notes', sideEffect: 'destructive', description: 'Delete a footnote or endnote by index\.' \}\]/);
  assert.match(js, /\['word\.update_comment', \{ category: 'Review', sideEffect: 'destructive', description: 'Reply, edit, delete, or reopen a comment thread\.' \}\]/);
  assert.match(js, /\['word\.set_change_tracking', \{ category: 'Review', sideEffect: 'mutating', description: 'Set Track Changes mode\.' \}\]/);
  assert.match(js, /\['word\.update_tracked_change', \{ category: 'Review', sideEffect: 'destructive', description: 'Accept, reject, or bulk-finalize tracked changes\.' \}\]/);
  assert.match(js, /case 'word\.get_header_footer':\s*data = await getHeaderFooter\(args\);/);
  assert.match(js, /case 'word\.update_header_footer':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateHeaderFooter\(args\);/);
  assert.match(js, /case 'word\.get_document_properties':\s*data = await getDocumentProperties\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.update_document_properties':\s*data = await updateDocumentProperties\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.insert_break':\s*data = await insertBreak\(args\);/);
  assert.match(js, /case 'word\.insert_page_break':\s*data = await insertPageBreak\(args\);/);
  assert.match(js, /case 'word\.list_sections':\s*data = await listSections\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.update_page_setup':\s*data = await updatePageSetup\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.list_lists':\s*data = await listLists\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.update_list':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateList\(args\);/);
  assert.match(js, /superseded_by: 'word\.insert_break'/);
  assert.match(js, /case 'word\.update_table':\s*data = await updateTable\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'delete_row':\s*return deleteTableRows\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'delete_column':\s*return deleteTableColumns\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'merge_cells':\s*return mergeTableCells\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'set_column_width':\s*return setTableColumnWidth\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'distribute_columns':\s*return distributeTableColumns\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'set_borders':\s*return setTableBorders\(args\);/);
  assert.match(functionBody(js, 'updateTable'), /case 'set_header_row':\s*return setTableHeaderRow\(args\);/);
  assert.match(js, /async function deleteTableRows\(args\)/);
  assert.match(js, /async function deleteTableColumns\(args\)/);
  assert.match(js, /async function mergeTableCells\(args\)/);
  assert.match(js, /async function setTableColumnWidth\(args\)/);
  assert.match(js, /async function distributeTableColumns\(args\)/);
  assert.match(js, /async function setTableBorders\(args\)/);
  assert.match(js, /async function setTableHeaderRow\(args\)/);
  assert.match(js, /case 'word\.list_content_controls':\s*data = await listContentControls\(args\);/);
  assert.match(js, /case 'word\.insert_content_control':\s*data = await insertContentControl\(args\);/);
  assert.match(js, /case 'word\.update_content_control':\s*data = await updateContentControl\(args\);/);
  assert.match(js, /case 'word\.delete_content_control':\s*data = await deleteContentControl\(args\);/);
  assert.doesNotMatch(js, /case 'word\.resize_image':\s*data = await resizeImage\(args\);/);
  assert.doesNotMatch(js, /case 'word\.delete_image':/);
  assert.match(js, /case 'word\.list_shapes':\s*data = await listShapes\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.insert_shape':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertShape\(args\);/);
  assert.match(js, /case 'word\.update_shape':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateShape\(args\);/);
  assert.match(js, /case 'word\.delete_shape':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await deleteShape\(args\);/);
  assert.match(js, /case 'word\.resolve_anchor':\s*data = await resolveAnchorTool\(args\);/);
  assert.match(js, /case 'word\.insert_bookmark':\s*data = await insertBookmark\(args\);/);
  assert.match(js, /case 'word\.list_bookmarks':\s*data = await listBookmarks\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.delete_bookmark':\s*data = await deleteBookmark\(args\);/);
  assert.match(js, /case 'word\.insert_note':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertNote\(args\);/);
  assert.match(js, /case 'word\.list_notes':\s*data = await listNotes\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.update_note':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateNote\(args\);/);
  assert.match(js, /case 'word\.delete_note':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await deleteNote\(args\);/);
  assert.match(js, /case 'word\.update_comment':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateComment\(args\);/);
  assert.match(js, /case 'word\.list_fields':\s*data = await listFields\(args \|\| \{\}\);/);
  assert.match(js, /case 'word\.insert_field':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await insertField\(args\);/);
  assert.match(js, /case 'word\.update_field':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await updateField\(args\);/);
  assert.match(js, /case 'word\.delete_field':\s*data = args\?\.validate_only \? await validateWordMutationOnly\(tool, args\) : await deleteField\(args\);/);
  assert.match(js, /case 'word\.set_change_tracking':\s*data = await setChangeTracking\(args\);/);
  assert.match(js, /case 'word\.update_tracked_change':\s*data = await updateTrackedChange\(args\);/);
  assert.match(js, /async function insertTable\(args\)/);
  assert.match(js, /table_index: tableIndex/);
  assert.match(js, /async function updateTable\(args\)/);
  assert.match(js, /async function listContentControls\(args\)/);
  assert.match(js, /async function insertContentControl\(args\)/);
  assert.match(js, /async function updateContentControl\(args\)/);
  assert.match(js, /async function deleteContentControl\(args\)/);
  assert.doesNotMatch(js, /async function resizeImage\(args\)/);
  assert.doesNotMatch(js, /async function deleteImage\(args\)/);
  assert.match(js, /async function insertBookmark\(args\)/);
  assert.match(js, /async function listBookmarks\(args\)/);
  assert.match(js, /async function deleteBookmark\(args\)/);
  assert.match(js, /async function insertNote\(args\)/);
  assert.match(js, /async function listNotes\(args\)/);
  assert.match(js, /async function updateNote\(args\)/);
  assert.match(js, /async function deleteNote\(args\)/);
  assert.match(js, /async function updateComment\(args\)/);
  assert.match(js, /async function validateUpdateCommentOnly\(args\)/);
  assert.match(js, /async function listLists\(args/);
  assert.match(js, /async function updateList\(args\)/);

  assert.match(js, /async function listFields\(args\)/);
  assert.match(js, /async function insertField\(args\)/);
  assert.match(js, /async function updateField\(args\)/);
  assert.match(js, /async function deleteField\(args\)/);
  assert.match(js, /async function getDocumentProperties\(args\)/);
  assert.match(js, /async function updateDocumentProperties\(args\)/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.insert_bookmark':\s*requireAnchor\(tool, args\.anchor\);\s*validateBookmarkName\(tool, args\.name\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.delete_bookmark':\s*validateBookmarkName\(tool, args\.name, \{ strictPattern: false \}\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.insert_note':\s*validateInsertNoteArgs\(tool, args\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.update_note':\s*validateUpdateNoteArgs\(tool, args\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.delete_note':\s*validateDeleteNoteArgs\(tool, args\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.update_comment':\s*validateUpdateCommentArgs\(tool, args\);/);
  assert.match(functionBody(js, 'getComments'), /replies: commentRepliesMetadata\(comment\)/);
  assert.match(js, /function commentRepliesMetadata\(comment\)/);
  assert.match(js, /function validateUpdateCommentArgs\(tool, args\)/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.insert_field':\s*validateInsertFieldArgs\(tool, args\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.update_field':\s*validateUpdateFieldArgs\(tool, args\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.delete_field':\s*validateDeleteFieldArgs\(tool, args\);/);
  assert.match(functionBody(js, 'preflightWordMutatingTool'), /case 'word\.update_document_properties':\s*validateUpdateDocumentPropertiesArgs\(tool, args\);/);
  assert.match(js, /function validateBookmarkName\(tool, name/);
  assert.match(js, /control\.load\('id'\);\s*await context\.sync\(\);\s*const id = control\.id;/);
  assert.match(js, /async function updateTrackedChange\(args\)/);
  assert.match(js, /async function setChangeTracking\(args\)/);
  assert.match(js, /return mutateTrackedChange\(args, action\);/);
  assert.match(js, /function targetContentControl\(context, args\)/);
  assert.match(js, /function contentControlMetadata\(control, index\)/);
  assert.match(js, /WordApi_1_5: requirements\.isSetSupported\('WordApi', '1\.5'\) \? '1\.5' : null/);
  assert.match(js, /WordApi_1_7: requirements\.isSetSupported\('WordApi', '1\.7'\) \? '1\.7' : null/);
  assert.match(js, /WordApi_1_9: requirements\.isSetSupported\('WordApi', '1\.9'\) \? '1\.9' : null/);
  assert.match(functionBody(js, 'contentControlTypeFrom'), /checkbox[\s\S]*CheckBox/);
  assert.match(functionBody(js, 'contentControlTypeFrom'), /dropdown_list[\s\S]*DropDownList/);
  assert.match(functionBody(js, 'contentControlTypeFrom'), /combo_box[\s\S]*ComboBox/);
  assert.match(functionBody(js, 'contentControlMetadata'), /checked: checkboxContentControlChecked\(control\)/);
  assert.match(functionBody(js, 'contentControlMetadata'), /list_items: contentControlListItems\(control\)/);
  assert.match(functionBody(js, 'contentControlMetadata'), /selected_text: contentControlSelectedText\(control\)/);
  assert.match(js, /function requireContentControlTypeCapability\(tool, type\)/);
  assert.match(functionBody(js, 'requireContentControlTypeCapability'), /WordApi 1\.7/);
  assert.match(functionBody(js, 'requireContentControlTypeCapability'), /WordApi 1\.9/);
  assert.match(functionBody(js, 'validateContentControlArgs'), /validateContentControlTypeSpecificArgs\(tool, args\)/);
  assert.match(js, /function applyTypedContentControlState\(context, control, args\)/);
  assert.match(functionBody(js, 'applyTypedContentControlState'), /checkboxContentControl\.isChecked = Boolean\(args\.checked\)/);
  assert.match(functionBody(js, 'applyTypedContentControlState'), /addContentControlListItems\(control, args\.list_items, args\.type\)/);
  assert.match(functionBody(js, 'applyTypedContentControlUpdate'), /selectContentControlListItem\(context, control, args\.selected_value\)/);
  assert.match(functionBody(js, 'applyTypedContentControlUpdate'), /deleteContentControlListItems\(control, args\.list_items_delete\)/);
  assert.match(js, /case 'update_cell':/);
  assert.match(js, /case 'add_row':/);
  assert.match(js, /case 'add_column':/);
  assert.match(js, /case 'format_cell':/);
  assert.match(js, /case 'delete':/);
  assert.match(js, /async function deleteTable\(args\)/);
  assert.match(js, /TOOL_PERMISSION_STORAGE_KEY/);
  assert.match(js, /TOOL_PERMISSION_MODE_STORAGE_KEY/);
  assert.match(js, /isToolAllowedByCapabilityMode,/);
  assert.match(js, /function isToolAllowedByMode\(tool\)/);
  assert.match(functionBody(js, 'isToolAllowedByMode'), /return isToolAllowedByCapabilityMode\(toolPermissionMode, sideEffect\);/);
  assert.match(js, /function handleToolModeChange\(event\)/);
  assert.match(js, /renderSharedToolModeControl\(toolModeControlEl, toolPermissionMode\)/);
  assert.doesNotMatch(functionBody(js, 'renderToolModeControl'), /querySelectorAll\('\[data-tool-mode\]'\)/);
  assert.match(js, /const toolListEl = document\.getElementById\('toolList'\)/);
  assert.match(js, /function renderToolSummary\(\)/);
  assert.match(js, /toolCountEl\.textContent = `\$\{effective\.length\}\/\$\{AVAILABLE_TOOLS\.length\}`/);
  assert.match(js, /\$\{enabledInGroup\.length\}\/\$\{tools\.length\}/);
  assert.doesNotMatch(js, /Enabled \$\{/);
  assert.match(js, /data-tool-group/);
  assert.match(js, /function handleToolGroupPermissionChange\(event\)/);
  assert.match(js, /role="switch"/);
  assert.match(js, /renderToolSummary\(\)/);
  assert.match(js, /function toolControlMarkup\(tool\)/);
  assert.doesNotMatch(js, /function renderToolPermissions\(\)/);
  assert.doesNotMatch(js, /toolPermissionListEl/);
  assert.match(js, /function effectiveTools\(\)/);
  assert.match(js, /available_tools: effectiveTools\(\)/);
  assert.match(js, /sessionUpdatedNotification\(\{/);
  assert.match(js, /patch: \{ available_tools: effectiveTools\(\) \}/);
  assert.match(js, /TOOL_DISABLED_BY_USER/);
  assert.match(js, /new TaskHistoryStore\(\{ redactText \}\)/);
  assert.match(js, /const \{ history, historyLimit \} = taskStore\.snapshot\(\)/);
  assert.match(js, /historyCountEl\.textContent = `\$\{history\.length\} \/ \$\{historyLimit\}`/);
  assert.match(js, /function taskMarkup\(task\)/);
  assert.match(js, /requestId: currentTask\.requestId/);
  assert.match(js, /const commandId = commandIdMarkup\(task\.requestId, \{ escapeHtml \}\)/);
  assert.match(js, /userIntent/);
  assert.match(js, /const requestId = message\.params\?\.request_id \|\| String\(message\.id\)/);
  assert.match(js, /startTask\(requestId, tool, message\.params \|\| \{\}, message\.params\.timeout_ms\)/);
  assert.match(js, /taskStore\.isCancelled\(requestId\)/);
  assert.match(js, /taskStore\.consumeCancellation\(requestId\)/);
  assert.match(js, /finishTask\(requestId, 'success'/);
  assert.match(js, /const metadata = taskMetadataMarkup\(task, \{ escapeHtml, formatTime, redactText, valueLabel: boolLabel \}\)/);
  assert.match(js, /taskStatusClass/);
  assert.match(js, /taskStatusLabel/);
  assert.match(js, /const tone = taskStatusClass\(task\.status \|\| 'running'\)/);
  assert.match(js, /<span class="status-badge \$\{tone\}">\$\{escapeHtml\(taskStatusLabel\(task\.status \|\| 'running'\)\)\}<\/span>/);
  assert.doesNotMatch(functionBody(js, 'taskMarkup'), /titleCase\(task\.status\)/);
  assert.match(js, /storeEndpointOverride\(value\)/);
  assert.doesNotMatch(js, /settingsToggleEl/);
  assert.match(js, /document\.addEventListener\('click', handleMetadataCopy\)/);
  assert.match(js, /setConnectionState: setSharedConnectionState/);
  assert.match(js, /setSharedConnectionState\(\{ badge: connectionBadgeEl, detail: connectionDetailEl, announcer: announcerEl \}, state, label\)/);
  assert.match(js, /async function handleMetadataCopy\(event\)/);
  assert.match(js, /copyMetadataValue\(event, \{ document, navigator, announcer: announcerEl, logger \}\)/);
  assert.doesNotMatch(js, /event\.target\.closest\('\[data-copy-target\], \[data-copy-value\]'\)/);
  assert.match(js, /setCopyableMetadata/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'main-ui.js'), 'utf8'), /function setCopyableMetadata\(element, value\)/);
  assert.match(js, /middleTruncate/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'main-ui.js'), 'utf8'), /function middleTruncate\(value, maxLength = 30\)/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'main-ui.js'), 'utf8'), /return `\$\{text\.slice\(0, head\)\}\$\{marker\}\$\{text\.slice\(text\.length - tail\)\}`/);
  assert.doesNotMatch(js, /function fallbackCopy\(value\)/);
  assert.doesNotMatch(js, /is-editing-settings/);
  assert.doesNotMatch(js, /is-editing-tools/);
  assert.doesNotMatch(js, /function activateSettingsWithKeyboard/);
  assert.match(js, /registerResult\(message, PROTOCOL_VERSION\)/);
  assert.match(js, /renderStaticMetadata\(\{ session: sessionEl, daemon: daemonEl, serverVersion: serverVersionEl, protocolVersion: protocolVersionEl, hostPlatform: hostPlatformEl \}, \{ sessionId, endpoint: configuredEndpoint\(\), serverInfo, protocolVersion: PROTOCOL_VERSION, defaultHost: 'Word' \}\)/);
  assert.doesNotMatch(functionBody(js, 'renderStaticState'), /setCopyableMetadata\(sessionEl/);
  assert.doesNotMatch(functionBody(js, 'renderStaticState'), /renderRuntimeVersions\(/);
  assert.doesNotMatch(functionBody(js, 'renderStaticState'), /officeHostSummary\(/);
  assert.match(js, /protectionLabel,/);
  assert.match(js, /documentStateLabel,/);
  assert.doesNotMatch(js, /function protectionLabel\(info\)/);
  assert.doesNotMatch(js, /function documentStateLabel\(info\)/);
  assert.doesNotMatch(js, /protectionEl\.textContent = documentInfo\.protection\?\.kind \|\| 'Unknown'/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'main-ui.js'), 'utf8'), /function protectionLabel\(info = \{\}\)/);
  assert.match(readFileSync(join(ADDIN_ROOT, '..', 'common', 'main-ui.js'), 'utf8'), /function documentStateLabel\(info = \{\}\)/);
  assert.doesNotMatch(js, /Dirty: \$\{boolLabel/);
  assert.match(js, /protocol_version/);
  assert.match(js, /beforeunload/);
  assert.match(js, /taskMetadataMarkup/);
  assert.match(js, /OfficeCtlCommon/);
  assert.match(js, /OfficeCtlAddinChannel/);
  assert.match(js, /createRequestId/);
  assert.match(js, /clearEndpointOverride/);
  assert.match(js, /currentOriginEndpoint/);
  assert.match(js, /function tryCurrentOriginEndpointFallback\(failedEndpoint\)/);
  assert.match(js, /function announce\(message\)/);
  assert.match(js, /OfficeCtlLogger/);
  assert.match(js, /OfficeCtlTaskHistory/);
  assert.match(js, /window\.__OFFICE_MCP_TASKPANE_READY__ = true/);
  assert.match(js, /function whenOfficeReady\(callback\)/);
  assert.match(js, /!window\.Office \|\| typeof Office\.onReady !== 'function'/);
  assert.match(js, /const requestId = createRequestId\(\)/);
  assert.match(js, /registerRequest\(requestId/);
  assert.match(js, /sessionAddedNotification\(\{/);
  assert.doesNotMatch(js, /function redactText/);
  assert.doesNotMatch(js, /function escapeHtml/);
  assert.doesNotMatch(js, /function configuredEndpoint/);
  assert.doesNotMatch(js, /const taskHistory = \[\]/);
  assert.doesNotMatch(js, /function sanitizeError/);
  assert.doesNotMatch(js, /console\.(log|warn|error)/);
  assert.doesNotMatch(js, /method: 'register'/);
  assert.doesNotMatch(js, /method: 'session\.added'/);
  assert.doesNotMatch(js, /message\.method === 'tool\.invoke'/);
  assert.doesNotMatch(js, /server_version/);
  assert.doesNotMatch(js, /localStorage\.setItem\('office-mcp\.addin-endpoint'/);
  assert.doesNotMatch(js, /sessionStorage\.setItem\('office-mcp\.instance-id'/);
});


test('Word task pane keeps settings inline and compact at narrow widths', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8');

  assert.match(css, /body \{[\s\S]*min-width: 320px;[\s\S]*overflow-x: hidden;/);
  assert.match(css, /\.taskpane-shell \{[\s\S]*align-content: start;[\s\S]*gap: 10px;[\s\S]*padding: 10px;/);
  assert.match(css, /\.summary-panel \{[\s\S]*display: grid;[\s\S]*gap: 10px;/);
  assert.match(css, /\.empty-state \{[\s\S]*padding: 10px;/);
  assert.doesNotMatch(css, /\.field-error/);
  assert.match(css, /#documentTitle \{[\s\S]*display: -webkit-box;[\s\S]*-webkit-line-clamp: 2;/);
  assert.doesNotMatch(css, /\b(min-)?height:\s*(1[2-9]\d|[2-9]\d{2,})px/);
  assert.doesNotMatch(cssRule(css, '.summary-panel'), /\bheight:/);
  assert.doesNotMatch(cssRule(css, '.current-task-panel'), /\bheight:/);
  assert.doesNotMatch(cssRule(css, '.history-panel'), /\bheight:/);
  assert.doesNotMatch(css, /overflow-x:\s*(auto|scroll)/);

  const summaryStart = html.indexOf('class="panel summary-panel"');
  const documentMetadataIndex = html.indexOf('class="metadata-section document-metadata"');
  const daemonMetadataIndex = html.indexOf('class="metadata-section daemon-metadata"');
  const daemonFormIndex = html.indexOf('class="daemon-endpoint-form"');
  const sessionIndex = html.indexOf('data-copy-target="session"');
  const runtimeIndex = html.indexOf('id="runtimeVersions"');
  const currentTaskIndex = html.indexOf('id="currentTaskHeading"');
  assert.ok(summaryStart !== -1 && daemonFormIndex !== -1, 'summary and daemon form exist');
  assert.ok(documentMetadataIndex > summaryStart, 'document metadata is in the summary panel');
  assert.ok(daemonMetadataIndex > documentMetadataIndex, 'daemon metadata is separate from document metadata');
  assert.ok(daemonFormIndex > daemonMetadataIndex, 'daemon endpoint stays inside daemon metadata');
  assert.ok(runtimeIndex > daemonMetadataIndex, 'runtime versions stay inside daemon metadata');
  assert.ok(sessionIndex > daemonMetadataIndex, 'session id stays inside daemon metadata');
  assert.ok(daemonFormIndex < currentTaskIndex, 'daemon settings appear before current task');
  assert.doesNotMatch(html, /<section id="settingsPanel"/);
});
test('Word task pane announces session only after successful register response', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const registerBody = functionBody(js, 'register');
  const responseBody = functionBody(js, 'handleRegisterResponse');
  const announceBody = functionBody(js, 'announceSession');

  assert.doesNotMatch(registerBody, /announceSession\(/);
  assert.match(responseBody, /serverInfo = registerResult\(message, PROTOCOL_VERSION\)/);
  assert.match(responseBody, /runtimeInstanceId = serverInfo\.assignedInstanceId \|\| instanceId/);
  assert.match(responseBody, /announceSession\(\)\.catch/);
  assert.match(responseBody, /session\.announce\.failed/);
  assert.match(announceBody, /instance_id: runtimeInstanceId/);
  assert.doesNotMatch(announceBody, /instance_id: instanceId/);
});

function functionBody(source, name) {
  let start = source.indexOf(`function ${name}(`);
  if (start === -1) start = source.indexOf(`async function ${name}(`);
  assert.notEqual(start, -1, `missing function ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(open + 1, index);
  }
  assert.fail(`unterminated function ${name}`);
}

function cssRule(source, selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = source.match(new RegExp(`${escaped}\\s*\\{([\\s\\S]*?)\\}`));
  return match?.[1] || '';
}

function pngDimensions(png) {
  return [png.readUInt32BE(16), png.readUInt32BE(20)];
}
