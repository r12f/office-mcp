import { execFileSync } from 'node:child_process';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const endpoint = process.argv[2] ?? 'http://127.0.0.1:8800/mcp';
const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '../../../../..');
let client = new Client({ name: 'office-mcp-smoke', version: '0.1.0' });
let transport = new StreamableHTTPClientTransport(new URL(endpoint));
const ONE_BY_ONE_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';

const mode = process.argv[3] ?? 'tools';
if (mode === 'stdio-sessions') {
  client = new Client({ name: 'office-mcp-stdio-smoke', version: '0.1.0' });
  const stdioTransport = new StdioClientTransport({ command: 'cargo', args: ['run', '-q', '-p', 'office-mcp-daemon', '--', 'stdio'], cwd: REPO_ROOT, stderr: 'pipe' });
  await client.connect(stdioTransport);
  const result = await client.callTool({ name: 'office.list_sessions', arguments: {} });
  console.log(JSON.stringify(result, null, 2));
  await client.close();
  process.exit(0);
}

await client.connect(transport);
if (mode === 'sessions') {
  const result = await client.callTool({ name: 'office.list_sessions', arguments: {} });
  console.log(JSON.stringify(result, null, 2));
} else if (mode === 'edit') {
  const sessionId = process.argv[4];
  if (!sessionId) throw new Error('edit mode requires a session ID argument');
  const text = `office-mcp MVP smoke ${Date.now()}`;
  const inserted = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: { session_id: sessionId, text, anchor: { kind: 'end_of_document' }, style: 'Normal' }
  });
  const read = await client.callTool({
    name: 'word.get_text',
    arguments: { session_id: sessionId, offset: 0, limit: 200 }
  });
  console.log(JSON.stringify({ text, inserted, read }, null, 2));
} else if (mode === 'word-core') {
  const sessionId = process.argv[4];
  if (!sessionId) throw new Error('word-core mode requires a session ID argument');
  const suffix = Date.now();
  const beforeStructure = resourceJson(await client.readResource({ uri: 'office://word/' + sessionId + '/structure' }));
  const tableIndex = Array.isArray(beforeStructure.tables) ? beforeStructure.tables.length : 0;
  const heading = await client.callTool({
    name: 'word.insert_heading',
    arguments: { session_id: sessionId, text: `Core Smoke ${suffix}`, level: 1, anchor: { kind: 'end_of_document' } }
  });
  const paragraph = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: { session_id: sessionId, text: `Body smoke ${suffix}`, anchor: { kind: 'end_of_document' }, style: 'Normal' }
  });
  const outline = await client.callTool({ name: 'word.get_outline', arguments: { session_id: sessionId, max_level: 3 } });
  const table = await client.callTool({
    name: 'word.insert_table',
    arguments: { session_id: sessionId, anchor: { kind: 'end_of_document' }, rows: 2, cols: 2, data: [['A', 'B'], ['C', 'D']] }
  });
  const cell = await client.callTool({ name: 'word.update_cell', arguments: { session_id: sessionId, table_index: tableIndex, row: 1, col: 1, text: `Z${suffix}` } });
  const addedRow = await client.callTool({ name: 'word.add_row', arguments: { session_id: sessionId, table_index: tableIndex, values: ['E', `F${suffix}`] } });
  const addedColumn = await client.callTool({ name: 'word.add_column', arguments: { session_id: sessionId, table_index: tableIndex, values: ['G', 'H', `I${suffix}`] } });
  const formattedCell = await client.callTool({
    name: 'word.format_cell',
    arguments: { session_id: sessionId, table_index: tableIndex, row: 0, col: 0, background_color: '#ffeecc', horizontal_alignment: 'center', vertical_alignment: 'center', padding_pt: 2, formatting: { bold: true } }
  });
  const readTable = await client.callTool({ name: 'word.read_table', arguments: { session_id: sessionId, table_index: tableIndex } });
  assertToolData('core add row smoke', addedRow, (data) => Number.isInteger(data.added_row_index));
  assertToolData('core add column smoke', addedColumn, (data) => Number.isInteger(data.added_column_index));
  assertToolData('core read table smoke', readTable, (data) => data.rows === 3 && data.cols === 3 && Array.isArray(data.data));
  console.log(JSON.stringify({ suffix, tableIndex, heading, paragraph, outline, table, cell, addedRow, addedColumn, formattedCell, readTable }, null, 2));
} else if (mode === 'word-formatting') {
  const sessionId = process.argv[4];
  if (!sessionId) throw new Error('word-formatting mode requires a session ID argument');
  const suffix = Date.now();
  const targetText = 'Formatting smoke original ' + suffix;
  const updatedText = 'Formatting smoke updated ' + suffix;
  const inserted = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: { session_id: sessionId, text: targetText, anchor: { kind: 'end_of_document' }, style: 'Normal' }
  });
  const before = toolJson(await client.callTool({ name: 'word.get_text', arguments: { session_id: sessionId, offset: 0, limit: 1000, include_metadata: true } }));
  const paragraph = before.paragraphs?.find((item: any) => item.text === targetText);
  if (!paragraph) throw new Error('Inserted formatting smoke paragraph was not found.');
  const updated = await client.callTool({ name: 'word.update_paragraph', arguments: { session_id: sessionId, index: paragraph.index, text: updatedText } });
  const appliedFormatting = await client.callTool({
    name: 'word.apply_formatting',
    arguments: { session_id: sessionId, anchor: { kind: 'after_text', text: updatedText }, extent: 'paragraph', formatting: { italic: true, highlight: '#ffff00' } }
  });
  const heading = await client.callTool({ name: 'word.set_heading_level', arguments: { session_id: sessionId, index: paragraph.index, level: 3 } });
  const styled = await client.callTool({ name: 'word.apply_style', arguments: { session_id: sessionId, anchor: { kind: 'paragraph_index', index: paragraph.index }, style: 'Normal' } });
  const pageBreak = await client.callTool({ name: 'word.insert_page_break', arguments: { session_id: sessionId, anchor: { kind: 'end_of_document' } } });
  const saved = await client.callTool({ name: 'word.save', arguments: { session_id: sessionId } });
  const after = await client.callTool({ name: 'word.get_paragraph', arguments: { session_id: sessionId, index: paragraph.index } });
  assertToolData('update paragraph smoke', after, (data) => String(data.text).replace(/\f/g, '') === updatedText && data.style === 'Normal');
  console.log(JSON.stringify({ suffix, inserted, paragraph, updated, appliedFormatting, heading, styled, pageBreak, saved, after }, null, 2));
} else if (mode === 'word-spec-args') {
  const sessionId = process.argv[4];
  if (!sessionId) throw new Error('word-spec-args mode requires a session ID argument');
  const suffix = Date.now();
  const headingText = 'Spec Args Heading ' + suffix;
  const firstTarget = 'scope target first ' + suffix;
  const secondTarget = 'scope target second ' + suffix;
  const heading = await client.callTool({
    name: 'word.insert_heading',
    arguments: { session_id: sessionId, text: headingText, level: 2, anchor: { kind: 'end_of_document' } }
  });
  const formatted = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: {
      session_id: sessionId,
      text: firstTarget,
      anchor: { kind: 'heading', text: headingText, level: 2 },
      style: 'Normal',
      formatting: { bold: true, color: '#336699' }
    }
  });
  const outside = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: { session_id: sessionId, text: secondTarget, anchor: { kind: 'end_of_document' }, style: 'Normal' }
  });
  const dryRun = await client.callTool({
    name: 'word.replace_text',
    arguments: { session_id: sessionId, find: 'scope target', replace: 'scoped replacement', dry_run: true, scope: { paragraph_range: [0, 999] }, partial_ok: true }
  });
  const replace = await client.callTool({
    name: 'word.replace_text',
    arguments: { session_id: sessionId, find: 'scope target', replace: 'scoped replacement', scope: { paragraph_range: [0, 999] }, partial_ok: true }
  });
  const info = toolJson(await client.callTool({ name: 'office.get_session_info', arguments: { session_id: sessionId } }));
  const bookmark = createComBookmark(info.document?.url, 'OfficeMcpSpecArgs' + String(suffix).slice(-8));
  const bookmarkComment = await client.callTool({
    name: 'word.add_comment',
    arguments: { session_id: sessionId, anchor: { kind: 'bookmark', name: bookmark.name }, text: 'Bookmark anchor smoke ' + suffix }
  });
  const listMarker = 'Spec list item ' + suffix;
  const list = await client.callTool({
    name: 'word.insert_list',
    arguments: { session_id: sessionId, anchor: { kind: 'end_of_document' }, items: [listMarker + ' A', listMarker + ' B'], kind: 'bulleted', level: 1 }
  });
  const sentenceMarker = 'Sentence delete target ' + suffix;
  const sentenceParagraph = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: { session_id: sessionId, text: `Keep before ${suffix}. ${sentenceMarker}. Keep after ${suffix}.`, anchor: { kind: 'end_of_document' }, style: 'Normal' }
  });
  const findWholeWord = await client.callTool({
    name: 'word.find_text',
    arguments: { session_id: sessionId, query: listMarker, whole_word: true, limit: 10 }
  });
  const metadataText = await client.callTool({
    name: 'word.get_text',
    arguments: { session_id: sessionId, offset: 0, limit: 1000, include_metadata: true }
  });
  const deletedSentence = await client.callTool({
    name: 'word.delete_range',
    arguments: { session_id: sessionId, anchor: { kind: 'after_text', text: sentenceMarker }, extent: 'sentence' }
  });
  const postDelete = await client.callTool({
    name: 'word.find_text',
    arguments: { session_id: sessionId, query: sentenceMarker, limit: 5 }
  });
  const structure = await client.readResource({ uri: 'office://word/' + sessionId + '/structure' });
  assertToolData('find whole-word list smoke', findWholeWord, (data) => data.count >= 1);
  assertToolData('metadata level smoke', metadataText, (data) => Array.isArray(data.paragraphs) && data.paragraphs.some((paragraph: any) => paragraph.text === headingText && paragraph.level === 2));
  assertToolData('sentence delete smoke', postDelete, (data) => data.count === 0);
  const structureJson = resourceJson(structure);
  if ((structureJson.lists?.filter((item: any) => String(item.text).includes(listMarker)).length ?? 0) < 2 || !structureJson.tables || !Array.isArray(structureJson.headings)) {
    throw new Error('Structure resource did not include expected headings/lists/tables shape.');
  }
  console.log(JSON.stringify({ suffix, heading, formatted, outside, dryRun, replace, bookmark, bookmarkComment, list, sentenceParagraph, findWholeWord, metadataText, deletedSentence, postDelete, structure }, null, 2));} else if (mode === 'word-review') {
  const sessionId = process.argv[4];
  if (!sessionId) throw new Error('word-review mode requires a session ID argument');
  const suffix = Date.now();
  const reviewText = 'Review smoke ' + suffix;
  const paragraph = await client.callTool({
    name: 'word.insert_paragraph',
    arguments: { session_id: sessionId, text: reviewText, anchor: { kind: 'end_of_document' }, style: 'Normal' }
  });
  const image = await client.callTool({
    name: 'word.insert_image',
    arguments: {
      session_id: sessionId,
      anchor: { kind: 'end_of_document' },
      image: { base64: ONE_BY_ONE_PNG_BASE64 },
      alt_text: 'office-mcp smoke ' + suffix,
      width_pt: 24,
      height_pt: 24
    }
  });
  const comment = await client.callTool({
    name: 'word.add_comment',
    arguments: { session_id: sessionId, anchor: { kind: 'after_text', text: reviewText }, text: 'Comment smoke ' + suffix }
  });
  const commentData = toolJson(comment);
  const resolve = await client.callTool({
    name: 'word.resolve_comment',
    arguments: { session_id: sessionId, comment_id: commentData.comment_id }
  });
  const comments = await client.readResource({ uri: 'office://word/' + sessionId + '/comments' });
  const trackChanges = await client.readResource({ uri: 'office://word/' + sessionId + '/track_changes' });
  console.log(JSON.stringify({ suffix, paragraph, image, comment, resolve, comments, trackChanges }, null, 2));
} else if (mode === 'word-track-change') {
  const sessionId = process.argv[4];
  const action = parseTrackChangeAction(process.argv[5] ?? 'accept');
  if (!sessionId) throw new Error('word-track-change mode requires a session ID argument');
  console.log(JSON.stringify(await mutateFirstTrackedChange(sessionId, action), null, 2));
} else if (mode === 'word-track-change-com') {
  const sessionId = process.argv[4];
  const action = parseTrackChangeAction(process.argv[5] ?? 'accept');
  if (!sessionId) throw new Error('word-track-change-com mode requires a session ID argument');
  const info = toolJson(await client.callTool({ name: 'office.get_session_info', arguments: { session_id: sessionId } }));
  const documentUrl = info.document?.url;
  if (!documentUrl) throw new Error('Selected session does not expose a local document URL for COM smoke.');
  const com = createComTrackedChange(documentUrl, action);
  const mutation = await mutateFirstTrackedChange(sessionId, action);
  if ((mutation as { skipped?: boolean }).skipped) {
    throw new Error(`COM tracked-change smoke created a revision but MCP mutation skipped it: ${JSON.stringify({ com, mutation })}`);
  }
  console.log(JSON.stringify({ com, mutation }, null, 2));
} else if (mode === 'word-resources') {
  const sessionId = process.argv[4];
  if (!sessionId) throw new Error('word-resources mode requires a session ID argument');
  const document = await client.readResource({ uri: 'office://word/' + sessionId + '/document?offset=0&limit=20' });
  const structure = await client.readResource({ uri: 'office://word/' + sessionId + '/structure' });
  const paragraph = await client.readResource({ uri: 'office://word/' + sessionId + '/paragraph/0' });
  const comments = await client.readResource({ uri: 'office://word/' + sessionId + '/comments' });
  const trackChanges = await client.readResource({ uri: 'office://word/' + sessionId + '/track_changes' });
  const selection = await client.readResource({ uri: 'office://word/' + sessionId + '/selection' });
  console.log(JSON.stringify({ document, structure, paragraph, comments, trackChanges, selection }, null, 2));
} else {
  const tools = await client.listTools();
  console.log(JSON.stringify({ tools: tools.tools.map((tool) => tool.name).sort() }, null, 2));
}
await client.close();




type TrackChangeAction = 'accept' | 'reject';

function parseTrackChangeAction(action: string): TrackChangeAction {
  if (action !== 'accept' && action !== 'reject') throw new Error('tracked change action must be accept or reject');
  return action;
}

async function mutateFirstTrackedChange(sessionId: string, action: TrackChangeAction): Promise<unknown> {
  const resource = await client.readResource({ uri: 'office://word/' + sessionId + '/track_changes' });
  const text = resource.contents[0] && 'text' in resource.contents[0] ? resource.contents[0].text : '{}';
  const data = JSON.parse(text);
  const change = data.changes?.[0];
  if (!change) return { skipped: true, reason: 'No tracked changes are present in the document.', trackChanges: data };
  const result = await client.callTool({
    name: action === 'accept' ? 'word.accept_change' : 'word.reject_change',
    arguments: { session_id: sessionId, change_index: change.index, expected_fingerprint: change.fingerprint }
  });
  return { skipped: false, action, before: change, result };
}

function createComBookmark(documentUrl: string | undefined, name: string): { document: string; name: string; text: string } {
  if (!documentUrl) throw new Error('Selected session does not expose a local document URL for COM bookmark smoke.');
  const text = 'Bookmark anchor target ' + Date.now();
  const script = [
    '$target = ' + psSingleQuote(documentUrl),
    '$name = ' + psSingleQuote(name),
    '$text = ' + psSingleQuote(text),
    '$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")',
    '$doc = $null',
    'foreach ($candidate in $word.Documents) { if ($candidate.FullName -eq $target) { $doc = $candidate; break } }',
    'if ($null -eq $doc) { throw "Target Word document not open: $target" }',
    '$doc.Activate()',
    '$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)',
    '$range.InsertAfter([Environment]::NewLine + $text)',
    '$bookmarkRange = $doc.Range($range.Start + 1, $range.Start + 1 + $text.Length)',
    '$doc.Bookmarks.Add($name, $bookmarkRange) | Out-Null',
    '[pscustomobject]@{ document = $doc.FullName; name = $name; text = $text } | ConvertTo-Json -Depth 4'
  ].join('; ');
  const output = execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], { encoding: 'utf8' });
  return JSON.parse(output);
}
function createComTrackedChange(documentUrl: string, action: TrackChangeAction): unknown {
  const marker = `${action === 'accept' ? 'Accept' : 'Reject'} COM smoke ${Date.now()}`;
  const script = [
    '$target = ' + psSingleQuote(documentUrl),
    '$marker = ' + psSingleQuote(marker),
    '$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")',
    '$doc = $null',
    'foreach ($candidate in $word.Documents) { if ($candidate.FullName -eq $target) { $doc = $candidate; break } }',
    'if ($null -eq $doc) { throw "Target Word document not open: $target" }',
    '$doc.Activate()',
    '$doc.TrackRevisions = $true',
    '$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)',
    '$range.InsertAfter([Environment]::NewLine + $marker)',
    '$doc.TrackRevisions = $false',
    '[pscustomobject]@{ document = $doc.FullName; marker = $marker; revisions = $doc.Revisions.Count } | ConvertTo-Json -Depth 4'
  ].join('; ');
  const output = execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], { encoding: 'utf8' });
  return JSON.parse(output);
}

function psSingleQuote(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

function toolJson(result: unknown): any {
  const content = (result as { content?: Array<{ text?: string }> }).content;
  if (!content?.[0]?.text) throw new Error('Tool result did not include JSON text content.');
  return JSON.parse(content[0].text);
}

function assertToolData(name: string, result: unknown, predicate: (data: any) => boolean): void {
  const data = toolJson(result);
  if (!predicate(data)) throw new Error(`${name} failed: ${JSON.stringify(data)}`);
}

function resourceJson(result: unknown): any {
  const content = (result as { contents?: Array<{ text?: string }> }).contents;
  if (!content?.[0]?.text) throw new Error('Resource result did not include JSON text content.');
  return JSON.parse(content[0].text);
}
