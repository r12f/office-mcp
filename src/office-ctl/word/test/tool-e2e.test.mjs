import test from 'node:test';
import {
  assertE2eCaseCoverage,
  e2eCase,
  officeE2eEnabled,
  requireOfficeE2eDriver,
  runOfficeToolE2e
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();

const WORD_E2E_CASES = Object.fromEntries([
  ['word.get_text', { verify: 'direct-result' }],
  ['word.get_outline', { verify: 'direct-result' }],
  ['word.get_paragraph', { verify: 'direct-result' }],
  ['word.find_text', { verify: 'direct-result' }],
  ['word.get_selection', { verify: 'direct-result' }],
  ['word.insert_paragraph', { args: { anchor: { kind: 'end_of_document' }, text: 'E2E paragraph' } }],
  ['word.insert_image', { args: { anchor: { kind: 'end_of_document' }, image: { base64: 'fixture' } } }],
  ['word.insert_table', { args: { anchor: { kind: 'end_of_document' }, rows: [['A', 'B']] } }],
  ['word.insert_page_break', { args: { anchor: { kind: 'end_of_document' } } }],
  ['word.insert_list', { args: { anchor: { kind: 'end_of_document' }, items: ['One', 'Two'] } }],
  ['word.replace_text', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'baseline marker' } }
      ]
    },
    args: { find: 'baseline marker', replace: 'updated marker' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_text',
      readbackArguments: { limit: 20 },
      expect: { contains: ['updated marker'], notContains: ['baseline marker'] }
    }
  }],
  ['word.update_paragraph', { args: { index: 0, text: 'Updated paragraph' } }],
  ['word.delete_range', { args: { anchor: { kind: 'paragraph_index', index: 0 }, extent: 'paragraph' } }],
  ['word.apply_formatting', { args: { anchor: { kind: 'paragraph_index', index: 0 }, formatting: { bold: true } } }],
  ['word.read_table', { verify: 'direct-result' }],
  ['word.update_table', { args: { table_index: 0, action: 'update_cell', row: 0, column: 0, text: 'Updated' } }],
  ['word.list_content_controls', { verify: 'direct-result' }],
  ['word.insert_content_control', { args: { anchor: { kind: 'paragraph_index', index: 0 }, tag: 'e2e' } }],
  ['word.update_content_control', { args: { tag: 'e2e', text: 'Updated control' } }],
  ['word.delete_content_control', { args: { tag: 'e2e', keep_content: true } }],
  ['word.apply_style', { args: { anchor: { kind: 'paragraph_index', index: 0 }, style: 'Heading 1' } }],
  ['word.add_comment', { args: { anchor: { kind: 'paragraph_index', index: 0 }, text: 'E2E comment' } }],
  ['word.resolve_comment', { args: { comment_index: 0 } }],
  ['word.update_tracked_change', { args: { change_index: 0, action: 'accept', expected_fingerprint: 'fixture' } }],
  ['word.save', {}]
].map(([tool, options]) => [tool, e2eCase(tool, options)]));

test('Word E2E case table covers every advertised tool', () => {
  assertE2eCaseCoverage({ addinRoot: ADDIN_ROOT, host: 'Word', cases: WORD_E2E_CASES });
});

test('Word Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'Word',
    cases: WORD_E2E_CASES,
    driver: requireOfficeE2eDriver('Word')
  });
});
