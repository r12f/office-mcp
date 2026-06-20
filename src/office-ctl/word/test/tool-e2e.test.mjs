import test from 'node:test';
import {
  assertE2eCaseCoverage,
  e2eCase,
  officeE2eEnabled,
  requireOfficeE2eDriver,
  runOfficeToolE2e
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();
const PNG_1X1_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';

const WORD_E2E_CASES = Object.fromEntries([
  ['word.get_text', { verify: 'direct-result' }],
  ['word.get_outline', { verify: 'direct-result' }],
  ['word.get_paragraph', { verify: 'direct-result' }],
  ['word.find_text', { verify: 'direct-result' }],
  ['word.get_selection', { verify: 'direct-result' }],
  ['word.insert_paragraph', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'insert anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'insert anchor marker' }, text: 'E2E paragraph' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_text',
      readbackArguments: { limit: 20 },
      expect: { contains: ['insert anchor marker', 'E2E paragraph'] }
    }
  }],
  ['word.insert_image', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'image anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'image anchor marker' }, image: { base64: PNG_1X1_BASE64 }, alt_text: 'E2E image', width_pt: 24, height_pt: 24 },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'inserted', value: true },
          { path: 'mime_type', value: 'image/png' }
        ]
      }
    }
  }],
  ['word.insert_table', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'table anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'table anchor marker' }, rows: [['E2E-A', 'E2E-B']] },
    verify: {
      kind: 'readback',
      readbackTool: 'word.read_table',
      readbackArguments: { table_index: 0 },
      expect: { contains: ['E2E-A', 'E2E-B'] }
    }
  }],
  ['word.insert_page_break', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'page break anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'page break anchor marker' } },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'inserted', value: true }, { path: 'break_type', value: 'page' }] }
    }
  }],
  ['word.insert_list', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'list anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'list anchor marker' }, items: ['E2E One', 'E2E Two'] },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_text',
      readbackArguments: { limit: 20 },
      expect: { contains: ['E2E One', 'E2E Two'] }
    }
  }],
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
  ['word.update_paragraph', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Paragraph before update' } }
      ]
    },
    args: { index: 0, text: 'Updated paragraph' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_text',
      readbackArguments: { limit: 20 },
      expect: { contains: ['Updated paragraph'], notContains: ['Paragraph before update'] }
    }
  }],
  ['word.delete_range', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Delete this E2E paragraph' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Delete this E2E paragraph' }, extent: 'paragraph' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_text',
      readbackArguments: { limit: 20 },
      expect: { notContains: ['Delete this E2E paragraph'] }
    }
  }],
  ['word.apply_formatting', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Format this E2E paragraph' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Format this E2E paragraph' }, formatting: { bold: true, color: '#1F4E79' } },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'formatted', value: true }] }
    }
  }],
  ['word.read_table', { verify: 'direct-result' }],
  ['word.update_table', {
    setup: {
      actions: [
        { tool: 'word.insert_table', arguments: { anchor: { kind: 'end_of_document' }, rows: [['Old', 'Value']] } }
      ]
    },
    args: { table_index: 0, action: 'update_cell', row: 0, col: 0, text: 'Updated table cell' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.read_table',
      readbackArguments: { table_index: 0 },
      expect: { contains: ['Updated table cell'], notContains: ['Old'] }
    }
  }],
  ['word.list_content_controls', { verify: 'direct-result' }],
  ['word.insert_content_control', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'content control anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'content control anchor marker' }, text: 'E2E controlled text', tag: 'e2e-insert-control', title: 'E2E Insert Control' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_content_controls',
      readbackArguments: { tag: 'e2e-insert-control' },
      expect: { contains: ['e2e-insert-control', 'E2E Insert Control'] }
    }
  }],
  ['word.update_content_control', {
    setup: {
      actions: [
        { tool: 'word.insert_content_control', saveAs: 'controlResult', arguments: { anchor: { kind: 'end_of_document' }, text: 'Before control update', tag: 'e2e-update-control', title: 'Before Control Update' } }
      ]
    },
    args: { content_control_id: '${controlResult.content_control.content_control_id}', text: 'Updated control', tag: 'e2e-updated-control', title: 'Updated Control' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_content_controls',
      readbackArguments: { tag: 'e2e-updated-control' },
      expect: { contains: ['e2e-updated-control', 'Updated Control'], notContains: ['Before Control Update'] }
    }
  }],
  ['word.delete_content_control', {
    setup: {
      actions: [
        { tool: 'word.insert_content_control', saveAs: 'controlResult', arguments: { anchor: { kind: 'end_of_document' }, text: 'Delete control payload', tag: 'e2e-delete-control', title: 'Delete Control' } }
      ]
    },
    args: { content_control_id: '${controlResult.content_control.content_control_id}', mode: 'keep_content' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_content_controls',
      readbackArguments: { tag: 'e2e-delete-control' },
      expect: { notContains: ['e2e-delete-control', 'Delete Control'] }
    }
  }],
  ['word.apply_style', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'E2E Styled Heading' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'E2E Styled Heading' }, heading_level: 1 },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_outline',
      readbackArguments: {},
      expect: { contains: ['E2E Styled Heading'], pathEquals: [{ path: 'headings.0.level', value: 1 }] }
    }
  }],
  ['word.add_comment', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Comment target paragraph' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Comment target paragraph' }, text: 'E2E comment' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'content', value: 'E2E comment' }, { path: 'resolved', value: false }] }
    }
  }],
  ['word.resolve_comment', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Resolve comment target' } },
        { tool: 'word.add_comment', saveAs: 'commentResult', arguments: { anchor: { kind: 'after_text', text: 'Resolve comment target' }, text: 'Resolve me E2E' } }
      ]
    },
    args: { comment_id: '${commentResult.comment_id}' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'resolved', value: true }] }
    }
  }],
  ['word.update_tracked_change', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Tracked change E2E paragraph' } },
        { resource: 'office://word/${session_id}/track_changes', saveAs: 'trackChanges' }
      ]
    },
    args: { change_index: '${trackChanges.changes.0.index}', action: 'accept', expected_fingerprint: '${trackChanges.changes.0.fingerprint}' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'action', value: 'accept' }] }
    }
  }],
  ['word.save', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Save dirty marker' } }
      ]
    },
    args: {},
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'saved', value: true }] }
    }
  }]
].map(([tool, options]) => [tool, e2eCase(tool, options)]));

test('Word E2E case table covers every advertised tool', () => {
  assertE2eCaseCoverage({ addinRoot: ADDIN_ROOT, host: 'Word', cases: WORD_E2E_CASES });
});

test('Word mutating E2E cases define concrete setup and readback checks', () => {
  assertConcreteReadback('word.replace_text');
  assertConcreteReadback('word.insert_paragraph');
  assertConcreteReadback('word.update_paragraph');
  assertConcreteReadback('word.insert_table');
  assertConcreteReadback('word.insert_list');
  assertConcreteReadback('word.update_table');
  assertConcreteReadback('word.delete_range');
  assertConcreteReadback('word.insert_content_control');
  assertConcreteReadback('word.update_content_control');
  assertConcreteReadback('word.delete_content_control');
  assertConcreteReadback('word.apply_style');
  assertDirectResult('word.insert_image');
  assertDirectResult('word.insert_page_break');
  assertDirectResult('word.apply_formatting');
  assertDirectResult('word.add_comment');
  assertDirectResult('word.resolve_comment');
  assertDirectResult('word.update_tracked_change');
  assertDirectResult('word.save');
});

test('Word Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'Word',
    cases: WORD_E2E_CASES,
    driver: requireOfficeE2eDriver('Word')
  });
});

function assertConcreteReadback(tool) {
  const toolCase = WORD_E2E_CASES[tool];
  if (!toolCase.setup?.actions?.length) throw new Error(`${tool} must define setup actions`);
  if (toolCase.verify?.kind !== 'readback') throw new Error(`${tool} must use readback verification`);
  if (!toolCase.verify.readbackTool) throw new Error(`${tool} must define a readback tool`);
  if (!toolCase.verify.expect) throw new Error(`${tool} must define readback expectations`);
}

function assertDirectResult(tool) {
  const toolCase = WORD_E2E_CASES[tool];
  if (!toolCase.setup?.actions?.length) throw new Error(`${tool} must define setup actions`);
  if (toolCase.verify?.kind !== 'direct-result') throw new Error(`${tool} must use direct-result verification`);
  if (!toolCase.verify.expect) throw new Error(`${tool} must define direct-result expectations`);
}
