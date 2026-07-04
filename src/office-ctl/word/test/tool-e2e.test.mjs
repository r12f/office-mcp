import test from 'node:test';
import { resolve } from 'node:path';
import {
  assertConcreteE2eCases,
  assertE2eCaseCoverage,
  e2eCase,
  officeE2eEnabled,
  requireOfficeE2eDriver,
  runOfficeToolE2e,
  wordReadback
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();
const REPORT_PATH = resolve(ADDIN_ROOT, '../../../artifacts/office-tool-e2e-word.json');
const PNG_1X1_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';

const WORD_E2E_CASES = Object.fromEntries([
  ['word.get_text', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Read text E2E paragraph' } }
      ]
    },
    args: { limit: 20 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['Read text E2E paragraph'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['word.get_outline', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'E2E Outline Heading', heading_level: 1 } }
      ]
    },
    args: { max_level: 2 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E Outline Heading'], pathEquals: [{ path: 'headings.0.level', value: 1 }] }
    }
  }],
  ['word.get_paragraph', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'start_of_document' }, text: 'Paragraph read target' } }
      ]
    },
    args: { index: 0 },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'index', value: 0 }, { path: 'text', value: 'Paragraph read target' }] }
    }
  }],
  ['word.find_text', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Find target E2E phrase' } }
      ]
    },
    args: { query: 'target E2E', limit: 10 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['target E2E'], pathEquals: [{ path: 'count', value: 1 }] }
    }
  }],
  ['word.resolve_anchor', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Resolve anchor E2E paragraph' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Resolve anchor E2E paragraph' } },
    verify: {
      kind: 'direct-result',
      expect: {
        contains: ['Resolve anchor E2E paragraph'],
        pathEquals: [
          { path: 'resolved', value: true },
          { path: 'object_type', value: 'Range' },
          { path: 'anchor_kind', value: 'after_text' },
          { path: 'untrusted_source', value: true }
        ]
      }
    }
  }],
  ['word.get_selection', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Selection structure baseline' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['word.get_header_footer', {
    setup: {
      actions: [
        { tool: 'word.update_header_footer', arguments: { location: 'header', action: 'set_text', text: 'Header read E2E' } }
      ]
    },
    args: { location: 'header', include_metadata: true },
    verify: {
      kind: 'direct-result',
      expect: {
        contains: ['Header read E2E'],
        pathEquals: [
          { path: 'location', value: 'header' },
          { path: 'header_footer_type', value: 'primary' },
          { path: 'untrusted_source', value: true }
        ]
      }
    }
  }],
  ['word.insert_paragraph', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'insert anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'insert anchor marker' }, text: 'E2E paragraph' },
    verify: wordReadback.documentText({ contains: ['insert anchor marker', 'E2E paragraph'] })
  }],
  ['word.insert_image', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'image anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_paragraph_index', index: 0 }, placement: 'new_paragraph_after', image: { base64: PNG_1X1_BASE64 }, alt_text: 'E2E image', width_pt: 24, height_pt: 24 },
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
  ['word.resize_image', {
    setup: {
      actions: [
        { tool: 'word.insert_image', arguments: { anchor: { kind: 'start_of_document' }, image: { base64: PNG_1X1_BASE64 }, alt_text: 'Resize image E2E', width_pt: 24, height_pt: 24 } }
      ]
    },
    args: { image: { kind: 'paragraph_index', index: 0, image_index: 0 }, width_pt: 48, preserve_aspect_ratio: true },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'resized', value: true },
          { path: 'image.new_width_pt', value: 48 },
          { path: 'image.new_height_pt', value: 48 },
          { path: 'image.old_width_pt', value: 24 },
          { path: 'image.old_height_pt', value: 24 }
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
    args: { anchor: { kind: 'after_text', text: 'table anchor marker' }, rows: 1, cols: 2, data: [['E2E-A', 'E2E-B']] },
    verify: wordReadback.table(0, { contains: ['E2E-A', 'E2E-B'] })
  }],
  ['word.insert_break', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'section break anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'section break anchor marker' }, break_type: 'section_next' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'inserted', value: true }, { path: 'break_type', value: 'section_next' }] }
    }
  }],
  ['word.list_sections', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'section listing marker' } },
        { tool: 'word.insert_break', arguments: { anchor: { kind: 'after_text', text: 'section listing marker' }, break_type: 'section_next' } }
      ]
    },
    args: {},
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'count', value: 2 }, { path: 'sections.0.index', value: 0 }, { path: 'sections.1.index', value: 1 }] }
    }
  }],
  ['word.update_page_setup', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'page setup anchor marker' } }
      ]
    },
    args: { orientation: 'landscape', margins_pt: { top: 72, bottom: 72, left: 54, right: 54 } },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_sections',
      readbackArguments: { include_page_setup: true },
      expect: { pathEquals: [{ path: 'sections.0.page_setup.orientation', value: 'landscape' }, { path: 'sections.0.page_setup.margins_pt.left', value: 54 }] }
    }
  }],
  ['word.update_header_footer', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'header footer update setup' } }
      ]
    },
    args: { location: 'footer', action: 'set_text', text: 'Footer update E2E' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_header_footer',
      readbackArguments: { location: 'footer' },
      expect: { contains: ['Footer update E2E'], pathEquals: [{ path: 'is_empty', value: false }] }
    }
  }],
  ['word.insert_list', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'list anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'list anchor marker' }, items: ['E2E One', 'E2E Two'] },
    verify: wordReadback.documentText({ contains: ['E2E One', 'E2E Two'] })
  }],
  ['word.replace_text', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'baseline marker' } }
      ]
    },
    args: { find: 'baseline marker', replace: 'updated marker' },
    verify: wordReadback.documentText({ contains: ['updated marker'], notContains: ['baseline marker'] })
  }],
  ['word.update_paragraph', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'start_of_document' }, text: 'Paragraph before update' } }
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
  ['word.read_table', {
    setup: {
      actions: [
        { tool: 'word.insert_table', saveAs: 'table', arguments: { anchor: { kind: 'end_of_document' }, rows: 1, cols: 2, data: [['ReadTable-A', 'ReadTable-B']] } }
      ]
    },
    args: { table_index: '${table.table_index}' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['ReadTable-A', 'ReadTable-B'], pathEquals: [{ path: 'rows', value: 1 }, { path: 'cols', value: 2 }] }
    }
  }],
  ['word.update_table', {
    setup: {
      actions: [
        { tool: 'word.insert_table', saveAs: 'table', arguments: { anchor: { kind: 'end_of_document' }, rows: 1, cols: 2, data: [['Old', 'Value']] } }
      ]
    },
    args: { table_index: '${table.table_index}', action: 'update_cell', row: 0, col: 0, text: 'Updated table cell' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.read_table',
      readbackArguments: { table_index: '${table.table_index}' },
      expect: { contains: ['Updated table cell'], notContains: ['Old'] }
    }
  }],
  ['word.list_content_controls', {
    setup: {
      actions: [
        { tool: 'word.insert_content_control', arguments: { anchor: { kind: 'end_of_document' }, text: 'List control payload', tag: 'e2e-list-control', title: 'E2E List Control' } }
      ]
    },
    args: { tag: 'e2e-list-control' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['e2e-list-control', 'E2E List Control'] }
    }
  }],
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
    verify: wordReadback.contentControls('e2e-delete-control', { notContains: ['e2e-delete-control', 'Delete Control'] })
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
    verify: wordReadback.comments({ contains: ['E2E comment'] })
  }],
  ['word.resolve_comment', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Resolve comment target' } },
        { tool: 'word.add_comment', saveAs: 'commentResult', arguments: { anchor: { kind: 'after_text', text: 'Resolve comment target' }, text: 'Resolve me E2E' } }
      ]
    },
    args: { comment_id: '${commentResult.comment_id}' },
    verify: wordReadback.comments({ contains: ['Resolve me E2E', 'true'] })
  }],
  ['word.update_tracked_change', {
    setup: {
      actions: [
        { driver: 'word.create_tracked_change', saveAs: 'trackedChangeSeed', arguments: { text: 'Tracked change E2E paragraph' } },
        { resource: 'office://word/${session_id}/track_changes', saveAs: 'trackChanges' }
      ]
    },
    args: { change_index: '${trackChanges.changes.0.index}', action: 'accept', expected_fingerprint: '${trackChanges.changes.0.fingerprint}' },
    verify: wordReadback.trackChanges({ notContains: ['Tracked change E2E paragraph'] })
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
  assertConcreteE2eCases({ host: 'Word', cases: WORD_E2E_CASES });
});

test('Word Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'Word',
    cases: WORD_E2E_CASES,
    driver: requireOfficeE2eDriver('Word'),
    reportPath: REPORT_PATH
  });
});
