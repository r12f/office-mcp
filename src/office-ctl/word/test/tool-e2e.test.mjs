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
  ['word.insert_bookmark', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Bookmark insert E2E paragraph' } }
      ]
    },
    args: { name: 'E2E_Insert_Bookmark', anchor: { kind: 'after_text', text: 'Bookmark insert E2E paragraph' } },
    verify: {
      kind: 'readback',
      readbackTool: 'word.resolve_anchor',
      readbackArguments: { anchor: { kind: 'bookmark', name: 'E2E_Insert_Bookmark' } },
      expect: { contains: ['Bookmark insert E2E paragraph'], pathEquals: [{ path: 'resolved', value: true }] }
    }
  }],
  ['word.list_bookmarks', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Bookmark list E2E paragraph' } },
        { tool: 'word.insert_bookmark', arguments: { name: 'E2E_List_Bookmark', anchor: { kind: 'after_text', text: 'Bookmark list E2E paragraph' } } }
      ]
    },
    args: {},
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E_List_Bookmark', 'Bookmark list E2E paragraph'] }
    }
  }],
  ['word.delete_bookmark', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Bookmark delete E2E paragraph' } },
        { tool: 'word.insert_bookmark', arguments: { name: 'E2E_Delete_Bookmark', anchor: { kind: 'after_text', text: 'Bookmark delete E2E paragraph' } } }
      ]
    },
    args: { name: 'E2E_Delete_Bookmark' },
    verify: wordReadback.bookmarks({ notContains: ['E2E_Delete_Bookmark'] })
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
  ['word.set_selection', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Selection navigation E2E target' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Selection navigation E2E target' }, mode: 'select' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_selection',
      readbackArguments: {},
      expect: {
        contains: ['Selection navigation E2E target'],
        pathEquals: [{ path: 'is_empty', value: false }]
      }
    }
  }],
  ['word.get_html', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'HTML read E2E paragraph' } }
      ]
    },
    args: {},
    verify: {
      kind: 'direct-result',
      expect: { contains: ['HTML read E2E paragraph'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['word.insert_html', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'HTML insert anchor marker' } }
      ]
    },
    args: {
      anchor: { kind: 'after_text', text: 'HTML insert anchor marker' },
      html: '<h2>HTML Interchange E2E</h2><p><strong>Bold HTML</strong> <a href="https://example.com">link</a></p><ul><li>List item</li></ul><table><tr><td>Cell A</td></tr></table>'
    },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_html',
      readbackArguments: {},
      expect: { contains: ['HTML Interchange E2E', 'Bold HTML', 'List item', 'Cell A'] }
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
  ['word.list_images', {
    setup: {
      actions: [
        { tool: 'word.insert_image', arguments: { anchor: { kind: 'start_of_document' }, image: { base64: PNG_1X1_BASE64 }, alt_text: 'List image E2E', width_pt: 24, height_pt: 24 } }
      ]
    },
    args: {},
    verify: {
      kind: 'direct-result',
      expect: {
        contains: ['List image E2E'],
        pathEquals: [
          { path: 'count', value: 1 },
          { path: 'images.0.paragraph_index', value: 0 },
          { path: 'images.0.image_index', value: 0 }
        ]
      }
    }
  }],
  ['word.get_image', {
    setup: {
      actions: [
        { tool: 'word.insert_image', arguments: { anchor: { kind: 'start_of_document' }, image: { base64: PNG_1X1_BASE64 }, alt_text: 'Get image E2E', width_pt: 24, height_pt: 24 } }
      ]
    },
    args: { image: { kind: 'paragraph_index', index: 0, image_index: 0 } },
    verify: {
      kind: 'direct-result',
      expect: {
        contains: ['Get image E2E'],
        pathEquals: [
          { path: 'paragraph_index', value: 0 },
          { path: 'image_index', value: 0 }
        ]
      }
    }
  }],
  ['word.update_image', {
    setup: {
      actions: [
        { tool: 'word.insert_image', arguments: { anchor: { kind: 'start_of_document' }, image: { base64: PNG_1X1_BASE64 }, alt_text: 'Old image E2E', width_pt: 24, height_pt: 24 } }
      ]
    },
    args: { image: { kind: 'paragraph_index', index: 0, image_index: 0 }, alt_text_description: 'Updated image E2E' },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'updated', value: true },
          { path: 'replaced', value: false },
          { path: 'image.paragraph_index', value: 0 },
          { path: 'image.image_index', value: 0 }
        ]
      }
    }
  }],
  ['word.delete_image', {
    setup: {
      actions: [
        { tool: 'word.insert_image', arguments: { anchor: { kind: 'start_of_document' }, image: { base64: PNG_1X1_BASE64 }, alt_text: 'Delete image E2E', width_pt: 24, height_pt: 24 } }
      ]
    },
    args: { image: { kind: 'paragraph_index', index: 0, image_index: 0 } },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'deleted', value: true },
          { path: 'image.paragraph_index', value: 0 },
          { path: 'image.image_index', value: 0 }
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
  ['word.insert_field', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'start_of_document' }, text: 'E2E Field Heading', heading_level: 1 } }
      ]
    },
    args: { anchor: { kind: 'start_of_document' }, field_type: 'toc' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_fields',
      readbackArguments: { limit: 10 },
      expect: { contains: ['toc'], pathEquals: [{ path: 'count', value: 1 }, { path: 'fields.0.index', value: 0 }] }
    }
  }],
  ['word.list_fields', {
    setup: {
      actions: [
        { tool: 'word.insert_field', arguments: { anchor: { kind: 'start_of_document' }, field_type: 'date' } }
      ]
    },
    args: { offset: 0, limit: 10 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['date'], pathEquals: [{ path: 'count', value: 1 }, { path: 'fields.0.index', value: 0 }] }
    }
  }],
  ['word.update_field', {
    setup: {
      actions: [
        { tool: 'word.insert_field', arguments: { anchor: { kind: 'start_of_document' }, field_type: 'date' } },
        { tool: 'word.list_fields', saveAs: 'fieldList', arguments: { limit: 10 } }
      ]
    },
    args: { action: 'refresh_all', expected_count: '${fieldList.count}' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'updated', value: true }, { path: 'action', value: 'refresh_all' }] }
    }
  }],
  ['word.delete_field', {
    setup: {
      actions: [
        { tool: 'word.insert_field', arguments: { anchor: { kind: 'start_of_document' }, field_type: 'date' } }
      ]
    },
    args: { field_index: 0 },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_fields',
      readbackArguments: { limit: 10 },
      expect: { pathEquals: [{ path: 'count', value: 0 }] }
    }
  }],
  ['word.list_styles', {
    setup: {
      actions: [
        { tool: 'word.create_style', arguments: { name: 'E2E ListStyles Custom', type: 'paragraph', font: { bold: true } } }
      ]
    },
    args: { type: 'paragraph' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E ListStyles Custom'] }
    }
  }],
  ['word.create_style', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Create style setup marker' } }
      ]
    },
    args: { name: 'E2E Created Style', type: 'paragraph', font: { bold: true, color: '#1F4E79' }, paragraph: { alignment: 'center' } },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_styles',
      readbackArguments: { type: 'paragraph' },
      expect: { contains: ['E2E Created Style'] }
    }
  }],
  ['word.update_style', {
    setup: {
      actions: [
        { tool: 'word.create_style', arguments: { name: 'E2E Updated Style', type: 'paragraph', font: { italic: true } } },
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Updated style paragraph' } },
        { tool: 'word.apply_style', arguments: { anchor: { kind: 'after_text', text: 'Updated style paragraph' }, style: 'E2E Updated Style' } }
      ]
    },
    args: { name: 'E2E Updated Style', font: { bold: true }, paragraph: { alignment: 'center' } },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_styles',
      readbackArguments: { type: 'paragraph' },
      expect: { contains: ['E2E Updated Style'] }
    }
  }],
  ['word.update_document_properties', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Document properties setup marker' } }
      ]
    },
    args: { title: 'E2E Document Properties Title', author: 'Office MCP E2E', custom_set: [{ key: 'OfficeMcpE2E', value: 'properties' }] },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_document_properties',
      readbackArguments: { include_custom: true },
      expect: { contains: ['E2E Document Properties Title', 'OfficeMcpE2E'], pathEquals: [{ path: 'title', value: 'E2E Document Properties Title' }, { path: 'author', value: 'Office MCP E2E' }] }
    }
  }],
  ['word.get_document_properties', {
    setup: {
      actions: [
        { tool: 'word.update_document_properties', arguments: { subject: 'E2E Properties Subject', custom_set: [{ key: 'OfficeMcpReadback', value: true }] } }
      ]
    },
    args: { include_custom: true },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E Properties Subject', 'OfficeMcpReadback'], pathEquals: [{ path: 'subject', value: 'E2E Properties Subject' }] }
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
  ['word.insert_hyperlink', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'hyperlink anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'hyperlink anchor marker' }, text: 'OpenAI', url: 'https://openai.com' },
    verify: wordReadback.hyperlinks({ contains: ['OpenAI', 'https://openai.com'], pathEquals: [{ path: 'count', value: 1 }] })
  }],
  ['word.list_hyperlinks', {
    setup: {
      actions: [
        { tool: 'word.insert_hyperlink', arguments: { anchor: { kind: 'end_of_document' }, text: 'List Link', url: 'mailto:test@example.com' } }
      ]
    },
    args: { offset: 0, limit: 10 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['List Link', 'mailto:test@example.com'], pathEquals: [{ path: 'count', value: 1 }, { path: 'hyperlinks.0.occurrence_in_paragraph', value: 0 }] }
    }
  }],
  ['word.remove_hyperlink', {
    setup: {
      actions: [
        { tool: 'word.insert_hyperlink', arguments: { anchor: { kind: 'end_of_document' }, text: 'Remove Link', url: 'https://example.com/remove' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Remove Link' }, keep_text: true },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_hyperlinks',
      readbackArguments: {},
      expect: { notContains: ['https://example.com/remove'] }
    }
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
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'start_of_document' }, text: 'Format this E2E paragraph' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'Format this E2E paragraph' }, paragraph: { alignment: 'center', left_indent_pt: 18, space_after_pt: 6 } },
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_paragraph',
      readbackArguments: { index: 0, include_formatting: true },
      expect: { contains: ['Format this E2E paragraph'], pathEquals: [{ path: 'formatting.alignment', value: 'center' }, { path: 'formatting.left_indent_pt', value: 18 }, { path: 'formatting.space_after_pt', value: 6 }] }
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
        { tool: 'word.insert_table', saveAs: 'table', arguments: { anchor: { kind: 'end_of_document' }, rows: 3, cols: 3, data: [['Old', 'Value', 'Delete Col'], ['Delete Row', 'Delete Row B', 'Delete Row C'], ['Keep Row', 'Keep Cell', 'Keep Delete Col']] } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'update_cell', row: 0, col: 0, text: 'Updated table cell' } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'delete_row', row: 1 } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'delete_column', col: 2 } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'set_header_row', header_row: true } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'set_column_width', col: 0, width_pt: 72 } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'set_borders', borders: { edges: ['all'], style: 'single', width_pt: 1, color: '#336699' } } },
        { tool: 'word.update_table', arguments: { table_index: '${table.table_index}', action: 'distribute_columns' } }
      ]
    },
    args: { table_index: '${table.table_index}', action: 'merge_cells', row_range: [0, 0], col_range: [0, 1] },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['merge_cells'], pathEquals: [{ path: 'merged', value: true }] }
    }
  }],
  ['word.list_content_controls', {
    setup: {
      actions: [
        { tool: 'word.insert_content_control', arguments: { anchor: { kind: 'end_of_document' }, type: 'checkbox', checked: false, tag: 'e2e-list-control', title: 'E2E List Control' } }
      ]
    },
    args: { tag: 'e2e-list-control' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['e2e-list-control', 'E2E List Control'], pathEquals: [{ path: 'content_controls.0.checked', value: false }] }
    }
  }],
  ['word.insert_content_control', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'content control anchor marker' } }
      ]
    },
    args: { anchor: { kind: 'after_text', text: 'content control anchor marker' }, type: 'checkbox', checked: false, tag: 'e2e-insert-control', title: 'E2E Insert Control' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_content_controls',
      readbackArguments: { tag: 'e2e-insert-control' },
      expect: { contains: ['e2e-insert-control', 'E2E Insert Control'], pathEquals: [{ path: 'content_controls.0.checked', value: false }] }
    }
  }],
  ['word.update_content_control', {
    setup: {
      actions: [
        { tool: 'word.insert_content_control', saveAs: 'controlResult', arguments: { anchor: { kind: 'end_of_document' }, type: 'dropdown_list', list_items: [{ display_text: 'Draft', value: 'draft' }, { display_text: 'Approved', value: 'approved' }, { display_text: 'Rejected', value: 'rejected' }], tag: 'e2e-update-control', title: 'Before Control Update' } }
      ]
    },
    args: { content_control_id: '${controlResult.content_control.content_control_id}', selected_value: 'approved', list_items_add: [{ display_text: 'Archived', value: 'archived' }], list_items_delete: ['rejected'], tag: 'e2e-updated-control', title: 'Updated Control' },
    verify: {
      kind: 'readback',
      readbackTool: 'word.list_content_controls',
      readbackArguments: { tag: 'e2e-updated-control' },
      expect: { contains: ['e2e-updated-control', 'Updated Control', 'Approved', 'Archived'], notContains: ['Before Control Update', 'Rejected'] }
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
  ['word.insert_note', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Footnote insert target' } }
      ]
    },
    args: { kind: 'footnote', anchor: { kind: 'after_text', text: 'Footnote insert target' }, text: 'Inserted footnote body' },
    verify: wordReadback.notes('footnote', { contains: ['Inserted footnote body'], pathEquals: [{ path: 'count', value: 1 }] })
  }],
  ['word.list_notes', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Endnote list target' } },
        { tool: 'word.insert_note', arguments: { kind: 'endnote', anchor: { kind: 'after_text', text: 'Endnote list target' }, text: 'Listed endnote body' } }
      ]
    },
    args: { kind: 'endnote', offset: 0, limit: 10 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['Listed endnote body'], pathEquals: [{ path: 'count', value: 1 }, { path: 'notes.0.kind', value: 'endnote' }, { path: 'notes.0.index', value: 0 }] }
    }
  }],
  ['word.update_note', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Footnote update target' } },
        { tool: 'word.insert_note', arguments: { kind: 'footnote', anchor: { kind: 'after_text', text: 'Footnote update target' }, text: 'Before note update' } }
      ]
    },
    args: { kind: 'footnote', index: 0, text: 'Updated footnote body' },
    verify: wordReadback.notes('footnote', { contains: ['Updated footnote body'], notContains: ['Before note update'] })
  }],
  ['word.delete_note', {
    setup: {
      actions: [
        { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'Endnote delete target' } },
        { tool: 'word.insert_note', arguments: { kind: 'endnote', anchor: { kind: 'after_text', text: 'Endnote delete target' }, text: 'Delete endnote body' } }
      ]
    },
    args: { kind: 'endnote', index: 0 },
    verify: wordReadback.notes('endnote', { notContains: ['Delete endnote body'], pathEquals: [{ path: 'count', value: 0 }] })
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
  ['word.set_change_tracking', {
    setup: {
      actions: [
        { tool: 'word.get_selection', arguments: {} }
      ]
    },
    args: { mode: 'track_all' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'mode', value: 'track_all' }] }
    },
    reset: {
      actions: [
        { tool: 'word.set_change_tracking', arguments: { mode: 'off' }, allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'] }
      ]
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
