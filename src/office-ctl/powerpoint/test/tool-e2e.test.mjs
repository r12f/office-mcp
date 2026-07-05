import test from 'node:test';
import { resolve } from 'node:path';
import {
  assertConcreteE2eCases,
  assertE2eCaseCoverage,
  e2eCase,
  officeE2eEnabled,
  requireOfficeE2eDriver,
  runOfficeToolE2e
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();
const REPORT_PATH = resolve(ADDIN_ROOT, '../../../artifacts/office-tool-e2e-powerpoint.json');
const PNG_1X1_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=';

const POWERPOINT_E2E_CASES = Object.fromEntries([
  ['powerpoint.get_presentation_info', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Presentation info baseline' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'host.app', value: 'powerpoint' },
          { path: 'is_read_only', value: false },
          { path: 'protection.kind', value: 'none' },
          { path: 'active_view_source', value: 'host' }
        ],
        pathMatches: [
          { path: 'active_view', pattern: '^(edit|read)$' }
        ]
      }
    }
  }],
  ['powerpoint.export_file', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Export file baseline' } }
      ]
    },
    args: { format: 'pdf' },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { pathEquals: [{ path: 'format', value: 'pdf' }, { path: 'mime_type', value: 'application/pdf' }] }
    }
  }],
  ['powerpoint.update_tags', {
    setup: {
      actions: [
        { tool: 'powerpoint.update_tags', arguments: { action: 'list' } }
      ]
    },
    args: { action: 'set', key: 'e2e-metadata', value: 'true' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.update_tags',
      readbackArguments: { action: 'list' },
      expect: { contains: ['e2e-metadata', 'true'] }
    }
  }],
  ['powerpoint.list_slides', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Title and Content', title: 'E2E Listed Slide' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'slides.0.slide_index', value: 0 }] }
    }
  }],
  ['powerpoint.add_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.list_slides', arguments: {} }
      ]
    },
    args: { layout: 'Title and Content', title: 'E2E Added Slide' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.read_text',
      readbackArguments: { slide_index: '${result.slide_index}' },
      expect: { contains: ['E2E Added Slide'] }
    }
  }],
  ['powerpoint.update_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Title and Content', title: 'Slide metadata target' } }
      ]
    },
    args: { slide_index: 0, action: 'set_tag', key: 'e2e-slide', value: 'updated' },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { contains: ['e2e-slide', 'updated'] }
    }
  }],
  ['powerpoint.delete_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', saveAs: 'deleteSlideTarget', arguments: { layout: 'Title and Content', title: 'Delete slide target' } },
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Title and Content', title: 'Keep slide target' } }
      ]
    },
    args: { slide_index: '${deleteSlideTarget.slide_index}' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.read_text',
      readbackArguments: {},
      expect: { contains: ['Keep slide target'], notContains: ['Delete slide target'] }
    }
  }],
  ['powerpoint.move_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', saveAs: 'moveFirst', arguments: { layout: 'Title and Content', title: 'Move source first' } },
        { tool: 'powerpoint.add_slide', saveAs: 'moveSecond', arguments: { layout: 'Title and Content', title: 'Move source second' } }
      ]
    },
    args: { slide_index: '${moveSecond.slide_index}', target_index: 0 },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'target_index', value: 0 }] }
    }
  }],
  ['powerpoint.export_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Export slide baseline' } }
      ]
    },
    args: { slide_index: 0 },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { pathEquals: [{ path: 'slide_index', value: 0 }, { path: 'mime_type', value: 'image/png' }] }
    }
  }],
  ['powerpoint.list_layouts', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Layout list baseline' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathMissing: ['masters.0.missing'] }
    }
  }],
  ['powerpoint.apply_layout', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank' } }
      ]
    },
    args: { slide_index: 0, layout_type: 'title' },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { pathEquals: [{ path: 'slide_index', value: 0 }] }
    }
  }],
  ['powerpoint.get_selection', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Selection read target' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { pathMissing: ['missing'] }
    }
  }],
  ['powerpoint.set_selection', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Selection target slide' } }
      ]
    },
    args: { slide_index: 1 },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { pathEquals: [{ path: 'selected', value: true }, { path: 'slide_index', value: 1 }] }
    }
  }],
  ['powerpoint.list_shapes', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank' } },
        { tool: 'powerpoint.add_shape', arguments: { slide_index: 0, type: 'rectangle', name: 'E2E Listed Shape' } }
      ]
    },
    args: { slide_index: 0 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E Listed Shape'], pathEquals: [{ path: 'slide_index', value: 0 }] }
    }
  }],
  ['powerpoint.add_text_box', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank' } }
      ]
    },
    args: { slide_index: 0, text: 'E2E text box' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.read_text',
      readbackArguments: { slide_index: 0 },
      expect: { contains: ['E2E text box'] }
    }
  }],
  ['powerpoint.add_shape', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank' } }
      ]
    },
    args: { slide_index: 0, type: 'rectangle' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['shape_id'] }
    }
  }],
  ['powerpoint.insert_image', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Image target slide' } }
      ]
    },
    args: { slide_index: 1, image: { base64: PNG_1X1_BASE64 }, width: 24, height: 24 },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: {
        pathEquals: [
          { path: 'inserted_image', value: true },
          { path: 'mime_type', value: 'image/png' }
        ]
      }
    }
  }],
  ['powerpoint.update_shape', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_shape', saveAs: 'shapeResult', arguments: { slide_index: 0, type: 'rectangle', name: 'Shape Before Update' } }
      ]
    },
    args: { slide_index: 0, shape_id: '${shapeResult.shape.shape_id}', name: 'Updated shape' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.list_shapes',
      readbackArguments: { slide_index: 0 },
      expect: { contains: ['Updated shape'], notContains: ['Shape Before Update'] }
    }
  }],
  ['powerpoint.read_text', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_text_box', arguments: { slide_index: 0, text: 'E2E read text baseline' } }
      ]
    },
    args: { slide_index: 0 },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E read text baseline'] }
    }
  }],
  ['powerpoint.replace_text', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_text_box', arguments: { slide_index: 0, text: 'baseline marker' } }
      ]
    },
    args: { search: 'baseline marker', replacement: 'updated marker' },
    verify: {
      kind: 'readback',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      readbackTool: 'powerpoint.read_text',
      readbackArguments: { slide_index: 0 },
      expect: { contains: ['updated marker'], notContains: ['baseline marker'] }
    }
  }],
  ['powerpoint.format_text', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_text_box', saveAs: 'textBoxResult', arguments: { slide_index: 0, text: 'Format me E2E' } }
      ]
    },
    args: { slide_index: 0, shape_id: '${textBoxResult.shape.shape_id}', bold: true, color: '#1F4E79' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.read_text',
      readbackArguments: { slide_index: 0 },
      expect: { contains: ['Format me E2E'] }
    }
  }],
  ['powerpoint.add_table', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank' } }
      ]
    },
    args: { slide_index: 0, values: [['A', 'B']] },
    verify: {
      kind: 'readback',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      readbackTool: 'powerpoint.read_table',
      readbackArguments: { slide_index: 0, shape_id: '${result.shape_id}' },
      expect: { contains: ['A', 'B'] }
    }
  }],
  ['powerpoint.read_table', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_table', saveAs: 'readTable', allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'], arguments: { slide_index: 0, values: [['Read', 'Table']] } }
      ]
    },
    args: { slide_index: 0, shape_id: '${readTable.shape_id}' },
    verify: {
      kind: 'direct-result',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      expect: { contains: ['Read', 'Table'], pathEquals: [{ path: 'rows', value: 1 }, { path: 'columns', value: 2 }] }
    }
  }],
  ['powerpoint.update_table', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_table', saveAs: 'table', allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'], arguments: { slide_index: 0, values: [['Old', 'Value']] } }
      ]
    },
    args: { slide_index: 0, shape_id: '${table.shape_id}', action: 'set_cell', row_index: 0, column_index: 0, value: 'Updated table cell' },
    verify: {
      kind: 'readback',
      allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
      readbackTool: 'powerpoint.read_table',
      readbackArguments: { slide_index: 0, shape_id: '${table.shape_id}' },
      expect: { contains: ['Updated table cell'], notContains: ['Old'] }
    }
  }]
].map(([tool, options]) => [tool, e2eCase(tool, options)]));

test('PowerPoint E2E case table covers every advertised tool', () => {
  assertE2eCaseCoverage({ addinRoot: ADDIN_ROOT, host: 'PowerPoint', cases: POWERPOINT_E2E_CASES });
});

test('PowerPoint mutating E2E cases define concrete setup and readback checks', () => {
  assertConcreteE2eCases({ host: 'PowerPoint', cases: POWERPOINT_E2E_CASES });
});

test('PowerPoint Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'PowerPoint',
    cases: POWERPOINT_E2E_CASES,
    driver: requireOfficeE2eDriver('PowerPoint'),
    reportPath: REPORT_PATH
  });
});
