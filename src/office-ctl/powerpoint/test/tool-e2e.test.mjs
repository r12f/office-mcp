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

const POWERPOINT_E2E_CASES = Object.fromEntries([
  ['powerpoint.get_presentation_info', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Presentation info baseline' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'host.app', value: 'powerpoint' }, { path: 'is_read_only', value: false }, { path: 'protection.kind', value: 'none' }] }
    }
  }],
  ['powerpoint.get_active_view', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Active view baseline' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'editable', value: true }] }
    }
  }],
  ['powerpoint.export_file', { verify: 'direct-result' }],
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
      expect: { contains: ['E2E Listed Slide'], pathEquals: [{ path: 'slides.0.slide_index', value: 0 }] }
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
      readbackTool: 'powerpoint.list_slides',
      readbackArguments: {},
      expect: { contains: ['E2E Added Slide'] }
    }
  }],
  ['powerpoint.update_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Title and Content', title: 'Slide Before Update' } }
      ]
    },
    args: { slide_index: 0, title: 'Updated slide' },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.list_slides',
      readbackArguments: {},
      expect: { contains: ['Updated slide'], notContains: ['Slide Before Update'] }
    }
  }],
  ['powerpoint.delete_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Delete slide target' } },
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Keep slide target' } }
      ]
    },
    args: { slide_index: 1 },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.list_slides',
      readbackArguments: {},
      expect: { pathEquals: [{ path: 'slides.1.slide_index', value: 1 }], pathMissing: ['slides.2'] }
    }
  }],
  ['powerpoint.move_slide', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Move source first' } },
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Move source second' } }
      ]
    },
    args: { slide_index: 1, target_index: 0 },
    verify: {
      kind: 'readback',
      readbackTool: 'powerpoint.read_text',
      readbackArguments: {},
      expect: { orderedContains: ['Move source second', 'Move source first'] }
    }
  }],
  ['powerpoint.export_slide', { verify: 'direct-result' }],
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
      kind: 'readback',
      readbackTool: 'powerpoint.list_slides',
      readbackArguments: {},
      expect: { contains: ['title'] }
    }
  }],
  ['powerpoint.get_selection', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Selection read target' } },
        { tool: 'powerpoint.set_selection', arguments: { slide_index: 1 } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'slides.0.slide_index', value: 1 }] }
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
      kind: 'readback',
      readbackTool: 'powerpoint.get_selection',
      readbackArguments: {},
      expect: { pathEquals: [{ path: 'slides.0.slide_index', value: 1 }] }
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
      kind: 'readback',
      readbackTool: 'powerpoint.list_shapes',
      readbackArguments: { slide_index: 0 },
      expect: { contains: ['rectangle'] }
    }
  }],
  ['powerpoint.insert_image', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_slide', arguments: { layout: 'Blank', title: 'Image target slide' } },
        { tool: 'powerpoint.set_selection', arguments: { slide_index: 1 } }
      ]
    },
    args: { slide_index: 1, image: { base64: PNG_1X1_BASE64 }, width: 24, height: 24 },
    verify: {
      kind: 'direct-result',
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
    args: { slide_index: 0, shape_id: '${shapeResult.shape.id}', name: 'Updated shape' },
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
      expect: { contains: ['E2E read text baseline'], pathEquals: [{ path: 'count', value: 1 }] }
    }
  }],
  ['powerpoint.replace_text', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_text_box', arguments: { slide_index: 0, text: 'baseline marker' } }
      ]
    },
    args: { find: 'baseline marker', replace: 'updated marker' },
    verify: {
      kind: 'readback',
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
      readbackTool: 'powerpoint.read_table',
      readbackArguments: { slide_index: 0, shape_id: '${result.shape_id}' },
      expect: { contains: ['A', 'B'] }
    }
  }],
  ['powerpoint.read_table', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_table', saveAs: 'readTable', arguments: { slide_index: 0, values: [['Read', 'Table']] } }
      ]
    },
    args: { slide_index: 0, shape_id: '${readTable.shape_id}' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['Read', 'Table'], pathEquals: [{ path: 'rows', value: 1 }, { path: 'columns', value: 2 }] }
    }
  }],
  ['powerpoint.update_table', {
    setup: {
      actions: [
        { tool: 'powerpoint.add_table', saveAs: 'table', arguments: { slide_index: 0, values: [['Old', 'Value']] } }
      ]
    },
    args: { slide_index: 0, shape_id: '${table.shape_id}', action: 'set_cell', row_index: 0, column_index: 0, value: 'Updated table cell' },
    verify: {
      kind: 'readback',
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
  assertDirectResult('powerpoint.get_presentation_info');
  assertDirectResult('powerpoint.get_active_view');
  assertDirectResult('powerpoint.list_slides');
  assertDirectResult('powerpoint.list_layouts');
  assertDirectResult('powerpoint.get_selection');
  assertDirectResult('powerpoint.list_shapes');
  assertDirectResult('powerpoint.read_text');
  assertDirectResult('powerpoint.read_table');
  assertConcreteReadback('powerpoint.replace_text');
  assertConcreteReadback('powerpoint.add_text_box');
  assertConcreteReadback('powerpoint.add_shape');
  assertConcreteReadback('powerpoint.add_slide');
  assertConcreteReadback('powerpoint.update_slide');
  assertConcreteReadback('powerpoint.add_table');
  assertConcreteReadback('powerpoint.update_table');
  assertConcreteReadback('powerpoint.update_shape');
  assertConcreteReadback('powerpoint.delete_slide');
  assertConcreteReadback('powerpoint.move_slide');
  assertConcreteReadback('powerpoint.apply_layout');
  assertConcreteReadback('powerpoint.update_tags');
  assertConcreteReadback('powerpoint.format_text');
  assertConcreteReadback('powerpoint.set_selection');
  assertDirectResult('powerpoint.insert_image');
});

test('PowerPoint Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'PowerPoint',
    cases: POWERPOINT_E2E_CASES,
    driver: requireOfficeE2eDriver('PowerPoint')
  });
});

function assertConcreteReadback(tool) {
  const toolCase = POWERPOINT_E2E_CASES[tool];
  if (!toolCase.setup?.actions?.length) throw new Error(`${tool} must define setup actions`);
  if (toolCase.verify?.kind !== 'readback') throw new Error(`${tool} must use readback verification`);
  if (!toolCase.verify.readbackTool) throw new Error(`${tool} must define a readback tool`);
  if (!toolCase.verify.expect) throw new Error(`${tool} must define readback expectations`);
}

function assertDirectResult(tool) {
  const toolCase = POWERPOINT_E2E_CASES[tool];
  if (!toolCase.setup?.actions?.length) throw new Error(`${tool} must define setup actions`);
  if (toolCase.verify?.kind !== 'direct-result') throw new Error(`${tool} must use direct-result verification`);
  if (!toolCase.verify.expect) throw new Error(`${tool} must define direct-result expectations`);
}
