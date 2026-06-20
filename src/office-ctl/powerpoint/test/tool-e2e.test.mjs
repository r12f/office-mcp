import test from 'node:test';
import {
  assertE2eCaseCoverage,
  e2eCase,
  officeE2eEnabled,
  requireOfficeE2eDriver,
  runOfficeToolE2e
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();

const POWERPOINT_E2E_CASES = Object.fromEntries([
  ['powerpoint.get_presentation_info', { verify: 'direct-result' }],
  ['powerpoint.get_active_view', { verify: 'direct-result' }],
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
  ['powerpoint.list_slides', { verify: 'direct-result' }],
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
  ['powerpoint.list_layouts', { verify: 'direct-result' }],
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
  ['powerpoint.get_selection', { verify: 'direct-result' }],
  ['powerpoint.set_selection', { args: { slide_index: 0 } }],
  ['powerpoint.list_shapes', { verify: 'direct-result' }],
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
  ['powerpoint.insert_image', { args: { slide_index: 0, image: { base64: 'fixture' } } }],
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
  ['powerpoint.read_text', { verify: 'direct-result' }],
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
  ['powerpoint.format_text', { args: { slide_index: 0, shape_id: 'fixture', formatting: { bold: true } } }],
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
  ['powerpoint.read_table', { verify: 'direct-result' }],
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
