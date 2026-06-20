import test from 'node:test';
import {
  assertE2eCaseCoverage,
  e2eCase,
  realOfficeE2eEnabled,
  requireRealOfficeE2eDriver
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();

const POWERPOINT_E2E_CASES = Object.fromEntries([
  ['powerpoint.get_presentation_info', { verify: 'direct-result' }],
  ['powerpoint.get_active_view', { verify: 'direct-result' }],
  ['powerpoint.export_file', { verify: 'direct-result' }],
  ['powerpoint.update_tags', { args: { action: 'set', key: 'e2e', value: 'true' } }],
  ['powerpoint.list_slides', { verify: 'direct-result' }],
  ['powerpoint.add_slide', { args: { layout: 'Title and Content' } }],
  ['powerpoint.update_slide', { args: { slide_index: 0, title: 'Updated slide' } }],
  ['powerpoint.delete_slide', { args: { slide_index: 1 } }],
  ['powerpoint.move_slide', { args: { slide_index: 1, target_index: 0 } }],
  ['powerpoint.export_slide', { verify: 'direct-result' }],
  ['powerpoint.list_layouts', { verify: 'direct-result' }],
  ['powerpoint.apply_layout', { args: { slide_index: 0, layout: 'Title Slide' } }],
  ['powerpoint.get_selection', { verify: 'direct-result' }],
  ['powerpoint.set_selection', { args: { slide_index: 0 } }],
  ['powerpoint.list_shapes', { verify: 'direct-result' }],
  ['powerpoint.add_text_box', { args: { slide_index: 0, text: 'E2E text box' } }],
  ['powerpoint.add_shape', { args: { slide_index: 0, shape_type: 'rectangle' } }],
  ['powerpoint.insert_image', { args: { slide_index: 0, image: { base64: 'fixture' } } }],
  ['powerpoint.update_shape', { args: { slide_index: 0, shape_id: 'fixture', text: 'Updated shape' } }],
  ['powerpoint.read_text', { verify: 'direct-result' }],
  ['powerpoint.replace_text', { args: { find: 'baseline', replace: 'updated' } }],
  ['powerpoint.format_text', { args: { slide_index: 0, shape_id: 'fixture', formatting: { bold: true } } }],
  ['powerpoint.add_table', { args: { slide_index: 0, rows: [['A', 'B']] } }],
  ['powerpoint.read_table', { verify: 'direct-result' }],
  ['powerpoint.update_table', { args: { slide_index: 0, table_id: 'fixture', row: 0, column: 0, text: 'Updated' } }]
].map(([tool, options]) => [tool, e2eCase(tool, options)]));

test('PowerPoint E2E case table covers every advertised tool', () => {
  assertE2eCaseCoverage({ addinRoot: ADDIN_ROOT, host: 'PowerPoint', cases: POWERPOINT_E2E_CASES });
});

test('PowerPoint real Office E2E driver', { skip: !realOfficeE2eEnabled() }, async () => {
  requireRealOfficeE2eDriver('PowerPoint');
});
