import test from 'node:test';
import {
  assertE2eCaseCoverage,
  e2eCase,
  realOfficeE2eEnabled,
  requireRealOfficeE2eDriver,
  runOfficeToolE2e
} from '../../common/test/tool-e2e-contract.mjs';

const ADDIN_ROOT = process.cwd();

const EXCEL_E2E_CASES = Object.fromEntries([
  ['excel.get_workbook_info', { verify: 'direct-result' }],
  ['excel.list_sheets', { verify: 'direct-result' }],
  ['excel.add_sheet', { args: { name: 'E2E Sheet' } }],
  ['excel.update_sheet', { args: { sheet: 'Sheet1', name: 'Renamed Sheet' } }],
  ['excel.delete_sheet', { args: { sheet: 'Delete Me' } }],
  ['excel.get_used_range', { verify: 'direct-result' }],
  ['excel.read_range', { verify: 'direct-result' }],
  ['excel.write_range', { args: { sheet: 'Sheet1', address: 'A1:B2', values: [['A', 'B'], ['C', 'D']] } }],
  ['excel.clear_range', { args: { sheet: 'Sheet1', address: 'A1:B2' } }],
  ['excel.find_replace_cells', { args: { sheet: 'Sheet1', find: 'baseline', replace: 'updated' } }],
  ['excel.set_formula', { args: { sheet: 'Sheet1', address: 'C1', formula: '=SUM(A1:B1)' } }],
  ['excel.format_range', { args: { sheet: 'Sheet1', address: 'A1:B2', format: { bold: true } } }],
  ['excel.sort_range', { args: { sheet: 'Sheet1', address: 'A1:B3', key_column: 0 } }],
  ['excel.apply_filter', { args: { sheet: 'Sheet1', address: 'A1:B3', column: 0, criterion: 'A' } }],
  ['excel.create_table', { args: { sheet: 'Sheet1', address: 'A1:B3', has_headers: true } }],
  ['excel.update_table', { args: { table: 'E2ETable', action: 'rename', name: 'E2ETableRenamed' } }],
  ['excel.create_chart', { args: { sheet: 'Sheet1', source_range: 'A1:B3', chart_type: 'columnClustered' } }],
  ['excel.update_chart', { args: { chart: 'E2EChart', title: 'Updated chart' } }],
  ['excel.create_pivot_table', { args: { sheet: 'Sheet1', source_range: 'A1:B4', destination: 'D1' } }],
  ['excel.update_pivot_table', { args: { pivot_table: 'E2EPivot', action: 'refresh' } }]
].map(([tool, options]) => [tool, e2eCase(tool, options)]));

test('Excel E2E case table covers every advertised tool', () => {
  assertE2eCaseCoverage({ addinRoot: ADDIN_ROOT, host: 'Excel', cases: EXCEL_E2E_CASES });
});

test('Excel real Office E2E driver', { skip: !realOfficeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'Excel',
    cases: EXCEL_E2E_CASES,
    driver: requireRealOfficeE2eDriver('Excel')
  });
});
