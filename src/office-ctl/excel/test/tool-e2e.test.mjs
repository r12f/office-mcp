import test from 'node:test';
import {
  assertE2eCaseCoverage,
  e2eCase,
  officeE2eEnabled,
  requireOfficeE2eDriver,
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
  ['excel.write_range', {
    setup: {
      actions: [
        { tool: 'excel.clear_range', arguments: { sheet: 'Sheet1', address: 'A1:B2' } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A1:B2', values: [['updated', 'marker'], ['3', '4']] },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B2' },
      expect: { contains: ['updated', 'marker', '3', '4'], notContains: ['baseline'] }
    }
  }],
  ['excel.clear_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B2', values: [['clear', 'me'], ['now', 'please']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A1:B2' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B2' },
      expect: { notContains: ['clear', 'please'] }
    }
  }],
  ['excel.find_replace_cells', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B1', values: [['baseline', 'marker']] } }
      ]
    },
    args: { sheet: 'Sheet1', find: 'baseline', replace: 'updated' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B1' },
      expect: { contains: ['updated', 'marker'], notContains: ['baseline'] }
    }
  }],
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

test('Excel mutating E2E cases define concrete setup and readback checks', () => {
  assertConcreteReadback('excel.write_range');
  assertConcreteReadback('excel.clear_range');
  assertConcreteReadback('excel.find_replace_cells');
});

test('Excel Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'Excel',
    cases: EXCEL_E2E_CASES,
    driver: requireOfficeE2eDriver('Excel')
  });
});

function assertConcreteReadback(tool) {
  const toolCase = EXCEL_E2E_CASES[tool];
  if (!toolCase.setup?.actions?.length) throw new Error(`${tool} must define setup actions`);
  if (toolCase.verify?.kind !== 'readback') throw new Error(`${tool} must use readback verification`);
  if (!toolCase.verify.readbackTool) throw new Error(`${tool} must define a readback tool`);
  if (!toolCase.verify.expect) throw new Error(`${tool} must define readback expectations`);
}
