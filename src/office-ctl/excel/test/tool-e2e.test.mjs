import test from 'node:test';
import assert from 'node:assert/strict';
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
const REPORT_PATH = resolve(ADDIN_ROOT, '../../../artifacts/office-tool-e2e-excel.json');

const EXCEL_E2E_CASES = Object.fromEntries([
  ['excel.get_workbook_info', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1', values: [['Workbook info marker']] } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'active_sheet.name', value: 'Sheet1' },
          { path: 'is_read_only', value: false },
          { path: 'is_protected', value: false }
        ]
      }
    }
  }],
  ['excel.save', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A2', values: [['Save marker']] } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'saved', value: true }] }
    }
  }],
  ['excel.calculate', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'B2:B3', values: [[2], [5]] } },
        { tool: 'excel.set_formula', arguments: { sheet: 'Sheet1', address: 'B4', formula: '=SUM(B2:B3)' } }
      ]
    },
    args: { type: 'recalculate' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'B4' },
      expect: { contains: ['7'] }
    }
  }],
  ['excel.list_named_items', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'Y1:Y2', values: [['Named'], ['Range']] } },
        { tool: 'excel.update_named_item', arguments: { action: 'add', name: 'E2EListNamedRange', reference: 'Sheet1!Y1:Y2', comment: 'E2E range name' } },
        { tool: 'excel.update_named_item', arguments: { action: 'add', name: 'E2EListNamedConstant', formula: '=42', comment: 'E2E constant name' } }
      ]
    },
    args: { scope: 'all' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2EListNamedRange', 'E2EListNamedConstant', 'Sheet1!Y1:Y2', '=42'] }
    }
  }],
  ['excel.update_named_item', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'Z1:Z2', values: [['Named'], ['Item']] } }
      ]
    },
    args: { action: 'add', name: 'E2ENamedItemRoundTrip', reference: 'Sheet1!Z1:Z2', comment: 'Initial name' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'E2ENamedItemRoundTrip' },
      expect: { contains: ['Named', 'Item'] }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_named_item', arguments: { action: 'edit', name: 'E2ENamedItemRoundTrip', formula: '=Sheet1!Z1:Z2', comment: 'Updated name' } },
        { tool: 'excel.update_named_item', arguments: { action: 'delete', name: 'E2ENamedItemRoundTrip' } }
      ]
    }
  }],
  ['excel.list_sheets', {
    setup: {
      actions: [
        { tool: 'excel.add_sheet', arguments: { name: 'E2E List Sheet' } }
      ]
    },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E List Sheet'], pathEquals: [{ path: 'sheets.1.name', value: 'E2E List Sheet' }] }
    }
  }],
  ['excel.add_sheet', {
    setup: {
      actions: [
        { tool: 'excel.list_sheets', arguments: {} }
      ]
    },
    args: { name: 'E2E Sheet' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_sheets',
      readbackArguments: {},
      expect: { contains: ['E2E Sheet'] }
    }
  }],
  ['excel.update_sheet', {
    setup: {
      actions: [
        { tool: 'excel.add_sheet', arguments: { name: 'Sheet To Rename' } }
      ]
    },
    args: { sheet: 'Sheet To Rename', name: 'Renamed Sheet' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_sheets',
      readbackArguments: {},
      expect: { contains: ['Renamed Sheet'], notContains: ['Sheet To Rename'] }
    }
  }],
  ['excel.delete_sheet', {
    setup: {
      actions: [
        { tool: 'excel.add_sheet', arguments: { name: 'Delete Me' } }
      ]
    },
    args: { sheet: 'Delete Me' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_sheets',
      readbackArguments: {},
      expect: { notContains: ['Delete Me'] }
    }
  }],
  ['excel.get_used_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'C3:D4', values: [['Used', 'Range'], ['E2E', 'Marker']] } }
      ]
    },
    args: { sheet: 'Sheet1' },
    verify: {
      kind: 'direct-result',
      expect: { pathEquals: [{ path: 'row_count', value: 4 }, { path: 'column_count', value: 4 }, { path: 'is_empty', value: false }] }
    }
  }],
  ['excel.read_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B2', values: [['Read', 'Range'], ['E2E', 'Marker']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A1:B2' },
    verify: {
      kind: 'direct-result',
      expect: {
        contains: ['Read', 'Range', 'E2E', 'Marker'],
        pathEquals: [
          { path: 'row_count', value: 2 },
          { path: 'column_count', value: 2 },
          { path: 'untrusted_source', value: true }
        ]
      }
    }
  }],
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
    args: { sheet: 'Sheet1', address: 'A1:B1', query: 'baseline', replacement: 'updated' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B1' },
      expect: { contains: ['updated', 'marker'], notContains: ['baseline'] }
    }
  }],
  ['excel.set_formula', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B1', values: [[2, 5]] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'C1', formula: '=SUM(A1:B1)' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'C1' },
      expect: { contains: ['7'] }
    }
  }],
  ['excel.format_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B2', values: [[12.5, 20], [3, 4]] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A1:B2', number_format: '0.00' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B2' },
      expect: { contains: ['0.00'] }
    }
  }],
  ['excel.sort_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B4', values: [['Name', 'Score'], ['Charlie', 3], ['Alpha', 1], ['Bravo', 2]] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A1:B4', fields: [{ key: 0, ascending: true }], has_headers: true },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B4' },
      expect: { orderedContains: ['Alpha', 'Bravo', 'Charlie'] }
    }
  }],
  ['excel.apply_filter', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B3', values: [['Name', 'Value'], ['Alpha', '1'], ['Beta', '2']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A1:B3', target_type: 'range', column_index: 0, criteria: { filter_on: 'values', values: ['Alpha'] } },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'target_type', value: 'range' },
          { path: 'address', value: 'A1:B3' },
          { path: 'column_index', value: 0 },
          { path: 'filtered', value: true }
        ]
      }
    }
  }],
  ['excel.create_table', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'AA1:AB3', values: [['Name', 'Value'], ['Alpha', '1'], ['Beta', '2']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'AA1:AB3', has_headers: true, name: 'E2ETable' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.update_table',
      readbackArguments: { table: 'E2ETable', action: 'metadata' },
      expect: { contains: ['E2ETable'] }
    }
  }],
  ['excel.update_table', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'D1:E3', values: [['Name', 'Value'], ['Alpha', '1'], ['Beta', '2']] } },
        { tool: 'excel.create_table', arguments: { sheet: 'Sheet1', address: 'D1:E3', has_headers: true, name: 'E2ETableToRename' } }
      ]
    },
    args: { table: 'E2ETableToRename', action: 'rename', name: 'E2ETableRenamed' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.update_table',
      readbackArguments: { table: 'E2ETableRenamed', action: 'metadata' },
      expect: { contains: ['E2ETableRenamed'], notContains: ['E2ETableToRename'] }
    }
  }],
  ['excel.create_chart', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'G1:H3', values: [['Label', 'Value'], ['Alpha', 1], ['Beta', 2]] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'G1:H3', type: 'columnClustered', title: 'E2E Chart' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.update_chart',
      readbackArguments: { sheet: 'Sheet1', chart: '${result.chart}', action: 'metadata' },
      expect: { contains: ['E2E Chart', 'ColumnClustered'] }
    }
  }],
  ['excel.update_chart', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'J1:K3', values: [['Label', 'Value'], ['Alpha', 3], ['Beta', 4]] } },
        { tool: 'excel.create_chart', saveAs: 'chartResult', arguments: { sheet: 'Sheet1', address: 'J1:K3', type: 'columnClustered', title: 'Chart Before Update' } }
      ]
    },
    args: { sheet: 'Sheet1', chart: '${chartResult.chart}', action: 'title', title: 'Updated chart' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.update_chart',
      readbackArguments: { sheet: 'Sheet1', chart: '${chartResult.chart}', action: 'metadata' },
      expect: { contains: ['Updated chart'], notContains: ['Chart Before Update'] }
    }
  }],
  ['excel.create_pivot_table', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'M1:N4', values: [['Region', 'Sales'], ['West', 10], ['East', 15], ['West', 20]] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'M1:N4', name: 'E2EPivotCreate', destination: 'P1' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.update_pivot_table',
      readbackArguments: { pivot_table: 'E2EPivotCreate', action: 'metadata' },
      expect: { contains: ['E2EPivotCreate'] }
    }
  }],
  ['excel.update_pivot_table', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'R1:S4', values: [['Region', 'Sales'], ['West', 10], ['East', 15], ['West', 20]] } },
        { tool: 'excel.create_pivot_table', arguments: { sheet: 'Sheet1', address: 'R1:S4', name: 'E2EPivotUpdate', destination: 'U1' } }
      ]
    },
    args: { pivot_table: 'E2EPivotUpdate', action: 'add_hierarchy', axis: 'row', hierarchy: 'Region' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.update_pivot_table',
      readbackArguments: { pivot_table: 'E2EPivotUpdate', action: 'metadata' },
      expect: { contains: ['E2EPivotUpdate', 'Region'] }
    }
  }]
].map(([tool, options]) => [tool, e2eCase(tool, options)]));

test('Excel E2E case table covers every advertised tool', () => {
  assertE2eCaseCoverage({ addinRoot: ADDIN_ROOT, host: 'Excel', cases: EXCEL_E2E_CASES });
});

test('Excel mutating E2E cases define concrete setup and readback checks', () => {
  assertConcreteE2eCases({ host: 'Excel', cases: EXCEL_E2E_CASES });
});

test('Excel E2E range filter does not overlap create_table fixture', () => {
  assert.notEqual(
    EXCEL_E2E_CASES['excel.apply_filter'].call.arguments.address,
    EXCEL_E2E_CASES['excel.create_table'].call.arguments.address,
    'range filters persist on the worksheet and must not share the create_table fixture range'
  );
});

test('Excel Office E2E driver', { skip: !officeE2eEnabled() }, async () => {
  await runOfficeToolE2e({
    host: 'Excel',
    cases: EXCEL_E2E_CASES,
    driver: requireOfficeE2eDriver('Excel'),
    reportPath: REPORT_PATH
  });
});
