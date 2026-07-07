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
      },
      contains: ['Normal']
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
  ['excel.update_document_properties', {
    setup: {
      actions: [
        { tool: 'excel.get_workbook_info', arguments: {} }
      ]
    },
    args: {
      title: 'E2E Workbook Properties Title',
      author: 'Office MCP E2E',
      custom_set: [{ key: 'OfficeMcpE2E', value: 'excel-properties' }]
    },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.get_document_properties',
      readbackArguments: { include_custom: true },
      expect: {
        contains: ['E2E Workbook Properties Title', 'Office MCP E2E', 'OfficeMcpE2E', 'excel-properties'],
        pathEquals: [{ path: 'untrusted_source', value: true }]
      }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_document_properties', arguments: { title: '', author: '', custom_delete: ['OfficeMcpE2E'] } }
      ]
    }
  }],
  ['excel.get_document_properties', {
    setup: {
      actions: [
        { tool: 'excel.update_document_properties', arguments: { subject: 'E2E Workbook Properties Subject', custom_set: [{ key: 'OfficeMcpReadback', value: true }] } }
      ]
    },
    args: { include_custom: true },
    verify: {
      kind: 'direct-result',
      expect: {
        contains: ['E2E Workbook Properties Subject', 'OfficeMcpReadback', 'true'],
        pathEquals: [{ path: 'untrusted_source', value: true }]
      }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_document_properties', arguments: { subject: '', custom_delete: ['OfficeMcpReadback'] } }
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
        { tool: 'excel.add_sheet', arguments: { name: 'Sheet To Rename' } },
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet To Rename', address: 'A1:C3', values: [['Header', 'One', 'Two'], ['Row', 1, 2], ['Row', 3, 4]] } }
      ]
    },
    args: { sheet: 'Sheet To Rename', name: 'Renamed Sheet', freeze: { rows: 1 }, show_gridlines: false, show_headings: false },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_sheets',
      readbackArguments: {},
      expect: {
        contains: ['Renamed Sheet', 'A2'],
        notContains: ['Sheet To Rename'],
        pathEquals: [
          { path: 'sheets.1.frozen.rows', value: 1 },
          { path: 'sheets.1.frozen.columns', value: 0 },
          { path: 'sheets.1.show_gridlines', value: false },
          { path: 'sheets.1.show_headings', value: false }
        ]
      }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_sheet', arguments: { sheet: 'Renamed Sheet', freeze: { unfreeze: true }, show_gridlines: true, show_headings: true } },
        { tool: 'excel.update_sheet', arguments: { sheet: 'Renamed Sheet', freeze: { at: 'B3' } } }
      ]
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
  ['excel.read_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:B2', values: [['Read', 'Range'], ['E2E', 'Marker']] } },
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'C3:D4', values: [['Used', 'Range'], ['E2E', 'Marker']] } }
      ]
    },
    args: { sheet: 'Sheet1', metadata_only: true },
    verify: {
      kind: 'direct-result',
      expect: {
        pathEquals: [
          { path: 'row_count', value: 4 },
          { path: 'column_count', value: 4 },
          { path: 'empty', value: false }
        ],
        notContains: ['Read', 'Range', 'E2E', 'Marker']
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
  ['excel.insert_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'A1:A3', values: [['before'], ['shift-me'], ['after']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'A2', shift: 'down', count: 1 },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:A4' },
      expect: { contains: ['before', 'shift-me', 'after'], pathEquals: [{ path: 'row_count', value: 4 }] }
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
  ['excel.set_hyperlink', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'H1', values: [['OpenAI']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'H1', action: 'set', url: 'https://openai.com/', text_to_display: 'OpenAI', screen_tip: 'OpenAI home' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'H1', include_hyperlinks: true },
      expect: {
        contains: ['https://openai.com/', 'OpenAI', 'OpenAI home'],
        pathEquals: [{ path: 'untrusted_source', value: true }]
      }
    },
    cleanup: {
      actions: [
        { tool: 'excel.set_hyperlink', arguments: { sheet: 'Sheet1', address: 'H1', action: 'clear' } }
      ]
    }
  }],
  ['excel.set_data_validation', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'I1:I3', values: [['Open'], ['Closed'], ['Open']] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'I1:I3', action: 'set', rule: { type: 'list', list_source: ['Open', 'Closed'], in_cell_dropdown: true }, error_alert: { style: 'stop', title: 'Invalid status', message: 'Choose Open or Closed.' }, input_prompt: { title: 'Status', message: 'Pick a status.' } },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'I1:I3', include_validation: true },
      expect: { contains: ['list', 'Open,Closed', 'Invalid status'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    },
    cleanup: {
      actions: [
        { tool: 'excel.set_data_validation', arguments: { sheet: 'Sheet1', address: 'I1:I3', action: 'clear' } }
      ]
    }
  }],
  ['excel.copy_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'J1:K2', values: [['Source', 'Copy'], [1, 2]] } },
        { tool: 'excel.set_formula', arguments: { sheet: 'Sheet1', address: 'K2', formula: '=J2*2' } }
      ]
    },
    args: { source_sheet: 'Sheet1', source_address: 'J1:K2', destination_sheet: 'Sheet1', destination_address: 'M1:N2', action: 'copy', copy_type: 'all' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'M1:N2', include_formulas: true },
      expect: { contains: ['Source', 'Copy', '=M2*2'] }
    }
  }],
  ['excel.copy_range', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'O1:O2', values: [[1], [2]] } }
      ]
    },
    args: { sheet: 'Sheet1', source_address: 'O1:O2', destination_address: 'O1:O6', action: 'autofill', autofill_type: 'series' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'O1:O6' },
      expect: { contains: ['1', '2', '6'] }
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
    args: {
      sheet: 'Sheet1',
      address: 'A1:B2',
      style: 'Normal',
      number_format: '0.00',
      column_width_pt: 72,
      row_height_pt: 24,
      hidden_columns: false,
      hidden_rows: false,
      merge: 'merge_across'
    },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.read_range',
      readbackArguments: { sheet: 'Sheet1', address: 'A1:B2' },
      expect: { contains: ['0.00'] }
    }
  }],
  ['excel.list_conditional_formats', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'F1:F3', values: [[-1], [2], [-3]] } },
        { tool: 'excel.update_conditional_format', arguments: { sheet: 'Sheet1', address: 'F1:F3', action: 'add', rule: { type: 'cell_value', operator: 'less_than', values: [0], format: { fill_color: '#FFC7CE', font_color: '#9C0006' } } } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'F1:F3' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['cell_value', 'less_than', '#FFC7CE'], pathEquals: [{ path: 'count', value: 1 }] }
    }
  }],
  ['excel.update_conditional_format', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'G1:G3', values: [[1], [2], [3]] } }
      ]
    },
    args: { sheet: 'Sheet1', address: 'G1:G3', action: 'add', rule: { type: 'color_scale', colors: ['#63BE7B', '#FFEB84', '#F8696B'] } },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_conditional_formats',
      readbackArguments: { sheet: 'Sheet1', address: 'G1:G3' },
      expect: { contains: ['color_scale', '#63BE7B', '#F8696B'] }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_conditional_format', arguments: { sheet: 'Sheet1', address: 'G1:G3', action: 'clear_range' } }
      ]
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
  }],
  ['excel.insert_image', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'AD1', values: [['Image anchor']] } }
      ]
    },
    args: {
      sheet: 'Sheet1',
      image: { base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/l6jLJwAAAABJRU5ErkJggg==' },
      left_pt: 24,
      top_pt: 24,
      width_pt: 48,
      height_pt: 48,
      alt_text: 'E2E Excel image'
    },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_shapes',
      readbackArguments: { sheet: 'Sheet1' },
      expect: { contains: ['E2E Excel image'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['excel.list_shapes', {
    setup: {
      actions: [
        { tool: 'excel.insert_image', arguments: { sheet: 'Sheet1', image: { base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/l6jLJwAAAABJRU5ErkJggg==' }, alt_text: 'Listed Excel image' } }
      ]
    },
    args: { sheet: 'Sheet1' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['Listed Excel image'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['excel.update_shape', {
    setup: {
      actions: [
        { tool: 'excel.insert_image', saveAs: 'shapeResult', arguments: { sheet: 'Sheet1', image: { base64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/l6jLJwAAAABJRU5ErkJggg==' }, alt_text: 'Shape before update' } }
      ]
    },
    args: { sheet: 'Sheet1', shape_id: '${shapeResult.shape_id}', action: 'set_alt_text', alt_text: 'Shape after update' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_shapes',
      readbackArguments: { sheet: 'Sheet1' },
      expect: { contains: ['Shape after update'], notContains: ['Shape before update'] }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_shape', arguments: { sheet: 'Sheet1', shape_id: '${shapeResult.shape_id}', action: 'delete' } }
      ]
    }
  }],
  ['excel.add_comment', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'AC1', values: [['Comment target']] } }
      ]
    },
    args: { sheet: 'Sheet1', cell: 'AC1', text: 'E2E Excel comment' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_comments',
      readbackArguments: { sheet: 'Sheet1' },
      expect: { contains: ['E2E Excel comment'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['excel.list_comments', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'AD1', values: [['List comment target']] } },
        { tool: 'excel.add_comment', arguments: { sheet: 'Sheet1', cell: 'AD1', text: 'E2E listed comment' } }
      ]
    },
    args: { sheet: 'Sheet1' },
    verify: {
      kind: 'direct-result',
      expect: { contains: ['E2E listed comment'], pathEquals: [{ path: 'untrusted_source', value: true }] }
    }
  }],
  ['excel.update_comment', {
    setup: {
      actions: [
        { tool: 'excel.write_range', arguments: { sheet: 'Sheet1', address: 'AE1', values: [['Update comment target']] } },
        { tool: 'excel.add_comment', saveAs: 'commentResult', arguments: { sheet: 'Sheet1', cell: 'AE1', text: 'E2E original comment' } }
      ]
    },
    args: { comment_id: '${commentResult.comment_id}', action: 'reply', text: 'E2E reply body' },
    verify: {
      kind: 'readback',
      readbackTool: 'excel.list_comments',
      readbackArguments: { sheet: 'Sheet1' },
      expect: { contains: ['E2E original comment', 'E2E reply body'] }
    },
    cleanup: {
      actions: [
        { tool: 'excel.update_comment', arguments: { comment_id: '${commentResult.comment_id}', action: 'delete' } }
      ]
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
