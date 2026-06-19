# 04 — Excel Capabilities

This document is the Excel v1 capability contract for `office-mcp`. Excel tools
run inside the document-scoped Excel Office.js add-in under `src/office-ctl/excel`
and are routed by the Rust daemon through the same add-in JSON-RPC channel as
Word tools.

## 1. Scope

v1 supports one connected workbook session per add-in runtime. The daemon
addresses the workbook by `session_id`; the add-in executes workbook operations
with `Excel.run` against either the active worksheet or an explicitly named
worksheet when the tool accepts `sheet`.

The target Excel surface is based on the Microsoft Excel add-in object model:
most useful workflows start at `Workbook`, move to `Worksheet`, operate on
`Range` values/formulas/formats, and then promote data into higher-level
`Table`, `Chart`, or `PivotTable` objects. `office-mcp` exposes task-oriented
tools for those workflows instead of mirroring every Excel.js class and method.

Out of scope for the core Excel surface: shapes and images, comments and notes,
slicers as first-class tools, events/subscriptions, custom XML, external data
connections, Power Query, Python/preview-only APIs, OLAP/Power Pivot, arbitrary
file import/export, workbook close, and save-as flows. These can be added later
only when there is a clear user workflow and host support evidence.

Implemented Excel v1 tools:

| Tool | Side effect | Minimum API | Summary |
|---|---|---|---|
| `excel.read_range` | read | `ExcelApi 1.1` | Read values, display text, dimensions, and number format for a range. |
| `excel.write_range` | edit | `ExcelApi 1.1` | Write a two-dimensional values matrix to a range. |
| `excel.add_sheet` | edit | `ExcelApi 1.1` | Add a worksheet and optionally activate it. |
| `excel.set_formula` | edit | `ExcelApi 1.1` | Fill a range with one formula. |
| `excel.format_range` | edit | `ExcelApi 1.1` | Apply basic font, fill, and number-format styling. |
| `excel.create_table` | edit | `ExcelApi 1.1` | Create a workbook table from a range. |
| `excel.create_chart` | edit | `ExcelApi 1.1` | Create a chart from a range on the active or named worksheet. |

The Excel task pane reports these tools in `session.added.available_tools` only
after the runtime registers successfully with the daemon.

Target core Excel tool surface:

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `excel.get_workbook_info` | planned | Workbook | read | verify during implementation | Return workbook identity, workbook-level state, active sheet name/id, and aggregate object counts; detailed sheet inventory belongs to `excel.list_sheets`. |
| `excel.list_sheets` | planned | Worksheet | read | `ExcelApi 1.1` | List worksheets with id, name, position, visibility, tab color, and active state; workbook metadata belongs to `excel.get_workbook_info`. |
| `excel.add_sheet` | implemented | Worksheet | edit | `ExcelApi 1.1` | Add a worksheet and optionally activate it. |
| `excel.update_sheet` | planned | Worksheet | edit | `ExcelApi 1.1` | Rename, activate, move, set visibility, and set tab color for a worksheet. |
| `excel.delete_sheet` | planned | Worksheet | destructive | `ExcelApi 1.1` | Delete a worksheet, rejecting attempts that would leave the workbook without sheets. |
| `excel.get_used_range` | planned | Range | read | `ExcelApi 1.1` | Return the used range address and dimensions for a sheet; cell values/text/formulas belong to `excel.read_range`. |
| `excel.read_range` | implemented | Range | read | `ExcelApi 1.1` | Read values, display text, formulas, dimensions, and number format for an explicit range. |
| `excel.write_range` | implemented | Range | edit | `ExcelApi 1.1` | Write a two-dimensional values matrix to a range. |
| `excel.clear_range` | planned | Range | destructive | `ExcelApi 1.1` | Clear contents, formats, hyperlinks, or all range data; optional cell deletion with shift direction. |
| `excel.find_replace_cells` | planned | Range | read/edit | verify during implementation | Search text, values, or formulas in a worksheet/range and optionally replace matches. |
| `excel.set_formula` | implemented | Formula | edit | `ExcelApi 1.1` | Fill a range with one formula or a formula matrix. |
| `excel.format_range` | implemented | Format | edit | `ExcelApi 1.1` | Apply font, fill, number format, borders, alignment, and autofit basics. |
| `excel.sort_range` | planned | Data | edit | verify during implementation | Sort a range or table body by one or more column keys; table structure changes belong to `excel.update_table`. |
| `excel.apply_filter` | planned | Data | edit | verify during implementation | Apply or clear worksheet range or table filter criteria; PivotTable filters belong to `excel.update_pivot_table`. |
| `excel.create_table` | implemented | Table | edit | `ExcelApi 1.1` | Create a workbook table from a range. |
| `excel.update_table` | planned | Table | read/edit/destructive | `ExcelApi 1.1` | Read table metadata/structure; add rows/columns; resize, rename, change table style/options, or delete a table. Table cell contents belong to `excel.read_range`. |
| `excel.create_chart` | implemented | Chart | edit | `ExcelApi 1.1` | Create a chart from a range or table. |
| `excel.update_chart` | planned | Chart | edit/read/destructive | verify during implementation | Update chart title, axes, legend, series, position, size, delete the chart, or export a chart image where supported. |
| `excel.create_pivot_table` | planned | PivotTable | edit | verify during implementation | Create a PivotTable from a range or table at a target destination. |
| `excel.update_pivot_table` | planned | PivotTable | edit/destructive | verify during implementation | Configure row, column, data, and filter hierarchies; set aggregation/calculation, refresh, apply PivotTable filters, or delete a PivotTable. |

The planned tools above are the Excel implementation backlog, not the current
runtime catalog. Before implementation, each planned tool must verify its
minimum requirement set against `@types/office-js` and the Microsoft API docs.

Tool ownership rules:

- One tool owns each common user intent. Do not add a second tool unless it has a
  different object owner, permission profile, or user-visible result.
- `excel.get_workbook_info` is workbook state only. It may include the active
  sheet id/name for orientation, but the worksheet list belongs to
  `excel.list_sheets`.
- `excel.get_used_range` locates the occupied sheet area. It does not return
  cell contents; callers use `excel.read_range` for values, text, formulas, and
  number formats.
- `excel.write_range` writes literal values. `excel.set_formula` writes formulas
  and formula matrices. Formula strings passed to `excel.write_range` remain
  literal input unless explicitly documented otherwise during implementation.
- `excel.format_range` owns cell-level visual formatting. Table style/options
  remain in `excel.update_table`; chart visual settings remain in
  `excel.update_chart`; PivotTable layout/format settings remain in
  `excel.update_pivot_table`.
- `excel.sort_range` and `excel.apply_filter` own sorting and filtering for both
  plain ranges and table bodies. `excel.update_table` must not duplicate table
  sort/filter behavior.
- `excel.update_table` may return table metadata and ranges, but it must not
  return table cell contents. Callers use `excel.read_range` against the table
  range for values, text, formulas, and number formats.
- `excel.update_table`, `excel.update_chart`, and
  `excel.update_pivot_table` are object-owner tools. They group lifecycle and
  configuration operations for their object type, but must not absorb generic
  range, formula, or cell-format operations.

## 2. Shared Arguments

Most Excel tools accept:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:C3"
}
```

- `session_id` is required by the daemon for all document-affecting calls.
- `address` is required for range tools and uses Excel A1 notation.
- `sheet` is optional. When omitted, the active worksheet is used.

The daemon serializes calls per `session_id`; the add-in uses Office.js host
errors for invalid sheet names, invalid ranges, protected workbook failures, and
other host-denied operations. These map through the standard error model in
[06-error-model.md](06-error-model.md).

## 3. Tool Contracts

### 3.1 `excel.read_range`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:B2"
}
```

Returns:

```json
{
  "address": "Sheet1!A1:B2",
  "values": [["Label", "Value"], ["Q1", 42]],
  "text": [["Label", "Value"], ["Q1", "42"]],
  "row_count": 2,
  "column_count": 2,
  "number_format": [["General", "General"], ["General", "General"]],
  "untrusted_source": true
}
```

`untrusted_source` is always `true` because workbook content may contain prompt
injection text.

### 3.2 `excel.write_range`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:B2",
  "values": [["Label", "Value"], ["Q1", 42]]
}
```

`values` MUST be a two-dimensional array. The host decides whether the shape is
valid for the target range.

Returns:

```json
{
  "address": "A1:B2",
  "row_count": 2,
  "column_count": 2,
  "wrote_values": true
}
```

### 3.3 `excel.add_sheet`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "name": "Analysis",
  "activate": true
}
```

- `name` is optional. If omitted, Excel chooses the worksheet name.
- `activate` defaults to `true`; pass `false` to add without activating.

Returns `{ "sheet": "Analysis", "activated": true }`.

### 3.4 `excel.set_formula`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "C2:C10",
  "formula": "=B2*2"
}
```

The same formula is written to every cell in the target range. Formula syntax is
the host Excel formula syntax.

Returns `{ "address": "C2:C10", "formula": "=B2*2", "wrote_formula": true }`.

### 3.5 `excel.format_range`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:C2",
  "bold": true,
  "italic": false,
  "font_color": "#17202A",
  "fill_color": "#DDEEFF",
  "number_format": "General"
}
```

All formatting fields are optional. `number_format`, when supplied, is written
as a matrix matching the target range dimensions.

Returns `{ "address": "A1:C2", "formatted": true }`.

### 3.6 `excel.create_table`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "has_headers": true,
  "name": "SalesTable"
}
```

- `has_headers` defaults to `true`.
- `name` is optional. If omitted, Excel chooses the table name.

Returns `{ "table": "SalesTable", "address": "A1:C10", "has_headers": true }`.

### 3.7 `excel.create_chart`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "type": "columnClustered",
  "title": "Quarterly Sales"
}
```

Supported `type` values in v1 are `area`, `barClustered`, `columnClustered`,
`doughnut`, `line`, `pie`, and `scatter`. Unknown values fall back to
`columnClustered`.

Returns `{ "chart": "Chart 1", "chart_type": "columnClustered", "source": "A1:C10" }`.

## 4. Planned Tool Contract Notes

Planned tools should keep contracts coarse and workflow-oriented:

- Workbook and worksheet tools own navigation, sheet lifecycle, and workbook
  metadata. They must not read workbook cell contents.
- Range tools own cell data, formulas, cell-level formatting, search/replace,
  sorting, filtering, and clearing. A single-cell operation is represented as a
  one-cell range.
- `excel.clear_range` is the only planned destructive cell primitive. It should
  cover content/format clearing and explicit delete-with-shift modes instead of
  adding a separate delete-cell tool.
- `excel.update_table`, `excel.update_chart`, and `excel.update_pivot_table`
  intentionally group object lifecycle and configuration by owner so the MCP
  catalog stays compact while still covering user-visible Excel workflows.
- PivotTable support must focus on normal workbook/range/table sources. OLAP,
  Power Pivot, slicer UI management, and preview-only PivotTable APIs remain
  deferred.

## 5. Limits And Validation

- The daemon validates session existence, session capability, queue depth,
  request size, timeout, and response size before or around forwarding.
- The add-in validates required scalar fields and obvious JSON shape errors.
- Excel remains the authority for range validity, formula syntax, table/chart
  constraints, protected workbook denial, and other workbook-specific errors.
- The server does not expose an Excel resource URI surface in v1. Clients use
  `excel.get_workbook_info`, `excel.list_sheets`, `excel.get_used_range`, and
  `excel.read_range` for workbook reads once the planned surface is implemented.

## 6. Evidence

Automated evidence exists at two levels:

- Rust daemon/unit evidence proves the Excel tool catalog is listed and that
  MCP calls are forwarded to an Excel add-in session.
- Runtime evidence supports a live workbook smoke with
  `npm run evidence:excel` from `src/office-mcp/daemon/evidence`, validated with
  `npm run evidence:validate -- --require-excel-smoke`.

Completion of the live Excel gate requires a connected Excel workbook and a
passed `excel.runtime_smoke` report.
