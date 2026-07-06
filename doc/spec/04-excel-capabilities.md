# 04 — Excel Capabilities

This document is the Excel v1 capability contract for `office-mcp`. Excel tools
run inside the document-scoped Excel Office.js add-in under `src/office-ctl/excel`
and are routed by the Rust daemon through the same add-in JSON-RPC channel as
Word tools. The tool list in this file is the source of truth for the daemon
catalog, add-in `available_tools`, task pane permission UI, and implementation
TODO list.

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

Research basis:

- Microsoft Learn's "Core Excel object model concepts"
  (https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts)
  is the starting point for the v1 surface. It describes the common add-in
  workflow as starting from a `Workbook`, moving to a `Worksheet`, working with
  one or more `Range` objects, and then creating higher-level `Table` or
  `Chart` objects from existing range data.
- The same page frames Excel add-ins around reading, writing, and visualizing
  workbook data through workbooks, worksheets, ranges, tables, and charts. This
  is the primary selection filter for v1 tools.
- The core concepts page states that a `Range` represents either one cell or a
  contiguous block of cells, and that Excel JavaScript API has no separate
  `Cell` object. Therefore v1 must not add separate cell CRUD tools. Single-cell
  reads, writes, formulas, formatting, clearing, search, sorting, and filtering
  are all range operations.
- The page calls out `values`, `formulas`, and `format` as the immediately
  useful range properties. v1 exposes those as separate value, formula, and
  formatting intents because they have different permission and validation
  profiles.
- Microsoft Learn's table guide focuses on creating, resizing, reading,
  sorting, filtering, and formatting tables. v1 keeps table lifecycle and
  table-specific style/options in `excel.update_table`, while generic table
  data, sort, and filter operations stay with range/data tools.
- Microsoft Learn's chart guide focuses on creating charts, adding series,
  formatting titles and axes, and exporting chart images. v1 exposes chart
  creation and one chart-owner update tool instead of separate axis, legend,
  series, and export tools.
- Microsoft Learn's PivotTable guide focuses on creating, configuring, and
  filtering PivotTables. PivotTables are not part of the core concepts page's
  first-level walkthrough, but they are included because they are a high-value
  Excel analysis workflow and are present in the stable Excel.js object model.
  v1 limits PivotTables to normal range/table sources and non-preview APIs.

The refined v1 target is 22 tools. The original 20-tool core covered workbook
orientation, worksheet lifecycle, range work, formulas, formatting, data
operations, tables, charts, and PivotTables. The workbook owner now also covers
two stable persistence/calculation tasks that cannot be represented by range or
object-owner tools: saving the current workbook and forcing calculation. Do not
expand the catalog by copying individual Excel.js methods into MCP tools. A
future tool can only be added after the selection matrix proves that the
existing tools cannot express a distinct object owner, permission profile, or
user-visible workflow safely.

The v1 priority order is: workbook orientation, worksheet lifecycle, range/cell
data CRUD, formulas, formatting, sort/filter, tables, charts, and PivotTables.
This follows the Excel core concepts path from workbook to worksheet to range,
then to table and chart objects; PivotTables are included as the only extra v1
analysis object because summarized workbook analysis is a core Excel user need.

Selection rules for the 15-20 tool budget:

- Start from the Microsoft Learn core object path: `Workbook` -> `Worksheet` ->
  `Range` -> `Table` / `Chart`.
- Add a tool only when it represents a user-level task, not a single Office.js
  property or method.
- Merge object lifecycle/configuration operations into object-owner update tools
  when they share the same owner and permission profile.
- Keep cell operations as range operations because Excel.js has no separate
  `Cell` object.
- Include PivotTables as the only v1 analysis object beyond the core concepts
  page because summarized analysis is a central Excel workflow; defer slicers,
  OLAP, Power Pivot, and preview-only APIs.

Out of scope for the core Excel surface: shapes and images, comments and notes,
slicers as first-class tools, events/subscriptions, custom XML, external data
connections, Power Query, Python/preview-only APIs, OLAP/Power Pivot, arbitrary
file import/export, workbook close, and save-as flows. Workbook save and
calculation are in scope because they are stable workbook-owner operations, not
file export, save-as, or close workflows.

Core tool selection matrix:

| User workflow | Object owner | v1 tools | Why this is enough |
|---|---|---|---|
| Inspect, persist, and calculate workbook state; navigate sheets | `Workbook` / `Worksheet` | `excel.get_workbook_info`, `excel.save`, `excel.calculate`, `excel.list_sheets`, `excel.add_sheet`, `excel.update_sheet`, `excel.delete_sheet` | Covers workspace-level sheet CRUD, workbook orientation, persistence, and formula recalculation without reading cell contents. |
| Locate and mutate cell data | `Range` | `excel.get_used_range`, `excel.read_range`, `excel.write_range`, `excel.clear_range`, `excel.find_replace_cells` | Cells are one-cell ranges, so separate cell CRUD would duplicate range tools. |
| Work with formulas | `Range` formulas | `excel.set_formula` | Formula writes are distinct from literal value writes and can support scalar or matrix input. |
| Apply user-visible cell formatting | `RangeFormat` | `excel.format_range` | Keeps font, fill, number format, borders, alignment, wrapping, and autofit under one cell-format owner. |
| Sort and filter data | `RangeSort` / table filter owners | `excel.sort_range`, `excel.apply_filter` | One data-operation owner covers both plain ranges and table bodies; table tools must not duplicate it. |
| Promote data into a structured table | `Table` | `excel.create_table`, `excel.update_table` | Creation is common enough to be direct; lifecycle, rows/columns, resize, rename, and table style/options share one object-owner update tool. |
| Visualize data | `Chart` | `excel.create_chart`, `excel.update_chart` | Creation is direct; title, axes, legend, series, size, position, delete, and supported export share one chart-owner update tool. |
| Analyze summarized data | `PivotTable` | `excel.create_pivot_table`, `excel.update_pivot_table` | Creation is direct; fields, filters, aggregation, refresh, metadata, and delete share one PivotTable-owner update tool. |

Final v1 tool set by category:

| Category | Tools | Count | User intent |
|---|---|---:|---|
| Workbook | `excel.get_workbook_info`, `excel.save`, `excel.calculate` | 3 | Inspect workbook-level state, persist changes, and recalculate formulas without reading cell contents. |
| Worksheet | `excel.list_sheets`, `excel.add_sheet`, `excel.update_sheet`, `excel.delete_sheet` | 4 | Sheet inventory and lifecycle. |
| Range / cell data | `excel.get_used_range`, `excel.read_range`, `excel.write_range`, `excel.clear_range`, `excel.find_replace_cells` | 5 | Locate, read, write, clear, and search cells through ranges. |
| Formula | `excel.set_formula` | 1 | Author formulas distinctly from literal value writes. |
| Format | `excel.format_range` | 1 | Apply user-visible cell formatting. |
| Data operations | `excel.sort_range`, `excel.apply_filter` | 2 | Sort and filter plain ranges or table bodies. |
| Table | `excel.create_table`, `excel.update_table` | 2 | Promote data to structured tables and manage table-owned lifecycle/options. |
| Chart | `excel.create_chart`, `excel.update_chart` | 2 | Visualize data and manage chart-owned configuration. |
| PivotTable | `excel.create_pivot_table`, `excel.update_pivot_table` | 2 | Create and configure summarized analysis views. |

Total: 22 tools.

Rejected tool families for v1:

- No `excel.read_cell`, `excel.write_cell`, or `excel.delete_cell`; cells are
  one-cell ranges.
- No separate worksheet format, freeze pane, protection, comments, shapes,
  images, slicer, event, binding, named-item, custom XML, external connection,
  Power Query, workbook import/export, save-as, or close tools. `excel.save`
  intentionally persists the current workbook through host save behavior only;
  it does not expose save-as, export, prompt, path, or close semantics.
- No separate table row, table column, table style, table sort, or table filter
  tools. Table structure/options belong to `excel.update_table`; table data,
  sort, and filter behavior belongs to range/data tools.
- No separate chart title, axis, legend, series, image-export, move, resize, or
  delete tools. Those actions belong to `excel.update_chart`.
- No separate PivotTable field, filter, refresh, or delete tools. Those actions
  belong to `excel.update_pivot_table`.

Implemented Excel v1 tools:

| Tool | Side effect | Minimum API | Summary |
|---|---|---|---|
| `excel.get_workbook_info` | read | `ExcelApi 1.1` | Return workbook identity, active sheet, and aggregate object counts. |
| `excel.save` | edit | `ExcelApi 1.11`; dirty state uses `ExcelApi 1.9` when available | Save the current workbook through host save behavior and report the pre-save dirty state when supported. |
| `excel.calculate` | edit | `ExcelApi 1.1`; `full_rebuild` requires `ExcelApi 1.2` | Recalculate workbook formulas and report the calculation mode. |
| `excel.list_sheets` | read | `ExcelApi 1.1` | List worksheets with id, name, position, visibility, tab color, and active state. |
| `excel.add_sheet` | edit | `ExcelApi 1.1` | Add a worksheet and optionally activate it. |
| `excel.update_sheet` | edit | `ExcelApi 1.1` | Rename, activate, move, set visibility, and set tab color for a worksheet. |
| `excel.delete_sheet` | destructive | `ExcelApi 1.1` | Delete a worksheet, rejecting attempts that would leave the workbook without sheets. |
| `excel.get_used_range` | read | `ExcelApi 1.1` | Return the used range address and dimensions for a sheet. |
| `excel.read_range` | read | `ExcelApi 1.1` | Read values, display text, dimensions, and number format for a range. |
| `excel.write_range` | edit | `ExcelApi 1.1` | Write a two-dimensional values matrix to a range. |
| `excel.clear_range` | destructive | `ExcelApi 1.1` | Clear contents, formats, all range data, or delete cells with a shift direction. |
| `excel.find_replace_cells` | read/edit | `ExcelApi 1.9` | Find the first matching cell in a range or replace matching cells. |
| `excel.set_formula` | edit | `ExcelApi 1.1` | Fill a range with one formula or a formula matrix. |
| `excel.format_range` | edit | `ExcelApi 1.1`; autofit requires `ExcelApi 1.2` | Apply font, fill, number formats, borders, alignment, wrapping, and autofit. |
| `excel.sort_range` | edit | `ExcelApi 1.2` | Sort a range or table body by one or more keys. |
| `excel.apply_filter` | edit | Range filters require `ExcelApi 1.9`; table column filters require `ExcelApi 1.2` | Apply, clear, remove, or reapply range and table filters. |
| `excel.create_table` | edit | `ExcelApi 1.1` | Create a workbook table from a range. |
| `excel.update_table` | read/edit/destructive | `ExcelApi 1.1`; visual options require `ExcelApi 1.3`; resize requires `ExcelApi 1.13` | Read or update table structure, options, and lifecycle. |
| `excel.create_chart` | edit | `ExcelApi 1.1` | Create a chart from a range on the active or named worksheet. |
| `excel.update_chart` | read/edit/destructive | `ExcelApi 1.1`; image export requires `ExcelApi 1.2`; axis selection and chart type metadata require `ExcelApi 1.7` | Read or update chart configuration, source, export, and lifecycle. |
| `excel.create_pivot_table` | edit | `ExcelApi 1.8` | Create a PivotTable from a range or table source. |
| `excel.update_pivot_table` | read/edit/destructive | `ExcelApi 1.3`; hierarchy/layout/delete require `ExcelApi 1.8`; filters require `ExcelApi 1.12` | Read or update PivotTable fields, layout, filters, refresh, and lifecycle. |

The Excel task pane reports these tools in `session.added.available_tools` only
after the runtime registers successfully with the daemon.

Target core Excel tool surface:

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `excel.get_workbook_info` | implemented | Workbook | read | `ExcelApi 1.1` | Return workbook identity, workbook-level state, active sheet name/id, and aggregate object counts; detailed sheet inventory belongs to `excel.list_sheets`. |
| `excel.save` | planned | Workbook | edit | `ExcelApi 1.11`; dirty state uses `ExcelApi 1.9` when available | Save the current workbook through host save behavior. Save-as, export, prompts, and close flows remain out of scope. |
| `excel.calculate` | planned | Workbook | edit | `ExcelApi 1.1`; `full_rebuild` requires `ExcelApi 1.2` | Recalculate workbook formulas with `recalculate`, `full`, or `full_rebuild` mode and return the calculation mode. |
| `excel.list_sheets` | implemented | Worksheet | read | `ExcelApi 1.1` | List worksheets with id, name, position, visibility, tab color, and active state; workbook metadata belongs to `excel.get_workbook_info`. |
| `excel.add_sheet` | implemented | Worksheet | edit | `ExcelApi 1.1` | Add a worksheet and optionally activate it. |
| `excel.update_sheet` | implemented | Worksheet | edit | `ExcelApi 1.1` | Rename, activate, move, set visibility, and set tab color for a worksheet. |
| `excel.delete_sheet` | implemented | Worksheet | destructive | `ExcelApi 1.1` | Delete a worksheet, rejecting attempts that would leave the workbook without sheets. |
| `excel.get_used_range` | implemented | Range | read | `ExcelApi 1.1` | Return the used range address and dimensions for a sheet; cell values/text/formulas belong to `excel.read_range`. |
| `excel.read_range` | implemented | Range | read | `ExcelApi 1.1` | Read values, display text, formulas, dimensions, and number format for an explicit range. |
| `excel.write_range` | implemented | Range | edit | `ExcelApi 1.1` | Write a two-dimensional values matrix to a range. |
| `excel.clear_range` | implemented | Range | destructive | `ExcelApi 1.1` | Clear contents, formats, or all range data; optional cell deletion with shift direction. Hyperlink-specific clear modes are deferred because they require `ExcelApi 1.7`. |
| `excel.find_replace_cells` | implemented | Range | read/edit | `ExcelApi 1.9` | Search cell contents in a range and optionally replace matches. |
| `excel.set_formula` | implemented | Formula | edit | `ExcelApi 1.1` | Fill a range with one formula or a formula matrix. |
| `excel.format_range` | implemented | Format | edit | `ExcelApi 1.1`; autofit requires `ExcelApi 1.2` | Apply font, fill, scalar or matrix number formats, borders, alignment, wrapping, and autofit. |
| `excel.sort_range` | implemented | Data | edit | `ExcelApi 1.2` | Sort a range or table body by one or more column keys; table structure changes belong to `excel.update_table`. |
| `excel.apply_filter` | implemented | Data | edit | Range filters require `ExcelApi 1.9`; table column filters require `ExcelApi 1.2` | Apply, clear, remove, or reapply worksheet range or table filter criteria; PivotTable filters belong to `excel.update_pivot_table`. |
| `excel.create_table` | implemented | Table | edit | `ExcelApi 1.1` | Create a workbook table from a range. |
| `excel.update_table` | implemented | Table | read/edit/destructive | `ExcelApi 1.1`; visual options require `ExcelApi 1.3`; resize requires `ExcelApi 1.13` | Read table metadata/structure; add rows/columns; resize, rename, change table style/options, or delete a table. Table cell contents belong to `excel.read_range`. |
| `excel.create_chart` | implemented | Chart | edit | `ExcelApi 1.1` | Create a chart from a range or table. |
| `excel.update_chart` | implemented | Chart | edit/read/destructive | `ExcelApi 1.1`; image export requires `ExcelApi 1.2`; axis selection and chart type metadata require `ExcelApi 1.7` | Read chart metadata; update chart title, axes, legend, source range, position, size, delete the chart, or export a chart image where supported. |
| `excel.create_pivot_table` | implemented | PivotTable | edit | `ExcelApi 1.8` | Create a PivotTable from a range or table at a target destination. |
| `excel.update_pivot_table` | implemented | PivotTable | edit/destructive | `ExcelApi 1.3`; hierarchy/layout/delete require `ExcelApi 1.8`; filters require `ExcelApi 1.12` | Read PivotTable metadata; configure row, column, data, and filter hierarchies; set aggregation and layout options, refresh, apply manual PivotTable filters, clear filters, or delete a PivotTable. |

The tools above are the Excel v1 contract. Implementation work must keep
the daemon catalog, MCP `tools/list`, Excel task pane `available_tools`, task
pane permission grouping, documentation, and tests aligned with this 22-tool
surface. Before implementing or changing a tool, verify its minimum requirement
set against `@types/office-js` and Microsoft API docs, then land the change as a
test-first implementation slice with daemon catalog coverage, task pane
contract coverage, and live Excel smoke evidence where the host API cannot be
fully proven statically.

Tool ownership rules:

- Excel follows the cross-app naming and split conventions in
  [03-mcp-tool-surface.md](03-mcp-tool-surface.md) §1.1. Workbook and worksheet
  state reads use `get_*` / `list_*`, range content reads use `read_*`, sheet
  collection operations use `add_*`, promoted workbook objects use `create_*`,
  and lifecycle/configuration changes use object-owner `update_*` tools.
- Excel follows the cross-app validation-only contract in
  [03-mcp-tool-surface.md](03-mcp-tool-surface.md) §6. Every advertised
  mutating Excel tool MUST accept `validate_only: true`, must advertise
  `_meta["com.office-mcp/supports_validate_only"]: true`, and must return the
  portable no-write preflight envelope before queuing workbook writes.
- One tool owns each common user intent. Do not add a second tool unless it has a
  different object owner, permission profile, or user-visible result.
- `excel.get_workbook_info` is workbook state only. It may include the active
  sheet id/name for orientation, but the worksheet list belongs to
  `excel.list_sheets`.
- `excel.save` is the workbook persistence owner. It calls the host save
  behavior for the current workbook only and must not accept path, format,
  prompt, export, save-as, or close options.
- `excel.calculate` is the workbook calculation owner. It may recalculate,
  fully calculate, or full-rebuild calculate formulas, and must gate
  `full_rebuild` behind `ExcelApi 1.2`.
- `excel.get_used_range` locates the occupied sheet area. It does not return
  cell contents; callers use `excel.read_range` for values, text, formulas, and
  number formats.
- `excel.find_replace_cells` is the Excel search owner. Omitting replacement
  arguments is the read-only find idiom; supplying replacement arguments makes
  it the range replacement owner.
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

Action side-effect maps:

| Tool | Read actions | Edit actions | Destructive actions |
|---|---|---|---|
| `excel.find_replace_cells` | omitted `replace` / find-only mode | replace mode | - |
| `excel.update_table` | `metadata`, `read` | `add_rows`, `add_columns`, `resize`, `rename`, `options`, `style` | `delete` |
| `excel.update_chart` | `metadata`, `read`, `export_image` | `title`, `legend`, `axis`, `data`, `series_source`, `position`, `size` | `delete` |
| `excel.update_pivot_table` | `metadata`, `read` | `refresh`, `add_hierarchy`, `remove_hierarchy`, `layout`, `filter`, `clear_filters` | `delete` |

These maps are part of the public contract described in
[03-mcp-tool-surface.md](03-mcp-tool-surface.md) §9. The daemon MUST expose the
same maps in MCP metadata and enforce Global Tool Access at action granularity
before dispatching to the Excel add-in.

API shape rules:

- Prefer explicit `action` fields for object-owner update tools when one tool
  owns multiple lifecycle operations, for example `{ "action": "rename" }`,
  `{ "action": "resize" }`, or `{ "action": "delete" }`.
- Prefer narrow structured option objects over free-form property bags. This
  keeps host validation understandable and avoids exposing unsupported Excel.js
  properties as accidental public API.
- Read operations must return metadata, addresses, dimensions, and object IDs
  where useful, but must not return large cell contents unless the tool is
  explicitly `excel.read_range`.
- Destructive actions must be explicit in the tool arguments and must set
  destructive metadata in the daemon catalog.

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

Pass `formula` to write the same formula to every cell in the target range, or
pass `formulas` as a two-dimensional matrix that exactly matches the target
range shape. Formula syntax is the host Excel formula syntax.

Matrix example:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "C2:D3",
  "formulas": [["=A2+B2", "=A2-B2"], ["=A3+B3", "=A3-B3"]]
}
```

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
  "number_format": "General",
  "horizontal_alignment": "center",
  "vertical_alignment": "top",
  "wrap_text": true,
  "borders": [
    { "side": "top", "style": "continuous", "weight": "thin", "color": "#17202A" }
  ],
  "autofit_columns": true
}
```

All formatting fields are optional. `number_format`, when supplied, is written
as a scalar matrix matching the target range dimensions. `number_formats` may be
supplied instead as a two-dimensional matrix that exactly matches the target
range shape. Supported alignment, border side, border style, and border weight
values are restricted to stable Office.js enum values. `autofit_rows` and
`autofit_columns` require `ExcelApi 1.2`; unsupported hosts return
`HOST_CAPABILITY_UNAVAILABLE` with no partial effect.

Returns `{ "address": "A1:C2", "formatted": true }`.

### 3.6 `excel.sort_range`

Arguments for a plain range sort:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:D20",
  "target_type": "range",
  "fields": [{ "key": 2, "ascending": false, "sort_on": "value" }],
  "has_headers": true,
  "orientation": "rows"
}
```

Arguments for a table sort:

```json
{
  "session_id": "excel-session-id",
  "target_type": "table",
  "table": "SalesTable",
  "fields": [{ "key": 1, "ascending": true }]
}
```

`fields[*].key` is the zero-based column or row offset within the sorted range
or table body. Supported actions are `apply`, `clear`, and `reapply`; `clear`
and `reapply` are table-only because Office.js exposes them on `TableSort`.
Range and table sorting require `ExcelApi 1.2`.

Returns `{ "target_type": "range", "address": "A1:D20", "sorted": true }`.

### 3.7 `excel.apply_filter`

Arguments for a range filter:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:D20",
  "target_type": "range",
  "column_index": 2,
  "criteria": { "filter_on": "values", "values": ["West", "East"] }
}
```

Arguments for a table filter:

```json
{
  "session_id": "excel-session-id",
  "target_type": "table",
  "table": "SalesTable",
  "column": "Region",
  "criteria": { "filter_on": "custom", "criterion1": "=*West" }
}
```

Supported actions are `apply`, `clear`, `remove`, and `reapply`. Range filters
use worksheet `AutoFilter` and require `ExcelApi 1.9`. Table column filters use
the table column `Filter` object and require `ExcelApi 1.2`; clearing or
reapplying a table filter uses the table `AutoFilter` object. PivotTable filters
belong to `excel.update_pivot_table`.

Returns `{ "target_type": "range", "address": "A1:D20", "filtered": true }`.

### 3.8 `excel.create_table`

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

### 3.9 `excel.create_chart`

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

## 4. Contract Ownership Notes

Planned tools should keep contracts coarse and workflow-oriented:

- Workbook and worksheet tools own navigation, sheet lifecycle, and workbook
  metadata. They must not read workbook cell contents.
- Range tools own cell data, formulas, cell-level formatting, search/replace,
  sorting, filtering, and clearing. A single-cell operation is represented as a
  one-cell range.
- `excel.clear_range` is the only destructive cell primitive. It covers
  content/format/all clearing and explicit delete-with-shift modes instead of
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
- Excel tools remain the authoritative execution surface. The daemon also
  exposes a limited read-only MCP resource fallback for clients that can read
  resources but cannot call dynamic app tools. The fallback forwards through
  the same add-in path as Excel tools and must respect daemon Global Tool
  Access policy and session `available_tools` capability checks.
- The Excel read-only resource templates are:

| URI template | Forwarded tool | Purpose |
|---|---|---|
| `office://excel/{session_id}/workbook` | `excel.get_workbook_info` | Workbook metadata and active sheet orientation. |
| `office://excel/{session_id}/sheets` | `excel.list_sheets` | Worksheet inventory. |
| `office://excel/{session_id}/used-range{?sheet}` | `excel.get_used_range` | Used range address and dimensions for the active or named sheet. |
| `office://excel/{session_id}/range/{address}{?sheet}` | `excel.read_range` | Values, display text, formulas, dimensions, and number formats for an explicit range. |

The fallback is intentionally read-only. Workbook mutation, formulas, table
updates, charts, PivotTables, and formatting remain tool-only operations.

## 6. Evidence

Automated evidence exists at two levels:

- Rust daemon/unit evidence proves the Excel tool catalog is listed and that
  MCP calls are forwarded to an Excel add-in session.
- Office tool E2E evidence runs the complete Excel tool catalog in one live
  workbook session with `npm run e2e:tools` from `src/office-ctl/excel` and
  writes `artifacts/office-tool-e2e-excel.json`.

Completion of the live Excel gate requires a passed `office_tool_e2e_report`
validated by `npm run evidence:validate -- --require-office-tool-e2e`.
