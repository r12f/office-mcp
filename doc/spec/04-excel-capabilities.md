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

The refined v1 target is 38 tools. The original 20-tool core covered workbook
orientation, worksheet lifecycle, range work, formulas, formatting, data
operations, tables, charts, and PivotTables. The workbook owner now also covers
stable persistence/calculation tasks, named-item lifecycle, and document
metadata. These cannot be represented by range or object-owner tools: saving
the current workbook, forcing calculation, managing workbook/sheet scoped names
that formulas and templates address indirectly, and reading or writing workbook
properties that document-management workflows key off. The Review owner now
covers threaded cell comments because comments are workbook review objects with a distinct
permission profile from range content. The Range owner now also covers
`excel.insert_range`, the non-destructive structural complement to
`excel.clear_range` delete-with-shift, so agents can insert cells, rows, and
columns without emulating that workflow through reads and writes. The Range
owner now also covers `excel.set_hyperlink` because cell hyperlinks are visible
cell metadata that need URL security validation and hyperlink-only clearing that
cannot be represented by generic value, formula, or formatting writes. The
Range owner now also covers `excel.set_data_validation` because in-cell
dropdowns and input constraints are range-owned metadata with typed rule
validation and readback needs before mutation. The Range owner now also covers
`excel.copy_range` because copy/autofill workflows must preserve formulas,
formats, relative references, and host fill semantics that cannot be safely
emulated through client-side read/write loops. The
Format owner now also covers conditional formatting because rule-based visual
formatting has a distinct lifecycle from static one-shot range formatting while
still sharing the cell-format permission profile. The Shapes owner now covers
worksheet image insertion, shape discovery, and shape update/delete workflows
because floating worksheet objects are not range, chart, or table content and
need their own geometry, metadata, text, and z-order contract. Do not
expand the catalog by copying individual Excel.js methods into MCP tools. A
future tool can only be added after the selection matrix proves that the
existing tools cannot express a distinct object owner, permission profile, or
user-visible workflow safely.

The v1 priority order is: workbook orientation, worksheet lifecycle, range/cell
data CRUD, formulas, formatting, sort/filter, tables, charts, and PivotTables.
This follows the Excel core concepts path from workbook to worksheet to range,
then to table and chart objects; PivotTables are included as the only extra v1
analysis object because summarized workbook analysis is a core Excel user need.

Selection rules for the core and follow-up tool budget:

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

Out of scope for the core Excel surface: legacy notes, slicers as first-class
tools, events/subscriptions, custom XML, external data
connections, Power Query, Python/preview-only APIs, OLAP/Power Pivot, arbitrary
file import/export, workbook close, and save-as flows. Workbook save and
calculation are in scope because they are stable workbook-owner operations, not
file export, save-as, or close workflows. Threaded comments are in scope as the
Review follow-up surface; legacy notes remain deferred until ExcelApi 1.18 host
coverage is verified. Range insertion is in scope as the structural complement
to range deletion because `Range.insert` is stable in `ExcelApi 1.1` and owns a
common cell/row/column workflow that cannot be represented by clearing. Cell
layout controls such as merge/unmerge, fixed row or column size, row or column
visibility, and named style application remain in scope as extensions of
`excel.format_range` because they share the existing RangeFormat owner and
formatting permission profile instead of introducing a new object owner. Cell
hyperlinks are in scope as a Range follow-up because Office.js exposes them as
`Range.hyperlink`, they require URL scheme validation, and agents need a
hyperlink-only clear path that preserves displayed cell text.
Data validation is in scope as a Range follow-up because Office.js exposes it
through `Range.dataValidation`, in-cell dropdowns and input rules are common
spreadsheet authoring workflows, and agents need readback before overwriting
cell input constraints.
Shapes and images are in scope as a Shapes follow-up because Office.js exposes
worksheet floating objects through `Worksheet.shapes`, image insertion is a
common workbook composition workflow, and shape geometry/metadata/text/z-order
cannot be represented by range, table, chart, or PivotTable owner tools.

Core tool selection matrix:

| User workflow | Object owner | v1 tools | Why this is enough |
|---|---|---|---|
| Inspect, persist, calculate, address workbook state, and manage workbook metadata; navigate sheets | `Workbook` / `Workbook.names` / `Workbook.properties` / `Worksheet` | `excel.get_workbook_info`, `excel.save`, `excel.calculate`, `excel.list_named_items`, `excel.update_named_item`, `excel.get_document_properties`, `excel.update_document_properties`, `excel.list_sheets`, `excel.add_sheet`, `excel.update_sheet`, `excel.delete_sheet` | Covers workspace-level sheet CRUD, worksheet view state, workbook orientation, style-name discovery, persistence, formula recalculation, template-stable named item addressing, and workbook core/custom document metadata without reading cell contents. |
| Locate and mutate cell data | `Range` | `excel.get_used_range`, `excel.read_range`, `excel.write_range`, `excel.insert_range`, `excel.clear_range`, `excel.find_replace_cells`, `excel.set_hyperlink`, `excel.set_data_validation`, `excel.copy_range` | Cells are one-cell ranges, so separate cell CRUD would duplicate range tools; insertion and delete-with-shift are the structural Range pair, and hyperlink/data-validation/copy/autofill workflows need typed validation while staying range-owned. |
| Work with formulas | `Range` formulas | `excel.set_formula` | Formula writes are distinct from literal value writes and can support scalar or matrix input. |
| Apply user-visible cell formatting | `RangeFormat` / `Range` style, merge, and conditional format APIs | `excel.format_range`, `excel.list_conditional_formats`, `excel.update_conditional_format` | Keeps static cell formatting in `excel.format_range` and rule-based conditional formatting in a dedicated list/update owner so agents can inspect, add, delete, and clear rules without duplicating value or table tools. |
| Sort and filter data | `RangeSort` / table filter owners | `excel.sort_range`, `excel.apply_filter` | One data-operation owner covers both plain ranges and table bodies; table tools must not duplicate it. |
| Promote data into a structured table | `Table` | `excel.create_table`, `excel.update_table` | Creation is common enough to be direct; lifecycle, rows/columns, resize, rename, and table style/options share one object-owner update tool. |
| Visualize data | `Chart` | `excel.create_chart`, `excel.update_chart` | Creation is direct; title, axes, legend, series, size, position, delete, and supported export share one chart-owner update tool. |
| Analyze summarized data | `PivotTable` | `excel.create_pivot_table`, `excel.update_pivot_table` | Creation is direct; fields, filters, aggregation, refresh, metadata, and delete share one PivotTable-owner update tool. |
| Work with floating worksheet objects | `Worksheet.shapes` / `Shape` | `excel.insert_image`, `excel.list_shapes`, `excel.update_shape` | Image insertion is direct; listing, geometry, text, alt text, z-order, and deletion use one Shapes owner so floating objects do not leak into range/chart/table tools. |
| Review workbook cells | `CommentCollection` / `Comment` / `CommentReply` | `excel.add_comment`, `excel.list_comments`, `excel.update_comment` | Threaded comments have a review permission profile and lifecycle separate from range content; legacy notes remain deferred. |

Final v1 tool set by category:

| Category | Tools | Count | User intent |
|---|---|---:|---|
| Workbook | `excel.get_workbook_info`, `excel.save`, `excel.calculate`, `excel.list_named_items`, `excel.update_named_item`, `excel.get_document_properties`, `excel.update_document_properties` | 7 | Inspect workbook-level state, persist changes, recalculate formulas, manage named ranges/items, and read or update workbook metadata without reading cell contents. |
| Worksheet | `excel.list_sheets`, `excel.add_sheet`, `excel.update_sheet`, `excel.delete_sheet` | 4 | Sheet inventory, lifecycle, and worksheet view state. |
| Range / cell data | `excel.get_used_range`, `excel.read_range`, `excel.write_range`, `excel.insert_range`, `excel.clear_range`, `excel.find_replace_cells`, `excel.set_hyperlink`, `excel.set_data_validation`, `excel.copy_range` | 9 | Locate, read, write, insert, clear/delete, search, manage cell hyperlinks and validation rules, and copy/autofill ranges. |
| Formula | `excel.set_formula` | 1 | Author formulas distinctly from literal value writes. |
| Format | `excel.format_range`, `excel.list_conditional_formats`, `excel.update_conditional_format` | 3 | Apply static cell formatting and manage rule-based conditional formatting. |
| Data operations | `excel.sort_range`, `excel.apply_filter` | 2 | Sort and filter plain ranges or table bodies. |
| Table | `excel.create_table`, `excel.update_table` | 2 | Promote data to structured tables and manage table-owned lifecycle/options. |
| Chart | `excel.create_chart`, `excel.update_chart` | 2 | Visualize data and manage chart-owned configuration. |
| PivotTable | `excel.create_pivot_table`, `excel.update_pivot_table` | 2 | Create and configure summarized analysis views. |
| Shapes | `excel.insert_image`, `excel.list_shapes`, `excel.update_shape` | 3 | Insert images and inspect or update floating worksheet shapes. |
| Review | `excel.add_comment`, `excel.list_comments`, `excel.update_comment` | 3 | Create, inspect, reply to, resolve, reopen, edit, and delete threaded cell comments. |

Total: 38 tools.

Rejected tool families for v1:

- No `excel.read_cell`, `excel.write_cell`, or `excel.delete_cell`; cells are
  one-cell ranges.
- No separate worksheet format, freeze pane, protection, legacy notes, slicer,
  event, binding, custom XML, external connection,
  Power Query, workbook import/export, save-as, or close tools. `excel.save`
  intentionally persists the current workbook through host save behavior only;
  it does not expose save-as, export, prompt, path, or close semantics. Freeze
  panes, gridline visibility, and heading visibility are worksheet view state
  owned by `excel.update_sheet` and read back through `excel.list_sheets`, not
  separate tool families.
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
| `excel.get_workbook_info` | read | `ExcelApi 1.1`; style names require `ExcelApi 1.7` | Return workbook identity, active sheet, aggregate object counts, and style names when supported. |
| `excel.save` | edit | `ExcelApi 1.11`; dirty state uses `ExcelApi 1.9` when available | Save the current workbook through host save behavior and report the pre-save dirty state when supported. |
| `excel.calculate` | edit | `ExcelApi 1.1`; `full_rebuild` requires `ExcelApi 1.2` | Recalculate workbook formulas and report the calculation mode. |
| `excel.list_named_items` | read | `ExcelApi 1.1`; sheet-scoped names require `ExcelApi 1.4` | List workbook and/or worksheet scoped named items, including non-range constants and formulas. |
| `excel.update_named_item` | edit/destructive | `ExcelApi 1.4`; editing formulas requires `ExcelApi 1.7` | Add, edit, or delete workbook/sheet scoped named items. |
| `excel.get_document_properties` | read | `ExcelApi 1.7` | Read workbook core document properties, read-only metadata, and optional custom properties. |
| `excel.update_document_properties` | edit | `ExcelApi 1.7` | Update writable workbook core document properties and upsert or delete custom properties. |
| `excel.list_sheets` | read | `ExcelApi 1.1`; freeze panes require `ExcelApi 1.7`; gridlines/headings require `ExcelApi 1.8` | List worksheets with id, name, position, visibility, tab color, active state, and optional view state. |
| `excel.add_sheet` | edit | `ExcelApi 1.1` | Add a worksheet and optionally activate it. |
| `excel.update_sheet` | edit | `ExcelApi 1.1`; freeze panes require `ExcelApi 1.7`; gridlines/headings require `ExcelApi 1.8` | Rename, activate, move, set visibility, set tab color, freeze/unfreeze panes, or set worksheet view flags. |
| `excel.delete_sheet` | destructive | `ExcelApi 1.1` | Delete a worksheet, rejecting attempts that would leave the workbook without sheets. |
| `excel.get_used_range` | read | `ExcelApi 1.1` | Return the used range address and dimensions for a sheet. |
| `excel.read_range` | read | `ExcelApi 1.1`; hyperlink readback requires `ExcelApi 1.7`; data-validation readback requires `ExcelApi 1.8` | Read values, display text, dimensions, number format, and optional hyperlink and data-validation metadata for a range. |
| `excel.write_range` | edit | `ExcelApi 1.1` | Write a two-dimensional values matrix to a range. |
| `excel.insert_range` | edit | `ExcelApi 1.1` | Insert cells, rows, or columns and shift existing content down or right. |
| `excel.clear_range` | destructive | `ExcelApi 1.1` | Clear contents, formats, all range data, or delete cells with a shift direction. |
| `excel.set_hyperlink` | edit | `ExcelApi 1.7` | Set or clear external or in-workbook hyperlinks while preserving displayed text on clear. |
| `excel.set_data_validation` | edit/destructive | `ExcelApi 1.8` | Set or clear range data-validation rules, including list dropdowns, numeric/date/time/text-length constraints, custom formulas, prompts, and error alerts. |
| `excel.copy_range` | edit | `ExcelApi 1.9` | Copy or autofill ranges using host semantics so formulas, references, formats, values, and series fills are preserved. |
| `excel.find_replace_cells` | read/edit | `ExcelApi 1.9` | Find the first matching cell in a range or replace matching cells. |
| `excel.set_formula` | edit | `ExcelApi 1.1` | Fill a range with one formula or a formula matrix. |
| `excel.format_range` | edit | `ExcelApi 1.1`; fixed sizing, hide/unhide, and autofit require `ExcelApi 1.2`; named styles require `ExcelApi 1.7` | Apply font, fill, number formats, borders, alignment, wrapping, merge/unmerge, fixed row or column size, row or column visibility, named styles, and autofit. |
| `excel.list_conditional_formats` | read | `ExcelApi 1.6` | List conditional formatting rules for a range or worksheet. |
| `excel.update_conditional_format` | edit/destructive | `ExcelApi 1.6` | Add, delete, or clear range conditional formatting rules. |
| `excel.sort_range` | edit | `ExcelApi 1.2` | Sort a range or table body by one or more keys. |
| `excel.apply_filter` | edit | Range filters require `ExcelApi 1.9`; table column filters require `ExcelApi 1.2` | Apply, clear, remove, or reapply range and table filters. |
| `excel.create_table` | edit | `ExcelApi 1.1` | Create a workbook table from a range. |
| `excel.update_table` | read/edit/destructive | `ExcelApi 1.1`; visual options require `ExcelApi 1.3`; resize requires `ExcelApi 1.13` | Read or update table structure, options, and lifecycle. |
| `excel.create_chart` | edit | `ExcelApi 1.1` | Create a chart from a range on the active or named worksheet. |
| `excel.update_chart` | read/edit/destructive | `ExcelApi 1.1`; image export requires `ExcelApi 1.2`; axis selection and chart type metadata require `ExcelApi 1.7` | Read or update chart configuration, source, export, and lifecycle. |
| `excel.create_pivot_table` | edit | `ExcelApi 1.8` | Create a PivotTable from a range or table source. |
| `excel.update_pivot_table` | read/edit/destructive | `ExcelApi 1.3`; hierarchy/layout/delete require `ExcelApi 1.8`; filters require `ExcelApi 1.12` | Read or update PivotTable fields, layout, filters, refresh, and lifecycle. |
| `excel.insert_image` | edit | `ExcelApi 1.9` | Insert a PNG or JPEG image as a floating worksheet shape from daemon-validated base64 or URL input. |
| `excel.list_shapes` | read | `ExcelApi 1.9` | List floating worksheet shapes with id, name, type, geometry, alt text, and bounded text previews. |
| `excel.update_shape` | edit/destructive | `ExcelApi 1.9` | Move, resize, update alt text or text, adjust z-order, or delete a floating worksheet shape. |
| `excel.add_comment` | comment | `ExcelApi 1.10` | Add a threaded comment to a cell as the signed-in Office user. |
| `excel.list_comments` | read | `ExcelApi 1.10`; resolved filtering requires `ExcelApi 1.11` | List threaded comments and replies, optionally filtered by resolved state. |
| `excel.update_comment` | comment/destructive | `ExcelApi 1.10`; resolve/reopen require `ExcelApi 1.11` | Reply to, edit, resolve, reopen, or delete a threaded comment or reply. |

The Excel task pane reports these tools in `session.added.available_tools` only
after the runtime registers successfully with the daemon.

Target core Excel tool surface:

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `excel.get_workbook_info` | implemented | Workbook | read | `ExcelApi 1.1`; style names require `ExcelApi 1.7` | Return workbook identity, workbook-level state, active sheet name/id, aggregate object counts, and style names when supported; detailed sheet inventory belongs to `excel.list_sheets`. |
| `excel.save` | implemented | Workbook | edit | `ExcelApi 1.11`; dirty state uses `ExcelApi 1.9` when available | Save the current workbook through host save behavior. Save-as, export, prompts, and close flows remain out of scope. |
| `excel.calculate` | implemented | Workbook | edit | `ExcelApi 1.1`; `full_rebuild` requires `ExcelApi 1.2` | Recalculate workbook formulas with `recalculate`, `full`, or `full_rebuild` mode and return the calculation mode. |
| `excel.list_named_items` | implemented | Workbook | read | `ExcelApi 1.1`; sheet-scoped names require `ExcelApi 1.4` | List workbook and/or worksheet scoped named items, returning range addresses when available and formulas for non-range items. |
| `excel.update_named_item` | implemented | Workbook | edit/destructive | `ExcelApi 1.4`; editing formulas requires `ExcelApi 1.7` | Add, edit, or delete named items. Duplicate adds and unknown edit/delete targets fail deterministically before mutation. |
| `excel.get_document_properties` | implemented | Workbook | read | `ExcelApi 1.7` | Read workbook core document properties, read-only metadata, and optional custom properties. |
| `excel.update_document_properties` | implemented | Workbook | edit | `ExcelApi 1.7` | Update writable workbook core document properties and upsert or delete custom properties. |
| `excel.list_sheets` | implemented | Worksheet | read | `ExcelApi 1.1`; freeze panes require `ExcelApi 1.7`; gridlines/headings require `ExcelApi 1.8` | List worksheets with id, name, position, visibility, tab color, active state, and optional view state; workbook metadata belongs to `excel.get_workbook_info`. |
| `excel.add_sheet` | implemented | Worksheet | edit | `ExcelApi 1.1` | Add a worksheet and optionally activate it. |
| `excel.update_sheet` | implemented | Worksheet | edit | `ExcelApi 1.1`; freeze panes require `ExcelApi 1.7`; gridlines/headings require `ExcelApi 1.8` | Rename, activate, move, set visibility, set tab color, freeze/unfreeze panes, or set worksheet view flags. |
| `excel.delete_sheet` | implemented | Worksheet | destructive | `ExcelApi 1.1` | Delete a worksheet, rejecting attempts that would leave the workbook without sheets. |
| `excel.get_used_range` | implemented | Range | read | `ExcelApi 1.1` | Return the used range address and dimensions for a sheet; cell values/text/formulas belong to `excel.read_range`. |
| `excel.read_range` | implemented | Range | read | `ExcelApi 1.1`; hyperlink readback requires `ExcelApi 1.7`; data-validation readback requires `ExcelApi 1.8` | Read values, display text, formulas, dimensions, number format, and optional hyperlink and data-validation metadata for an explicit range. |
| `excel.write_range` | implemented | Range | edit | `ExcelApi 1.1` | Write a two-dimensional values matrix to a range. |
| `excel.insert_range` | implemented | Range | edit | `ExcelApi 1.1` | Insert cells, whole rows, or whole columns from a range address, shifting existing content down or right. |
| `excel.clear_range` | implemented | Range | destructive | `ExcelApi 1.1` | Clear contents, formats, or all range data; optional cell deletion with shift direction. Hyperlink-only clearing belongs to `excel.set_hyperlink`. |
| `excel.set_hyperlink` | implemented | Range | edit | `ExcelApi 1.7` | Set or clear external or in-workbook cell hyperlinks with deterministic scheme and argument validation. |
| `excel.set_data_validation` | implemented | Range | edit/destructive | `ExcelApi 1.8` | Set or clear range data-validation rules with deterministic type/operator/payload validation and optional prompt/error-alert metadata. |
| `excel.copy_range` | implemented | Range | edit | `ExcelApi 1.9` | Copy a source range to a destination range or autofill a containing destination range with deterministic action and containment validation. |
| `excel.find_replace_cells` | implemented | Range | read/edit | `ExcelApi 1.9` | Search cell contents in a range and optionally replace matches. |
| `excel.set_formula` | implemented | Formula | edit | `ExcelApi 1.1` | Fill a range with one formula or a formula matrix. |
| `excel.format_range` | implemented | Format | edit | `ExcelApi 1.1`; fixed sizing, hide/unhide, and autofit require `ExcelApi 1.2`; named styles require `ExcelApi 1.7` | Apply font, fill, scalar or matrix number formats, borders, alignment, wrapping, merge/unmerge, fixed row or column size, row or column visibility, named styles, and autofit. |
| `excel.list_conditional_formats` | implemented | Format | read | `ExcelApi 1.6` | List conditional formatting rules for an explicit range or the active/named worksheet. |
| `excel.update_conditional_format` | implemented | Format | edit/destructive | `ExcelApi 1.6` | Add typed conditional formatting rules, delete a rule by id, or clear all conditional formats for a range. |
| `excel.sort_range` | implemented | Data | edit | `ExcelApi 1.2` | Sort a range or table body by one or more column keys; table structure changes belong to `excel.update_table`. |
| `excel.apply_filter` | implemented | Data | edit | Range filters require `ExcelApi 1.9`; table column filters require `ExcelApi 1.2` | Apply, clear, remove, or reapply worksheet range or table filter criteria; PivotTable filters belong to `excel.update_pivot_table`. |
| `excel.create_table` | implemented | Table | edit | `ExcelApi 1.1` | Create a workbook table from a range. |
| `excel.update_table` | implemented | Table | read/edit/destructive | `ExcelApi 1.1`; visual options require `ExcelApi 1.3`; resize requires `ExcelApi 1.13` | Read table metadata/structure; add rows/columns; resize, rename, change table style/options, or delete a table. Table cell contents belong to `excel.read_range`. |
| `excel.create_chart` | implemented | Chart | edit | `ExcelApi 1.1` | Create a chart from a range or table. |
| `excel.update_chart` | implemented | Chart | edit/read/destructive | `ExcelApi 1.1`; image export requires `ExcelApi 1.2`; axis selection and chart type metadata require `ExcelApi 1.7` | Read chart metadata; update chart title, axes, legend, source range, position, size, delete the chart, or export a chart image where supported. |
| `excel.create_pivot_table` | implemented | PivotTable | edit | `ExcelApi 1.8` | Create a PivotTable from a range or table at a target destination. |
| `excel.update_pivot_table` | implemented | PivotTable | edit/destructive | `ExcelApi 1.3`; hierarchy/layout/delete require `ExcelApi 1.8`; filters require `ExcelApi 1.12` | Read PivotTable metadata; configure row, column, data, and filter hierarchies; set aggregation and layout options, refresh, apply manual PivotTable filters, clear filters, or delete a PivotTable. |
| `excel.insert_image` | implemented | Shapes | edit | `ExcelApi 1.9` | Insert a daemon-validated PNG or JPEG image onto the active or named worksheet with optional geometry and alt text. |
| `excel.list_shapes` | implemented | Shapes | read | `ExcelApi 1.9` | List worksheet shapes with stable IDs, names, types, geometry, alt text, and bounded text previews. |
| `excel.update_shape` | implemented | Shapes | edit/destructive | `ExcelApi 1.9` | Move, resize, set alt text, set text, set z-order, or delete a worksheet shape. |
| `excel.add_comment` | implemented | Review | comment | `ExcelApi 1.10` | Add a threaded cell comment as the signed-in Office user. |
| `excel.list_comments` | implemented | Review | read | `ExcelApi 1.10`; resolved filtering requires `ExcelApi 1.11` | List threaded cell comments and replies, marking workbook-authored text as untrusted source content. |
| `excel.update_comment` | implemented | Review | comment/destructive | `ExcelApi 1.10`; resolve/reopen require `ExcelApi 1.11` | Reply to, edit, resolve, reopen, or delete a threaded comment or reply. |

The tools above are the Excel v1 contract. Implementation work must keep
the daemon catalog, MCP `tools/list`, Excel task pane `available_tools`, task
pane permission grouping, documentation, and tests aligned with this 38-tool
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
- `excel.list_sheets` and `excel.update_sheet` own worksheet view state.
  `excel.list_sheets` returns `frozen` when `ExcelApi 1.7` is available and
  `show_gridlines` / `show_headings` when `ExcelApi 1.8` is available.
  `excel.update_sheet.freeze` accepts exactly one freeze mode: `rows`,
  `columns`, `at`, or `unfreeze`. Combining modes fails with
  `INVALID_ARGUMENT` and `partial_effect: "none"`. A rows+columns freeze is
  represented by `freeze.at` at the cell below and to the right of the frozen
  panes. Unsupported per-field host APIs fail with `HOST_CAPABILITY_UNAVAILABLE`
  before mutation.
- `excel.save` is the workbook persistence owner. It calls the host save
  behavior for the current workbook only and must not accept path, format,
  prompt, export, save-as, or close options.
- `excel.calculate` is the workbook calculation owner. It may recalculate,
  fully calculate, or full-rebuild calculate formulas, and must gate
  `full_rebuild` behind `ExcelApi 1.2`.
- `excel.list_named_items` and `excel.update_named_item` are the named-item
  owners. They cover workbook-scoped and sheet-scoped names, range-backed names,
  and non-range constants/formulas. Deleting a name is destructive because
  formulas that reference it can break.
- `excel.get_document_properties` and `excel.update_document_properties` own
  workbook document metadata. They cover writable core fields such as title,
  subject, author, keywords, category, comments, company, and manager;
  read-only metadata such as last author, creation date, revision number, and
  last save time; and custom property upsert/delete. Metadata values are
  workbook-authored content and MUST be returned with `untrusted_source: true`.
  Read-only fields are intentionally absent from the update schema.
- `excel.get_used_range` locates the occupied sheet area. It does not return
  cell contents; callers use `excel.read_range` for values, text, formulas, and
  number formats.
- Range tools that accept `address` (`excel.get_used_range` excluded) may accept
  a named item as an address. If an address is not parsed by `Worksheet.getRange`,
  the task pane resolves it through workbook names and then worksheet names where
  a sheet scope is provided. Unknown or non-range named items fail with
  `INVALID_ARGUMENT` and `partial_effect: "none"` before mutation.
- `excel.find_replace_cells` is the Excel search owner. Omitting replacement
  arguments is the read-only find idiom; supplying replacement arguments makes
  it the range replacement owner.
- `excel.set_hyperlink` is the Excel cell hyperlink owner. It handles external
  URL links, in-workbook document references, and hyperlink-only clearing while
  preserving displayed text on clear. `excel.clear_range` must not grow a
  duplicate hyperlink-clear mode.
- `excel.set_data_validation` is the Excel data-validation owner. It handles
  in-cell list dropdowns, numeric/date/time/text-length constraints, custom
  validation formulas, prompts, alerts, and validation clearing. `excel.read_range`
  owns validation readback through `include_validation`; validation formulas,
  prompts, and alert text are workbook-authored untrusted source content.
- `excel.copy_range` is the Excel copy/autofill owner. It handles host-native
  `Range.copyFrom` and `Range.autoFill` workflows, including formula reference
  adjustment, format preservation, linked copies, transposition, skipped blanks,
  and series fills. Agents should read destination ranges first when overwrites
  matter; destination overwrites are inherent to the host copy/fill operations.
- `excel.write_range` writes literal values. `excel.set_formula` writes formulas
  and formula matrices. Formula strings passed to `excel.write_range` remain
  literal input unless explicitly documented otherwise during implementation.
- `excel.format_range` owns cell-level visual formatting and layout. This
  includes font, fill, number format, borders, alignment, wrapping, autofit,
  merge/unmerge, fixed row or column size, row or column visibility, and named
  style application. Style names are discoverable through
  `excel.get_workbook_info` because they are workbook state, while applying a
  style to cells remains range formatting. Table style/options remain in
  `excel.update_table`; chart visual settings remain in `excel.update_chart`;
  PivotTable layout/format settings remain in `excel.update_pivot_table`.
- `excel.list_conditional_formats` and `excel.update_conditional_format` own
  rule-based cell formatting. Static one-shot formatting remains in
  `excel.format_range`; conditional formatting rule lifecycle must not be
  represented as table, data-filter, or static range-format operations.
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
- `excel.insert_image`, `excel.list_shapes`, and `excel.update_shape` own
  floating worksheet objects. They must not read or mutate cell contents, chart
  internals, table structure, or PivotTable configuration. Image input is
  normalized by the daemon before forwarding, and workbook-authored shape text
  returned by `excel.list_shapes` MUST be marked as untrusted source content.
- `excel.add_comment`, `excel.list_comments`, and `excel.update_comment` own
  threaded cell comments. Comments are authored as the signed-in Office user;
  comment and reply content is workbook content and MUST be returned with
  `untrusted_source: true`. Legacy notes (`Workbook.notes`, ExcelApi 1.18) are
  not part of this owner until host coverage is verified.

Action side-effect maps:

| Tool | Read actions | Edit actions | Destructive actions |
|---|---|---|---|
| `excel.find_replace_cells` | omitted `replace` / find-only mode | replace mode | - |
| `excel.copy_range` | - | `copy`, `autofill` | - |
| `excel.update_document_properties` | - | core property writes, `custom_set`, `custom_delete` | - |
| `excel.set_hyperlink` | - | `set`, `clear` | - |
| `excel.set_data_validation` | - | `set` | `clear` |
| `excel.update_conditional_format` | - | `add` | `delete`, `clear_range` |
| `excel.update_table` | `metadata`, `read` | `add_rows`, `add_columns`, `resize`, `rename`, `options`, `style` | `delete` |
| `excel.update_chart` | `metadata`, `read`, `export_image` | `title`, `legend`, `axis`, `data`, `series_source`, `position`, `size` | `delete` |
| `excel.update_pivot_table` | `metadata`, `read` | `refresh`, `add_hierarchy`, `remove_hierarchy`, `layout`, `filter`, `clear_filters` | `delete` |
| `excel.update_shape` | - | `move`, `resize`, `set_alt_text`, `set_text`, `set_z_order` | `delete` |
| `excel.update_comment` | - | `reply`, `edit`, `resolve`, `reopen` | `delete` |

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

### 3.0A `excel.get_document_properties`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "include_custom": true
}
```

Returns writable core fields (`title`, `subject`, `author`, `keywords`,
`category`, `comments`, `company`, and `manager`), read-only metadata
(`last_author`, `revision_number`, `creation_date`, and `last_save_time` when
Office returns them), and, when `include_custom` is not false,
`custom: [{ key, type, value }]`. The response normalizes Office camelCase
names to snake_case and includes `untrusted_source: true` because workbook
metadata and custom property values are workbook-authored content.

### 3.0B `excel.update_document_properties`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "title": "Quarterly Report",
  "subject": "Revenue",
  "author": "Finance Team",
  "keywords": "finance,quarterly",
  "category": "Reporting",
  "comments": "Prepared for review.",
  "company": "Contoso",
  "manager": "Finance Ops",
  "custom_set": [{ "key": "Workflow", "value": "review" }],
  "custom_delete": ["ObsoleteFlag"],
  "validate_only": false
}
```

Callers must provide at least one writable core property, one `custom_set`
entry, or one `custom_delete` key. Requests with no writable operation fail
with `INVALID_ARGUMENT` and `partial_effect: "none"`. Read-only fields such as
`last_author`, `revision_number`, `creation_date`, and `last_save_time` are not
accepted by the schema. `custom_set` upserts by deleting any existing custom
property with the same key before adding the new value through
`CustomPropertyCollection.add(key, value)`. `custom_delete` deletes matching
keys and reports unknown keys as successful no-ops in `custom_missing`.
`deleteAll` is not exposed.

| Operation | API | Requirement set |
|---|---|---|
| Read/write core workbook properties | `Workbook.properties` / `DocumentProperties` | `ExcelApi 1.7` |
| Enumerate custom workbook properties | `DocumentProperties.custom` / `CustomPropertyCollection` | `ExcelApi 1.7` |
| Upsert/delete custom workbook properties | `CustomPropertyCollection.add(key, value)` / `CustomProperty.delete()` | `ExcelApi 1.7` |

### 3.1 `excel.read_range`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:B2",
  "include_hyperlinks": false,
  "include_validation": false
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

`include_hyperlinks` defaults to `false`. When `true`, the response includes a
`hyperlinks` matrix with the same shape as `values`; each cell is either `null`
or `{ "url", "document_reference", "text_to_display", "screen_tip" }`.
Hyperlink readback requires `ExcelApi 1.7`. `include_validation` defaults to
`false`. When `true`, the response includes a `validation` object for the
target range with `{ "type", "rule_summary", "ignore_blanks", "valid" }` when
the host exposes those fields. Data-validation readback requires
`ExcelApi 1.8`. `untrusted_source` is always `true` because workbook content,
including hyperlink URLs, display text, validation formulas, and validation
messages, may contain prompt injection text.

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

### 3.3 `excel.insert_range`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "3:3",
  "shift": "down",
  "count": 2,
  "validate_only": false
}
```

`address` accepts cell ranges such as `B2:C3`, whole-row addresses such as
`3:5`, and whole-column addresses such as `B:B`. `shift` is required and must
be `down` or `right`. Whole-row addresses require `shift: "down"`; whole-column
addresses require `shift: "right"`; incompatible combinations fail with
`INVALID_ARGUMENT` and `partial_effect: "none"` before mutation. `count`
defaults to `1` and expands the target before a single `Range.insert` call, so
multi-row or multi-column insertion remains one host operation. The host adjusts
formulas and shifted references exactly as if the user inserted cells
interactively. Requires `ExcelApi 1.1`.

Returns:

```json
{
  "address": "Sheet1!3:4",
  "shift": "down",
  "count": 2,
  "inserted": true,
  "new_used_range": "Sheet1!A1:D20"
}
```

### 3.4 `excel.set_hyperlink`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "B2:B5",
  "action": "set",
  "url": "https://example.com",
  "text_to_display": "Example",
  "screen_tip": "Open example",
  "validate_only": false
}
```

`action` is required and must be `set` or `clear`. `set` requires exactly one
of `url` or `document_reference`. External URLs must use `https`, `http`, or
`mailto`; other schemes fail with `INVALID_ARGUMENT` and
`partial_effect: "none"` before mutation. `document_reference` maps to the
Office.js hyperlink `subAddress` field for in-workbook targets such as
`Sheet2!A1` or a named range. `text_to_display` and `screen_tip` are optional.
Multi-cell addresses apply the same hyperlink to every cell in the range using
host semantics. `clear` removes hyperlinks only, preserving displayed text, and
must not accept `url`, `document_reference`, `text_to_display`, or `screen_tip`.
The tool requires `ExcelApi 1.7` and supports the shared `validate_only`
mutation preflight contract.

Returns:

```json
{
  "address": "Sheet1!B2:B5",
  "action": "set",
  "updated": true
}
```

### 3.5 `excel.set_data_validation`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "B2:B20",
  "action": "set",
  "rule": {
    "type": "list",
    "list_source": ["Open", "Closed"],
    "in_cell_dropdown": true
  },
  "ignore_blanks": true,
  "error_alert": {
    "style": "stop",
    "title": "Invalid status",
    "message": "Choose a value from the list.",
    "show_alert": true
  },
  "input_prompt": {
    "title": "Status",
    "message": "Pick the current status.",
    "show_prompt": true
  },
  "validate_only": false
}
```

`action` is `set` or `clear`. `set` requires `rule`; `clear` removes validation
rules from the range without changing cell values. Supported rule types are
`list`, `whole_number`, `decimal`, `date`, `time`, `text_length`, and `custom`.
List rules require `list_source`, either an inline non-empty string array or an
A1/named-reference string. Numeric/date/time/text-length rules require an
`operator` and `value1`; `between` and `not_between` also require `value2`.
Custom rules require `formula`. Invalid type/operator/payload combinations fail
with `INVALID_ARGUMENT` and `partial_effect: "none"` before mutation. Requires
`ExcelApi 1.8`.

Returns:

```json
{
  "address": "Sheet1!B2:B20",
  "action": "set",
  "updated": true,
  "validation": {
    "type": "list",
    "rule_summary": { "source": "Open,Closed", "in_cell_dropdown": true },
    "ignore_blanks": true,
    "valid": null,
    "untrusted_source": true
  }
}
```

### 3.6 `excel.copy_range`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "action": "copy",
  "source_sheet": "Sheet1",
  "source_address": "A1:B10",
  "destination_sheet": "Summary",
  "destination_address": "A1:B10",
  "copy_type": "all",
  "skip_blanks": false,
  "transpose": false,
  "autofill_type": "default",
  "validate_only": false
}
```

`action` is `copy` or `autofill`. Both actions require `source_address` and
`destination_address`; `source_sheet` and `destination_sheet` default to the
active worksheet when omitted. `copy` executes host `Range.copyFrom` from the
source to the destination and supports `copy_type` values `all`, `values`,
`formulas`, `formats`, and `link`, plus `skip_blanks` and `transpose`.
`autofill` executes host `Range.autoFill` and supports `autofill_type` values
`default`, `copy`, `series`, `formats`, `values`, and `flash_fill`. The
autofill destination must contain the source range according to the Office.js
contract; deterministic containment failures return `INVALID_ARGUMENT` with
`partial_effect: "none"` before mutation. Requires `ExcelApi 1.9` and supports
the shared `validate_only` mutation preflight contract.

Returns:

```json
{
  "action": "copy",
  "source": "Sheet1!A1:B10",
  "destination": "Summary!A1:B10",
  "copy_type": "all",
  "copied": true
}
```

For `autofill`, the response uses the same source/destination shape and
returns `autofill_type` plus `copied: true`. Workbook-authored formulas copied
or filled by the host remain workbook content; callers that read formulas after
copy/fill receive them through `excel.read_range` with `untrusted_source: true`.

### 3.7 `excel.add_sheet`

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

### 3.8 `excel.set_formula`

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

### 3.9 `excel.format_range`

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
  "merge": "merge_across",
  "column_width_pt": 72,
  "row_height_pt": 24,
  "hidden_columns": false,
  "hidden_rows": false,
  "style": "Heading 1",
  "borders": [
    { "side": "top", "style": "continuous", "weight": "thin", "color": "#17202A" }
  ],
  "autofit_columns": true
}
```

All formatting fields are optional. `number_format`, when supplied, is written
as a scalar matrix matching the target range dimensions. `number_formats` may be
supplied instead as a two-dimensional matrix that exactly matches the target
range shape. Supported alignment, border side, border style, border weight, and
merge values are restricted to stable Office.js enum values. `merge` accepts
`merge`, `merge_across`, or `unmerge`; host merge semantics keep the upper-left
cell value when merged cells overlap existing values, so callers should read the
target range first when data preservation matters. `column_width_pt` and
`row_height_pt` set fixed sizes for every column or row intersecting the target
range. `hidden_columns` and `hidden_rows` hide or unhide every intersecting
column or row. Passing a fixed column width with `autofit_columns: true`, or a
fixed row height with `autofit_rows: true`, fails with `INVALID_ARGUMENT` and
`partial_effect: "none"` before mutation. `style` applies a workbook style name
before explicit formatting fields in the same request, so explicit fields win.
`autofit_rows`, `autofit_columns`, fixed sizing, and hide/unhide require
`ExcelApi 1.2`; `style` requires `ExcelApi 1.7`. Unsupported hosts return
`HOST_CAPABILITY_UNAVAILABLE` with no partial effect.

Returns `{ "address": "A1:C2", "formatted": true }`.

### 3.10 `excel.list_conditional_formats`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:D20"
}
```

`address` is optional. When omitted, the tool lists conditional formats for
the active or named worksheet. When supplied, it lists rules attached to the
target range. Requires `ExcelApi 1.6`.

Returns `{ "conditional_formats": [{ "id", "type", "address", "priority", "stop_if_true", "rule_summary" }], "count": 1, "untrusted_source": true }`.
Rule formulas, text criteria, and addresses are workbook-authored content and
MUST be treated as untrusted source text.

### 3.11 `excel.update_conditional_format`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "address": "A1:D20",
  "action": "add",
  "rule": {
    "type": "cell_value",
    "operator": "less_than",
    "values": [0],
    "format": { "font_color": "#9C0006", "fill_color": "#FFC7CE" }
  },
  "priority": 0,
  "stop_if_true": false,
  "validate_only": false
}
```

Supported actions are `add`, `delete`, and `clear_range`. `add` requires
`address` and `rule`; `delete` requires `id`; `clear_range` requires `address`
and removes every conditional format attached to that target range. `delete`
and `clear_range` are destructive actions because they remove rules; `add` is
mutating. Supported rule types are `cell_value`, `color_scale`, `data_bar`,
`icon_set`, `top_bottom`, `preset_criteria`, `contains_text`, and
`custom_formula`. Invalid rule/type/argument combinations fail with
`INVALID_ARGUMENT` and `partial_effect: "none"` before mutation where they are
deterministically checkable. Requires `ExcelApi 1.6` and supports the shared
`validate_only` mutation preflight contract.

Returns `{ "action": "add", "id": "conditional-format-id", "address": "Sheet1!A1:D20", "updated": true }`.

### 3.12 `excel.sort_range`

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

### 3.13 `excel.apply_filter`

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

### 3.14 `excel.create_table`

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

### 3.15 `excel.create_chart`

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

### 3.16 `excel.add_comment`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "cell": "B2",
  "text": "Review this value"
}
```

`cell` MUST identify a single cell. Comments are authored as the signed-in
Office user. Plain text comments are supported first; mention-rich comment
content is deferred. Requires `ExcelApi 1.10`.

Returns `{ "comment_id": "...", "cell": "Sheet1!B2", "commented": true }`.

### 3.17 `excel.list_comments`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "sheet": "Sheet1",
  "resolved": false
}
```

`sheet` and `resolved` are optional. Omitting `sheet` lists workbook comments;
supplying `sheet` lists comments for that worksheet. `resolved` filters the
returned list and requires the `ExcelApi 1.11` resolved-state capability.

Returns:

```json
{
  "comments": [
    {
      "comment_id": "...",
      "cell": "Sheet1!B2",
      "author_name": "Office User",
      "created": "2026-07-06T00:00:00.000Z",
      "content": "Review this value",
      "resolved": false,
      "replies": [
        {
          "reply_id": "...",
          "author_name": "Office User",
          "created": "2026-07-06T00:01:00.000Z",
          "content": "Updated context"
        }
      ]
    }
  ],
  "count": 1,
  "untrusted_source": true
}
```

`untrusted_source` is always `true` because workbook comments may contain
prompt injection text.

### 3.18 `excel.update_comment`

Arguments:

```json
{
  "session_id": "excel-session-id",
  "comment_id": "comment-id",
  "action": "reply",
  "text": "Reply body",
  "reply_id": "reply-id",
  "validate_only": false
}
```

Supported actions are `reply`, `edit`, `resolve`, `reopen`, and `delete`.
`reply` requires `text`. `edit` requires `text` and edits either the thread
content or the targeted `reply_id`. `delete` removes either the targeted reply
or the whole thread. `resolve` and `reopen` set the thread resolved state and
require `ExcelApi 1.11`; the other actions require `ExcelApi 1.10`. The tool
supports the shared `validate_only` mutation preflight contract.

Returns `{ "comment_id": "comment-id", "action": "reply", "updated": true }`.

## 4. Contract Ownership Notes

Planned tools should keep contracts coarse and workflow-oriented:

- Workbook and worksheet tools own navigation, sheet lifecycle, and workbook
  metadata. They must not read workbook cell contents.
- Range tools own cell data, formulas, cell-level formatting, search/replace,
  sorting, filtering, and clearing. A single-cell operation is represented as a
  one-cell range.
- `excel.insert_range` and `excel.clear_range` are the structural Range pair.
  `excel.insert_range` inserts cells, whole rows, and whole columns with a
  required shift direction; `excel.clear_range` remains the only destructive
  cell primitive and covers content/format/all clearing plus explicit
  delete-with-shift modes instead of adding a separate delete-cell tool.
- `excel.set_hyperlink` owns cell hyperlink writes and hyperlink-only clearing.
  `excel.read_range` owns hyperlink readback through `include_hyperlinks`; the
  default read payload remains unchanged.
- `excel.set_data_validation` owns range validation writes and clearing.
  `excel.read_range` owns validation readback through `include_validation`; the
  default read payload remains unchanged.
- `excel.copy_range` owns range copy and autofill. It uses host copy/fill
  behavior instead of client-side read/write emulation so relative formulas,
  formats, linked references, transposition, skipped blanks, and series fills
  stay consistent with Excel.
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
