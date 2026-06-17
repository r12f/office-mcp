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

## 4. Limits And Validation

- The daemon validates session existence, session capability, queue depth,
  request size, timeout, and response size before or around forwarding.
- The add-in validates required scalar fields and obvious JSON shape errors.
- Excel remains the authority for range validity, formula syntax, table/chart
  constraints, protected workbook denial, and other workbook-specific errors.
- The server does not expose an Excel resource URI surface in v1. Clients use
  `excel.read_range` for workbook reads.

## 5. Evidence

Automated evidence exists at two levels:

- Rust daemon/unit evidence proves the Excel tool catalog is listed and that
  MCP calls are forwarded to an Excel add-in session.
- Runtime evidence supports a live workbook smoke with
  `npm run evidence:excel` from `src/office-mcp/daemon/evidence`, validated with
  `npm run evidence:validate -- --require-excel-smoke`.

Completion of the live Excel gate requires a connected Excel workbook and a
passed `excel.runtime_smoke` report.
