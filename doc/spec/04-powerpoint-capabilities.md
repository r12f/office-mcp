# 04 — PowerPoint Capabilities

This document is the PowerPoint v1 capability contract for `office-mcp`.
PowerPoint tools run inside the document-scoped PowerPoint Office.js add-in
under `src/office-ctl/powerpoint` and are routed by the Rust daemon through the
same add-in JSON-RPC channel as Word and Excel tools. The tool list in this file
is the source of truth for the daemon catalog, add-in `available_tools`, task
pane permission UI, and implementation TODO list.

## 1. Scope

v1 supports one connected presentation session per add-in runtime. The daemon
addresses the presentation by `session_id`; the add-in executes presentation
operations with `PowerPoint.run` when the PowerPoint-specific object model owns
the operation, and with the Office Common API when the capability is explicitly
document-level, such as active view and file export.

The target PowerPoint surface is based on Microsoft Learn's PowerPoint add-in
core concepts page:
https://learn.microsoft.com/en-us/office/dev/add-ins/powerpoint/core-concepts.
That page identifies two API layers and the primary object path:

- PowerPoint JavaScript API owns strongly typed presentation, slide, table,
  shape, and formatting objects.
- Office Common API owns runtime context, requirements probing, document file
  access, and active view detection.
- A `Presentation` contains slides plus presentation-level entities such as
  settings and custom XML parts.
- A `Slide` contains content such as shapes, text, and tables.
- A `Layout` determines how slide content is organized and displayed.

`office-mcp` exposes workflow-oriented tools for those objects. It must not
mirror every PowerPoint.js class, property, or method. The target v1 catalog is
25 tools: large enough to cover presentation orientation, slide CRUD, layout,
selection, shape/text/table authoring, visual formatting, metadata, and export,
but small enough to keep each tool's owner and permission profile clear.

Selection rules for the 25-tool budget:

- Start from the Microsoft Learn object path: `Presentation` -> `Slide` ->
  `Shape` / `TextRange` / `Table`, with `Layout` as the slide organization
  owner.
- Add a tool only when it represents a user-level task, not a single Office.js
  method.
- Merge lifecycle and configuration operations into object-owner update tools
  when they share the same owner and permission profile.
- Keep text-in-shape operations as text-range operations; do not add separate
  placeholder text, title text, and body text tools unless a future revision
  proves that the generic text owner is unsafe.
- Keep table row, column, cell value, merge, clear, and style operations under
  the table owner. Do not add separate table-row or table-cell tools.
- Keep shape move, resize, rotate, fill, line, z-order, name/alt-text, grouping,
  and deletion under the shape owner. Do not add method-level shape formatting
  tools.
- Include presentation-level metadata, tags, active view, and export
  because those are high-value automation workflows and are explicitly surfaced
  by the PowerPoint core concepts page or the stable Office.js typings.
- Defer charts, SmartArt, media, animations, transitions, comments, speaker
  notes, slide show control, external data, and preview-only APIs until a later
  user workflow proves that the 25-tool surface cannot express the need safely.

Candidate selection matrix before final pruning:

| User workflow | Object owner | Candidate tools | Why this is enough |
|---|---|---|---|
| Inspect and export the deck | `Presentation` / Common `Document` | `powerpoint.get_presentation_info`, `powerpoint.get_active_view`, `powerpoint.export_file` | Covers orientation, edit/read mode, and file/PDF/PPTX export without separate export tools per format. |
| Manage presentation metadata | `TagCollection` | `powerpoint.update_tags` | Keeps app-owned presentation tags under one explicit metadata owner; document-property writes and custom XML are deferred. |
| Read and mutate slides | `SlideCollection` / `Slide` | `powerpoint.list_slides`, `powerpoint.add_slide`, `powerpoint.update_slide`, `powerpoint.delete_slide`, `powerpoint.duplicate_slide`, `powerpoint.move_slide`, `powerpoint.export_slide` | Covers slide inventory, CRUD, ordering, duplication, and single-slide image/PPTX export without adding separate slide selection/export variants. |
| Use layouts and masters | `SlideMaster` / `SlideLayout` | `powerpoint.list_layouts`, `powerpoint.apply_layout`, `powerpoint.set_slide_background` | Layout and background are the slide organization/visual base owners; shape content remains separate. |
| Work from the user's selection | `Presentation` selection APIs | `powerpoint.get_selection`, `powerpoint.set_selection` | One read and one write owner for selected slides/shapes/text ranges. |
| Author and manage shapes | `ShapeCollection` / `Shape` | `powerpoint.list_shapes`, `powerpoint.add_text_box`, `powerpoint.add_shape`, `powerpoint.insert_image`, `powerpoint.update_shape` | Shape creation is split by common user intent; lifecycle/configuration is consolidated under `update_shape`. |
| Edit and format text | `TextRange` / `ShapeFont` / `ParagraphFormat` | `powerpoint.read_text`, `powerpoint.replace_text`, `powerpoint.format_text` | Text reads, replacement, and formatting have distinct validation and permission profiles. |
| Author tables | `Table` | `powerpoint.add_table`, `powerpoint.read_table`, `powerpoint.update_table` | Table creation, reading, and table-owned mutation cover row/column/cell/style/merge/clear without method-level sprawl. |

Candidate tool pressure before Occam reduction:

| Category | Tools | Count | User intent |
|---|---|---:|---|
| Presentation | `powerpoint.get_presentation_info`, `powerpoint.get_active_view`, `powerpoint.export_file` | 3 | Inspect and export the connected deck. |
| Metadata | `powerpoint.update_properties`, `powerpoint.update_tags`, `powerpoint.update_custom_xml` | 3 | Manage presentation-level metadata and structured add-in state. |
| Slides | `powerpoint.list_slides`, `powerpoint.add_slide`, `powerpoint.update_slide`, `powerpoint.delete_slide`, `powerpoint.duplicate_slide`, `powerpoint.move_slide`, `powerpoint.export_slide` | 7 | Slide inventory, lifecycle, ordering, duplication, and slide-level export. |
| Layout | `powerpoint.list_layouts`, `powerpoint.apply_layout`, `powerpoint.set_slide_background` | 3 | Discover and apply slide organization and base visual styling. |
| Selection | `powerpoint.get_selection`, `powerpoint.set_selection` | 2 | Read and update current slide/shape/text selection. |
| Shapes | `powerpoint.list_shapes`, `powerpoint.add_text_box`, `powerpoint.add_shape`, `powerpoint.insert_image`, `powerpoint.update_shape` | 5 | Shape inventory, creation, image insertion, and shape-owned lifecycle/configuration. |
| Text | `powerpoint.read_text`, `powerpoint.replace_text`, `powerpoint.format_text` | 3 | Read, replace, and style text ranges inside shapes. |
| Tables | `powerpoint.add_table`, `powerpoint.read_table`, `powerpoint.update_table` | 3 | Create, inspect, and mutate slide tables. |

Total: 29. The table above intentionally shows candidate group pressure.
The accepted v1 budget is 25 tools, so the final accepted set removes four
candidate tools that overlap with stronger owners:

- Remove `powerpoint.update_properties`; fold safe title/subject/author metadata
  reads into `powerpoint.get_presentation_info` and defer property writes.
- Remove `powerpoint.update_custom_xml`; custom XML is high-power structured
  metadata and is deferred until a concrete integration workflow needs it.
- Remove `powerpoint.duplicate_slide`; the current `@types/office-js` baseline
  does not expose a stable PowerPoint slide duplication method. Add it only if a
  later host-supported path is proven.
- Remove `powerpoint.set_slide_background`; background changes belong to
  `powerpoint.update_slide` because the slide owns its background.

Accepted v1 tool set:

| Category | Tools | Count |
|---|---|---:|
| Presentation | `powerpoint.get_presentation_info`, `powerpoint.get_active_view`, `powerpoint.export_file` | 3 |
| Metadata | `powerpoint.update_tags` | 1 |
| Slides | `powerpoint.list_slides`, `powerpoint.add_slide`, `powerpoint.update_slide`, `powerpoint.delete_slide`, `powerpoint.move_slide`, `powerpoint.export_slide` | 6 |
| Layout | `powerpoint.list_layouts`, `powerpoint.apply_layout` | 2 |
| Selection | `powerpoint.get_selection`, `powerpoint.set_selection` | 2 |
| Shapes | `powerpoint.list_shapes`, `powerpoint.add_text_box`, `powerpoint.add_shape`, `powerpoint.insert_image`, `powerpoint.update_shape` | 5 |
| Text | `powerpoint.read_text`, `powerpoint.replace_text`, `powerpoint.format_text` | 3 |
| Tables | `powerpoint.add_table`, `powerpoint.read_table`, `powerpoint.update_table` | 3 |

Total: 25 tools.

Rejected v1 expansions:

- No separate `powerpoint.export_pdf` after refinement. It is superseded by
  `powerpoint.export_file` with `format: "pdf" | "pptx"` so export behavior has
  one owner.
- No separate `powerpoint.apply_shape_fill`, `powerpoint.move_shape`,
  `powerpoint.resize_shape`, `powerpoint.delete_shape`, or
  `powerpoint.group_shapes`; those belong to `powerpoint.update_shape`.
- No separate `powerpoint.add_row`, `powerpoint.add_column`,
  `powerpoint.update_cell`, `powerpoint.merge_cells`, or
  `powerpoint.clear_table`; those belong to `powerpoint.update_table`.
- No separate title/body/placeholder text tools. Text inside shapes is owned by
  `powerpoint.read_text`, `powerpoint.replace_text`, and
  `powerpoint.format_text`.
- No `powerpoint.save` in v1. `@types/office-js` exposes stable save APIs for
  Word and Excel, but not a PowerPoint presentation save API. Settings
  `saveAsync` is only for the add-in settings bag and must not be presented as a
  deck save.
- No `powerpoint.duplicate_slide` in v1 unless a stable host-supported
  duplication path is proven against typings and live PowerPoint evidence.
- No chart, SmartArt, media, animation, transition, comments, speaker notes,
  slide show, PowerPoint Designer, or macro tools in v1.

## 2. Target Tool Catalog

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `powerpoint.get_presentation_info` | implemented | Presentation | read | `PowerPointApi 1.0`; richer properties require `PowerPointApi 1.5+` / `1.7` | Return presentation title/id, host metadata, slide count, selection summary, active view when available, and capability gates. |
| `powerpoint.get_active_view` | implemented | Presentation | read | Common `Office.Document.getActiveViewAsync` | Return whether the presentation is in editable or read-only presentation view. |
| `powerpoint.export_file` | implemented | Presentation | read | Common `Office.Document.getFileAsync`; PDF/PPTX format support is host-gated | Export the current presentation as PDF or PPTX base64 slices through one export owner, or return explicit host-capability rejection where the host blocks export. |
| `powerpoint.update_tags` | implemented | Metadata | read/edit/destructive | `PowerPointApi 1.3` | Read, set, or delete presentation tags. |
| `powerpoint.list_slides` | implemented | Slides | read | `PowerPointApi 1.2` | List slides with id, index, layout/master ids, shape count, tags, and optional thumbnail metadata. |
| `powerpoint.add_slide` | implemented | Slides | edit | `PowerPointApi 1.3` | Add a slide, optionally using a layout and initial title/body text boxes. |
| `powerpoint.update_slide` | implemented | Slides | edit | `PowerPointApi 1.3`; background requires `PowerPointApi 1.10` | Update slide tags and background where supported; slide content belongs to shape/text/table tools. |
| `powerpoint.delete_slide` | implemented | Slides | destructive | `PowerPointApi 1.3` | Delete a slide, rejecting attempts that would leave the deck empty. |
| `powerpoint.move_slide` | implemented | Slides | edit | `PowerPointApi 1.8` | Move a slide to a target index. |
| `powerpoint.export_slide` | implemented | Slides | read | `PowerPointApi 1.8` | Export one slide as PNG image data or a one-slide presentation, depending on arguments and host support. |
| `powerpoint.list_layouts` | implemented | Layout | read | `PowerPointApi 1.3` | List slide masters and layouts with id, name, and type. |
| `powerpoint.apply_layout` | implemented | Layout | edit | `PowerPointApi 1.8` | Apply a layout to a target slide by id, name, or type. |
| `powerpoint.get_selection` | implemented | Selection | read | `PowerPointApi 1.5` | Return selected slides, shapes, and text range metadata. |
| `powerpoint.set_selection` | implemented | Selection | edit | `PowerPointApi 1.5` | Select slides or a text range; shape selection support is host-gated and must be verified before advertisement. |
| `powerpoint.list_shapes` | implemented | Shapes | read | `PowerPointApi 1.3`; richer shape metadata requires `PowerPointApi 1.4+` | List shapes on a slide with id, type, position, size, text/table presence, and accessibility metadata. |
| `powerpoint.add_text_box` | implemented | Shapes | edit | `PowerPointApi 1.4` | Add a text box to a slide. This supersedes title/body-specific text insertion tools. |
| `powerpoint.add_shape` | implemented | Shapes | edit | `PowerPointApi 1.4` | Add a geometric shape or line to a slide with explicit type and geometry. |
| `powerpoint.insert_image` | implemented | Shapes | edit | Current implementation uses Common API image insertion; shape-owned implementation requires PowerPoint image API verification | Insert an image on a slide or current selection from validated base64 or daemon-fetched HTTPS URL. |
| `powerpoint.update_shape` | implemented | Shapes | edit/destructive | `PowerPointApi 1.4`; grouping and advanced fields require higher sets | Read/update shape position, size, rotation, name, alt text, fill, line, z-order, grouping, and delete through one shape owner. |
| `powerpoint.read_text` | implemented | Text | read | `PowerPointApi 1.4` | Read text from selected text, a shape, one slide, or all slides. |
| `powerpoint.replace_text` | implemented | Text | edit | `PowerPointApi 1.4` | Replace matching text in slide shape text ranges, with optional scope and dry-run support. |
| `powerpoint.format_text` | implemented | Text | edit | `PowerPointApi 1.4`; advanced font features require `PowerPointApi 1.8+` | Apply font and paragraph formatting to selected text or a target shape text range. |
| `powerpoint.add_table` | implemented | Tables | edit | `PowerPointApi 1.8` | Add a table to a slide with initial dimensions and optional values, or return explicit host-capability rejection where table objects are unavailable. |
| `powerpoint.read_table` | implemented | Tables | read | `PowerPointApi 1.8` | Read table dimensions, values, merged areas, and basic cell metadata. |
| `powerpoint.update_table` | implemented | Tables | edit/destructive | `PowerPointApi 1.8`; rows/columns/clear require `PowerPointApi 1.9` | Update table cell values, row/column dimensions, add/delete rows/columns, merge cells, clear values/formatting, apply style settings, or delete the table shape. |

The tools above are the target PowerPoint v1 contract. Implementation work must
keep the daemon catalog, MCP `tools/list`, PowerPoint task pane
`available_tools`, task pane permission grouping, documentation, and tests
aligned with this 25-tool surface. Before implementing or changing a tool,
verify its minimum requirement set against `@types/office-js` and Microsoft API
docs, then land the change as a test-first implementation slice with daemon
catalog coverage, task pane contract coverage, and live PowerPoint smoke
evidence where the host API cannot be fully proven statically.

## 3. Tool Ownership Rules

- One common user intent has one tool owner. Add a new tool only when it has a
  different object owner, permission profile, or user-visible result.
- `powerpoint.get_presentation_info` is orientation only. It may include counts
  and capability gates, but detailed slide inventory belongs to
  `powerpoint.list_slides`. Its `slide_count` field must be populated from the
  current presentation slide collection when that collection is available, and
  must match the number of entries returned by `powerpoint.list_slides` for the
  same session.
- `powerpoint.export_file` owns full-presentation export for PDF/PPTX. Do not
  keep a separate `powerpoint.export_pdf` in the refined catalog.
- `office-mcp` must not advertise a PowerPoint deck save tool until stable
  typings expose a host-supported presentation save API. Office settings
  persistence is not deck save behavior.
- `powerpoint.export_slide` owns single-slide export. Full-deck export belongs
  to `powerpoint.export_file`.
- `powerpoint.update_slide` owns slide-level metadata, tags, hidden state where
  supported, and background. It must not edit shapes or text content.
- `powerpoint.apply_layout` owns layout assignment. `powerpoint.list_layouts`
  owns layout discovery. Do not add layout-name-specific tools.
- `powerpoint.update_shape` owns shape lifecycle and visual geometry. It must
  not edit text contents or table contents except when deleting the owning
  shape explicitly.
- `powerpoint.read_text`, `powerpoint.replace_text`, and
  `powerpoint.format_text` own text inside shapes. Do not add separate title,
  subtitle, body, placeholder, or selected-text mutation tools.
- `powerpoint.update_table` owns table row, column, cell, merge, style, clear,
  and table deletion behavior. `powerpoint.read_table` owns table value reads.
- `powerpoint.update_tags` owns presentation tags only; it must not become a
  generic metadata bag for custom XML or arbitrary document properties.

## 4. API Shape Rules

- Prefer explicit `action` fields for object-owner update tools when one tool
  owns multiple lifecycle operations, for example `{ "action": "rename" }`,
  `{ "action": "set_background" }`, `{ "action": "delete" }`, or
  `{ "action": "merge_cells" }`.
- Prefer narrow structured option objects over free-form property bags. This
  avoids exposing unsupported PowerPoint.js properties as accidental public API.
- Read operations must return metadata, ids, indices, dimensions, and object
  owners where useful, but must not return large deck text unless the tool is
  explicitly `powerpoint.read_text`.
- Destructive actions must be explicit in the tool arguments and must set
  destructive metadata in the daemon catalog.
- Host-gated operations must fail with `HOST_CAPABILITY_UNAVAILABLE` when the
  current PowerPoint runtime lacks the required API. Do not silently no-op or
  return partial success.

## 5. Shared Arguments

Most PowerPoint tools accept:

```json
{
  "session_id": "powerpoint-session-id",
  "slide_id": "256#123",
  "slide_index": 0,
  "shape_id": "5"
}
```

- `session_id` is required by the daemon for all presentation-affecting calls.
- Tools that target one slide should accept either `slide_id` or `slide_index`.
  `slide_id` is preferred when clients have it.
- Tools that target a shape should accept `shape_id` and either `slide_id` or
  `slide_index` unless the target is the current selection.
- Selection-scoped tools should make that scope explicit with
  `{ "target": { "kind": "selection" } }`.

The daemon serializes calls per `session_id`; the add-in uses Office.js host
errors for invalid slide IDs, invalid indices, protected presentation failures,
and host-denied operations. These map through the standard error model in
[06-error-model.md](06-error-model.md).

## 6. Limits And Validation

- The daemon validates session existence, session capability, queue depth,
  request size, timeout, and response size before or around forwarding.
- The add-in validates required scalar fields, enum values, coordinate numbers,
  and obvious JSON shape errors.
- PowerPoint remains the authority for slide/layout/shape existence, invalid
  geometry, unsupported views, and protected presentation denial.
- Export operations must enforce the shared `MAX_RESPONSE_BYTES` limit and may
  return chunk metadata or a host-capability error instead of oversized inline
  payloads.
- PowerPoint tools remain the authoritative execution surface. The daemon also
  exposes a limited read-only MCP resource fallback for clients that can read
  resources but cannot call dynamic app tools. The fallback forwards through
  the same add-in path as PowerPoint tools and must respect daemon Global Tool
  Access policy and session `available_tools` capability checks.
- The PowerPoint read-only resource templates are:

| URI template | Forwarded tool | Purpose |
|---|---|---|
| `office://powerpoint/{session_id}/presentation` | `powerpoint.get_presentation_info` | Presentation metadata, counts, selection summary, and capability gates. |
| `office://powerpoint/{session_id}/slides` | `powerpoint.list_slides` | Slide inventory with IDs, indices, layouts, tags, and shape counts. |
| `office://powerpoint/{session_id}/slide/{index}/text` | `powerpoint.read_text` | Text for one slide. |
| `office://powerpoint/{session_id}/slide/{index}/shapes` | `powerpoint.list_shapes` | Shape inventory for one slide. |

The fallback is intentionally read-only. Presentation export, slide mutation,
metadata writes, shape creation or updates, text replacement or formatting, and
table reads or mutations remain tool-only operations.

## 7. Evidence

Automated evidence should exist at two levels:

- Rust daemon/unit evidence proves the PowerPoint tool catalog is listed and
  that MCP calls are forwarded to a PowerPoint add-in session.
- Office tool E2E evidence runs the complete PowerPoint tool catalog in one live
  presentation session with `npm run e2e:tools` from
  `src/office-ctl/powerpoint` and writes
  `artifacts/office-tool-e2e-powerpoint.json`.

Completion of the live PowerPoint gate requires a connected PowerPoint
presentation and a passed `office_tool_e2e_report` validated by
`npm run evidence:validate -- --require-office-tool-e2e`.
