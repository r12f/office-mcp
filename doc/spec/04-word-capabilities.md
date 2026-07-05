# 04 — Word Capabilities (v1)

Per-tool JSON Schemas. All tools take a top-level `session_id` (string, UUID, required).
Common keys (anchor, etc.) are defined in [03-mcp-tool-surface.md §6](03-mcp-tool-surface.md).

The canonical generated tool schemas use JSON Schema 2020-12 and set
`"additionalProperties": false` on argument objects. The fragments below are
assembled with these shared definitions:

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "$defs": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": {
      "oneOf": [
        { "type": "object", "required": ["kind"], "properties": { "kind": { "const": "selection" } }, "additionalProperties": false },
        { "type": "object", "required": ["kind"], "properties": { "kind": { "enum": ["start_of_document", "end_of_document"] } }, "additionalProperties": false },
        { "type": "object", "required": ["kind", "index"], "properties": { "kind": { "enum": ["paragraph_index", "before_paragraph_index", "after_paragraph_index"] }, "index": { "type": "integer", "minimum": 0 } }, "additionalProperties": false },
        { "type": "object", "required": ["kind", "text"], "properties": { "kind": { "enum": ["after_text", "before_text"] }, "text": { "type": "string", "minLength": 1 }, "occurrence": { "type": "integer", "minimum": 1, "default": 1 } }, "additionalProperties": false },
        { "type": "object", "required": ["kind", "text"], "properties": { "kind": { "const": "heading" }, "text": { "type": "string", "minLength": 1 }, "level": { "type": "integer", "minimum": 1, "maximum": 9 } }, "additionalProperties": false },
        { "type": "object", "required": ["kind", "name"], "properties": { "kind": { "const": "bookmark" }, "name": { "type": "string", "minLength": 1 } }, "additionalProperties": false }
      ]
    },
    "extent": { "enum": ["paragraph", "sentence", "selection"] },
    "run_formatting": {
      "type": "object",
      "properties": {
        "bold": { "type": "boolean" },
        "italic": { "type": "boolean" },
        "underline": { "type": "boolean" },
        "strikethrough": { "type": "boolean" },
        "font_name": { "type": "string", "minLength": 1 },
        "font_size_pt": { "type": "number", "exclusiveMinimum": 0 },
        "color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
        "highlight": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" }
      },
      "additionalProperties": false
    },
    "paragraph_formatting": {
      "type": "object",
      "properties": {
        "alignment": { "enum": ["left", "center", "right", "justified"] },
        "left_indent_pt": { "type": "number", "minimum": 0 },
        "right_indent_pt": { "type": "number", "minimum": 0 },
        "first_line_indent_pt": { "type": "number" },
        "line_spacing_pt": { "type": "number", "exclusiveMinimum": 0 },
        "line_unit_before": { "type": "number", "minimum": 0 },
        "line_unit_after": { "type": "number", "minimum": 0 },
        "space_before_pt": { "type": "number", "minimum": 0 },
        "space_after_pt": { "type": "number", "minimum": 0 },
        "outline_level": { "type": "integer", "minimum": 0, "maximum": 9 }
      },
      "additionalProperties": false
    }
  }
}
```

The `$ref` values below refer to these shared definitions.

## 1. Overview

### 1.1 Word tool catalog

The current advertised Word v1 tool surface has 54 tools, grouped by object-owner
category. Categories are not permission tiers and must not be action buckets such
as `Read`, `Insert`, and `Edit`; side-effect level is tracked separately per
tool so the UI can apply Read/Write/All permission modes without hiding the
object model.

The per-tool JSON Schemas follow in §2-§9.

| Category | Tools |
|---|---|
| **Document & structure** | `word.get_text`, `word.get_outline`, `word.get_header_footer`, `word.update_header_footer`, `word.get_document_properties`, `word.update_document_properties`, `word.insert_break`, `word.list_sections`, `word.update_page_setup`, `word.list_fields`, `word.insert_field`, `word.update_field`, `word.delete_field`, `word.list_styles`, `word.create_style`, `word.update_style`, `word.save` |
| **Range & selection** | `word.get_selection`, `word.set_selection`, `word.get_html`, `word.insert_html`, `word.find_text`, `word.resolve_anchor`, `word.insert_bookmark`, `word.list_bookmarks`, `word.delete_bookmark`, `word.insert_hyperlink`, `word.list_hyperlinks`, `word.remove_hyperlink`, `word.replace_text`, `word.delete_range`, `word.apply_formatting`, `word.apply_style` |
| **Paragraphs & lists** | `word.get_paragraph`, `word.insert_paragraph`, `word.update_paragraph`, `word.insert_list`, `word.list_lists`, `word.update_list` |
| **Tables** | `word.insert_table`, `word.read_table`, `word.update_table` |
| **Media** | `word.insert_image`, `word.list_images`, `word.get_image`, `word.update_image`, `word.list_shapes`, `word.insert_shape`, `word.update_shape`, `word.delete_shape` |
| **Content controls** | `word.list_content_controls`, `word.insert_content_control`, `word.update_content_control`, `word.delete_content_control` |
| **Notes** | `word.insert_note`, `word.list_notes`, `word.update_note`, `word.delete_note` |
| **Review** | `word.add_comment`, `word.resolve_comment`, `word.update_comment`, `word.set_change_tracking`, `word.update_tracked_change` |

### 1.2 Target refined Word tool surface

The target refined Word surface is based on the Microsoft Word add-in object
model: `Document` contains sections and document-level state; a section has a
`Body`; `Body` and `Range` own most text operations; higher-level objects such
as paragraphs, lists, tables, content controls, comments, and tracked changes
own object-specific lifecycle and review workflows.

The target surface has 59 tools. It deliberately consolidates specialized tools
that perform the same user intent under a single owner. Superseded
compatibility tools remain documented below for migration history, but they are
not advertised by the daemon catalog or task pane available-tools metadata.

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `word.get_text` | implemented | Document & structure | read | `WordApi 1.3` | Read paginated document body text; paragraph metadata is optional. |
| `word.get_outline` | implemented | Document & structure | read | `WordApi 1.3` | Read headings and lightweight document structure without body text. |
| `word.get_header_footer` | implemented | Document & structure | read | `WordApi 1.1` | Read section-scoped header or footer text and optional paragraph metadata; non-primary layout validation uses `WordApiDesktop 1.3` when required. |
| `word.update_header_footer` | implemented | Document & structure | edit/destructive | `WordApi 1.1` | Replace, append to, or clear a section-scoped header or footer body; non-primary layout validation uses `WordApiDesktop 1.3` when required. |
| `word.get_document_properties` | implemented | Document & structure | read | `WordApi 1.3` | Read writable core document properties, read-only metadata, and optionally custom properties. |
| `word.update_document_properties` | implemented | Document & structure | edit | `WordApi 1.3` | Update writable core document properties and upsert or delete custom properties. |
| `word.get_paragraph` | implemented | Paragraphs & lists | read | `WordApi 1.3` | Read one paragraph by index, optionally including direct paragraph formatting metadata. |
| `word.find_text` | implemented | Range & selection | read | `WordApi 1.3` | Search text with Word search options and return portable paragraph-relative matches. |
| `word.resolve_anchor` | implemented | Range & selection | read | `WordApi 1.3` | Resolve an anchor to safe diagnostic metadata without returning full document text. |
| `word.insert_bookmark` | implemented | Range & selection | edit | `WordApi 1.4` | Create or move a named bookmark at an anchored range with explicit duplicate handling. |
| `word.list_bookmarks` | implemented | Range & selection | read | `WordApi 1.4` | List bookmark names and bounded location previews without returning full document text. |
| `word.delete_bookmark` | implemented | Range & selection | destructive | `WordApi 1.4` | Delete a bookmark marker without deleting the bookmarked text. |
| `word.get_selection` | implemented | Range & selection | read | `WordApi 1.3` | Read current selection text and simple selection metadata. |
| `word.set_selection` | implemented | Range & selection | edit | `WordApi 1.3` | Resolve an anchor and set the current selection or cursor position. |
| `word.get_html` | implemented | Range & selection | read | `WordApi 1.3`; underlying HTML read APIs are `WordApi 1.1` | Read body or anchored range HTML, subject to the large-result cap. |
| `word.insert_html` | implemented | Range & selection | edit | `WordApi 1.3`; underlying HTML insert APIs are `WordApi 1.1` | Insert sanitized HTML at an anchored range with explicit insert-location semantics. |
| `word.insert_hyperlink` | implemented | Range & selection | edit | `WordApi 1.3` | Create a hyperlink on an anchored range or inserted text with URL scheme validation. |
| `word.list_hyperlinks` | implemented | Range & selection | read | `WordApi 1.3` | List hyperlink text and URLs with paragraph-relative locations. |
| `word.remove_hyperlink` | implemented | Range & selection | edit | `WordApi 1.3` | Remove hyperlink targets from an anchored range while preserving text by default. |
| `word.insert_paragraph` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Insert a paragraph at an anchor; also owns heading insertion through style or heading-level arguments after migration. |
| `word.insert_table` | implemented | Tables | edit | `WordApi 1.3` | Insert a table with optional initial data and style. |
| `word.insert_image` | implemented | Media | edit | `WordApi 1.3` | Insert a validated image from base64 or a daemon-fetched HTTPS URL. |
| `word.list_images` | implemented | Media | read | `WordApi 1.3` | List inline images with paragraph and image indexes, dimensions, alt text, and hyperlink presence. |
| `word.get_image` | implemented | Media | read | `WordApi 1.3` | Export one inline image as base64 with dimensions and metadata, subject to large-result limits. |
| `word.update_image` | implemented | Media | edit/destructive | `WordApi 1.3` | Resize, update metadata, replace bytes, or delete one inline image through a single owner. |
| `word.list_shapes` | implemented | Media | read | `WordApiDesktop 1.2` | List inline and floating body shapes, including text boxes, geometric shapes, pictures, groups, and canvases, with bounded text previews. |
| `word.insert_shape` | implemented | Media | edit | `WordApiDesktop 1.2` | Insert a text box, geometric shape, or floating picture at an anchored paragraph or range. |
| `word.update_shape` | implemented | Media | edit | `WordApiDesktop 1.2` | Update shape text, geometry, alt text, fill, line, wrapping, or visibility through one shape owner. |
| `word.delete_shape` | implemented | Media | destructive | `WordApiDesktop 1.2` | Delete one shape by current shape id. |
| `word.insert_break` | implemented | Document & structure | edit | `WordApi 1.3` | Insert a page, line, or section break at an anchor. |
| `word.list_sections` | implemented | Document & structure | read | `WordApi 1.3` | List document sections with paragraph range and header/footer metadata. |
| `word.update_page_setup` | implemented | Document & structure | edit | `WordApiDesktop 1.3` | Update document or section page setup such as orientation, margins, and page size. |
| `word.list_fields` | implemented | Document & structure | read | `WordApi 1.4` | List document fields with current indices, code, result preview, locked state, and paragraph location hints. |
| `word.insert_field` | implemented | Document & structure | edit | `WordApi 1.5` | Insert curated Word fields, including a default table of contents, at an anchored range. |
| `word.update_field` | implemented | Document & structure | edit | `WordApi 1.5` | Refresh, lock, or unlock one field, or refresh all fields after an expected-count stale check. |
| `word.delete_field` | implemented | Document & structure | destructive | `WordApi 1.5` | Delete one field and its current result by current field index. |
| `word.list_styles` | implemented | Document & structure | read | `WordApi 1.5` | List built-in and custom document styles with type, built-in, in-use, base-style, and priority metadata. |
| `word.create_style` | implemented | Document & structure | edit | `WordApi 1.5` | Create a paragraph, character, table, or list style with optional base style and formatting. |
| `word.update_style` | implemented | Document & structure | edit | `WordApi 1.5` | Update an existing style's base style, font formatting, or paragraph formatting. |
| `word.insert_list` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Insert a numbered or bulleted list. |
| `word.list_lists` | implemented | Paragraphs & lists | read | `WordApi 1.3` | List existing document lists and their paragraph items so callers can address list mutations. |
| `word.update_list` | implemented | Paragraphs & lists | edit/destructive | `WordApi 1.3` | Add items, change item levels, attach/detach paragraphs, and update list level formatting for existing lists. |
| `word.replace_text` | implemented | Range & selection | edit | `WordApi 1.3` | Find and replace text, with dry-run support. |
| `word.update_paragraph` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Replace one paragraph's text wholesale. |
| `word.delete_range` | implemented | Range & selection | destructive | `WordApi 1.3` | Delete an anchored paragraph, sentence, or selection. |
| `word.apply_formatting` | implemented | Range & selection | edit | `WordApi 1.3` | Apply direct run and/or paragraph formatting to an anchored range. |
| `word.apply_style` | implemented | Range & selection | edit | `WordApi 1.3` | Apply an Office style to an anchored range; also owns heading-level changes after migration. |
| `word.read_table` | implemented | Table | read | `WordApi 1.3`; merged-cell diagnostics use `WordApi 1.4` when available | Read table dimensions, header state, cell text, and merged-table diagnostics. |
| `word.update_table` | implemented | Table | edit/destructive | `WordApi 1.3`; `merge_cells` requires `WordApi 1.4` | Update table cells, rows, columns, widths, borders, header state, merge state, table/cell formatting, and table deletion through one table-owner tool. |
| `word.list_content_controls` | implemented | Content controls | read | `WordApi 1.5` | List content controls with id/tag/title/type metadata, without duplicating document text reads. |
| `word.insert_content_control` | implemented | Content controls | edit | `WordApi 1.5` | Create a content control around an anchored range or inserted placeholder content. |
| `word.update_content_control` | implemented | Content controls | edit | `WordApi 1.5` | Update content control metadata, locked state, or contained text through the content-control owner. |
| `word.delete_content_control` | implemented | Content controls | destructive | `WordApi 1.5` | Delete a content control, preserving or deleting contents according to an explicit mode. |
| `word.insert_note` | implemented | Notes | edit | `WordApi 1.5` | Insert a footnote or endnote at an anchored range. |
| `word.list_notes` | implemented | Notes | read | `WordApi 1.5` | List footnotes or endnotes with bounded note text and reference paragraph locations. |
| `word.update_note` | implemented | Notes | edit | `WordApi 1.5` | Replace one footnote or endnote body by current note index. |
| `word.delete_note` | implemented | Notes | destructive | `WordApi 1.5` | Delete one footnote or endnote reference and body by current note index. |
| `word.add_comment` | implemented | Review | comment | `WordApi 1.4` | Add a comment to an anchored range as the signed-in Office user. |
| `word.resolve_comment` | implemented | Review | comment | `WordApi 1.4` | Resolve an existing comment. |
| `word.update_comment` | implemented | Review | comment/destructive | `WordApi 1.4` | Reply to, edit, delete, or reopen an existing comment thread or reply. |
| `word.set_change_tracking` | implemented | Review | edit | `WordApi 1.4` | Set Track Changes mode and report the previous mode. |
| `word.update_tracked_change` | implemented | Review | edit/destructive | `WordApi 1.6` | Accept or reject one tracked change by current index and expected fingerprint, or bulk accept/reject after an expected-count stale check. |
| `word.save` | implemented | Document & structure | edit | `WordApi 1.1` | Save the current document with the host save behavior. |

Superseded target-surface tools:

| Current tool | Target owner | Reason |
|---|---|---|
| `word.insert_heading` | `word.insert_paragraph` | Heading insertion is paragraph insertion with a heading style or level. |
| `word.set_heading_level` | `word.apply_style` | Heading level changes are style changes. |
| `word.update_cell` | `word.update_table` | Cell updates are table-owned mutations. |
| `word.add_row` | `word.update_table` | Row insertion is a table-owned mutation. |
| `word.add_column` | `word.update_table` | Column insertion is a table-owned mutation. |
| `word.format_cell` | `word.update_table` | Cell formatting is table-owned formatting, distinct from generic run formatting. |
| `word.accept_change` / `word.reject_change` | `word.update_tracked_change` | Accept/reject are one tracked-change action with an explicit `action` argument. |
| `word.insert_page_break` | `word.insert_break` | Page breaks are one break kind in the generalized break owner. |
| `word.resize_image` | `word.update_image` | Resize is an inline-image mutation action, not a distinct media owner. |
| `word.delete_image` | `word.update_image` | Delete is an inline-image mutation action with destructive side effect. |

Tool ownership rules:

- One common user intent has one tool owner. Add a new tool only when it has a
  different object owner, permission profile, or user-visible result.
- Read resources and read tools may expose the same underlying document data,
  but mutation tools must not duplicate each other's writes.
- `word.insert_paragraph` owns paragraph creation, including headings. Do not add
  a separate heading insertion tool after migration.
- `word.list_styles`, `word.create_style`, and `word.update_style` own the
  document style catalog. `word.apply_style` owns applying an existing semantic
  Office style to a range, including heading level. `word.apply_formatting`
  owns direct character/run and paragraph layout formatting. Direct paragraph
  layout includes alignment, indentation, spacing, and outline level; named
  style definitions remain owned by the style-catalog tools.
- `word.insert_hyperlink`, `word.list_hyperlinks`, and
  `word.remove_hyperlink` own hyperlink lifecycle. Generic run styling remains
  owned by `word.apply_formatting`, and bookmark creation/deletion remains out
  of scope for these tools.
- `word.read_table` owns table content reads and merged-table diagnostics.
  `word.update_table` owns table structure, cell value, table/cell formatting,
  width, borders, header-row state, merge, and deletion mutations.
- `word.insert_list` owns new list creation. `word.list_lists` owns list
  discovery. `word.update_list` owns mutations of existing list membership,
  item level, and level formatting. Deleting a list item remains paragraph
  deletion and is owned by `word.delete_range`.
- Media tools own inline image lifecycle. `word.insert_image` owns new image
  insertion. `word.list_images` owns inline image discovery, and
  `word.get_image` owns bounded image byte export. `word.update_image` owns
  inline-image resize, alt text, hyperlink, byte replacement, and deletion for
  existing images. `word.list_shapes`, `word.insert_shape`, `word.update_shape`,
  and `word.delete_shape` own desktop-tier floating and shape-backed media,
  including text boxes, geometric shapes, floating pictures, groups, and
  canvases. Inline image tools must not silently mutate floating pictures; shape
  tools must not duplicate inline-picture byte export.
- Content-control tools own content-control lifecycle and metadata. Generic text
  edits inside a known range remain owned by range/paragraph tools unless the
  caller is explicitly targeting a content control.
- Note tools own footnote and endnote lifecycle. Main body range tools may place
  the reference anchor, but note body creation, enumeration, body replacement,
  and deletion are owned by `word.insert_note`, `word.list_notes`,
  `word.update_note`, and `word.delete_note`.
- Review comment tools own comment thread lifecycle. `word.add_comment` creates
  a top-level thread, `word.resolve_comment` remains the compatibility owner for
  resolving a thread, and `word.update_comment` owns replies, edits, deletes,
  and reopening.
- `word.set_change_tracking` owns document Track Changes mode. The session
  descriptor and tracked-change resource may expose the current mode as
  read-only metadata, but mode mutation stays with this review tool.
- `word.update_tracked_change` owns tracked-change accept/reject actions,
  including bulk accept/reject. The tracked-change resource remains the read
  owner for current indices, fingerprints, and counts.
- `word.get_header_footer` and `word.update_header_footer` own header/footer
  body reads and writes. Body, range, and paragraph tools operate on the main
  document body and must not silently reach into headers or footers.
- `word.get_document_properties` and `word.update_document_properties` own
  document metadata. The read tool may return writable core fields, read-only
  metadata such as creation and save timestamps, and custom properties. The
  update tool owns writable core fields plus custom property upsert/delete; it
  must not expose read-only fields as writable arguments.
- `word.insert_break` owns all Word break insertion. The superseded
  `word.insert_page_break` compatibility tool MAY remain callable for older
  clients but MUST NOT be advertised in the daemon catalog or task pane
  available-tools metadata.
- `word.list_sections` owns section structure reads. Header/footer tools own
  section header/footer body content; `word.list_sections` reports only bounded
  structural metadata.
- `word.update_page_setup` owns page layout mutations for the document or a
  single section and is advertised only after a successful
  `WordApiDesktop 1.3` probe.
- Field tools own Word field lifecycle. `word.list_fields` reads current field
  indices and bounded previews; `word.insert_field` owns curated field
  insertion, including the table-of-contents convenience path; `word.update_field`
  owns refresh/lock/unlock; and `word.delete_field` owns removal. Paragraph,
  style, and header/footer tools may create text that fields reference, but they
  must not create, refresh, or delete fields implicitly.

### 1.3 Runtime capability tiers

The base manifest requires `WordApi 1.3`. The add-in probes higher sets at
runtime and advertises only tools whose complete implementation is supported:

| Tier | Requirement | Additional tools/features |
|---|---|---|
| Core | `WordApi 1.3` | text, paragraphs, search, insert/edit, tables at start/end, styles, selection, document properties |
| Review | `WordApi 1.4` | comments, bookmark anchors, bookmark lifecycle tools, and field listing |
| Notes, content controls, fields, and styles | `WordApi 1.5` | footnote/endnote lifecycle, rich-text/plain-text content-control lifecycle, field insertion/update/deletion tools, and style-catalog tools |
| Tracked changes | `WordApi 1.6` | tracked-change resource and accept/reject |
| Checkbox content controls | `WordApi 1.7` | checkbox content-control insertion, checked-state listing, and checked-state updates |
| List content controls | `WordApi 1.9` | dropdown-list and combo-box insertion, item listing, item mutation, and selected-item updates |
| Desktop shapes | `WordApiDesktop 1.2` | shape and text-box listing, insertion, mutation, deletion, and text-box preview reachability |
| Host-specific | explicit successful probe | page setup, active-window, and protection metadata |

Header/footer body access and `word.save` use production Word APIs available
before the base manifest requirement and are available whenever the core tier is
available. Preview APIs are excluded from v1.

`word.update_page_setup` requires `WordApiDesktop 1.3` and is absent from the
catalog when that probe fails. `word.insert_break` and `word.list_sections` are
core-tier tools because their required Office.js APIs are available within the
base Word API tier.

`word.list_shapes`, `word.insert_shape`, `word.update_shape`, and
`word.delete_shape` require `WordApiDesktop 1.2` and are absent from the catalog
when that probe fails, including Word on the web. These tools use the desktop
shape APIs exposed from `Body.shapes`, `Paragraph.shapes`, `Range.shapes`,
`ShapeCollection.getById`, `Paragraph.insertTextBox`,
`Paragraph.insertGeometricShape`, `Paragraph.insertPictureFromBase64`, and the
matching `Range` insertion APIs. Shape ids are stable only within the current
document session and callers must re-run `word.list_shapes` after insertion,
deletion, or grouping changes.

### 1.4 Word resources

Declarative read-only resources (mutations go through tools in §2–§8).
URI scheme and cross-cutting semantics are defined in
[03-mcp-tool-surface.md §1–§3](03-mcp-tool-surface.md).

| URI pattern | Returns | Notes |
|---|---|---|
| `office://word/<session_id>/document?offset=0&limit=200` | Paginated plain text | Honors IRM; denial carries `IRM_DENIED` |
| `office://word/<session_id>/structure` | JSON outline (headings, lists, tables) | Lightweight |
| `office://word/<session_id>/paragraph/<index>` | Single paragraph | |
| `office://word/<session_id>/comments` | All comments JSON | Includes resolved state and reply ids/text so mutation tools can target threads and replies. |
| `office://word/<session_id>/track_changes` | Tracked changes JSON | |
| `office://word/<session_id>/selection` | Currently selected range text + metadata | |

## 2. Read

### 2.1 `word.get_text`

Return the document's plain text.

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "offset": { "type": "integer", "minimum": 0, "default": 0 },
    "limit":  { "type": "integer", "minimum": 1, "maximum": 1000, "default": 200 },
    "include_metadata": { "type": "boolean", "default": false }
  }
}
```

Returns:

```json
{
  "text": "...",
  "paragraph_count": 247,
  "returned_paragraphs": { "offset": 0, "limit": 200 },
  "truncated": false
}
```

When `include_metadata: true`, each paragraph is wrapped:

```json
{
  "paragraphs": [
    { "index": 0, "text": "Title", "style": "Title", "level": 0 },
    { "index": 1, "text": "Para 1", "style": "Normal" }
  ]
}
```

IRM: requires `extract` right.

### 2.2 `word.get_outline`

Return headings + structure, no body text.

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string" },
    "max_level": { "type": "integer", "default": 6 }
  }
}
```

Returns nested tree of `{ text, level, paragraph_index, children: [...] }`.

### 2.3 `word.get_paragraph`

```json
{
  "type": "object",
  "required": ["session_id", "index"],
  "properties": {
    "session_id": { "type": "string" },
    "index": { "type": "integer", "minimum": 0 },
    "include_formatting": { "type": "boolean", "default": false }
  }
}
```

When `include_formatting` is true, the response includes a `formatting` object
with the same direct paragraph-layout fields accepted by
`word.apply_formatting.paragraph`, plus `style` when available. This read-back
shape is intended for round-trip verification; it is not a substitute for
semantic style inspection.

### 2.4 `word.find_text`

```json
{
  "type": "object",
  "required": ["session_id", "query"],
  "properties": {
    "session_id": { "type": "string" },
    "query": { "type": "string" },
    "match_case": { "type": "boolean", "default": false },
    "whole_word": { "type": "boolean", "default": false },
    "wildcards": { "type": "boolean", "default": false },
    "limit": { "type": "integer", "default": 50 }
  }
}
```

`wildcards` selects Word's wildcard syntax; it is not JavaScript regular
expression syntax. Returns matches as
`[{ paragraph_index, occurrence_in_paragraph, text, snippet }]`. Paragraph
location is computed by searching paragraph ranges; character offsets are not
promised because Word does not expose portable document offsets.

### 2.5 `word.get_selection`

```json
{ "type": "object", "required": ["session_id"], "properties": { "session_id": { "type": "string" } } }
```

Returns `{ text, paragraph_count, is_empty }`. A selection may span multiple
paragraphs; portable character offsets are intentionally not exposed.

### 2.5A `word.set_selection`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "$ref": "#/$defs/extent" },
    "mode": { "enum": ["select", "cursor_start", "cursor_end"], "default": "select" }
  },
  "additionalProperties": false
}
```

Resolves the shared Word anchor vocabulary, optionally applies the same extent
rules as other anchored range tools, and calls Word's selection API on the
resolved range. `mode: "select"` highlights and scrolls to the range;
`cursor_start` and `cursor_end` collapse the selection to the range boundary.
The response returns `{ selected_text_preview, paragraph_index, is_empty }`,
where the preview is bounded to avoid turning selection control into a bulk
document-read path.

Although `word.set_selection` does not edit document contents, it changes UI
state and affects later `selection`-anchored mutating calls. It is therefore
classified as `edit` and is not exposed under the Read permission ceiling.

### 2.5B `word.get_html`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "$ref": "#/$defs/extent" }
  },
  "additionalProperties": false
}
```

Omitting `anchor` reads the whole document body with Word's HTML interchange
API. Supplying `anchor` resolves the shared Word anchor vocabulary and applies
the same optional `extent` rules used by other range-scoped tools before
calling the range HTML API. The response returns `{ html, byte_length,
truncated: false }` and is marked as untrusted document content.

`word.get_html` is a rich-content interchange read, not a high-fidelity export
format. Word's HTML conversion is host-defined and may normalize styles,
attributes, whitespace, lists, and tables. Responses are subject to the daemon's
`MAX_RESPONSE_BYTES` cap; oversized results fail with `MAX_RESPONSE_SIZE` and
`max_response_bytes` rather than returning a partial HTML document.

### 2.5C `word.insert_html`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "html"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "html": { "type": "string", "minLength": 1, "maxLength": 1000000 },
    "insert_location": { "enum": ["replace", "before", "after", "start", "end"], "default": "after" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

`word.insert_html` resolves the target anchor and calls Word's HTML insertion
API. `replace` replaces the resolved range; `before` and `after` insert adjacent
to the resolved range; `start` and `end` insert inside the resolved range when
the host supports those locations. `validate_only: true` performs anchor and
HTML-policy validation and returns a no-write verdict.

The daemon and task pane MUST reject unsafe HTML before mutation with
`INVALID_ARGUMENT` and `partial_effect: none`. The policy rejects script blocks,
inline event-handler attributes, `javascript:` URLs, CSS `url(...)` references,
and external resource-bearing attributes that Word would fetch such as `src`,
`srcset`, `poster`, and similar media or stylesheet references. Plain links such
as `https://` anchors are allowed because they are hyperlink targets rather than
resources fetched during insertion. OOXML interchange is out of scope for this
tool pair because it has a larger injection and compatibility surface.

### 2.6 `word.list_bookmarks`

List named bookmark markers without dumping the bookmarked text body.

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "include_hidden": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Returns `{ bookmarks, count }`. Each bookmark includes `name`, a best-effort
`paragraph_index` when the referenced range can be located in the current body,
and a bounded `text_preview` of at most 80 characters. Names are reported as
returned by Word. The tool MUST NOT return full bookmarked text; callers that
need body text use `word.get_text` or `word.resolve_anchor` with a bookmark
anchor.

Hidden bookmarks are omitted by default. `include_hidden: true` passes through
Word's hidden-bookmark enumeration support and may reveal host-generated names;
clients should treat those names as implementation details unless the user
explicitly requested them.

## 3. Insert

### 3.1 `word.insert_paragraph`

```json
{
  "type": "object",
  "required": ["session_id", "text", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "text": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "style": { "type": "string", "default": "Normal" },
    "heading_level": { "type": "integer", "minimum": 0, "maximum": 9 },
    "formatting": { "$ref": "#/$defs/run_formatting" }
  }
}
```

`heading_level` is the target replacement for `word.insert_heading`. Level `0`
inserts body text; levels `1`-`9` apply the corresponding heading style.

### 3.2 `word.insert_heading`

Superseded compatibility contract. This tool is no longer advertised; use
`word.insert_paragraph` with `heading_level`.

```json
{
  "type": "object",
  "required": ["session_id", "text", "level", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "text": { "type": "string" },
    "level": { "type": "integer", "minimum": 1, "maximum": 9 },
    "anchor": { "$ref": "#/$defs/anchor" }
  }
}
```

### 3.3 `word.insert_table`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "rows", "cols"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "rows": { "type": "integer", "minimum": 1 },
    "cols": { "type": "integer", "minimum": 1 },
    "data": {
      "type": "array",
      "items": { "type": "array", "items": { "type": "string" } }
    },
    "header_row": { "type": "boolean", "default": false },
    "style": { "type": "string" }
  }
}
```

If `data` is provided, its dimensions must match `rows × cols`.

### 3.4 `word.insert_image`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "image"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "placement": {
      "enum": [
        "inline",
        "before_paragraph",
        "after_paragraph",
        "new_paragraph_before",
        "new_paragraph_after",
        "replace_paragraph",
        "selection"
      ],
      "default": "inline"
    },
    "image": {
      "oneOf": [
        { "type": "object", "required": ["base64"], "properties": { "base64": { "type": "string" } } },
        { "type": "object", "required": ["url"],    "properties": { "url":    { "type": "string", "format": "uri" } } }
      ]
    },
    "alt_text": { "type": "string" },
    "width_pt":  { "type": "number" },
    "height_pt": { "type": "number" },
    "validate_only": { "type": "boolean", "default": false }
  }
}
```

`url` MUST be `https://` and is fetched server-side (not by the add-in) to avoid
mixed-content issues in the add-in webview. The daemon applies the fetch policy
in [05-security.md §6.1](05-security.md): no cookies, auth headers, private
addresses, unvalidated redirects, oversized bodies, or non-image payloads.
Base64 input is subject to the same decoded-byte and image-format limits.

Before invoking any Office.js mutating API, `word.insert_image` MUST preflight
the post-daemon argument shape. That preflight includes the presence of
`image.base64`, positive `width_pt` and `height_pt` values when provided, and
the compatibility between `placement` and `anchor.kind`. Invalid image
arguments MUST fail with `INVALID_ARGUMENT` and `partial_effect: "none"` before
the add-in queues a Word mutation.

`placement` controls how the resolved anchor is used. `inline` preserves the
legacy behavior and inserts the image at the resolved range, before/after the
range according to the anchor direction. `selection` inserts into the current
selection and requires `anchor.kind="selection"`. `before_paragraph`,
`after_paragraph`, `new_paragraph_before`, `new_paragraph_after`, and
`replace_paragraph` require a paragraph-resolving anchor (`paragraph_index`,
`before_paragraph_index`, `after_paragraph_index`, or `heading`).

For paragraph placements, `before_paragraph` and `after_paragraph` insert the
image directly before or after the resolved paragraph. `new_paragraph_before`
and `new_paragraph_after` insert the image into a clean adjacent paragraph. The
resolved anchor paragraph's text and style MUST remain unchanged; for example,
inserting after a heading must not append the image into the heading paragraph
itself. `replace_paragraph` replaces the resolved paragraph contents with the
image. If the host cannot support the requested paragraph placement without
mutation, the tool MUST return a specific `INVALID_ARGUMENT` error with
`partial_effect: "none"`.

With `validate_only: true`, the tool MUST run the same daemon preprocessing,
argument preflight, anchor resolution, and placement compatibility checks, then
return without inserting an image. A valid response includes `valid: true`,
`operation: "word.insert_image"`, and a `resolved_target` summary with the
resolved anchor kind, target object type when known, and placement. If the
anchor/placement combination is unsupported, the error SHOULD include a
`suggestion.placement` value such as `new_paragraph_after` when that correction
is deterministic.

### 3.5 `word.list_images`

List inline pictures in the main document body without exporting image bytes.

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string" }
  },
  "additionalProperties": false
}
```

The response shape is:

```json
{
  "images": [
    {
      "paragraph_index": 0,
      "image_index": 0,
      "width_pt": 96,
      "height_pt": 48,
      "alt_text_title": "Logo",
      "alt_text_description": "Company logo",
      "has_hyperlink": true
    }
  ],
  "count": 1
}
```

`paragraph_index` is the current zero-based body paragraph index, and
`image_index` is the zero-based inline-picture index within that paragraph.
Callers must re-run `word.list_images` after insertion, replacement, or
deletion before using previously observed indexes.

### 3.6 `word.get_image`

Export one inline picture by the same locator used by `word.update_image`.

```json
{
  "type": "object",
  "required": ["session_id", "image"],
  "properties": {
    "session_id": { "type": "string" },
    "image": {
      "type": "object",
      "required": ["kind", "index"],
      "properties": {
        "kind": { "const": "paragraph_index" },
        "index": { "type": "integer", "minimum": 0 },
        "image_index": { "type": "integer", "minimum": 0, "default": 0 }
      },
      "additionalProperties": false
    }
  },
  "additionalProperties": false
}
```

The tool returns `base64`, `width_pt`, `height_pt`, `alt_text_title`,
`alt_text_description`, and `has_hyperlink`. Because this exports document
content, it requires the extract/read permission ceiling used for content reads.
The add-in MUST reject responses that exceed the daemon's large-result limit
with a structured result-size error instead of allowing an oversized transport
payload.

### 3.7 `word.update_image`

Mutate one existing inline image by current paragraph/image index.

```json
{
  "type": "object",
  "required": ["session_id", "image", "action"],
  "properties": {
    "session_id": { "type": "string" },
    "image": {
      "type": "object",
      "required": ["kind", "index"],
      "properties": {
        "kind": { "const": "paragraph_index" },
        "index": { "type": "integer", "minimum": 0 },
        "image_index": { "type": "integer", "minimum": 0, "default": 0 }
      },
      "additionalProperties": false
    },
    "action": { "enum": ["resize", "set_alt_text", "set_hyperlink", "replace", "delete"] },
    "width_pt": { "type": "number", "exclusiveMinimum": 0 },
    "height_pt": { "type": "number", "exclusiveMinimum": 0 },
    "preserve_aspect_ratio": { "type": "boolean", "default": true },
    "alt_text_title": { "type": "string" },
    "alt_text_description": { "type": "string" },
    "hyperlink": { "type": "string", "format": "uri" },
    "base64": { "type": "string" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Per-action requirements:

- `resize`: requires `width_pt` and/or `height_pt`; when
  `preserve_aspect_ratio` is true and exactly one dimension is provided, the
  add-in derives the other dimension from current image dimensions. The action
  preserves paragraph placement, alt text, relationship identity, and adjacent
  text.
- `set_alt_text`: requires `alt_text_title` and/or `alt_text_description`.
- `set_hyperlink`: requires `hyperlink`, using the same URL scheme allowlist as
  the Word hyperlink tools.
- `replace`: requires `base64`; replacement bytes use the same decoded-byte and
  image-format validation as `word.insert_image`.
- `delete`: deletes the inline picture only and preserves surrounding paragraph
  text; this action is destructive.

With `validate_only: true`, the tool resolves the target image and validates
arguments without changing image metadata or bytes. Invalid locators or action
arguments fail with `INVALID_ARGUMENT` and `partial_effect: "none"`.

### 3.8 `word.list_shapes`

List desktop-tier Word shapes from the main body by default. Shapes include
text boxes, geometric shapes, groups, pictures, and canvases exposed by
`Body.shapes`, `Paragraph.shapes`, and `Range.shapes`.

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "scope": { "enum": ["body", "paragraph", "anchor"], "default": "body" },
    "paragraph_index": { "type": "integer", "minimum": 0 },
    "anchor": { "$ref": "#/$defs/anchor" },
    "include_text": { "type": "boolean", "default": true }
  },
  "additionalProperties": false
}
```

For `scope: "paragraph"`, `paragraph_index` is required. For
`scope: "anchor"`, `anchor` is required. The response returns:

```json
{
  "shapes": [
    {
      "shape_id": 42,
      "type": "TextBox",
      "name": "Text Box 1",
      "text_preview": "Bounded text-box text",
      "left_pt": 72,
      "top_pt": 96,
      "width_pt": 240,
      "height_pt": 80,
      "relative_horizontal_position": "Margin",
      "relative_vertical_position": "Paragraph",
      "alt_text_description": "Callout text",
      "untrusted_source": true
    }
  ],
  "count": 1,
  "untrusted_source": true
}
```

`text_preview` is populated only from shape body or text-frame text that the
desktop API exposes and must be bounded by the same preview discipline as other
document text reads. `word.get_text` continues to read the main document body;
text inside text boxes is reachable through `word.list_shapes` and
`word.update_shape` rather than silently mixed into body text.

### 3.9 `word.insert_shape`

Insert a desktop-tier text box, geometric shape, or floating picture anchored
at an existing paragraph or range. The implementation uses
`Paragraph.insertTextBox`, `Paragraph.insertGeometricShape`,
`Paragraph.insertPictureFromBase64`, or the equivalent `Range` APIs after
resolving the anchor.

```json
{
  "type": "object",
  "required": ["session_id", "shape_type"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "shape_type": { "enum": ["text_box", "rectangle", "ellipse", "rounded_rectangle", "line", "picture"] },
    "anchor": { "$ref": "#/$defs/anchor" },
    "text": { "type": "string" },
    "image": { "$ref": "#/$defs/image_input" },
    "left_pt": { "type": "number" },
    "top_pt": { "type": "number" },
    "width_pt": { "type": "number", "exclusiveMinimum": 0 },
    "height_pt": { "type": "number", "exclusiveMinimum": 0 },
    "alt_text_description": { "type": "string" },
    "fill_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "line_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

`shape_type: "text_box"` uses `text` as the inserted text-box body, defaulting
to an empty string. Geometric shapes map only the small portable enum above to
Office geometric shape values; unsupported shapes fail `INVALID_ARGUMENT` before
mutation. `shape_type: "picture"` requires `image` and uses the same daemon
fetch, byte-size, MIME, and decoded-image validation policy as
`word.insert_image`. The response returns the new `shape` metadata from
`word.list_shapes` plus `created: true`.

### 3.10 `word.update_shape`

Mutate one existing desktop-tier shape by current `shape_id`.

```json
{
  "type": "object",
  "required": ["session_id", "shape_id", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "shape_id": { "type": "integer" },
    "action": { "enum": ["move", "resize", "set_text", "set_alt_text", "set_fill", "set_line", "set_wrap", "set_visibility"] },
    "left_pt": { "type": "number" },
    "top_pt": { "type": "number" },
    "width_pt": { "type": "number", "exclusiveMinimum": 0 },
    "height_pt": { "type": "number", "exclusiveMinimum": 0 },
    "text": { "type": "string" },
    "alt_text_description": { "type": "string" },
    "fill_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "line_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "wrap_type": { "enum": ["inline", "square", "tight", "behind", "front", "top_bottom"] },
    "visible": { "type": "boolean" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Per-action required arguments are preflighted with `INVALID_ARGUMENT` and
`partial_effect: "none"`. `set_text` applies only to shapes with a body or text
frame, and unsupported shape types fail before mutation where the host exposes
enough metadata to decide. `move` changes `left` and/or `top`; `resize` changes
`width` and/or `height`; `set_wrap` maps the portable wrap enum to
`Shape.textWrap`. `validate_only: true` resolves the shape and validates the
arguments without changing it. The response returns `{ action, shape, updated:
true }`.

### 3.11 `word.delete_shape`

Delete one existing desktop-tier shape by current `shape_id`.

```json
{
  "type": "object",
  "required": ["session_id", "shape_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "shape_id": { "type": "integer" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

The tool deletes only the target shape and does not delete surrounding body
paragraphs or inline-picture indexes. With `validate_only: true`, the tool
resolves the shape and returns without deleting it.

### 3.12 `word.insert_break`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "break_type": { "enum": ["page", "line", "section_next", "section_continuous", "section_even", "section_odd"], "default": "page" }
  },
  "additionalProperties": false
}
```

`word.insert_break` resolves the anchor and calls Word's break insertion API
with the requested break kind. `break_type: "page"` is the owner for page break
insertion. Section breaks are inserted at the resolved range boundary and rely on
Word's host layout semantics for the newly created section.

### 3.7 `word.insert_page_break`

Superseded compatibility contract. This tool is no longer advertised; use
`word.insert_break` with `break_type: "page"`.

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" }
  },
  "additionalProperties": false
}
```

### 3.8 `word.list_sections`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "include_page_setup": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Returns:

```json
{
  "sections": [
    {
      "index": 0,
      "first_paragraph_index": 0,
      "paragraph_count": 12,
      "has_header": true,
      "has_footer": false,
      "page_setup": {
        "orientation": "portrait",
        "paper_size": "letter",
        "margins_pt": { "top": 72, "bottom": 72, "left": 72, "right": 72 },
        "page_width_pt": 612,
        "page_height_pt": 792
      }
    }
  ],
  "count": 1
}
```

`page_setup` is included only when `include_page_setup: true` and the
`WordApiDesktop 1.3` probe succeeds. Otherwise the tool still returns section
structure without layout fields.

### 3.9 `word.update_page_setup`

Desktop-tier tool, advertised only after a successful `WordApiDesktop 1.3`
probe.

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "section_index": { "type": "integer", "minimum": 0 },
    "orientation": { "enum": ["portrait", "landscape"] },
    "paper_size": { "type": "string", "minLength": 1 },
    "margins_pt": {
      "type": "object",
      "properties": {
        "top": { "type": "number", "minimum": 0 },
        "bottom": { "type": "number", "minimum": 0 },
        "left": { "type": "number", "minimum": 0 },
        "right": { "type": "number", "minimum": 0 }
      },
      "additionalProperties": false
    },
    "page_width_pt": { "type": "number", "exclusiveMinimum": 0 },
    "page_height_pt": { "type": "number", "exclusiveMinimum": 0 }
  },
  "additionalProperties": false
}
```

Omitting `section_index` applies the requested page setup to the document-level
page setup. Supplying `section_index` applies only to that section. Calls that
provide no mutable page-setup fields fail with `INVALID_ARGUMENT` and
`partial_effect: "none"` before any write is queued.

### 3.10 `word.get_header_footer`

```json
{
  "type": "object",
  "required": ["session_id", "location"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "location": { "enum": ["header", "footer"] },
    "header_footer_type": { "enum": ["primary", "first_page", "even_pages"], "default": "primary" },
    "section_index": { "type": "integer", "minimum": 0, "default": 0 },
    "include_metadata": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Returns `{ text, is_empty, section_count }`. When `include_metadata` is true,
the response also includes `paragraphs: [{ index, text, style }]` for the target
header or footer body. `section_index` is zero-based and must be within the
current document's section collection.

`header_footer_type: "first_page"` requires the target section's different-first
page layout to be enabled. `header_footer_type: "even_pages"` requires odd/even
header-footer layout to be enabled. If the layout is disabled, the tool fails
with `INVALID_ARGUMENT` rather than reading a header/footer slot that Word will
not render for the user.

### 3.11 `word.insert_bookmark`

Create a named bookmark around an anchored range.

```json
{
  "type": "object",
  "required": ["session_id", "name", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "name": { "type": "string", "minLength": 1, "pattern": "^[A-Za-z_][A-Za-z0-9_]{0,39}$" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "$ref": "#/$defs/extent" },
    "overwrite": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

The name pattern mirrors Word bookmark names: it must start with a letter or
underscore, contain only letters, digits, and underscores, and be at most 40
characters. Word treats bookmark names case-insensitively; the add-in MUST
perform duplicate checks case-insensitively before queuing a write.

When `overwrite` is false and a bookmark with the same name already exists, the
call fails with `INVALID_ARGUMENT` and `partial_effect: "none"`. When
`overwrite` is true, the existing bookmark marker may be moved to the newly
resolved range. The tool returns `{ bookmark: { name, paragraph_index,
text_preview }, overwritten }` after the write is synchronized.

`extent` follows the shared anchored range contract. For point-like anchors
such as `start_of_document` and `end_of_document`, the bookmark is collapsed at
the resolved location unless the implementation can safely expand a caller
provided extent.

### 3.8 `word.update_header_footer`

```json
{
  "type": "object",
  "required": ["session_id", "location", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "location": { "enum": ["header", "footer"] },
    "header_footer_type": { "enum": ["primary", "first_page", "even_pages"], "default": "primary" },
    "section_index": { "type": "integer", "minimum": 0, "default": 0 },
    "action": { "enum": ["set_text", "append_paragraph", "clear"] },
    "text": { "type": "string" },
    "style": { "type": "string" },
    "formatting": { "$ref": "#/$defs/run_formatting" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Action semantics:

- `set_text` requires `text` and replaces the whole target header/footer body
  with `Body.insertText(text, "Replace")`.
- `append_paragraph` requires `text`, appends a paragraph to the target
  header/footer body, and may apply `style` and run `formatting` to the inserted
  paragraph.
- `clear` removes all content from the target header/footer body with
  `Body.clear()` and is classified destructive.

`validate_only: true` performs argument, section, layout, and action validation
without invoking a mutating Office.js API. Successful validation returns
`{ valid: true, operation: "word.update_header_footer", partial_effect: "none", resolved_target: { section_index, location, header_footer_type } }`.
Failures return `INVALID_ARGUMENT` with `partial_effect: "none"`.

### 3.9 `word.insert_list`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "items"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "items": {
      "type": "array",
      "minItems": 1,
      "items": { "type": "string" }
    },
    "kind": { "enum": ["bulleted", "numbered"], "default": "bulleted" },
    "level": { "type": "integer", "minimum": 0, "maximum": 8, "default": 0 }
  }
}
```

`word.insert_list` creates a new list from plain text items at an anchor. Existing
list membership and formatting changes belong to `word.update_list`.

### 3.10 `word.list_lists`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "offset": { "type": "integer", "minimum": 0, "default": 0 },
    "limit": { "type": "integer", "minimum": 1, "maximum": 200, "default": 50 }
  },
  "additionalProperties": false
}
```

Returns:

```json
{
  "lists": [
    {
      "list_id": 1,
      "kind": "bulleted",
      "item_count": 2,
      "first_paragraph_index": 3,
      "items": [
        { "paragraph_index": 3, "text": "First", "level": 0, "list_string": "•" }
      ]
    }
  ],
  "count": 1,
  "truncated": false,
  "untrusted_source": true
}
```

`list_id` is the host list id for the current document session. It must be
treated like other runtime identifiers: re-read `word.list_lists` before
mutating after large document edits. Empty documents return `count: 0`.

### 3.11 `word.update_list`

```json
{
  "type": "object",
  "required": ["session_id", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "action": { "enum": ["add_item", "set_item_level", "attach_paragraph", "detach_paragraph", "set_level_format"] },
    "list_id": { "type": "integer", "minimum": 0 },
    "paragraph_index": { "type": "integer", "minimum": 0 },
    "text": { "type": "string" },
    "position": { "enum": ["start", "end", "after_paragraph"], "default": "end" },
    "level": { "type": "integer", "minimum": 0, "maximum": 8 },
    "numbering": { "enum": ["bullet", "arabic", "upper_roman", "lower_roman", "upper_letter", "lower_letter", "none"] },
    "bullet_char": { "type": "string", "maxLength": 1 },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Action semantics:

- `add_item` requires `list_id` and `text`. It inserts a paragraph into the
  target list at `position`, defaulting to the end of the list. `level` defaults
  to `0`.
- `set_item_level` requires `paragraph_index` and `level`. The paragraph must be
  a list item.
- `attach_paragraph` requires `paragraph_index`; with `list_id` it attaches the
  paragraph to that list, and without `list_id` it starts a new list from the
  paragraph. `level` defaults to `0`.
- `detach_paragraph` requires `paragraph_index` and converts a list item back to
  body text without deleting text.
- `set_level_format` requires `list_id`, `level`, and `numbering`; `numbering:
  "bullet"` may use `bullet_char`, while numbered formats map to Word list
  numbering styles where supported.

Unknown list ids, stale paragraph indices, non-list paragraph targets for
list-only actions, and action/argument mismatches return `INVALID_ARGUMENT` or
`INDEX_OUT_OF_RANGE` with `partial_effect: "none"`. `validate_only: true`
performs argument and target validation without invoking mutating Office.js list
APIs and returns `{ valid: true, operation: "word.update_list", partial_effect:
"none", action }`.

## 4. Edit

### 4.1 `word.replace_text`

```json
{
  "type": "object",
  "required": ["session_id", "find", "replace"],
  "properties": {
    "session_id": { "type": "string" },
    "find": { "type": "string" },
    "replace": { "type": "string" },
    "match_case": { "type": "boolean", "default": false },
    "whole_word": { "type": "boolean", "default": false },
    "wildcards": { "type": "boolean", "default": false },
    "scope": {
      "type": "object",
      "properties": {
        "paragraph_range": { "type": "array", "items": { "type": "integer" }, "minItems": 2, "maxItems": 2 },
        "selection_only":  { "type": "boolean", "default": false }
      }
    },
    "dry_run": { "type": "boolean", "default": false },
    "validate_only": { "type": "boolean", "default": false },
    "partial_ok": { "type": "boolean", "default": false }
  }
}
```

Returns `{ replaced_count: N, matches: [...] }`.

`wildcards` uses Word wildcard syntax, not JavaScript regular expressions.

`dry_run: true` returns matches without modifying the document. **Highly
recommended pattern for agents**: dry-run first, present to user, then run again
without `dry_run`.

`validate_only: true` is accepted as a synonym for `dry_run: true` and returns
the same bounded `matches`, `skipped_count`, and `replaced_count: 0` shape plus
`valid: true`. If both flags are provided, either truthy value selects the
no-mutation path.

### 4.2 `word.insert_hyperlink`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "url"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "url": { "type": "string" },
    "text": { "type": "string" },
    "extent": { "$ref": "#/$defs/extent" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

If `text` is provided, the add-in inserts that text at the resolved anchor and
hyperlinks the inserted range. Otherwise it applies the hyperlink to the
resolved anchor range. URL schemes are restricted to `https`, `http`, `mailto`,
and in-document bookmark targets beginning with `#`. `file:`, `javascript:`,
and other schemes fail with `INVALID_ARGUMENT` and `partial_effect: "none"`
before any write is queued.

With `validate_only: true`, the add-in resolves the anchor and URL policy and
returns `valid: true`, `operation: "word.insert_hyperlink"`,
`partial_effect: "none"`, and a `resolved_target` summary without setting
`Range.hyperlink` or inserting text.

### 4.3 `word.list_hyperlinks`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "offset": { "type": "integer", "minimum": 0, "default": 0 },
    "limit": { "type": "integer", "minimum": 1, "maximum": 200, "default": 50 }
  },
  "additionalProperties": false
}
```

Returns `{ hyperlinks: [{ paragraph_index, occurrence_in_paragraph, text, url }],
count, truncated }`. The implementation enumerates paragraphs and uses
`Range.getHyperlinkRanges()` so locations stay paragraph-relative rather than
promising stable character offsets.

### 4.4 `word.remove_hyperlink`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "keep_text": { "type": "boolean", "default": true }
  },
  "additionalProperties": false
}
```

With `keep_text: true`, the add-in clears `Range.hyperlink` and preserves the
range text. With `keep_text: false`, it deletes the resolved range using the
same preflight and partial-effect rules as other range deletion paths.

### 4.5 `word.update_paragraph`

Replace one paragraph's text wholesale.

```json
{
  "type": "object",
  "required": ["session_id", "index", "text"],
  "properties": {
    "session_id": { "type": "string" },
    "index": { "type": "integer" },
    "text": { "type": "string" },
    "validate_only": { "type": "boolean", "default": false }
  }
}
```

Replacing paragraph text preserves paragraph-level style but does not promise
to preserve mixed character-run formatting inside the old text.

With `validate_only: true`, the add-in resolves the paragraph by index, loads
safe metadata such as the current paragraph text length and style when
available, and returns `valid: true`, `operation: "word.update_paragraph"`, and
`resolved_target.paragraph_index` without calling `insertText`.

### 4.6 `word.delete_range`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "enum": ["paragraph", "sentence", "selection"] },
    "validate_only": { "type": "boolean", "default": false }
  }
}
```

With `validate_only: true`, the add-in resolves the target range or current
selection and returns `valid: true`, `operation: "word.delete_range"`, and a
`resolved_target` summary that includes the requested extent and target object
type when known. It MUST NOT call `delete`.

### 4.4 `word.delete_bookmark`

Delete a bookmark marker without deleting the bookmarked text.

```json
{
  "type": "object",
  "required": ["session_id", "name"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "name": { "type": "string", "minLength": 1 }
  },
  "additionalProperties": false
}
```

Missing bookmarks fail with `INVALID_ARGUMENT` and `partial_effect: "none"`
before any write is queued. Name matching is case-insensitive. The success
response includes `{ deleted: true, name, count }`, where `count` is the number
of remaining visible bookmarks after deletion.

### 4.7 `word.apply_formatting`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "$ref": "#/$defs/extent" },
    "formatting": { "$ref": "#/$defs/run_formatting" },
    "paragraph": { "$ref": "#/$defs/paragraph_formatting" }
  },
  "anyOf": [
    { "required": ["formatting"] },
    { "required": ["paragraph"] }
  ],
  "additionalProperties": false
}
```

`run_formatting`:

```json
{
  "type": "object",
  "properties": {
    "bold": { "type": "boolean" },
    "italic": { "type": "boolean" },
    "underline": { "type": "boolean" },
    "strikethrough": { "type": "boolean" },
    "font_name": { "type": "string" },
    "font_size_pt": { "type": "number" },
    "color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "highlight": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" }
  }
}
```

`paragraph_formatting`:

```json
{
  "type": "object",
  "properties": {
    "alignment": { "enum": ["left", "center", "right", "justified"] },
    "left_indent_pt": { "type": "number", "minimum": 0 },
    "right_indent_pt": { "type": "number", "minimum": 0 },
    "first_line_indent_pt": { "type": "number" },
    "line_spacing_pt": { "type": "number", "exclusiveMinimum": 0 },
    "line_unit_before": { "type": "number", "minimum": 0 },
    "line_unit_after": { "type": "number", "minimum": 0 },
    "space_before_pt": { "type": "number", "minimum": 0 },
    "space_after_pt": { "type": "number", "minimum": 0 },
    "outline_level": { "type": "integer", "minimum": 0, "maximum": 9 }
  },
  "additionalProperties": false
}
```

At least one of `formatting` or `paragraph` is required. A call that passes
neither block, or passes an empty block, fails with `INVALID_ARGUMENT` and
`partial_effect: "none"`. When both blocks are present, run formatting applies
to the resolved range font and paragraph formatting applies to every paragraph
intersecting the resolved range. Negative `first_line_indent_pt` expresses a
hanging indent using Word's native `Paragraph.firstLineIndent` semantics.

## 5. Tables

### 5.1 `word.read_table`

```json
{
  "type": "object",
  "required": ["session_id", "table_index"],
  "properties": {
    "session_id": { "type": "string" },
    "table_index": { "type": "integer" }
  }
}
```

Returns `{ rows, cols, data: string[][], header_row: boolean, merged?: object[] }`.
On uniform tables, `data` is the rectangular cell text matrix. On merged or
otherwise non-uniform tables where Word cannot provide a rectangular value
matrix, the tool returns the dimensions plus `merged` diagnostics when the host
supports the `WordApi 1.4` table APIs; hosts that cannot inspect the merged
geometry fail with `HOST_CAPABILITY_UNAVAILABLE` instead of returning a
misleading matrix.

### 5.2 `word.update_table`

`word.update_table` is the target table mutation owner. The action field selects
one table-owned mutation while preserving the existing v1 table behavior.

```json
{
  "type": "object",
  "required": ["session_id", "table_index", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "table_index": { "type": "integer", "minimum": 0 },
    "action": {
      "enum": [
        "update_cell", "add_row", "add_column", "format_cell", "delete",
        "delete_row", "delete_column", "merge_cells", "set_column_width",
        "distribute_columns", "set_borders", "set_header_row"
      ]
    },
    "row": { "type": "integer", "minimum": 0 },
    "col": { "type": "integer", "minimum": 0 },
    "text": { "type": "string" },
    "index": { "type": "integer", "minimum": 0 },
    "row_range": { "type": "array", "items": { "type": "integer", "minimum": 0 }, "minItems": 2, "maxItems": 2 },
    "col_range": { "type": "array", "items": { "type": "integer", "minimum": 0 }, "minItems": 2, "maxItems": 2 },
    "values": { "type": "array", "items": { "type": "string" } },
    "width_pt": { "type": "number", "exclusiveMinimum": 0 },
    "header_row": { "type": "boolean" },
    "borders": {
      "type": "object",
      "properties": {
        "edges": { "type": "array", "items": { "enum": ["top", "bottom", "left", "right", "inside_horizontal", "inside_vertical", "all"] } },
        "style": { "enum": ["single", "double", "dotted", "dashed", "none"] },
        "width_pt": { "type": "number", "minimum": 0 },
        "color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" }
      },
      "additionalProperties": false
    },
    "background_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "horizontal_alignment": { "enum": ["left", "center", "right"] },
    "vertical_alignment": { "enum": ["top", "center", "bottom"] },
    "padding_pt": { "type": "number", "minimum": 0 },
    "formatting": { "$ref": "#/$defs/run_formatting" }
  },
  "additionalProperties": false
}
```

Action semantics:

- `update_cell` requires `row`, `col`, and `text`; `formatting` is optional.
- `add_row` accepts optional `index` and `values`; omitting `index` appends.
- `add_column` accepts optional `index` and `values`; omitting `index` appends.
- `format_cell` requires `row` and `col`; cell background, alignment, padding,
  and run formatting are optional.
- `delete_row` requires `row` or `row_range`; it deletes one row or an inclusive
  row span after validating table bounds. This is destructive.
- `delete_column` requires `col` or `col_range`; it deletes one column or an
  inclusive column span after validating table bounds. This is destructive.
- `merge_cells` requires `row_range` and `col_range`; it merges the rectangular
  span and concatenates cell contents using Word's native table semantics. It
  requires `WordApi 1.4` host support.
- `set_column_width` requires `col` and `width_pt`; it updates the target column
  width without changing cell text.
- `distribute_columns` distributes columns evenly across the table.
- `set_borders` requires `borders` and updates table-level borders. When `row`
  and `col` are provided, it updates the addressed cell borders instead.
- `set_header_row` requires `header_row` and toggles a single header row.
- `delete` deletes the whole table and must be requested explicitly.

All new actions preflight live table bounds before mutation. Invalid indices,
inverted ranges, or missing per-action arguments fail with `INVALID_ARGUMENT`
and `partial_effect: "none"`. `delete_row` and `delete_column` are destructive
actions and require the All ceiling; the other new actions are edit actions.

The compatibility tools in §5.3-§5.6 remain documented until the catalog
migration retires them from advertisement.

### 5.3 `word.update_cell`

Superseded compatibility contract. This tool is no longer advertised; use
`word.update_table` with `action: "update_cell"`.

```json
{
  "type": "object",
  "required": ["session_id", "table_index", "row", "col", "text"],
  "properties": {
    "session_id": { "type": "string" },
    "table_index": { "type": "integer" },
    "row": { "type": "integer" },
    "col": { "type": "integer" },
    "text": { "type": "string" },
    "formatting": { "$ref": "#/$defs/run_formatting" }
  }
}
```

### 5.4 `word.add_row`

Superseded compatibility contract. This tool is no longer advertised; use
`word.update_table` with `action: "add_row"`.

```json
{
  "type": "object",
  "required": ["session_id", "table_index"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "table_index": { "type": "integer", "minimum": 0 },
    "index": { "type": "integer", "minimum": 0 },
    "values": { "type": "array", "items": { "type": "string" } }
  }
}
```

Omitting `index` appends the row. An interior insertion uses
`TableCell.insertRows("Before", ...)`; non-uniform tables that cannot satisfy
that API return `HOST_CAPABILITY_UNAVAILABLE`. If `values` is provided, its
length must equal the table's current column count.

### 5.5 `word.add_column`

Superseded compatibility contract. This tool is no longer advertised; use
`word.update_table` with `action: "add_column"`.

```json
{
  "type": "object",
  "required": ["session_id", "table_index"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "table_index": { "type": "integer", "minimum": 0 },
    "index": { "type": "integer", "minimum": 0 },
    "values": { "type": "array", "items": { "type": "string" } }
  }
}
```

Omitting `index` appends the column. An interior insertion uses
`TableCell.insertColumns("Before", ...)`; non-uniform tables that cannot
satisfy that API return `HOST_CAPABILITY_UNAVAILABLE`. If `values` is
provided, its length must equal the table's current row count.

### 5.6 `word.format_cell`

Superseded compatibility contract. This tool is no longer advertised; use
`word.update_table` with `action: "format_cell"`.

```json
{
  "type": "object",
  "required": ["session_id", "table_index", "row", "col"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "table_index": { "type": "integer", "minimum": 0 },
    "row": { "type": "integer", "minimum": 0 },
    "col": { "type": "integer", "minimum": 0 },
    "background_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "horizontal_alignment": { "enum": ["left", "center", "right"] },
    "vertical_alignment": { "enum": ["top", "center", "bottom"] },
    "padding_pt": { "type": "number", "minimum": 0 },
    "formatting": { "$ref": "#/$defs/run_formatting" }
  }
}
```

Cell merging and arbitrary border editing are now owned by `word.update_table`
through `merge_cells` and `set_borders`; the superseded compatibility tool
remains documented only for historical argument compatibility.

## 6. Structure

### 6.1 `word.set_heading_level`

Superseded compatibility contract. This tool is no longer advertised; use
`word.apply_style` with `heading_level`.

```json
{
  "type": "object",
  "required": ["session_id", "index", "level"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "index": { "type": "integer", "minimum": 0 },
    "level": { "type": "integer", "minimum": 0, "maximum": 9 }
  }
}
```

Level `0` converts the paragraph to body text.

### 6.2 `word.apply_style`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "style": { "type": "string", "minLength": 1 },
    "heading_level": { "type": "integer", "minimum": 0, "maximum": 9 }
  }
}
```

Callers must provide either `style` or `heading_level`. `heading_level` is the
target replacement for `word.set_heading_level`.

### 6.3 `word.list_styles`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "type": { "enum": ["paragraph", "character", "table", "list"] },
    "built_in": { "type": "boolean" },
    "in_use_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Returns `{ styles, count }`. Each style item includes `name_local`, `type`,
`built_in`, `in_use`, `base_style`, and `priority`. The tool uses
`Document.getStyles()` and is advertised only when the `WordApi 1.5` probe
passes.

### 6.4 `word.create_style`

```json
{
  "type": "object",
  "required": ["session_id", "name", "type"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "name": { "type": "string", "minLength": 1 },
    "type": { "enum": ["paragraph", "character", "table", "list"] },
    "base_style": { "type": "string", "minLength": 1 },
    "font": { "$ref": "#/$defs/run_formatting" },
    "paragraph": { "$ref": "#/$defs/paragraph_formatting" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Creates a document style via `Document.addStyle(name, type)`, then applies
optional font and paragraph-formatting properties. Duplicate style names fail
before mutation with `INVALID_ARGUMENT` and `partial_effect: none`.
`base_style` is accepted only when a `WordApi 1.6` probe succeeds because
Office.js exposes `Style.baseStyle` in 1.5 but only supports setting it in 1.6;
on 1.5-only hosts it fails with `HOST_CAPABILITY_UNAVAILABLE` before mutation.
`validate_only: true` resolves existing styles and validates arguments without
creating the style.

### 6.5 `word.update_style`

```json
{
  "type": "object",
  "required": ["session_id", "name"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "name": { "type": "string", "minLength": 1 },
    "base_style": { "type": "string", "minLength": 1 },
    "font": { "$ref": "#/$defs/run_formatting" },
    "paragraph": { "$ref": "#/$defs/paragraph_formatting" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Callers must provide at least one of `base_style`, `font`, or `paragraph`.
Unknown style names and unknown `base_style` names fail before mutation with
`INVALID_ARGUMENT` and `partial_effect: none`. `base_style` writes require a
successful `WordApi 1.6` probe; other style updates remain available at
`WordApi 1.5`. Built-in styles may be updated when Word allows it; responses
include `built_in` so clients can present an appropriate warning. Style
deletion remains out of scope for this tool because deleting an in-use style
changes document content fallback behavior and needs a separate destructive
contract.

## 6A. Document Properties

Document-property tools own Word document metadata. Core properties map to
`Document.properties`; custom properties map to
`Document.properties.customProperties`. The tool surface keeps read-only Office
metadata observable but not writable.

### 6A.1 `word.get_document_properties`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "include_custom": { "type": "boolean", "default": true }
  },
  "additionalProperties": false
}
```

Returns writable core fields (`title`, `subject`, `author`, `keywords`,
`category`, `comments`, `company`, and `manager`), read-only metadata
(`last_author`, `revision_number`, `creation_date`, `last_save_time`, and
`security` when Office returns it), and, when `include_custom` is not false,
`custom: [{ key, type, value }]`.

The response normalizes Office camelCase names to snake_case. Date values are
returned as ISO-8601 strings when Word returns a `Date`; missing host values are
omitted rather than filled with synthetic defaults.

### 6A.2 `word.update_document_properties`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "title": { "type": "string" },
    "subject": { "type": "string" },
    "author": { "type": "string" },
    "keywords": { "type": "string" },
    "category": { "type": "string" },
    "comments": { "type": "string" },
    "company": { "type": "string" },
    "manager": { "type": "string" },
    "custom_set": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["key", "value"],
        "properties": {
          "key": { "type": "string", "minLength": 1 },
          "value": {
            "oneOf": [
              { "type": "string" },
              { "type": "number" },
              { "type": "boolean" }
            ]
          }
        },
        "additionalProperties": false
      }
    },
    "custom_delete": {
      "type": "array",
      "items": { "type": "string", "minLength": 1 }
    }
  },
  "additionalProperties": false
}
```

Callers must provide at least one writable core property, one `custom_set`
entry, or one `custom_delete` key. Requests that contain no writable operation
fail with `INVALID_ARGUMENT` and `partial_effect: "none"`. Read-only fields
such as `last_author`, `revision_number`, `creation_date`, and
`last_save_time` are intentionally absent from the schema, so attempts to write
them fail schema validation.

`custom_set` upserts by deleting any existing custom property with the same key
before adding the new value through `CustomPropertyCollection.add(key, value)`.
`custom_delete` deletes matching keys and reports unknown keys in the response;
unknown-key deletes are successful no-ops. `deleteAll` is not exposed because it
is too broad for the normal metadata-edit workflow.

| Operation | API | Requirement set |
|---|---|---|
| Read/write core properties | `Document.properties` / `DocumentProperties` | `WordApi 1.3` |
| Read read-only metadata | `DocumentProperties.lastAuthor`, `creationDate`, `lastSaveTime`, `revisionNumber`, `security` | `WordApi 1.3` |
| Enumerate custom properties | `DocumentProperties.customProperties` / `CustomPropertyCollection` | `WordApi 1.3` |
| Upsert/delete custom properties | `CustomPropertyCollection.add(key, value)` / `CustomProperty.delete()` | `WordApi 1.3` |

Document-property tools are core-tier tools. `word.get_document_properties`
requires the same read ceiling as document text extraction because metadata can
contain sensitive author, keyword, and comment values. `word.update_document_properties`
requires edit permission and has no destructive variant.

## 7. Content Controls

The content-control tools own content-control lifecycle, metadata, and
type-specific form state. v1 supports rich text and plain text controls on
`WordApi 1.5`, checkbox controls on `WordApi 1.7`, and dropdown-list / combo-box
controls on `WordApi 1.9`. The four existing owner tools remain the complete
surface; checkbox and list controls extend their schemas instead of adding new
tools.

The task pane advertises the content-control tools when the base `WordApi 1.5`
content-control tier is present. Type-specific arguments that require a higher
tier MUST fail before mutation with `HOST_CAPABILITY_UNAVAILABLE` and
`partial_effect: none` when the host does not support the required Word API.
Session-specific schema narrowing may be added later, but the stable daemon
schema documents the full cross-host contract.

Picture, date-picker, repeating-section, and group content controls remain
deferred. Picture controls overlap the Media owner and need explicit byte/fetch
safety policy; date-picker support is desktop-specific; repeating-section and
group controls do not yet have enough stable portable CRUD behavior for the v1
contract.

### 7.1 `word.list_content_controls`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "type": { "enum": ["rich_text", "plain_text", "checkbox", "dropdown_list", "combo_box"] },
    "tag": { "type": "string" },
    "title": { "type": "string" }
  },
  "additionalProperties": false
}
```

Returns `{ content_controls, count }`. Each item includes
`content_control_id`, `tag`, `title`, `type`, `subtype`, `cannot_delete`, and
`cannot_edit`. Checkbox items also include `checked`. Dropdown-list and
combo-box items include `list_items: [{ display_text, value }]` and
`selected_text`. The tool does not return arbitrary contained document text for
rich/plain text controls; callers use read tools for body text.

### 7.2 `word.insert_content_control`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "type": { "enum": ["rich_text", "plain_text", "checkbox", "dropdown_list", "combo_box"] },
    "text": { "type": "string" },
    "checked": { "type": "boolean" },
    "list_items": {
      "type": "array",
      "minItems": 1,
      "items": {
        "type": "object",
        "required": ["display_text"],
        "properties": {
          "display_text": { "type": "string", "minLength": 1 },
          "value": { "type": "string" }
        },
        "additionalProperties": false
      }
    },
    "tag": { "type": "string" },
    "title": { "type": "string" },
    "cannot_delete": { "type": "boolean" },
    "cannot_edit": { "type": "boolean" },
    "appearance": { "enum": ["bounding_box", "tags", "hidden"] },
    "color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "placeholder_text": { "type": "string" }
  },
  "additionalProperties": false
}
```

When `anchor` is omitted, the current selection is used. When `text` is
provided for rich/plain text controls, the anchored range is replaced with that
text before wrapping it. `checked` is valid only with `type: "checkbox"`.
`list_items` is valid only with `type: "dropdown_list"` or `type: "combo_box"`
and must provide at least one item. Invalid type-specific arguments fail with
`INVALID_ARGUMENT` and no mutation.

### 7.3 `word.update_content_control`

```json
{
  "type": "object",
  "required": ["session_id", "content_control_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "content_control_id": { "type": "integer", "minimum": 0 },
    "text": { "type": "string" },
    "checked": { "type": "boolean" },
    "selected_value": { "type": "string" },
    "list_items_add": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["display_text"],
        "properties": {
          "display_text": { "type": "string", "minLength": 1 },
          "value": { "type": "string" },
          "index": { "type": "integer", "minimum": 0 }
        },
        "additionalProperties": false
      }
    },
    "list_items_delete": { "type": "array", "items": { "type": "string" } },
    "list_items_clear": { "type": "boolean", "default": false },
    "tag": { "type": "string" },
    "title": { "type": "string" },
    "cannot_delete": { "type": "boolean" },
    "cannot_edit": { "type": "boolean" },
    "appearance": { "enum": ["bounding_box", "tags", "hidden"] },
    "color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
    "placeholder_text": { "type": "string" }
  },
  "additionalProperties": false
}
```

`checked` is valid only for checkbox controls. `selected_value`,
`list_items_add`, `list_items_delete`, and `list_items_clear` are valid only for
dropdown-list or combo-box controls. `selected_value` matches either item value
or display text and selects exactly one item; ambiguous or missing matches fail
without mutation. List item deletion matches values or display text and fails if
any requested item is absent unless a future explicit partial mode is added.

### 7.4 `word.delete_content_control`

```json
{
  "type": "object",
  "required": ["session_id", "content_control_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "content_control_id": { "type": "integer", "minimum": 0 },
    "mode": { "enum": ["keep_content", "delete_content"], "default": "keep_content" }
  },
  "additionalProperties": false
}
```

The default preserves contents while removing the content-control wrapper.

### 7.5 Office.js Mapping

| Operation | API | Requirement set |
|---|---|---|
| Rich/plain content controls | `Range.insertContentControl("RichText" | "PlainText")`, `Body.getContentControls()` | `WordApi 1.5` |
| Checkbox state | `ContentControl.checkboxContentControl.isChecked` | `WordApi 1.7` |
| Insert checkbox | `Range.insertContentControl("CheckBox")` | `WordApi 1.7` |
| Dropdown/combobox object | `ContentControl.dropDownListContentControl` / `comboBoxContentControl` | `WordApi 1.9` |
| List items | `addListItem(displayText, value, index)`, `deleteAllListItems()`, `ContentControlListItem.delete()` / `select()` | `WordApi 1.9` |

## 7A. Bookmark Lifecycle

Bookmark lifecycle tools are Range & selection tools because bookmarks are
named range markers, not review annotations. They are listed here together to
make their Office.js mapping explicit.

| Operation | API | Requirement set |
|---|---|---|
| Insert | `Range.insertBookmark(name)` | `WordApi 1.4` |
| Enumerate | `Range.getBookmarks(includeHidden, includeAdjacent)` | `WordApi 1.4` |
| Resolve to range | `Document.getBookmarkRange(name)` / `getBookmarkRangeOrNullObject(name)` | `WordApi 1.4` |
| Delete | `Document.deleteBookmark(name)` | `WordApi 1.4` |

The lifecycle tools share the same name validation and case-insensitive
existence checks described in §3.11 and §4.4. A bookmark anchor remains part of
the shared anchor vocabulary whether or not the lifecycle tools are enabled;
however, the lifecycle tools themselves are advertised only when the `WordApi
1.4` review tier is available.

Selection-setting stays in Range & selection because it targets the same shared
anchor vocabulary as `word.resolve_anchor` and anchored mutation tools. It is
implemented with `Range.select("Select" | "Start" | "End")`; the tool itself
uses the existing Word core anchor-resolution tier (`WordApi 1.3`) rather than
introducing a separate selection capability gate.

## 7B. Note Lifecycle

Footnotes and endnotes share one lifecycle owner because they use the same Word
note object model and differ only by collection. Note indices are current
collection positions. Clients that need stable addressing across edits must
re-read `word.list_notes` after any insertion or deletion because Word shifts
indices when note references move or disappear.

### 7B.1 `word.insert_note`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "kind", "text"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "kind": { "enum": ["footnote", "endnote"] },
    "text": { "type": "string", "minLength": 1 },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

The add-in resolves the anchor, inserts the note reference at the resolved
range, and sets the note body text with `Range.insertFootnote(text)` or
`Range.insertEndnote(text)`. With `validate_only: true`, the add-in resolves
the anchor and validates `kind` and `text`, then returns `valid: true`,
`operation: "word.insert_note"`, and `partial_effect: "none"` without creating
a note.

### 7B.2 `word.list_notes`

```json
{
  "type": "object",
  "required": ["session_id", "kind"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "kind": { "enum": ["footnote", "endnote"] },
    "offset": { "type": "integer", "minimum": 0, "default": 0 },
    "limit": { "type": "integer", "minimum": 1, "maximum": 200, "default": 50 }
  },
  "additionalProperties": false
}
```

Returns `{ notes: [{ index, kind, text, reference_paragraph_index }], count,
truncated }` from `Body.footnotes` or `Body.endnotes`. Note body text is the
note content, not surrounding body text. `reference_paragraph_index` is
best-effort paragraph-relative metadata for the note reference location.

### 7B.3 `word.update_note`

```json
{
  "type": "object",
  "required": ["session_id", "kind", "index", "text"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "kind": { "enum": ["footnote", "endnote"] },
    "index": { "type": "integer", "minimum": 0 },
    "text": { "type": "string" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Replaces the addressed note body text wholesale. Out-of-range indices fail with
`INVALID_ARGUMENT` and `partial_effect: "none"` before any write is queued. With
`validate_only: true`, the add-in resolves the note and returns safe current
metadata without replacing note body text.

### 7B.4 `word.delete_note`

```json
{
  "type": "object",
  "required": ["session_id", "kind", "index"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "kind": { "enum": ["footnote", "endnote"] },
    "index": { "type": "integer", "minimum": 0 },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Deletes the addressed note reference and note body with `NoteItem.delete()`.
Out-of-range indices fail with `INVALID_ARGUMENT` and `partial_effect: "none"`
before any deletion. The success response includes `{ deleted: true, kind,
index, count }`, where `count` is the remaining number of notes of that kind.
With `validate_only: true`, the add-in resolves the target note and returns the
same current metadata without deleting it.

| Operation | API | Requirement set |
|---|---|---|
| Insert footnote/endnote | `Range.insertFootnote(text)` / `Range.insertEndnote(text)` | `WordApi 1.5` |
| Enumerate notes | `Body.footnotes` / `Body.endnotes` | `WordApi 1.5` |
| Read/edit note body | `NoteItem.body` | `WordApi 1.5` |
| Locate note reference | `NoteItem.reference` | `WordApi 1.5` |
| Delete note | `NoteItem.delete()` | `WordApi 1.5` |

Note lifecycle tools are advertised only when the `WordApi 1.5` notes and
content-controls tier is available.

## 7C. Field Lifecycle

Word fields are document-structure objects used for generated furniture such as
tables of contents, page numbers, dates, cross-references, sequence numbers,
and style references. Field tools own field lifecycle. The field allowlist is
curated; field types that can import external content or execute unsafe host
behavior, such as `INCLUDETEXT` or `IMPORT`, are not exposed.

Field indices are current collection positions. Clients that need stable
addressing across edits must re-read `word.list_fields` after any insert,
delete, or document edit that can add or remove fields.

### 7C.1 `word.list_fields`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "type": { "type": "string" },
    "offset": { "type": "integer", "minimum": 0, "default": 0 },
    "limit": { "type": "integer", "minimum": 1, "maximum": 200, "default": 50 }
  },
  "additionalProperties": false
}
```

Returns `{ fields: [{ index, type, code, result_preview, locked,
paragraph_index }], count, truncated }` from `Document.fields`. The optional
`type` filter accepts the normalized field type names returned by this tool;
unknown filters return an empty list rather than an error. `code` and
`result_preview` are bounded previews and must not duplicate full document text.

### 7C.2 `word.insert_field`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "field_type"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "field_type": { "enum": ["toc", "page", "num_pages", "date", "time", "ref", "hyperlink", "seq", "styleref"] },
    "code_options": { "type": "string" },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

The add-in resolves the anchor and maps the curated `field_type` to a Word
field insertion. `field_type: "toc"` uses a safe default table-of-contents
field code when `code_options` is omitted: `\\o "1-3" \\h \\z \\u`, producing
a hyperlinkable three-level TOC. Other field types use either the matching
`Word.FieldType` value or a safe field-code construction when Office.js requires
code text.

`code_options` is limited to field switches and arguments for the selected
allowlisted type. The add-in rejects unsupported field types, external-content
field names, and malformed options with `INVALID_ARGUMENT` and
`partial_effect: "none"` before queuing writes. With `validate_only: true`, the
add-in resolves the anchor and validates the field type and options, then
returns `{ valid: true, operation: "word.insert_field", partial_effect: "none" }`
without inserting a field.

### 7C.3 `word.update_field`

```json
{
  "type": "object",
  "required": ["session_id", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "action": { "enum": ["refresh", "refresh_all", "lock", "unlock"] },
    "field_index": { "type": "integer", "minimum": 0 },
    "expected_count": { "type": "integer", "minimum": 0 },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false,
  "allOf": [
    {
      "if": { "properties": { "action": { "enum": ["refresh", "lock", "unlock"] } } },
      "then": { "required": ["field_index"] }
    },
    {
      "if": { "properties": { "action": { "const": "refresh_all" } } },
      "then": { "required": ["expected_count"] }
    }
  ]
}
```

`refresh` calls `Field.updateResult()` for one current field index. `lock` and
`unlock` set `Field.locked`. `refresh_all` reloads `Document.fields`, compares
the live count to `expected_count`, and refreshes every field only when the
count still matches. Count mismatches fail with `STALE_INDEX` and no mutation.
Out-of-range indices fail with `INVALID_ARGUMENT` and `partial_effect: "none"`
before any write is queued. With `validate_only: true`, the add-in resolves the
target or count guard and returns current metadata without changing the field.

### 7C.4 `word.delete_field`

```json
{
  "type": "object",
  "required": ["session_id", "field_index"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "field_index": { "type": "integer", "minimum": 0 },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

Deletes the addressed field with `Field.delete()`. Deletion removes the field
and its current result text; it does not preserve a plain-text copy. Out-of-range
indices fail with `INVALID_ARGUMENT` and `partial_effect: "none"` before any
deletion. The success response includes `{ deleted: true, field_index, count }`,
where `count` is the remaining number of fields. With `validate_only: true`, the
add-in resolves the target field and returns the same current metadata without
deleting it.

| Operation | API | Requirement set |
|---|---|---|
| Enumerate fields | `Document.fields` / `Range.fields` | `WordApi 1.4` |
| Read field metadata | `Field.type`, `Field.code`, `Field.result`, `Field.locked` | `WordApi 1.4` |
| Insert field | `Range.insertField(insertLocation, fieldType, text, removeFormatting)` | `WordApi 1.5` |
| Refresh field result | `Field.updateResult()` | `WordApi 1.5` |
| Lock/unlock field | `Field.locked` | `WordApi 1.5` |
| Delete field | `Field.delete()` | `WordApi 1.5` |

`word.list_fields` is advertised when `WordApi 1.4` is available. Mutating
field tools are advertised only when the `WordApi 1.5` notes, content-controls,
and fields tier is available.

## 8. Review

### 8.1 `word.add_comment`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "text"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "text": { "type": "string" }
  }
}
```

Comments are authored as the signed-in Office user. The agent operating on
behalf of the user is the user, in the same way that a macro the user runs is
the user. No "AI watermark" is added: the value of office-mcp comes from being
indistinguishable from the user doing it themselves.

### 8.2 `word.resolve_comment`

```json
{
  "type": "object",
  "required": ["session_id", "comment_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "comment_id": { "type": "string", "minLength": 1 }
  }
}
```

### 8.3 `word.update_comment`

```json
{
  "type": "object",
  "required": ["session_id", "comment_id", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "comment_id": { "type": "string", "minLength": 1 },
    "action": { "enum": ["reply", "edit", "delete", "reopen"] },
    "text": { "type": "string" },
    "reply_id": { "type": "string", "minLength": 1 },
    "validate_only": { "type": "boolean", "default": false }
  },
  "additionalProperties": false
}
```

`reply` requires `text` and appends to the thread. `edit` requires `text` and
edits the top-level comment unless `reply_id` is supplied, in which case it
edits that reply. `delete` removes either the top-level thread or the selected
reply and is classified destructive. `reopen` clears the resolved state on the
top-level comment and is the inverse of `word.resolve_comment`.

The comments resource returns each thread as `{ comment_id, content, resolved,
author, created_at, replies: [{ reply_id, content, author, created_at }] }`.
Unknown `comment_id` or `reply_id` fails with `INVALID_ARGUMENT` and
`partial_effect: "none"` before mutation. `validate_only: true` resolves the
target thread or reply and returns the current metadata without mutating.

| Operation | API | Requirement set |
|---|---|---|
| Reply | `Comment.reply(text)` | `WordApi 1.4` |
| Edit thread | `Comment.content` | `WordApi 1.4` |
| Edit reply | `CommentReply.content` | `WordApi 1.4` |
| Delete thread/reply | `Comment.delete()` / `CommentReply.delete()` | `WordApi 1.4` |
| Reopen/read state | `Comment.resolved` | `WordApi 1.4` |
| Enumerate replies | `Comment.replies` | `WordApi 1.4` |

### 8.4 `word.set_change_tracking`

```json
{
  "type": "object",
  "required": ["session_id", "mode"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "mode": { "enum": ["off", "track_all", "track_mine_only"] }
  },
  "additionalProperties": false
}
```

The tool maps to `Document.changeTrackingMode` (`WordApi 1.4`) and returns
`{ previous_mode, mode }`. The accepted modes map to Office.js values as
follows: `off` -> `Word.ChangeTrackingMode.off`, `track_all` ->
`trackAll`, and `track_mine_only` -> `trackMineOnly`. Unsupported host
capabilities return `HOST_CAPABILITY_UNAVAILABLE` before mutation.

### 8.5 `word.update_tracked_change`

Stable Office.js tracked-change objects do not expose an ID. The tracked
changes resource therefore returns each item as
`{ index, author, date, type, text, fingerprint }`, where `index` is the
current collection index and `fingerprint` is a daemon-defined hash of the
loaded fields.

```json
{
  "type": "object",
  "required": ["session_id", "action"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "action": { "enum": ["accept", "reject", "accept_all", "reject_all"] },
    "change_index": { "type": "integer", "minimum": 0 },
    "expected_fingerprint": { "type": "string", "minLength": 1 },
    "expected_count": { "type": "integer", "minimum": 0 }
  },
  "additionalProperties": false,
  "allOf": [
    {
      "if": { "properties": { "action": { "enum": ["accept", "reject"] } } },
      "then": { "required": ["change_index", "expected_fingerprint"] }
    },
    {
      "if": { "properties": { "action": { "enum": ["accept_all", "reject_all"] } } },
      "then": { "required": ["expected_count"] }
    }
  ]
}
```

The add-in reloads the collection immediately before mutation. An index or
fingerprint mismatch returns `STALE_INDEX`; clients must re-read the resource.
For bulk actions, the add-in compares the live tracked-change count to
`expected_count` before calling `TrackedChangeCollection.acceptAll()` or
`rejectAll()`. A count mismatch returns `STALE_INDEX` with no mutation.
Single-change `accept` and `reject` remain edit actions guarded by current
index and fingerprint. Bulk `accept_all` and `reject_all` are destructive
actions because they finalize every tracked revision in the document and are
available only under the All permission ceiling.

### 8.5 `word.accept_change` and `word.reject_change`

Compatibility tools retained until the target-surface migration removes them
from the advertised catalog. New clients should call `word.update_tracked_change`.

Stable Office.js tracked-change objects do not expose an ID. The tracked
changes resource therefore returns each item as
`{ index, author, date, type, text, fingerprint }`, where `index` is the
current collection index and `fingerprint` is a daemon-defined hash of the
loaded fields.

Each tool uses:

```json
{
  "type": "object",
  "required": ["session_id", "change_index", "expected_fingerprint"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "change_index": { "type": "integer", "minimum": 0 },
    "expected_fingerprint": { "type": "string", "minLength": 1 }
  }
}
```

The add-in reloads the collection immediately before mutation. An index or
fingerprint mismatch returns `STALE_INDEX`; clients must re-read the resource.

## 9. Document

### 9.1 `word.save`

```json
{ "type": "object", "required": ["session_id"], "properties": { "session_id": { "type": "string" } } }
```

Calls `Word.Document.save(Word.SaveBehavior.save)`. `Word.Document.saved`
provides the before/after dirty-state signal in the core tier.

`word.save_as` and `word.export_pdf` are not v1 tools. Stable Office.js can
prompt/save the current document and can name a new document, but it does not
provide a portable arbitrary-path Save As or Word-to-PDF byte export. These
names remain reserved for a future API-backed design.

## 10. Behavior contracts

### 10.1 Mutation consistency

Office.js batches queued operations at `context.sync()`, but does not provide a
general transaction or rollback guarantee. A mutating tool MUST:

1. Resolve anchors and preflight known permissions, indices, and constraints.
2. Queue related writes together where the API permits.
3. Minimize the number of `context.sync()` calls after the first write.
4. Report `data.partial_effect` as `none`, `possible`, or `unknown` on failure.

Tools MUST NOT claim atomicity unless the specific Office.js API used documents
that guarantee.

### 10.2 Mutating tool preflight

Mutating Word tools MUST validate all deterministic argument errors before
queuing Office.js writes. This validation is split across two layers:

- The daemon validates JSON schema shape, rejects unsupported fields, enforces
  advertised anchor-kind support, and normalizes daemon-owned inputs such as
  fetched image URLs.
- The Word add-in validates semantic rules that depend on the final tool
  arguments, including required anchors, numeric bounds, mutually exclusive
  options, placement compatibility, and target-object capabilities that can be
  checked before mutation.

When either layer rejects a mutating call before any write is queued, the error
MUST use `INVALID_ARGUMENT` and MUST report `partial_effect: "none"`. Error
messages SHOULD name the tool, field, and expected corrective action when that
information is available.

### 10.3 Validation-only mode

Mutating tools that advertise `validate_only: true`, including
`word.insert_image`, `word.replace_text`, `word.update_paragraph`,
`word.delete_range`, note lifecycle mutations, header/footer updates, and field
lifecycle mutations, MUST implement validation-only mode as a no-write preflight.
Validation-only mode is not a transaction preview. The add-in MAY call read-only
Office.js APIs and `context.sync()` to resolve anchors, count matches, or load
target metadata, but it MUST NOT queue a write before returning.

Successful validation-only responses MUST include:

- `valid: true`
- `operation`: the tool name
- `partial_effect: "none"`
- `resolved_target` or equivalent tool-specific planning metadata when a target
  can be resolved

Invalid validation-only requests use the normal MCP error envelope with
`INVALID_ARGUMENT` and `partial_effect: "none"`. If the add-in can identify a
safe correction, it SHOULD include a small structured `suggestion` object.

### 10.4 Undo grouping

The add-in does not control Word's undo stack or custom undo labels through a
portable Office.js API. Implementations SHOULD group writes into as few sync
boundaries as possible, but the number and labels of undo entries are
host-defined. Tool results do not promise one-keystroke rollback.

### 10.5 IRM enforcement

The following table describes the conceptual right required by each category:

| Category | Required right |
|---|---|
| Read | `extract` (or `view` if no body text returned) |
| Insert | `edit` |
| Edit | `edit` |
| Tables | `edit` |
| Structure | `edit` |
| Notes | `edit` |
| Review | `comment` (or `edit` if `comment` not granted by policy) |
| Document.save | `edit` |

If the host exposes effective rights, the add-in SHOULD fail fast. Production
Word APIs do not currently guarantee a complete effective-rights enumeration,
so the add-in otherwise attempts the operation and maps a host access-denied
failure to `IRM_DENIED`. `denied_rights` and `granted_rights` are included
only when known; clients MUST tolerate their absence.

### 10.6 Track-changes interaction

If the document has Track Changes ON, Word decides how edits are recorded.
When `WordApi 1.6` is available, the add-in may re-read tracked changes after
the edit and return current indices and fingerprints it can correlate
reliably. Tool success does not depend on producing revision references.

### 10.7 What the tool does NOT do

- It does NOT print.
- It does NOT mail-merge.
- It does NOT expose a VBA execution tool.
- It does NOT change document protection settings.
- It does NOT bypass IRM.

These exclusions are deliberate; see [00-overview.md §2](00-overview.md).
