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

The current advertised Word v1 tool surface has 31 tools, grouped by object-owner
category. Categories are not permission tiers and must not be action buckets such
as `Read`, `Insert`, and `Edit`; side-effect level is tracked separately per
tool so the UI can apply Read/Write/All permission modes without hiding the
object model.

The per-tool JSON Schemas follow in §2-§9.

| Category | Tools |
|---|---|
| **Document & structure** | `word.get_text`, `word.get_outline`, `word.get_header_footer`, `word.update_header_footer`, `word.insert_break`, `word.list_sections`, `word.update_page_setup`, `word.save` |
| **Range & selection** | `word.get_selection`, `word.find_text`, `word.resolve_anchor`, `word.insert_bookmark`, `word.list_bookmarks`, `word.delete_bookmark`, `word.insert_hyperlink`, `word.list_hyperlinks`, `word.remove_hyperlink`, `word.replace_text`, `word.delete_range`, `word.apply_formatting`, `word.apply_style` |
| **Paragraphs & lists** | `word.get_paragraph`, `word.insert_paragraph`, `word.update_paragraph`, `word.insert_list` |
| **Tables** | `word.insert_table`, `word.read_table`, `word.update_table` |
| **Media** | `word.insert_image`, `word.resize_image` |
| **Content controls** | `word.list_content_controls`, `word.insert_content_control`, `word.update_content_control`, `word.delete_content_control` |
| **Notes** | `word.insert_note`, `word.list_notes`, `word.update_note`, `word.delete_note` |
| **Review** | `word.add_comment`, `word.resolve_comment`, `word.update_comment`, `word.update_tracked_change` |

### 1.2 Target refined Word tool surface

The target refined Word surface is based on the Microsoft Word add-in object
model: `Document` contains sections and document-level state; a section has a
`Body`; `Body` and `Range` own most text operations; higher-level objects such
as paragraphs, lists, tables, content controls, comments, and tracked changes
own object-specific lifecycle and review workflows.

The target surface has 42 tools. It deliberately consolidates specialized tools
that perform the same user intent under a single owner. Superseded
compatibility tools remain documented below for migration history, but they are
not advertised by the daemon catalog or task pane available-tools metadata.

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `word.get_text` | implemented | Document & structure | read | `WordApi 1.3` | Read paginated document body text; paragraph metadata is optional. |
| `word.get_outline` | implemented | Document & structure | read | `WordApi 1.3` | Read headings and lightweight document structure without body text. |
| `word.get_header_footer` | implemented | Document & structure | read | `WordApi 1.1` | Read section-scoped header or footer text and optional paragraph metadata; non-primary layout validation uses `WordApiDesktop 1.3` when required. |
| `word.update_header_footer` | implemented | Document & structure | edit/destructive | `WordApi 1.1` | Replace, append to, or clear a section-scoped header or footer body; non-primary layout validation uses `WordApiDesktop 1.3` when required. |
| `word.get_paragraph` | implemented | Paragraphs & lists | read | `WordApi 1.3` | Read one paragraph by index, optionally including direct paragraph formatting metadata. |
| `word.find_text` | implemented | Range & selection | read | `WordApi 1.3` | Search text with Word search options and return portable paragraph-relative matches. |
| `word.resolve_anchor` | implemented | Range & selection | read | `WordApi 1.3` | Resolve an anchor to safe diagnostic metadata without returning full document text. |
| `word.insert_bookmark` | implemented | Range & selection | edit | `WordApi 1.4` | Create or move a named bookmark at an anchored range with explicit duplicate handling. |
| `word.list_bookmarks` | implemented | Range & selection | read | `WordApi 1.4` | List bookmark names and bounded location previews without returning full document text. |
| `word.delete_bookmark` | implemented | Range & selection | destructive | `WordApi 1.4` | Delete a bookmark marker without deleting the bookmarked text. |
| `word.get_selection` | implemented | Range & selection | read | `WordApi 1.3` | Read current selection text and simple selection metadata. |
| `word.insert_hyperlink` | implemented | Range & selection | edit | `WordApi 1.3` | Create a hyperlink on an anchored range or inserted text with URL scheme validation. |
| `word.list_hyperlinks` | implemented | Range & selection | read | `WordApi 1.3` | List hyperlink text and URLs with paragraph-relative locations. |
| `word.remove_hyperlink` | implemented | Range & selection | edit | `WordApi 1.3` | Remove hyperlink targets from an anchored range while preserving text by default. |
| `word.insert_paragraph` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Insert a paragraph at an anchor; also owns heading insertion through style or heading-level arguments after migration. |
| `word.insert_table` | implemented | Tables | edit | `WordApi 1.3` | Insert a table with optional initial data and style. |
| `word.insert_image` | implemented | Media | edit | `WordApi 1.3` | Insert a validated image from base64 or a daemon-fetched HTTPS URL. |
| `word.resize_image` | implemented | Media | edit | `WordApi 1.3` | Resize an existing inline image in place by paragraph index and image index. |
| `word.insert_break` | implemented | Document & structure | edit | `WordApi 1.3` | Insert a page, line, or section break at an anchor. |
| `word.list_sections` | implemented | Document & structure | read | `WordApi 1.3` | List document sections with paragraph range and header/footer metadata. |
| `word.update_page_setup` | implemented | Document & structure | edit | `WordApiDesktop 1.3` | Update document or section page setup such as orientation, margins, and page size. |
| `word.insert_list` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Insert a numbered or bulleted list. |
| `word.replace_text` | implemented | Range & selection | edit | `WordApi 1.3` | Find and replace text, with dry-run support. |
| `word.update_paragraph` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Replace one paragraph's text wholesale. |
| `word.delete_range` | implemented | Range & selection | destructive | `WordApi 1.3` | Delete an anchored paragraph, sentence, or selection. |
| `word.apply_formatting` | implemented | Range & selection | edit | `WordApi 1.3` | Apply direct run and/or paragraph formatting to an anchored range. |
| `word.apply_style` | implemented | Range & selection | edit | `WordApi 1.3` | Apply an Office style to an anchored range; also owns heading-level changes after migration. |
| `word.read_table` | implemented | Table | read | `WordApi 1.3` | Read table dimensions, header state, and cell text. |
| `word.update_table` | implemented | Table | edit/destructive | `WordApi 1.3` | Update table cells, rows, columns, table/cell formatting, and table deletion through one table-owner tool. |
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
| `word.update_tracked_change` | implemented | Review | edit/destructive | `WordApi 1.6` | Accept or reject one tracked change by current index and expected fingerprint. |
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

Tool ownership rules:

- One common user intent has one tool owner. Add a new tool only when it has a
  different object owner, permission profile, or user-visible result.
- Read resources and read tools may expose the same underlying document data,
  but mutation tools must not duplicate each other's writes.
- `word.insert_paragraph` owns paragraph creation, including headings. Do not add
  a separate heading insertion tool after migration.
- `word.apply_style` owns semantic Office style changes, including heading level.
  `word.apply_formatting` owns direct character/run and paragraph layout
  formatting. Direct paragraph layout includes alignment, indentation, spacing,
  and outline level; named styles remain owned by `word.apply_style`.
- `word.insert_hyperlink`, `word.list_hyperlinks`, and
  `word.remove_hyperlink` own hyperlink lifecycle. Generic run styling remains
  owned by `word.apply_formatting`, and bookmark creation/deletion remains out
  of scope for these tools.
- `word.read_table` owns table content reads. `word.update_table` owns table
  structure, cell value, table/cell formatting, and deletion mutations.
- `word.insert_image` owns new image insertion. `word.resize_image` owns in-place
  resizing of an existing inline image and must not require re-uploading image
  bytes or alter surrounding paragraphs.
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
- `word.update_tracked_change` owns tracked-change accept/reject actions. The
  tracked-change resource remains the read owner for current indices and
  fingerprints.
- `word.get_header_footer` and `word.update_header_footer` own header/footer
  body reads and writes. Body, range, and paragraph tools operate on the main
  document body and must not silently reach into headers or footers.
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

### 1.3 Runtime capability tiers

The base manifest requires `WordApi 1.3`. The add-in probes higher sets at
runtime and advertises only tools whose complete implementation is supported:

| Tier | Requirement | Additional tools/features |
|---|---|---|
| Core | `WordApi 1.3` | text, paragraphs, search, insert/edit, tables at start/end, styles, selection |
| Review | `WordApi 1.4` | comments, bookmark anchors and bookmark lifecycle tools |
| Notes and content controls | `WordApi 1.5` | footnote/endnote lifecycle and content-control lifecycle tools |
| Tracked changes | `WordApi 1.6` | tracked-change resource and accept/reject |
| Host-specific | explicit successful probe | page setup, active-window, and protection metadata |

Header/footer body access and `word.save` use production Word APIs available
before the base manifest requirement and are available whenever the core tier is
available. Preview APIs are excluded from v1.

`word.update_page_setup` requires `WordApiDesktop 1.3` and is absent from the
catalog when that probe fails. `word.insert_break` and `word.list_sections` are
core-tier tools because their required Office.js APIs are available within the
base Word API tier.

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

### 3.5 `word.resize_image`

Resize an existing inline image in place without deleting and reinserting its
binary data.

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
    },
    "width_pt": { "type": "number", "exclusiveMinimum": 0 },
    "height_pt": { "type": "number", "exclusiveMinimum": 0 },
    "preserve_aspect_ratio": { "type": "boolean", "default": true }
  },
  "additionalProperties": false
}
```

Callers must provide at least one of `width_pt` or `height_pt`. When
`preserve_aspect_ratio` is true and exactly one dimension is provided, the
add-in derives the other dimension from the current image dimensions. The tool
returns old and new dimensions in points and preserves paragraph placement, alt
text, relationship identity, and adjacent text.

### 3.6 `word.insert_break`

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

Returns `{ rows, cols, data: string[][], header_row: boolean }`.

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
    "action": { "enum": ["update_cell", "add_row", "add_column", "format_cell", "delete"] },
    "row": { "type": "integer", "minimum": 0 },
    "col": { "type": "integer", "minimum": 0 },
    "text": { "type": "string" },
    "index": { "type": "integer", "minimum": 0 },
    "values": { "type": "array", "items": { "type": "string" } },
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
- `delete` deletes the whole table and must be requested explicitly.

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

Cell merging and arbitrary border editing are deferred until their
cross-platform Office.js behavior is verified.

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

`word.create_style` is reserved for a future capability set after its
cross-platform Office.js behavior is verified.

## 7. Content Controls

The content-control tools use generic `Word.ContentControl` APIs only. v1
supports rich text and plain text controls; checkbox, dropdown, combo box,
picture, date picker, repeating-section, group, and desktop-only specialized
control behavior is deferred.

### 7.1 `word.list_content_controls`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "type": { "enum": ["rich_text", "plain_text"] },
    "tag": { "type": "string" },
    "title": { "type": "string" }
  },
  "additionalProperties": false
}
```

Returns `{ content_controls, count }`. Each item includes
`content_control_id`, `tag`, `title`, `type`, `subtype`, `cannot_delete`, and
`cannot_edit`. It does not return the contained document text; callers use read
tools for text.

### 7.2 `word.insert_content_control`

```json
{
  "type": "object",
  "required": ["session_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "type": { "enum": ["rich_text", "plain_text"] },
    "text": { "type": "string" },
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
provided, the anchored range is replaced with that text before wrapping it.

### 7.3 `word.update_content_control`

```json
{
  "type": "object",
  "required": ["session_id", "content_control_id"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "content_control_id": { "type": "integer", "minimum": 0 },
    "text": { "type": "string" },
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

### 8.4 `word.update_tracked_change`

Stable Office.js tracked-change objects do not expose an ID. The tracked
changes resource therefore returns each item as
`{ index, author, date, type, text, fingerprint }`, where `index` is the
current collection index and `fingerprint` is a daemon-defined hash of the
loaded fields.

```json
{
  "type": "object",
  "required": ["session_id", "action", "change_index", "expected_fingerprint"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "action": { "enum": ["accept", "reject"] },
    "change_index": { "type": "integer", "minimum": 0 },
    "expected_fingerprint": { "type": "string", "minLength": 1 }
  },
  "additionalProperties": false
}
```

The add-in reloads the collection immediately before mutation. An index or
fingerprint mismatch returns `STALE_INDEX`; clients must re-read the resource.

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

`word.insert_image`, `word.replace_text`, `word.update_paragraph`, and
`word.delete_range` MUST support `validate_only: true`. Validation-only mode is
not a transaction preview; it is a no-write preflight. The add-in MAY call
read-only Office.js APIs and `context.sync()` to resolve anchors, count matches,
or load target metadata, but it MUST NOT queue a write before returning.

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
