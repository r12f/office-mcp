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
    }
  }
}
```

The `$ref` values below refer to these shared definitions.

## 1. Overview

### 1.1 Word tool catalog

The current advertised Word v1 tool surface has 26 tools, grouped by object-owner
category. Categories are not permission tiers and must not be action buckets such
as `Read`, `Insert`, and `Edit`; side-effect level is tracked separately per
tool so the UI can apply Read/Write/All permission modes without hiding the
object model.

The per-tool JSON Schemas follow in §2-§8.

| Category | Tools |
|---|---|
| **Document & structure** | `word.get_text`, `word.get_outline`, `word.insert_page_break`, `word.save` |
| **Range & selection** | `word.get_selection`, `word.find_text`, `word.replace_text`, `word.delete_range`, `word.apply_formatting`, `word.apply_style` |
| **Paragraphs & lists** | `word.get_paragraph`, `word.insert_paragraph`, `word.update_paragraph`, `word.insert_list` |
| **Tables** | `word.insert_table`, `word.read_table`, `word.update_table` |
| **Media** | `word.insert_image`, `word.resize_image` |
| **Content controls** | `word.list_content_controls`, `word.insert_content_control`, `word.update_content_control`, `word.delete_content_control` |
| **Review** | `word.add_comment`, `word.resolve_comment`, `word.update_tracked_change` |

### 1.2 Target refined Word tool surface

The target refined Word surface is based on the Microsoft Word add-in object
model: `Document` contains sections and document-level state; a section has a
`Body`; `Body` and `Range` own most text operations; higher-level objects such
as paragraphs, lists, tables, content controls, comments, and tracked changes
own object-specific lifecycle and review workflows.

The target surface has 26 tools. It deliberately consolidates specialized tools
that perform the same user intent under a single owner. Superseded
compatibility tools remain documented below for migration history, but they are
not advertised by the daemon catalog or task pane available-tools metadata.

| Tool | Status | Category | Side effect | Minimum API | Summary |
|---|---|---|---|---|---|
| `word.get_text` | implemented | Document & structure | read | `WordApi 1.3` | Read paginated document body text; paragraph metadata is optional. |
| `word.get_outline` | implemented | Document & structure | read | `WordApi 1.3` | Read headings and lightweight document structure without body text. |
| `word.get_paragraph` | implemented | Paragraphs & lists | read | `WordApi 1.3` | Read one paragraph by index. |
| `word.find_text` | implemented | Range & selection | read | `WordApi 1.3` | Search text with Word search options and return portable paragraph-relative matches. |
| `word.get_selection` | implemented | Range & selection | read | `WordApi 1.3` | Read current selection text and simple selection metadata. |
| `word.insert_paragraph` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Insert a paragraph at an anchor; also owns heading insertion through style or heading-level arguments after migration. |
| `word.insert_table` | implemented | Tables | edit | `WordApi 1.3` | Insert a table with optional initial data and style. |
| `word.insert_image` | implemented | Media | edit | `WordApi 1.3` | Insert a validated image from base64 or a daemon-fetched HTTPS URL. |
| `word.resize_image` | implemented | Media | edit | `WordApi 1.3` | Resize an existing inline image in place by paragraph index and image index. |
| `word.insert_page_break` | implemented | Document & structure | edit | `WordApi 1.3` | Insert a page break at an anchor. |
| `word.insert_list` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Insert a numbered or bulleted list. |
| `word.replace_text` | implemented | Range & selection | edit | `WordApi 1.3` | Find and replace text, with dry-run support. |
| `word.update_paragraph` | implemented | Paragraphs & lists | edit | `WordApi 1.3` | Replace one paragraph's text wholesale. |
| `word.delete_range` | implemented | Range & selection | destructive | `WordApi 1.3` | Delete an anchored paragraph, sentence, or selection. |
| `word.apply_formatting` | implemented | Range & selection | edit | `WordApi 1.3` | Apply character/run formatting to an anchored range. |
| `word.apply_style` | implemented | Range & selection | edit | `WordApi 1.3` | Apply an Office style to an anchored range; also owns heading-level changes after migration. |
| `word.read_table` | implemented | Table | read | `WordApi 1.3` | Read table dimensions, header state, and cell text. |
| `word.update_table` | implemented | Table | edit/destructive | `WordApi 1.3` | Update table cells, rows, columns, table/cell formatting, and table deletion through one table-owner tool. |
| `word.list_content_controls` | implemented | Content controls | read | `WordApi 1.5` | List content controls with id/tag/title/type metadata, without duplicating document text reads. |
| `word.insert_content_control` | implemented | Content controls | edit | `WordApi 1.5` | Create a content control around an anchored range or inserted placeholder content. |
| `word.update_content_control` | implemented | Content controls | edit | `WordApi 1.5` | Update content control metadata, locked state, or contained text through the content-control owner. |
| `word.delete_content_control` | implemented | Content controls | destructive | `WordApi 1.5` | Delete a content control, preserving or deleting contents according to an explicit mode. |
| `word.add_comment` | implemented | Review | comment | `WordApi 1.4` | Add a comment to an anchored range as the signed-in Office user. |
| `word.resolve_comment` | implemented | Review | comment | `WordApi 1.4` | Resolve an existing comment. |
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

Tool ownership rules:

- One common user intent has one tool owner. Add a new tool only when it has a
  different object owner, permission profile, or user-visible result.
- Read resources and read tools may expose the same underlying document data,
  but mutation tools must not duplicate each other's writes.
- `word.insert_paragraph` owns paragraph creation, including headings. Do not add
  a separate heading insertion tool after migration.
- `word.apply_style` owns semantic Office style changes, including heading level.
  `word.apply_formatting` owns direct character/run formatting only.
- `word.read_table` owns table content reads. `word.update_table` owns table
  structure, cell value, table/cell formatting, and deletion mutations.
- `word.insert_image` owns new image insertion. `word.resize_image` owns in-place
  resizing of an existing inline image and must not require re-uploading image
  bytes or alter surrounding paragraphs.
- Content-control tools own content-control lifecycle and metadata. Generic text
  edits inside a known range remain owned by range/paragraph tools unless the
  caller is explicitly targeting a content control.
- `word.update_tracked_change` owns tracked-change accept/reject actions. The
  tracked-change resource remains the read owner for current indices and
  fingerprints.

### 1.3 Runtime capability tiers

The base manifest requires `WordApi 1.3`. The add-in probes higher sets at
runtime and advertises only tools whose complete implementation is supported:

| Tier | Requirement | Additional tools/features |
|---|---|---|
| Core | `WordApi 1.3` | text, paragraphs, search, insert/edit, tables at start/end, styles, selection |
| Review | `WordApi 1.4` | comments, bookmark anchors |
| Tracked changes | `WordApi 1.6` | tracked-change resource and accept/reject |
| Host-specific | explicit successful probe | active-window and protection metadata |

`word.save` uses the production `WordApi 1.1` API and is available whenever
the core tier is available. Preview APIs are excluded from v1.

### 1.4 Word resources

Declarative read-only resources (mutations go through tools in §2–§8).
URI scheme and cross-cutting semantics are defined in
[03-mcp-tool-surface.md §1–§3](03-mcp-tool-surface.md).

| URI pattern | Returns | Notes |
|---|---|---|
| `office://word/<session_id>/document?offset=0&limit=200` | Paginated plain text | Honors IRM; denial carries `IRM_DENIED` |
| `office://word/<session_id>/structure` | JSON outline (headings, lists, tables) | Lightweight |
| `office://word/<session_id>/paragraph/<index>` | Single paragraph | |
| `office://word/<session_id>/comments` | All comments JSON | |
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
    "index": { "type": "integer", "minimum": 0 }
  }
}
```

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
    "image": {
      "oneOf": [
        { "type": "object", "required": ["base64"], "properties": { "base64": { "type": "string" } } },
        { "type": "object", "required": ["url"],    "properties": { "url":    { "type": "string", "format": "uri" } } }
      ]
    },
    "alt_text": { "type": "string" },
    "width_pt":  { "type": "number" },
    "height_pt": { "type": "number" }
  }
}
```

`url` MUST be `https://` and is fetched server-side (not by the add-in) to avoid
mixed-content issues in the add-in webview. The daemon applies the fetch policy
in [05-security.md §6.1](05-security.md): no cookies, auth headers, private
addresses, unvalidated redirects, oversized bodies, or non-image payloads.
Base64 input is subject to the same decoded-byte and image-format limits.

For paragraph-resolving anchors (`paragraph_index`, `before_paragraph_index`,
`after_paragraph_index`, and `heading`), `word.insert_image` MUST insert the
image into a clean paragraph adjacent to the resolved paragraph instead of
calling the inline-picture API directly on the `Paragraph` object. The resolved
anchor paragraph's text and style MUST remain unchanged; for example, inserting
after a heading must not append the image into the heading paragraph itself.
If the host cannot support the requested paragraph placement without mutation,
the tool MUST return a specific `INVALID_ARGUMENT` error with
`partial_effect: "none"`.

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

### 3.6 `word.insert_page_break`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string", "format": "uuid" },
    "anchor": { "$ref": "#/$defs/anchor" }
  }
}
```

### 3.7 `word.insert_list`

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
    "partial_ok": { "type": "boolean", "default": false }
  }
}
```

Returns `{ replaced_count: N, matches: [...] }`.

`wildcards` uses Word wildcard syntax, not JavaScript regular expressions.

`dry_run: true` returns matches without modifying the document. **Highly
recommended pattern for agents**: dry-run first, present to user, then run again
without `dry_run`.

### 4.2 `word.update_paragraph`

Replace one paragraph's text wholesale.

```json
{
  "type": "object",
  "required": ["session_id", "index", "text"],
  "properties": {
    "session_id": { "type": "string" },
    "index": { "type": "integer" },
    "text": { "type": "string" }
  }
}
```

Replacing paragraph text preserves paragraph-level style but does not promise
to preserve mixed character-run formatting inside the old text.

### 4.3 `word.delete_range`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "enum": ["paragraph", "sentence", "selection"] }
  }
}
```

### 4.4 `word.apply_formatting`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "formatting"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/$defs/anchor" },
    "extent": { "$ref": "#/$defs/extent" },
    "formatting": { "$ref": "#/$defs/run_formatting" }
  }
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

### 8.3 `word.update_tracked_change`

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

### 8.4 `word.accept_change` and `word.reject_change`

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

### 9.2 Undo grouping

The add-in does not control Word's undo stack or custom undo labels through a
portable Office.js API. Implementations SHOULD group writes into as few sync
boundaries as possible, but the number and labels of undo entries are
host-defined. Tool results do not promise one-keystroke rollback.

### 9.3 IRM enforcement

The following table describes the conceptual right required by each category:

| Category | Required right |
|---|---|
| Read | `extract` (or `view` if no body text returned) |
| Insert | `edit` |
| Edit | `edit` |
| Tables | `edit` |
| Structure | `edit` |
| Review | `comment` (or `edit` if `comment` not granted by policy) |
| Document.save | `edit` |

If the host exposes effective rights, the add-in SHOULD fail fast. Production
Word APIs do not currently guarantee a complete effective-rights enumeration,
so the add-in otherwise attempts the operation and maps a host access-denied
failure to `IRM_DENIED`. `denied_rights` and `granted_rights` are included
only when known; clients MUST tolerate their absence.

### 9.4 Track-changes interaction

If the document has Track Changes ON, Word decides how edits are recorded.
When `WordApi 1.6` is available, the add-in may re-read tracked changes after
the edit and return current indices and fingerprints it can correlate
reliably. Tool success does not depend on producing revision references.

### 9.5 What the tool does NOT do

- It does NOT print.
- It does NOT mail-merge.
- It does NOT expose a VBA execution tool.
- It does NOT change document protection settings.
- It does NOT bypass IRM.

These exclusions are deliberate; see [00-overview.md §2](00-overview.md).
