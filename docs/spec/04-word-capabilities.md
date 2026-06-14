# 04 — Word Capabilities (v1)

Per-tool JSON Schemas. All tools take a top-level `session_id` (string, UUID, required).
Common keys (anchor, etc.) are defined in [03-mcp-tool-surface.md §6](03-mcp-tool-surface.md).

## 1. Read

### 1.1 `word.get_text`

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

### 1.2 `word.get_outline`

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

### 1.3 `word.get_paragraph`

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

### 1.4 `word.find_text`

```json
{
  "type": "object",
  "required": ["session_id", "query"],
  "properties": {
    "session_id": { "type": "string" },
    "query": { "type": "string" },
    "match_case": { "type": "boolean", "default": false },
    "whole_word": { "type": "boolean", "default": false },
    "regex": { "type": "boolean", "default": false },
    "limit": { "type": "integer", "default": 50 }
  }
}
```

Returns matches as `[{ paragraph_index, start_offset, end_offset, snippet }]`.

### 1.5 `word.get_selection`

```json
{ "type": "object", "required": ["session_id"], "properties": { "session_id": { "type": "string" } } }
```

Returns `{ text, paragraph_index, start_offset, end_offset, is_empty }`.

## 2. Insert

### 2.1 `word.insert_paragraph`

```json
{
  "type": "object",
  "required": ["session_id", "text", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "text": { "type": "string" },
    "anchor": { "$ref": "#/definitions/anchor" },
    "style": { "type": "string", "default": "Normal" },
    "formatting": { "$ref": "#/definitions/run_formatting" }
  }
}
```

### 2.2 `word.insert_heading`

```json
{
  "type": "object",
  "required": ["session_id", "text", "level", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "text": { "type": "string" },
    "level": { "type": "integer", "minimum": 1, "maximum": 9 },
    "anchor": { "$ref": "#/definitions/anchor" }
  }
}
```

### 2.3 `word.insert_table`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "rows", "cols"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/definitions/anchor" },
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

### 2.4 `word.insert_image`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "image"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/definitions/anchor" },
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
mixed-content issues in the add-in webview.

### 2.5 `word.insert_page_break` and `word.insert_list`

Schemas in same vein, omitted here for brevity; see reference implementation.

## 3. Edit

### 3.1 `word.replace_text`

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
    "regex": { "type": "boolean", "default": false },
    "scope": {
      "type": "object",
      "properties": {
        "paragraph_range": { "type": "array", "items": { "type": "integer" }, "minItems": 2, "maxItems": 2 },
        "selection_only":  { "type": "boolean", "default": false }
      }
    },
    "dry_run": { "type": "boolean", "default": false }
  }
}
```

Returns `{ replaced_count: N, matches: [...] }`.

`dry_run: true` returns matches without modifying the document. **Highly
recommended pattern for agents**: dry-run first, present to user, then run again
without `dry_run`.

### 3.2 `word.update_paragraph`

Replace one paragraph's text wholesale.

```json
{
  "type": "object",
  "required": ["session_id", "index", "text"],
  "properties": {
    "session_id": { "type": "string" },
    "index": { "type": "integer" },
    "text": { "type": "string" },
    "preserve_formatting": { "type": "boolean", "default": true }
  }
}
```

### 3.3 `word.delete_range`

```json
{
  "type": "object",
  "required": ["session_id", "anchor"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/definitions/anchor" },
    "extent": {
      "oneOf": [
        { "const": "paragraph" },
        { "const": "sentence" },
        { "type": "object", "required": ["chars"], "properties": { "chars": { "type": "integer" } } }
      ]
    }
  }
}
```

### 3.4 `word.apply_formatting`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "formatting"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/definitions/anchor" },
    "extent": { "$ref": "#/definitions/extent" },
    "formatting": { "$ref": "#/definitions/run_formatting" }
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

## 4. Tables

### 4.1 `word.read_table`

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

### 4.2 `word.update_cell`

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
    "formatting": { "$ref": "#/definitions/run_formatting" }
  }
}
```

### 4.3 `word.add_row` and `word.add_column`

Trivial schemas; add at index or end.

### 4.4 `word.format_cell`

Background color, padding, alignment, borders, merging. Schema mirrors
python-docx's table cell API for v1 to ease porting of existing tools.

## 5. Structure

### 5.1 `word.set_heading_level` / `word.apply_style` / `word.create_style`

Allow agents to maintain document structure declaratively without touching
formatting directly.

## 6. Review

### 6.1 `word.add_comment`

```json
{
  "type": "object",
  "required": ["session_id", "anchor", "text"],
  "properties": {
    "session_id": { "type": "string" },
    "anchor": { "$ref": "#/definitions/anchor" },
    "text": { "type": "string" },
    "author_display_name": { "type": "string" }
  }
}
```

If `author_display_name` is omitted, the add-in uses `"AI agent (via office-mcp)"`.
This is a deliberate UX choice: comments authored by the AI MUST be visually
distinguishable from user comments. Servers MAY enforce a `[AI]` prefix in v2.

### 6.2 `word.resolve_comment` / `word.accept_change` / `word.reject_change`

Standard review-pane operations. Each takes a `comment_id` or `change_id`
discoverable via the corresponding read resource.

## 7. Document

### 7.1 `word.save`

```json
{ "type": "object", "required": ["session_id"], "properties": { "session_id": { "type": "string" } } }
```

Triggers `Office.context.document.save()`. No-op if already clean.

### 7.2 `word.save_as`

```json
{
  "type": "object",
  "required": ["session_id", "path"],
  "properties": {
    "session_id": { "type": "string" },
    "path": { "type": "string" },
    "format": { "enum": ["docx", "pdf", "rtf", "txt", "html"], "default": "docx" }
  }
}
```

`path` is on the user's machine, NOT a server path. The add-in resolves it.
The add-in MUST refuse paths it cannot validate (no traversal beyond the user's
documents folder unless explicitly confirmed by the user via Office's file picker).

### 7.3 `word.export_pdf`

Convenience wrapper around `save_as` with `format: "pdf"`. Returns the file path.

## 8. Behavior contracts

### 8.1 Atomicity

Each tool call corresponds to one Office.js batch (one `context.sync()`).
Either the entire call's effect lands or none of it does. Within a call, the
add-in MUST not commit a partial effect on error.

### 8.2 Undo grouping

Each tool call creates one undo entry in Word, labeled with the tool name
(e.g. "AI: Insert paragraph"). The user can `Ctrl+Z` the entire tool call
in one keystroke.

### 8.3 IRM enforcement

The add-in MUST check the active document's effective rights before executing
each call. Mapping from tool category to required right:

| Category | Required right |
|---|---|
| Read | `extract` (or `view` if no body text returned) |
| Insert | `edit` |
| Edit | `edit` |
| Tables | `edit` |
| Structure | `edit` |
| Review | `comment` (or `edit` if `comment` not granted by policy) |
| Document.save_as PDF | `export` |
| Document.save_as DOCX (different path) | `extract` + `edit` |

Denied calls return `error.code = -32403` with `data.denied_rights`.

### 8.4 Track-changes interaction

If the document has Track Changes ON, tools that edit content (`insert_paragraph`,
`replace_text`, etc.) produce tracked changes, NOT direct edits. The tool result
includes `tracked_change_ids: [...]` so the agent can refer to them later.

### 8.5 What the tool does NOT do

- It does NOT print.
- It does NOT mail-merge.
- It does NOT execute VBA macros (refused with `-32601`).
- It does NOT change document protection settings.
- It does NOT bypass IRM.

These exclusions are deliberate; see [00-overview.md §2](00-overview.md).
