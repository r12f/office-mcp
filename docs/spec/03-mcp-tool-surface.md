# 03 — MCP Tool Surface

What the `office-mcp` server exposes to MCP clients (the AI side).

## 1. Naming & namespacing

- Tool names are `<app>.<verb_object>` — e.g. `word.insert_paragraph`,
  `word.replace_text`, `excel.read_range`. `<app>` is the Office application
  the tool targets (`word`, `excel`, `outlook`, ...).
- A small set of cross-app management tools live under `office.*`:
  `office.list_sessions`, `office.get_session_info`, `office.activate_session`.
- Resource URIs use a custom scheme: `office://<app>/<session_id>/<path>`.
  `<app>` occupies the URI authority slot purely as a namespace —
  there is no network host; the daemon is a local singleton (see
  [07-deployment.md](07-deployment.md) §0).
  Example: `office://word/44444444-4444.../document` (whole doc),
  `office://word/44444444-4444.../paragraph/12` (one paragraph),
  `office://word/44444444-4444.../comments` (all comments).
  `office://excel/...` and `office://outlook/...` are reserved for future
  apps and follow the same shape.

## 2. Sessions as the unit of addressing

Every Office app instance ↔ one or more **document sessions**. A session ID is
required for any document-affecting tool call.

Discovery flow for an MCP client:

1. Call `office.list_sessions` → get back an array of session descriptors.
2. Pick one (typically the user's foreground document, identifiable via
   `is_active: true`).
3. Pass that `session_id` to subsequent tool calls.

The session ID is stable for the life of the document session and is the
**same** ID used in the add-in protocol ([02-registration-protocol.md](02-registration-protocol.md)).

## 3. Resources

MCP resources allow clients to read document state declaratively. All resources
are read-only; mutations happen via tools.

| URI pattern | Returns | Notes |
|---|---|---|
| `office://sessions` | List of session descriptors | Roughly `office.list_sessions` as a resource |
| `office://<app>/<session_id>/document` | Full plain text | Honors IRM: returns 403 if `extract` right denied |
| `office://<app>/<session_id>/structure` | JSON outline (headings, lists, tables) | Lightweight |
| `office://<app>/<session_id>/paragraph/<index>` | Single paragraph | |
| `office://word/<session_id>/comments` | All comments JSON | |
| `office://word/<session_id>/track_changes` | Tracked changes JSON | |
| `office://word/<session_id>/selection` | Currently selected range text + metadata | |

## 4. Management tools

### 4.1 `office.list_sessions`

No arguments. Returns:

```json
{
  "sessions": [
    {
      "session_id": "44444444-...",
      "instance_id": "22222222-...",
      "app": "word",
      "document": {
        "title": "Q3 Report",
        "url": "C:\\Users\\riff\\Documents\\Q3-Report.docx",
        "is_dirty": false,
        "is_protected": true,
        "protection_kind": "irm",
        "rights": ["view", "edit", "extract"]
      },
      "is_active": true,
      "registered_at": "2026-06-14T01:25:01Z"
    }
  ]
}
```

### 4.2 `office.get_session_info`

Args: `{ "session_id": "..." }`. Returns the same descriptor shape, useful when
the client only kept the ID.

### 4.3 `office.activate_session`

Args: `{ "session_id": "..." }`. Brings the document to the foreground in its
Office instance (calls `Office.context.document.bringToFront()` / equivalent).
Use case: the agent wants the user to look at what it's about to change.

## 5. Word tool catalog — see [04-word-capabilities.md](04-word-capabilities.md)

This file lists the categories; the per-tool schemas are in `04-word-capabilities.md`.

| Category | Tools |
|---|---|
| **Read** | `word.get_text`, `word.get_outline`, `word.get_paragraph`, `word.find_text`, `word.get_selection` |
| **Insert** | `word.insert_paragraph`, `word.insert_heading`, `word.insert_table`, `word.insert_image`, `word.insert_page_break`, `word.insert_list` |
| **Edit** | `word.replace_text`, `word.update_paragraph`, `word.delete_range`, `word.apply_formatting` |
| **Tables** | `word.read_table`, `word.update_cell`, `word.add_row`, `word.add_column`, `word.format_cell` |
| **Structure** | `word.set_heading_level`, `word.apply_style`, `word.create_style` |
| **Review** | `word.add_comment`, `word.resolve_comment`, `word.accept_change`, `word.reject_change` |
| **Document** | `word.save`, `word.save_as`, `word.export_pdf` |

## 6. Anchor model

Many tools take an `anchor` argument that describes *where* in the document to act.
Anchor variants:

```json
{ "kind": "selection" }
{ "kind": "start_of_document" }
{ "kind": "end_of_document" }
{ "kind": "paragraph_index", "index": 12 }
{ "kind": "before_paragraph_index", "index": 12 }
{ "kind": "after_paragraph_index", "index": 12 }
{ "kind": "after_text", "text": "Introduction", "occurrence": 1 }
{ "kind": "before_text", "text": "Conclusion", "occurrence": 1 }
{ "kind": "heading", "text": "Methodology", "level": 2 }
{ "kind": "named_range", "name": "ResultsSection" }
```

Paragraph indices are 0-based and refer to the *current* document state at the
moment the tool call is processed. Clients must not cache indices across edits.

`after_text` / `before_text` resolution rules:

- `occurrence: 1` = first match, scanning top-to-bottom.
- Case-insensitive by default; pass `match_case: true` in the surrounding tool args.
- Returns error `-32422` if no match.
- If `occurrence > 1` and fewer matches exist, returns error `-32422`.

## 7. Concurrent calls and ordering

- Calls to the same `session_id` are serialized inside the add-in (Office.js
  requires single-threaded batched access).
- Calls to different sessions run in parallel.
- The MCP client can issue parallel calls to one session; the server enqueues
  them in arrival order. There is no transaction; intermediate state may be
  observable to other clients sharing the session.

If two clients race on the same document, **last-write-wins**. v2 may add a
`acquire_edit_lock` tool that returns a lock token required for subsequent
edits.

## 8. Pagination

Read operations that may return large bodies (`word.get_text`,
`office://.../document`) MUST support:

- `offset` (paragraph offset, 0-based)
- `limit` (max paragraphs, default 200)

The server enforces a hard cap of `MAX_RESPONSE_BYTES` (default 1 MiB);
exceeding it returns `error.code = -32600` with `data.max_response_bytes`.

## 9. Tool metadata

Each tool's MCP metadata includes:

```json
{
  "since": "0.1.0",
  "side_effects": "edit" | "read" | "save",
  "irm_rights_required": ["edit"],
  "estimated_duration_ms": 200,
  "user_visible_change": true
}
```

`user_visible_change: true` means the tool will produce a flash on the
user's screen (the document changes). MCP clients are encouraged to surface this
in their UI ("Word document will be modified").

## 10. Prompts (templated, optional)

The server exposes a small set of MCP prompts that demonstrate idiomatic use:

| Prompt | Args | Body sketch |
|---|---|---|
| `summarize_active_document` | (none) | "Read the active Word session via `office://.../document`, summarize in 200 words, post the summary as a comment via `word.add_comment` on paragraph 0." |
| `polish_section` | `heading: string` | "Find heading, polish prose, present changes for user accept before applying." |
| `extract_action_items` | (none) | "Read document, extract action items, return as JSON; do not modify the doc." |

Prompts are optional sugar; everything they do is achievable via raw tool calls.
