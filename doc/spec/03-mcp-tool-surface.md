# 03 — MCP Tool Surface

What the `office-mcp` server exposes to MCP clients (the AI side).

## 1. Naming & namespacing

- Tool names are `<app>.<verb_object>` — e.g. `word.insert_paragraph`,
  `word.replace_text`, `excel.read_range`, `outlook.draft_reply`.
  `<app>` is the Office application the tool targets
  (`word`, `excel`, `outlook`, ...).
- A small set of cross-app management tools live under `office.*`:
  `office.list_sessions`, `office.get_session_info`.
- Resource URIs use a custom scheme:
  `office://<app>/<session_id>/<app-specific-path>`.
  The `<app>` segment is the URI authority — purely a namespace, not a
  network host (the daemon is a strictly local singleton; see
  [07-deployment.md](07-deployment.md) §0). Resources under each app are
  defined by that app's spec — Word resources are listed in §3 below. Excel v1
  intentionally uses tools rather than MCP resources; see
  [04-excel-capabilities.md](04-excel-capabilities.md). Outlook resources will
  be added with the future Outlook capability doc.

## 2. Sessions as the unit of addressing

Every Office app instance ↔ one or more **document sessions**. A session ID is
required for any document-affecting tool call.

Discovery flow for an MCP client:

1. Call `office.list_sessions` → get back an array of session descriptors.
2. Pick one. `is_active` may help on hosts that expose active-window state,
   but it is nullable; clients must ask the user when selection is ambiguous.
3. Pass that `session_id` to subsequent tool calls.

The session ID is stable for the life of the document session and is the
**same** ID used in the add-in protocol ([02-registration-protocol.md](02-registration-protocol.md)).

## 3. Resources (Word v1)

MCP resources allow clients to read document state declaratively. All resources
are read-only; mutations happen via tools.

The table below is the **Word v1** resource surface. Excel v1 does not expose
MCP resources; clients read workbook state through `excel.read_range` in
[04-excel-capabilities.md](04-excel-capabilities.md). Outlook will ship its own
resource table with a future capability doc; app resource types do not share a
generic "document" abstraction.

| URI pattern | Returns | Notes |
|---|---|---|
| `office://sessions` | List of session descriptors across all apps | Roughly `office.list_sessions` as a resource |
| `office://word/<session_id>/document?offset=0&limit=200` | Paginated plain text | Honors IRM; denial is an MCP resource error carrying `IRM_DENIED` |
| `office://word/<session_id>/structure` | JSON outline (headings, lists, tables) | Lightweight |
| `office://word/<session_id>/paragraph/<index>` | Single paragraph | |
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
        "rights": null,
        "rights_source": "unavailable"
      },
      "is_active": null,
      "capability_tiers": ["core", "review", "tracked_changes"],
      "available_tool_count": 27,
      "registered_at": "2026-06-14T01:25:01Z"
    }
  ]
}
```

### 4.2 `office.get_session_info`

Args: `{ "session_id": "..." }`. Returns the descriptor plus the full
`available_tools` array, useful when the client only kept the ID or needs to
plan around host capabilities.

### 4.3 `office.activate_session`

Deferred from v1. Office.js does not provide a portable API that reliably
brings an arbitrary document window to the foreground. Clients may use
`word.select_range` in a future capability set to make a location visible
inside an already-active document.

## 5. Per-app tool catalogs

This file defines the *cross-cutting* surface (URI scheme, sessions,
management tools, anchor model, pagination, concurrency, metadata, prompts).
Per-app tool catalogs and their JSON Schemas live in each app's capability
doc:

- Word: [04-word-capabilities.md](04-word-capabilities.md)
- Excel: [04-excel-capabilities.md](04-excel-capabilities.md)
- Outlook: future Outlook capability doc

Tool names are not re-listed here on purpose — the per-app doc is the single
source of truth, and any list in this file would drift on the first PR that
adds or renames a tool.

## 6. Anchor model

Many tools take an `anchor` argument that describes *where* in the document to act.
Anchor variants:

```jsonc
{ "kind": "selection" }
{ "kind": "start_of_document" }
{ "kind": "end_of_document" }
{ "kind": "paragraph_index", "index": 12 }
{ "kind": "before_paragraph_index", "index": 12 }
{ "kind": "after_paragraph_index", "index": 12 }
{ "kind": "after_text", "text": "Introduction", "occurrence": 1 }
{ "kind": "before_text", "text": "Conclusion", "occurrence": 1 }
{ "kind": "heading", "text": "Methodology", "level": 2 }
{ "kind": "bookmark", "name": "ResultsSection" }
```

Paragraph indices are 0-based and refer to the *current* document state at the
moment the tool call is processed. Clients must not cache indices across edits.

`after_text` / `before_text` resolution rules:

- `occurrence: 1` = first match, scanning top-to-bottom.
- Case-insensitive by default; pass `match_case: true` in the surrounding tool args.
- Returns `ANCHOR_NOT_FOUND` if no match.
- If `occurrence > 1` and fewer matches exist, returns `ANCHOR_NOT_FOUND`.

## 7. Concurrent calls and ordering

- Calls to the same `session_id` are serialized by the daemon. One call is
  dispatched and up to `MAX_PENDING_PER_SESSION` additional calls wait in a
  FIFO queue.
- Calls to different sessions run in parallel.
- The MCP client can issue parallel calls to one session; the server enqueues
  them in arrival order. There is no transaction; intermediate state may be
  observable to other clients sharing the session.

If two clients race on the same document, **last-write-wins**. v2 may add a
`acquire_edit_lock` tool that returns a lock handle required for subsequent
edits.

## 8. Pagination

Read operations that may return large bodies (`word.get_text`,
`office://.../document`) MUST support:

- `offset` (paragraph offset, 0-based)
- `limit` (max paragraphs, default 200)

Tools carry these as arguments. Resources carry them as URI query parameters.
There is no unbounded "read the entire document" response.

The server enforces a hard cap of `MAX_RESPONSE_BYTES` (default 1 MiB);
exceeding it returns `MAX_RESPONSE_SIZE` with `max_response_bytes`.

## 9. Tool metadata

Each tool's project metadata is carried under MCP `_meta` so the standard tool
shape remains valid:

```jsonc
{
  "_meta": {
    "com.office-mcp/since": "0.1.0",
    "com.office-mcp/side_effects": "edit",
    "com.office-mcp/irm_rights_required": ["edit"],
    "com.office-mcp/minimum_requirement_sets": { "WordApi": "1.3" },
    "com.office-mcp/estimated_duration_ms": 200
  },
  "annotations": {
    "readOnlyHint": false,
    "destructiveHint": false,
    "idempotentHint": false,
    "openWorldHint": false
  }
}
```

The daemon lists the stable server-wide catalog. Session-specific support is
reported by `available_tools`; invoking a listed tool against an incompatible
session yields a tool execution error, not a protocol error.

## 10. Prompts (templated, optional)

The server exposes a small set of MCP prompts that demonstrate idiomatic use:

| Prompt | Args | Body sketch |
|---|---|---|
| `summarize_document` | `session_id: string` | "Read the selected Word session via `office://.../document`, summarize in 200 words, post the summary as a comment via `word.add_comment` on paragraph 0." |
| `polish_section` | `session_id: string, heading: string` | "Find heading, polish prose, present changes for user accept before applying." |
| `extract_action_items` | `session_id: string` | "Read document, extract action items, return as JSON; do not modify the doc." |

Prompts are optional sugar; everything they do is achievable via raw tool calls.
