# 03 — MCP Tool Surface

What the `office-mcp` server exposes to MCP clients (the AI side).

## 1. Naming & namespacing

- Canonical tool names are `<app>.<verb_object>` — e.g.
  `word.insert_paragraph`, `word.replace_text`, `excel.read_range`,
  `outlook.draft_reply`. `<app>` is the Office application the tool targets
  (`word`, `excel`, `outlook`, ...). Canonical names are the names used in
  add-in `available_tools`, audit logs, errors, and daemon-to-add-in forwarding.
- MCP `tools/list` also exposes a client-safe underscore alias for every public
  tool, formed by replacing the namespace dot with `_`, for example
  `word_insert_paragraph`, `word_replace_text`, and `excel_read_range`. These
  aliases exist for MCP clients that only surface function-like tool names.
  Alias entries MUST carry the same schema, annotations, side-effect metadata,
  examples, and common errors as their canonical tool, plus
  `_meta["com.office-mcp/alias_for"]` and
  `_meta["com.office-mcp/canonical_name"]` pointing at the canonical dotted
  name. A `tools/call` using either spelling MUST execute the same canonical
  tool and forward the canonical dotted name to the add-in.
- A small set of cross-app management tools live under `office.*`:
  `office.list_sessions`, `office.get_session_info`, `office.describe_tools`.
- Resource URIs use a custom scheme:
  `office://<app>/<session_id>/<app-specific-path>`.
  The `<app>` segment is the URI authority — purely a namespace, not a
  network host (the daemon is a strictly local singleton; see
  [07-deployment.md](07-deployment.md) §0). Resources under each app are
  defined by that app's spec — Word and PowerPoint resources are listed in §3
  below. Excel v1 tools remain the authoritative execution surface, but the
  daemon also exposes a limited Excel read-only fallback resource surface for
  resource-only clients; see [04-excel-capabilities.md](04-excel-capabilities.md).
  Outlook resources will be added with the future Outlook capability doc.

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

## 3. Resources (Word, Excel, and PowerPoint v1)

MCP resources allow clients to read document state declaratively. All resources
are read-only; mutations happen via tools.

The table below is the **Word v1**, **Excel v1 read-only fallback**, and
**PowerPoint v1 read-only fallback** resource surface. Excel and PowerPoint
tools remain the authoritative execution surface, but the daemon also exposes
read-only resources so MCP clients that can read resources but cannot call
dynamic app tools can still inspect active workbooks and decks. Outlook will
ship its own resource table with a future capability doc; app resource types do
not share a generic "document" abstraction.

The resource fallback is additive. When a host session reports tools in
`available_tools`, those public tools must still be exposed as callable MCP
tools through `tools/list` whenever the daemon Global Tool Access policy allows
them. In particular, PowerPoint action tools such as `powerpoint.add_slide`,
`powerpoint.add_text_box`, `powerpoint.update_shape`,
`powerpoint.replace_text`, and `powerpoint.format_text` are not resource-only
capabilities; clients invoke them with `tools/call` and the daemon then checks
the target session's `available_tools` before forwarding to the add-in.

| URI pattern | Returns | Notes |
|---|---|---|
| `office://sessions` | List of session descriptors across all apps | Roughly `office.list_sessions` as a resource |
| `office://word/<session_id>/document?offset=0&limit=200` | Paginated plain text | Honors IRM; denial is an MCP resource error carrying `IRM_DENIED` |
| `office://word/<session_id>/structure` | JSON outline (headings, lists, tables) | Lightweight |
| `office://word/<session_id>/paragraph/<index>` | Single paragraph | |
| `office://word/<session_id>/comments` | All comments JSON | |
| `office://word/<session_id>/track_changes` | Tracked changes JSON | |
| `office://word/<session_id>/selection` | Currently selected range text + metadata | |
| `office://excel/<session_id>/workbook` | Workbook metadata JSON | Forwards to `excel.get_workbook_info` and requires that capability in the session. |
| `office://excel/<session_id>/sheets` | Worksheet inventory JSON | Forwards to `excel.list_sheets` and requires that capability in the session. |
| `office://excel/<session_id>/used-range?sheet=<name>` | Used range metadata JSON | Forwards to `excel.get_used_range` with optional `sheet`. |
| `office://excel/<session_id>/range/<address>?sheet=<name>` | Range values, text, formulas, dimensions, and formats JSON | Forwards to `excel.read_range` with `address` and optional `sheet`. |
| `office://powerpoint/<session_id>/presentation` | Presentation metadata JSON | Forwards to `powerpoint.get_presentation_info` and requires that capability in the session. |
| `office://powerpoint/<session_id>/slides` | Slide inventory JSON | Forwards to `powerpoint.list_slides` and requires that capability in the session. |
| `office://powerpoint/<session_id>/slides/text?start=0&end=<slide_count>` | Text JSON for a slide range | Forwards to `powerpoint.read_text` with optional `start` and `end`; the range is half-open `[start, end)`, defaults to the full deck, and returns per-slide text groups. |
| `office://powerpoint/<session_id>/slide/<index>/text` | Slide text JSON | Forwards to `powerpoint.read_text` with `slide_index`. |
| `office://powerpoint/<session_id>/slide/<index>/shapes` | Shape inventory JSON | Forwards to `powerpoint.list_shapes` for one slide. |

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
      "available_tool_count": 25,
      "registered_at": "2026-06-14T01:25:01Z"
    }
  ]
}
```

### 4.2 `office.get_session_info`

Args: `{ "session_id": "..." }`. Returns the descriptor plus the full
`available_tools` array, useful when the client only kept the ID or needs to
plan around host capabilities.

### 4.3 `office.describe_tools`

Args: `{ "tools": ["word.insert_image", "excel.read_range"] }`. Returns
runtime contracts for multiple public Office MCP tools in one call. Each
`input_schema` field is the same JSON Schema object advertised by MCP
`tools/list`; clients can use it without reading daemon or add-in source code.
Each contract also includes a client-friendly top-level `parameters` array,
side-effect classification, curated examples for complex tools, and common
error hints.

Example response shape:

```json
{
  "tools": [
    {
      "name": "word.insert_image",
      "input_schema": { "type": "object" },
      "parameters": [
        {
          "name": "session_id",
          "required": true,
          "schema": { "type": "string" }
        }
      ],
      "examples": [
        {
          "description": "Insert a PNG as a new paragraph after paragraph 2.",
          "arguments": { "session_id": "session-1" }
        }
      ],
      "side_effect": "mutating",
      "app": "word",
      "category": "Media",
      "common_errors": [
        {
          "code": "INVALID_ARGUMENTS",
          "cause": "The arguments do not match the advertised input schema."
        }
      ]
    }
  ]
}
```

Unknown requested tool names return structured per-entry `UNKNOWN_TOOL`
results instead of failing the whole batch:

```json
{
  "tools": [
    {
      "name": "word.future_tool",
      "error": {
        "office_mcp_code": "UNKNOWN_TOOL",
        "message": "Unknown tool word.future_tool."
      }
    }
  ]
}
```

### 4.4 `office.activate_session`

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
- PowerPoint: [04-powerpoint-capabilities.md](04-powerpoint-capabilities.md)
- Outlook: future Outlook capability doc

Tool names are not re-listed here on purpose — the per-app doc is the single
source of truth, and any list in this file would drift on the first PR that
adds or renames a tool.

The daemon implementation MUST keep Word, Excel, and PowerPoint public tool
catalogs on a shared catalog representation or shared helper path. App-specific
catalog data may differ, but common operations such as iterating canonical tool
names, checking membership, generating MCP-safe aliases, and comparing parity
with task pane `AVAILABLE_TOOLS` arrays MUST NOT depend on duplicated per-app
definition structs or duplicated wrapper logic unless an app has real
app-specific fields or behavior.

## 6. Anchor model

Many tools take an `anchor` argument that describes *where* in the document to act.
The public contract is a shared typed `Anchor` schema. Tools that accept an
`anchor` argument MUST expose that argument as a JSON Schema `oneOf` over the
anchor variants they support, and MUST omit unsupported variants from that
tool's advertised schema. This keeps client planning machine-readable: a client
can inspect `tools/list` to learn whether a tool supports a paragraph, text,
selection, heading, bookmark, start-of-document, or end-of-document target.

The complete Word anchor vocabulary is:

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

Per-tool schemas intentionally narrow this vocabulary. For example,
`word.insert_paragraph`, `word.insert_table`, `word.insert_page_break`,
`word.insert_list`, `word.delete_range`, `word.apply_formatting`,
`word.apply_style`, `word.insert_content_control`, and `word.add_comment`
support the full anchor vocabulary. `word.insert_image` also supports the full
anchor vocabulary, but its advertised schema includes a `placement` argument so
callers can distinguish inline insertion from paragraph-relative insertion such
as `new_paragraph_after`.

The daemon validates the advertised anchor kind before forwarding a mutating
call to the add-in. Unsupported anchor kinds fail with `INVALID_ARGUMENTS` and
`partial_effect: none`; they must not reach Office.js and must not mutate the
document.

Mutating tools MAY advertise `validate_only: true` when they can check the
request without applying a document write. The daemon still performs the same
schema, policy, session, and daemon-owned preprocessing checks as a normal
call, then forwards the request to the add-in with `validate_only` preserved.
The add-in MUST run deterministic semantic validation and target resolution,
but MUST NOT queue mutating Office.js calls such as `insertText`, `delete`,
`insertInlinePictureFromBase64`, or `context.document.save`.

For Word v1, `word.insert_image`, `word.delete_range`, `word.replace_text`,
and `word.update_paragraph` MUST support `validate_only`. `word.replace_text`
keeps the existing `dry_run` alias for compatibility; `validate_only: true`
has the same no-mutation behavior and SHOULD be preferred by new clients for a
consistent contract across mutating tools.

A validation-only success returns tool-specific planning metadata and includes
`valid: true`. A validation-only failure uses the same MCP tool error envelope
as normal preflight failures, with a stable `INVALID_ARGUMENT` or
`INVALID_ARGUMENTS` code and `partial_effect: none`. When the implementation
can infer a correction, the structured response SHOULD include a `suggestion`
object, for example a compatible image `placement`.

Paragraph indices are 0-based and refer to the *current* document state at the
moment the tool call is processed. Clients must not cache indices across edits.

`after_text` / `before_text` resolution rules:

- `occurrence: 1` = first match, scanning top-to-bottom.
- Case-insensitive by default; pass `match_case: true` in the surrounding tool args.
- Returns `ANCHOR_NOT_FOUND` if no match.
- If `occurrence > 1` and fewer matches exist, returns `ANCHOR_NOT_FOUND`.

Clients can call `word.resolve_anchor` to inspect the same anchor resolution
path without mutating the document. The response returns safe metadata such as
`object_type`, paragraph index when known, supported and unsupported operation
hints, tool suitability for image insertion, text replacement, deletion, and
formatting, and a bounded text preview unless `include_text_preview: false` is
set. It must not return the full document body.

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

Every tool returned by `tools/list` MUST include rich, machine-readable
contract metadata for that specific tool. The contract is part of the public
client surface, not debug documentation. Clients must be able to plan valid
calls from `tools/list` alone without reading add-in source code, per-host
JavaScript, or first making a secondary `office.describe_tools` call.

Required `tools/list` metadata for every public tool:

- `inputSchema`: a tool-specific JSON Schema object.
- `annotations`: MCP tool annotations derived from the tool side-effect model.
- `_meta["com.office-mcp/side_effects"]`: one of `read`, `mutating`, or
  `destructive`.
- `_meta["com.office-mcp/common_errors"]`: stable, sanitized error hints for
  client planning and recovery.
- `_meta["com.office-mcp/examples"]`: at least one valid example for complex
  tools whose arguments contain nested objects, action-dependent fields, image
  inputs, anchors, table operations, chart options, PivotTable descriptors, or
  slide/shape selectors. Simple tools may expose an empty example array.
- `_meta["com.office-mcp/app"]` and `_meta["com.office-mcp/category"]` for
  Office-host tools, matching the daemon UI permission catalog.

Schema requirements:

- Top-level schemas are JSON objects with explicit `required` fields and
  `additionalProperties: false`, unless the tool section documents a deliberate
  extension bag.
- Each document-affecting tool requires `session_id` at the top level. Cross-app
  management tools only require `session_id` when they target one session.
- Primitive fields declare their JSON type; integer offsets, indices, row and
  column positions declare a lower bound; bounded strings use `enum` or `const`.
- Nested objects such as anchors, image inputs, table ranges, formatting blocks,
  slide/shape selectors, chart options, and PivotTable descriptors are modeled
  with nested schemas and reject unknown keys wherever the operation has a
  closed contract.
- Reused structures may be represented through `$defs` and `$ref` inside a
  tool schema, but each advertised tool still carries a complete schema graph
  that an MCP client can validate locally.

Invalid argument names such as `paragraph_index` for a tool that requires
`index` are rejected by the advertised schema before dispatch. The daemon may
also validate server-side, but client-visible `tools/list` schemas are the
source of truth for contract discovery.

Tool failure results must preserve structured, sanitized host diagnostics when
the add-in can provide them. For Office.js failures, the MCP
`structuredContent.error.debug` object may include the host error code,
reported error location, anchor kind, target object type, placement, and a
short hint. It must not include document text, image bytes/base64, complete raw
arguments, file contents, auth material, or arbitrary host exception dumps.
Clients should treat these fields as optional diagnostic aids and continue to
branch on stable `office_mcp_code` values first.

The daemon also exposes `office.describe_tools` as a read-only batch discovery
helper for clients that want to retrieve multiple detailed tool contracts at
runtime. Its `tools` argument names public Office MCP tools, and each result
entry MUST return the same input schema graph, examples, side-effect
classification, app/category metadata, and common error hints advertised for
that tool in `tools/list`, plus a top-level `parameters` array derived from the
schema. Unknown requested tool names return structured per-entry `UNKNOWN_TOOL`
results rather than being forwarded to an Office session. The helper is for
contract discovery only; it does not require a document session and must not
mutate host state. The previous single-tool `office.describe_tool` helper is
not a public tool.

Transport rate limiting must distinguish discovery/read-only requests from
document-operation requests. `initialize`, `tools/list`, `resources/list`,
`resources/templates/list`, `resources/read office://sessions`,
`office.list_sessions`, `office.get_session_info`, and `office.describe_tools`
are discovery/read-only traffic. These requests use a separate discovery budget
so client introspection cannot consume the same strict budget as mutating Office
operations. Other `resources/read` requests and document-affecting
`tools/call` requests remain protected by the normal operation budget.

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

The daemon lists the current daemon-allowed server-wide catalog. The underlying
host catalog is stable, but `tools/list` is filtered by the daemon global tool
access policy before it is returned to MCP clients. Tools disabled globally are
not listed, which reduces client context and prevents agents from planning with
tools the daemon will reject. Session-specific support is reported by
`available_tools`; invoking a listed tool against an incompatible session yields
a tool execution error, not a protocol error.

The repository MUST keep a permanent Office tool exposure parity gate that
compares every implemented Word, Excel, and PowerPoint tool across the app
task pane `AVAILABLE_TOOLS` arrays, the daemon canonical catalog, the daemon
tool metadata catalog, the implemented capability tables, and MCP `tools/list`.
That gate MUST fail with app- and layer-specific missing-tool details when any
public Office tool appears in one layer but not another. It also MUST prove that
every canonical Office tool has exactly one MCP-safe underscore alias in
`tools/list`, that aliases do not collide, and that representative mutating
aliases for all hosts remain exposed, including `word_update_table`,
`excel_write_range`, and `powerpoint_add_slide`.

Every `tools/call` is still enforced server-side even if the client skipped
refreshing discovery data. The daemon first checks the global tool access policy
and returns `TOOL_NOT_AVAILABLE` with `refresh_tools: true` when the tool is no
longer daemon-allowed. Only after that passes does it check the target session's
`available_tools`; if the tool is absent, it returns
`TOOL_NOT_ENABLED_FOR_DOCUMENT` with `refresh_session_info: true` and the target
`session_id`. Clients SHOULD refresh `tools/list` after the first error and
`office.get_session_info` or `office.list_sessions` after the second error.
Both access-control errors use the same human-readable prefix,
`Tool <name> is disabled.`, then name the rejecting layer and refresh action.

## 10. Tool-level E2E contract

Runtime evidence and smoke scripts are not substitutes for tool-level E2E
coverage. Every advertised MCP tool MUST have one non-evidence E2E case before
it is treated as release-ready for a host.

The host E2E harness owns the common lifecycle:

1. Start the production MCP daemon.
2. Create a fresh empty Office file for the host: document, workbook, or
   presentation.
3. Open the file in the real Office host, activate/connect the Office MCP add-in,
   and wait until `office.list_sessions` reports the expected active session.
4. Discover the host's advertised tool catalog from `tools/list` or the
   session's `available_tools`.
5. Run one table-driven loop over the exact advertised tools.
6. Close the driver-owned Office file and delete it during cleanup.

Cleanup is part of the E2E contract, not a convenience. The live driver MUST
close every driver-owned file it opened, including Office sideload copies such
as `Word add-in ...docx`, `Excel add-in ...xlsx`, or `PowerPoint add-in
...pptx`, and it MUST quit the Office application when no non-driver-owned
documents remain. A failed activation, failed session wait, failed tool call,
or failed verifier MUST still run cleanup in `finally`. The command is not
release-ready if it leaves driver-owned Word, Excel, or PowerPoint windows,
Office processes, or temporary Office files for the developer to close by hand.

The normal E2E gate MUST use one Office application/document lifecycle per host
run: open the host program and driver-owned test file once, connect one add-in
session, run every advertised tool for that session, and then perform cleanup
once. The per-tool loop may only reset/setup deterministic content, call the
tool, and verify the result. It MUST NOT start another daemon, open another
Office program instance, create another test file, activate the add-in again,
or wait for a new session for each tool. Restarting Office or recreating the
file per tool is only a diagnostic fallback because it hides session-lifecycle
bugs and makes Office automation less stable.

When the live Office driver runs, it MUST write a structured JSON report under
`artifacts/office-tool-e2e-<host>.json`. The report records the host, daemon
endpoint, driver-owned document path, session ID, advertised tool list, session
`available_tools`, the ordered executed tool list, per-tool verifier kind and
pass/fail result, lifecycle counters, and cleanup proof. Cleanup proof MUST
include `deleted_paths` listing the concrete driver-owned original file and any
Office sideload copies that were closed and deleted. A count-only cleanup report
is not release-ready because it cannot prove which driver-owned files were closed and deleted. A passing report must show exactly one daemon start, one
`tools/list`, one document creation, one add-in activation attempt, one session
wait, one document cleanup, and one daemon stop for the host run.
Release validation consumes these reports through
`npm run evidence:validate -- --require-office-tool-e2e` with explicit Word,
Excel, and PowerPoint report paths. The validator fails if any report is
missing, does not pass, has more than one lifecycle for the host run, omits a
tool, or records a failed per-tool verifier.

Each tool case supplies only the tool-specific pieces:

- Fixed setup content to add before the tool call.
- MCP tool arguments for operating on that fixed content.
- A verifier for the expected result.

The shared loop performs the repeated work for every case: reset or recreate the
baseline content, call `tools/call`, and verify the outcome. Read tools verify
their direct result. Mutating tools verify by reading the document back through
the most appropriate read tool or resource. Destructive tools verify that the
target content or object is absent after readback. The harness MUST fail when a
new advertised tool has no case, when setup content cannot be created, when the
daemon/add-in session does not connect, or when cleanup cannot prove it only
closed/deleted driver-owned test files.

The E2E command is a separate developer gate from evidence generation. Evidence
recorders may consume its output later, but `npm run evidence:*` passing does
not satisfy this contract.

## 11. Prompts (templated, optional)

The server exposes a small set of MCP prompts that demonstrate idiomatic use:

| Prompt | Args | Body sketch |
|---|---|---|
| `summarize_document` | `session_id: string` | "Read the selected Word session via `office://.../document`, summarize in 200 words, post the summary as a comment via `word.add_comment` on paragraph 0." |
| `polish_section` | `session_id: string, heading: string` | "Find heading, polish prose, present changes for user accept before applying." |
| `extract_action_items` | `session_id: string` | "Read document, extract action items, return as JSON; do not modify the doc." |

Prompts are optional sugar; everything they do is achievable via raw tool calls.
