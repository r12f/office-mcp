# 06 - Error Model

office-mcp has two error layers. Numeric JSON-RPC codes describe protocol
failures. Stable symbolic `office_mcp_code` values describe operational
failures that an agent can act on.

## 1. Symbolic code table

| Code | Meaning |
|---|---|
| `GENERIC_FAILURE` | Catch-all execution failure. Use sparingly. |
| `NO_SESSIONS` | No Office document sessions are connected. |
| `SESSION_LOST` | The session disappeared during a call. |
| `SESSION_NOT_FOUND` | The `session_id` was never known or has expired. |
| `SESSION_STALE` | Session is in its reconnect grace period. |
| `HOST_BUSY` | Office is blocked by a modal dialog or equivalent state. |
| `HOST_ERROR` | Host API rejected the operation and no more specific symbolic code is known. |
| `MAX_PENDING_EXCEEDED` | The per-session FIFO queue is full. |
| `FORBIDDEN_ORIGIN` | A browser-facing request or add-in WSS upgrade came from an untrusted `Origin`. |
| `IRM_DENIED` | The document policy denied the requested operation. |
| `DOCUMENT_READ_ONLY` | The document cannot be edited. |
| `PROTECTION_BLOCKS` | Word protection blocked the operation. |
| `ANCHOR_NOT_FOUND` | The requested anchor could not be resolved. |
| `NO_MATCHES` | Search or replacement found no required match. |
| `INDEX_OUT_OF_RANGE` | A paragraph, table, row, or column index is invalid. |
| `STALE_INDEX` | The document changed after an index precondition was read. |
| `INVALID_ARGUMENT` | Tool arguments violate the tool schema or a cross-field rule. |
| `UNSUPPORTED_FORMAT` | The selected host cannot produce the requested format. |
| `PATH_REFUSED` | A daemon-side file destination violates path policy. |
| `IMAGE_FETCH_FAILED` | The daemon could not fetch or validate an image URL. |
| `HOST_CAPABILITY_UNAVAILABLE` | The session lacks a required API set or verified host capability. |
| `TOOL_NOT_AVAILABLE` | The daemon global tool access policy disabled this tool, so it is not currently discoverable or callable. |
| `TOOL_NOT_ENABLED_FOR_DOCUMENT` | The target document session has disabled this tool by omitting it from `available_tools`. |
| `TOOL_DISABLED_BY_USER` | The user disabled this tool for the current add-in session. |
| `CANCELLED` | The client or daemon cancelled the operation. |
| `TIMEOUT` | The operation exceeded its deadline. |
| `MAX_RESPONSE_SIZE` | The result exceeded `MAX_RESPONSE_BYTES`. |
| `PROTOCOL_VERSION_MISMATCH` | Add-in protocol major versions differ. |
| `HEARTBEAT_MISSED` | The add-in missed the heartbeat threshold. |
| `INTERNAL_BUG` | An invariant failed. |

## 2. MCP mapping

Unknown MCP methods/tools, malformed MCP requests, and internal protocol
failures use standard JSON-RPC errors such as `-32601`, `-32602`, and
`-32603`.

Expected tool execution failures are MCP tool results:

```json
{
  "jsonrpc": "2.0",
  "id": 4,
  "result": {
    "content": [
      {
        "type": "text",
        "text": "The selected Word session does not support tracked changes."
      }
    ],
    "structuredContent": {
      "ok": false,
      "error": {
        "office_mcp_code": "HOST_CAPABILITY_UNAVAILABLE",
        "session_id": "44444444-...",
        "tool": "word.accept_change",
        "retriable": false,
        "required_requirement_sets": { "WordApi": "1.6" }
      }
    },
    "isError": true
  }
}
```

For backwards compatibility with clients that ignore `structuredContent`, the
text content must summarize the same failure. If an `outputSchema` is declared,
the error result must conform to it.

Tool availability failures must include client refresh guidance so agents do
not continue planning from stale capability data. When the daemon global tool
access policy rejects a call with `TOOL_NOT_AVAILABLE`, the structured error
must include `refresh_tools: true` and the text message must tell the client to
refresh `tools/list`. When the session preflight rejects a call with
`TOOL_NOT_ENABLED_FOR_DOCUMENT`, the structured error must include
`refresh_session_info: true`, the affected `session_id`, and a message telling
the client to refresh `office.get_session_info` or `office.list_sessions` before
retrying. If both daemon policy and session state may have changed, include both
refresh flags. These hints are advisory for client UX; the daemon still enforces
both checks on every `tools/call`.

The human-facing message prefix for both tool availability failures must use
the same wording: `Tool <name> is disabled.` Do not use `not enabled`,
`unsupported`, or `not available` as the main sentence for these access-control
failures. Add the layer and refresh action after that prefix, for example:
`Tool word.get_text is disabled by daemon access policy. Refresh tools/list
before retrying.` or `Tool word.get_text is disabled for this document session.
Refresh office.get_session_info or office.list_sessions before retrying.`

`resources/read` does not return a tool result. An operational resource failure
uses JSON-RPC `-32000` with `error.data.office_mcp_code` and the same structured
fields. Browser `Origin` failures and TLS/front-door failures use their normal
HTTP status codes before MCP dispatch.

HTTP rate-limit failures use status `429` before MCP dispatch and must include
`Retry-After` so clients can back off without guessing. Where the request body
contains JSON-RPC, the response should also preserve JSON-RPC shape with
`error.code = -32000` and `error.data.office_mcp_code = "RATE_LIMITED"`.
Discovery/read-only traffic uses a separate transport rate-limit budget from
document-operation traffic so recovery calls such as `office://sessions`,
`resources/list`, and `resources/templates/list` remain available during normal
client introspection bursts.

Expected tool execution failures MAY include a `debug` object with sanitized,
machine-readable context. The object is for diagnosis and recovery, not for
free-form logs. Allowed fields include Office.js error metadata such as
`office_error_code`, `office_error_location`, `error_location`, selected tool
argument categories such as `anchor_kind`, `target_object_type`, `placement`,
and a short `hint`. The add-in and daemon MUST NOT include full document text,
image base64, file contents, auth material, or unsanitized raw argument objects
in `debug`. When Office.js reports an `InvalidArgument` style failure, the
add-in MUST map it to `INVALID_ARGUMENT` when the rejection is actionable by
changing tool arguments; otherwise it should use `HOST_ERROR` with the same
safe Office.js details.

## 3. Add-in protocol mapping

Malformed or unknown add-in JSON-RPC messages use standard JSON-RPC errors.
Registration rejection uses JSON-RPC `-32000` with a symbolic code in
`error.data.office_mcp_code`.

An accepted `tool.invoke` always receives a JSON-RPC result. Operational
failure is represented as:

```json
{
  "jsonrpc": "2.0",
  "id": "55555555-...",
  "result": {
    "ok": false,
    "error": {
      "office_mcp_code": "IRM_DENIED",
      "message": "Word denied the requested edit.",
      "session_id": "44444444-...",
      "tool": "word.replace_text",
      "retriable": false,
      "partial_effect": "none"
    },
    "elapsed_ms": 18
  }
}
```

`retriable` is required. When retrying is useful, include
`retry_after_ms`. Mutation failures also include `partial_effect` as `none`,
`possible`, or `unknown`.

Add-ins SHOULD preserve safe Office.js diagnostics on operational failures by
including `debug` under `error`. The daemon MUST forward this sanitized
`debug` object into the MCP tool result unchanged except for enforcing the same
redaction boundary. If the add-in provides `partial_effect`, the daemon MUST
preserve it; otherwise daemon-mapped add-in failures default to `unknown`.

Example Office.js rejection payload:

```json
{
  "jsonrpc": "2.0",
  "id": "55555555-...",
  "result": {
    "ok": false,
    "error": {
      "office_mcp_code": "INVALID_ARGUMENT",
      "message": "Word.js InvalidArgument while running word.insert_image.",
      "session_id": "44444444-...",
      "tool": "word.insert_image",
      "retriable": false,
      "partial_effect": "none",
      "debug": {
        "office_error_code": "InvalidArgument",
        "error_location": "Range.insertInlinePictureFromBase64",
        "anchor_kind": "after_paragraph_index",
        "target_object_type": "Paragraph",
        "placement": "inline",
        "hint": "Use a paragraph-relative placement such as new_paragraph_after."
      }
    },
    "elapsed_ms": 18
  }
}
```

Schema validation, daemon argument normalization, and add-in semantic preflight
for mutating tools all happen before Office.js writes are queued. Failures from
those phases MUST return `INVALID_ARGUMENT` with `partial_effect: "none"` so
clients can safely correct the request and retry.

Validation-only tool calls follow the same error model. If `validate_only: true`
is set, the daemon and add-in MUST return preflight failures before any write is
queued, and successful responses MUST include `valid: true` and
`partial_effect: "none"`. Validation-only responses may include safe planning
metadata such as resolved target type, paragraph index, match count, or a small
`suggestion` object; they must not include document body text beyond existing
bounded preview fields.

## 4. Partial-success semantics

The default is to preflight every deterministic target and argument constraint
before mutation. Office.js does not provide a general transaction, so a host
failure during `context.sync()` can still leave `partial_effect: "unknown"`.

Tools that explicitly support `partial_ok: true` may return `ok: true` with a
`skipped` array when at least one target succeeds. If none succeeds, return the
most actionable symbolic failure.

## 5. Failure modes

### 5.1 Office is killed mid-operation

- The WebSocket drops while `tool.invoke` is in flight.
- The daemon returns an MCP tool error with `SESSION_LOST`.
- `partial_effect` is `unknown`.
- The session remains stale through the reconnect grace period, then expires.

The client must call `office.list_sessions`; a reopened document has a new
session ID.

### 5.2 User closes the document mid-operation

Word does not expose a portable document-close event to task-pane add-ins.
The daemon detects WebSocket loss or heartbeat failure. A best-effort page
unload notification may accelerate cleanup but is not authoritative. The
in-flight call returns `SESSION_LOST`.

### 5.3 Office shows a modal dialog

Office.js work may stall until the user dismisses the dialog. The daemon
deadline expires and returns `HOST_BUSY` when host-blocking evidence is known,
otherwise `TIMEOUT`. Both include `user_action_required` when appropriate.

### 5.4 Two clients edit the same paragraph

The add-in serializes calls. The second call sees the first call's effects.
Tools with an explicit state precondition may return `STALE_INDEX`.

## 6. Retry guidance

| Code | Retry guidance |
|---|---|
| `NO_SESSIONS` | Wait for the user to activate the add-in. |
| `SESSION_LOST` | Do not retry the old ID; list sessions. |
| `SESSION_STALE` | Retry during the grace period with 2s, 5s, 10s backoff. |
| `HOST_BUSY` | Retry after `retry_after_ms` and user action. |
| `MAX_PENDING_EXCEEDED` | Back off briefly. |
| `IRM_DENIED` | Do not retry unchanged. |
| `ANCHOR_NOT_FOUND` | Re-read and alter arguments. |
| `STALE_INDEX` | Re-read state; retry at most once automatically. |
| `TOOL_NOT_AVAILABLE` | Refresh `tools/list`; do not retry unchanged unless the daemon policy changed. |
| `TOOL_NOT_ENABLED_FOR_DOCUMENT` | Refresh `office.get_session_info` or `office.list_sessions`; do not retry unchanged unless the session's `available_tools` changed. |
| `TOOL_DISABLED_BY_USER` | Do not retry unchanged; ask the user to re-enable the tool in the add-in task pane settings. |
| `TIMEOUT` | Retry only when the operation is known idempotent or after re-reading state. |

Humans should see the symbolic code, session ID, and actionable message.
Effective IRM rights are included only when Word actually reports them.
