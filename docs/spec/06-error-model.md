# 06 — Error Model

JSON-RPC 2.0 errors. Codes < -32000 are reserved by JSON-RPC; we use the
application range `-32000` to `-32999`.

## 1. Code table

| Code | Name | Meaning |
|---|---|---|
| -32000 | `GENERIC_FAILURE` | Catch-all server-side failure. Use sparingly. |
| -32001 | `NO_SESSIONS` | No Office instances connected. |
| -32002 | `SESSION_LOST` | The session was alive but is now gone (Office crashed, document closed mid-call). |
| -32003 | `SESSION_NOT_FOUND` | The `session_id` was never known to the server. |
| -32004 | `SESSION_STALE` | Session is in grace period, awaiting reconnect; retry in a few seconds. |
| -32005 | `HOST_BUSY` | Office is in a modal dialog or other state that blocks add-in calls. |
| -32006 | `MAX_INFLIGHT_EXCEEDED` | Too many concurrent calls for one session; client should back off. |
| -32401 | `AUTH_FAILED` | (Add-in registration, only when `addin.shared_secret` is configured.) Shared secret missing or mismatched. |
| -32402 | `AUTH_KEY_MISSING` | (HTTP transport with `--api-key`.) Client did not present a key. |
| -32403 | `IRM_DENIED` | Requested operation requires a right not granted by the IRM policy. |
| -32404 | `DOCUMENT_READ_ONLY` | Doc is read-only (file attribute, share lock, or "Mark as Final"). |
| -32405 | `PROTECTION_BLOCKS` | Restricted Editing / form protection blocks the change. |
| -32421 | `ANCHOR_NOT_FOUND` | `anchor.kind = after_text` and the text isn't there. |
| -32422 | `NO_MATCHES` | `find_text` / `replace_text` found zero matches but caller required at least one. |
| -32423 | `INDEX_OUT_OF_RANGE` | Paragraph / table index does not exist in current document state. |
| -32424 | `STALE_INDEX` | Index was valid when call started but the document changed concurrently. |
| -32501 | `INVALID_ARGUMENT` | Tool arguments failed JSON Schema validation. |
| -32502 | `UNSUPPORTED_FORMAT` | E.g. `save_as` requested a format the add-in cannot produce on this Office version. |
| -32503 | `PATH_REFUSED` | `save_as` path was rejected by add-in policy (traversal, sensitive dir). |
| -32504 | `IMAGE_FETCH_FAILED` | `insert_image` with URL: server could not fetch. |
| -32601 | `METHOD_NOT_FOUND` | Tool name unknown OR explicitly disabled (e.g. macro execution). |
| -32604 | `CANCELLED` | Operation was cancelled (client disconnect, timeout). `data.undo_applied: boolean` indicates whether changes were rolled back. |
| -32605 | `TIMEOUT` | Server-side timeout (`tool.timeout_ms`) exceeded. |
| -32606 | `MAX_RESPONSE_SIZE` | Read result exceeds `MAX_RESPONSE_BYTES`. `data.max_response_bytes` echoes the limit; client should retry with `offset` / `limit`. |
| -32701 | `PROTOCOL_VERSION_MISMATCH` | Add-in's `protocol_version` major differs from server's. |
| -32702 | `HEARTBEAT_MISSED` | Add-in stopped responding to pings. |
| -32801 | `CLIENT_DISCONNECTED` | Inflight call's MCP client went away; equivalent to cancellation. |
| -32999 | `INTERNAL_BUG` | Unrecoverable internal error. Should never appear; if it does, file an issue. |

## 2. Error envelope

```json
{
  "jsonrpc": "2.0",
  "id": "...",
  "error": {
    "code": -32403,
    "message": "Document IRM policy denies edit; only 'view' and 'extract' are granted.",
    "data": {
      "tool": "word.replace_text",
      "denied_rights": ["edit"],
      "granted_rights": ["view", "extract"],
      "session_id": "...",
      "retriable": false
    }
  }
}
```

- `message` is human-readable English. Localization is out of scope for v1.
- `data.retriable` is the **single most important field for agents**:
  - `true` → client may retry after the suggested delay (`data.retry_after_ms`).
  - `false` → no retry will succeed; alter the request or give up.

## 3. Partial-success semantics

Some tools can do partial work — e.g. `word.replace_text` with multiple matches
where one match is in a read-only section. Policy:

- **Default**: atomic. If any match would fail, abort, return error, no edits land.
- **Opt-in best-effort**: pass `tool args { "partial_ok": true }` to allow
  successful matches to land. Response carries:
  ```json
  {
    "ok": true,
    "data": {
      "replaced_count": 3,
      "skipped": [
        { "paragraph_index": 17, "reason": "read_only_section" }
      ]
    }
  }
  ```
  with overall `ok: true` if at least one match succeeded; otherwise `ok: false`
  and a `-32000` family error.

## 4. Failure modes by scenario

### 4.1 Office is killed mid-operation

- The WS connection drops while a `tool.invoke` is in flight.
- Server replies to the MCP client with `-32002 SESSION_LOST`.
- `data.partial_effect: "unknown"` because we cannot ask Office anymore.
- Session enters `stale` state. After grace period, removed.

Recovery: agent should call `office.list_sessions` to discover the user's
current state. If the user reopened the document, a *new* session ID will
be assigned; the agent must NOT assume the old session ID is now valid.

### 4.2 User closes the document mid-operation

- The add-in's `DocumentClose` handler fires.
- Add-in sends `session.removed` with `reason: "closed"`.
- The in-flight `tool.invoke` reply is `-32002`.

### 4.3 Office shows a modal dialog (e.g. unsaved changes prompt on Save)

- Office.js calls block until the modal is dismissed.
- The add-in detects this via call duration and replies with `-32005 HOST_BUSY`.
- `data.retry_after_ms: 5000` and `data.user_action_required: true`.

### 4.4 Document is in Track Changes mode and edit would conflict

- The add-in adds the change as a tracked revision instead of committing.
- Response is `ok: true` with `data.tracked_change_ids: [...]`.
- No error is raised — Track Changes is the user's chosen workflow.

### 4.5 Two clients edit the same paragraph in parallel

- The add-in serializes them; second call sees the first's effect.
- If the second call had a `STALE_INDEX` precondition (rare; only some tools
  validate), it returns `-32424` with `data.actual_state` describing what
  the add-in observed.

## 5. Retry guidance

| Code | Retry? | When | Backoff |
|---|---|---|---|
| `-32001 NO_SESSIONS` | No | (Wait for user to open Office) | — |
| `-32002 SESSION_LOST` | No (session is gone) | Call `list_sessions` first | — |
| `-32004 SESSION_STALE` | Yes | Up to grace period | 2s, 5s, 10s |
| `-32005 HOST_BUSY` | Yes | After user dismisses modal | `retry_after_ms` |
| `-32006 MAX_INFLIGHT_EXCEEDED` | Yes | After server signals capacity | 100ms, 500ms, 1s |
| `-32403 IRM_DENIED` | No | — | — |
| `-32421 ANCHOR_NOT_FOUND` | No | Alter args | — |
| `-32424 STALE_INDEX` | Yes | After re-reading current state | Immediate, max 1 retry |
| `-32605 TIMEOUT` | Yes | Increase `timeout_ms` | Backoff per request |

## 6. Surfacing errors to humans

The MCP protocol gives clients little guidance on how to display errors.
This project recommends:

- Distinguish **user-actionable** (`-3240*`, `-3242*`, `-3250*`) from
  **infrastructure** (`-32001`, `-3270*`).
- Always show `data.session_id` so the user can find the right Office window.
- For IRM denials, surface `data.granted_rights` so the user knows what IS
  possible.
