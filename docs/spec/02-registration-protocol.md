# 02 — Add-in ↔ Server Registration Protocol

Wire-level spec for the WebSocket channel between Office add-ins and the
`office-mcp` server. JSON-RPC 2.0, one JSON object per WS text frame, UTF-8.

## 1. Conventions

- All method names are lowercase, dot-separated, namespaced
  (`register`, `session.added`, `tool.invoke`, `event.selection_changed`).
- All IDs are stringly-typed UUIDv4.
- All timestamps are RFC 3339 strings.
- All sizes are bytes; durations are milliseconds.
- Unknown fields MUST be ignored by both parties (forward compat).

## 2. Handshake

### 2.1 Discovery file

When the server starts, it writes:

```
%LOCALAPPDATA%\office-mcp\handshake.json     (Windows)
~/Library/Application Support/office-mcp/handshake.json  (macOS)
~/.config/office-mcp/handshake.json          (Linux — for cross-platform testing)
```

with permissions `0600`:

```json
{
  "version": 1,
  "server_pid": 14823,
  "ws_url": "ws://127.0.0.1:8765",
  "bearer_token": "ot_8f3e2c1a9d4b5e6f7a8b9c0d1e2f3a4b",
  "issued_at": "2026-06-14T01:23:45Z",
  "expires_at": "2026-06-15T01:23:45Z"
}
```

- `bearer_token` is a random 32-byte hex with `ot_` prefix.
- The file is rewritten on every server start (new port may be assigned).
- The add-in reads this file on its first attempt to connect.

### 2.2 First message from add-in: `register`

```json
{
  "jsonrpc": "2.0",
  "id": "11111111-1111-1111-1111-111111111111",
  "method": "register",
  "params": {
    "bearer_token": "ot_8f3e2c1a9d4b5e6f7a8b9c0d1e2f3a4b",
    "instance_id": "22222222-2222-2222-2222-222222222222",
    "host": {
      "app": "word",
      "version": "16.0.18025.20096",
      "platform": "windows",
      "build": "Desktop"
    },
    "add_in": {
      "version": "0.1.0",
      "protocol_version": "1.0",
      "supported_features": [
        "doc.read",
        "doc.write",
        "doc.tables",
        "doc.comments",
        "doc.selection_events",
        "doc.irm_metadata"
      ]
    },
    "user": {
      "display_name": "Riff",
      "upn": "riff@contoso.com",
      "tenant_id": "33333333-3333-3333-3333-333333333333"
    }
  }
}
```

- `instance_id` MUST be stable for the lifetime of the Office instance and
  unique across instances. Generated on first add-in load, stored in
  `Office.context.document.settings` (persists for the document) or in
  `localStorage` (persists for the user on this Office install).
- `user.upn` and `user.tenant_id` come from `Office.context.auth.getAccessToken`
  or, if unavailable, from `Office.context.mailbox.userProfile` (Outlook only)
  or are omitted (Word/Excel without identity).

### 2.3 Server reply: `register.result` or `register.error`

Success:

```json
{
  "jsonrpc": "2.0",
  "id": "11111111-1111-1111-1111-111111111111",
  "result": {
    "server_version": "0.1.0",
    "protocol_version": "1.0",
    "session_grace_sec": 60,
    "heartbeat_interval_sec": 30,
    "max_inflight_per_session": 4,
    "assigned_instance_id": "22222222-2222-2222-2222-222222222222"
  }
}
```

Error (e.g. token mismatch, protocol version skew):

```json
{
  "jsonrpc": "2.0",
  "id": "11111111-1111-1111-1111-111111111111",
  "error": {
    "code": -32401,
    "message": "Invalid bearer token. Restart office-mcp server to refresh handshake.",
    "data": { "reason": "token_mismatch" }
  }
}
```

After a fatal error the server MUST close the WS with code 4001.

## 3. Session events (add-in → server)

### 3.1 `session.added`

Emitted when a document opens in the registered instance.

```json
{
  "jsonrpc": "2.0",
  "method": "session.added",
  "params": {
    "session_id": "44444444-4444-4444-4444-444444444444",
    "instance_id": "22222222-2222-2222-2222-222222222222",
    "document": {
      "title": "Q3 Report",
      "url": "C:\\Users\\riff\\Documents\\Q3-Report.docx",
      "filename": "Q3-Report.docx",
      "is_dirty": false,
      "is_read_only": false,
      "is_protected": true,
      "protection": {
        "kind": "irm",
        "rights": ["view", "edit", "extract"],
        "policy_id": "..."
      },
      "word_count": 2417,
      "page_count": 8,
      "opened_at": "2026-06-14T01:25:01Z"
    }
  }
}
```

- `protection.kind` ∈ `none | password | restricted_editing | irm | signed`.
- `protection.rights` for IRM mirrors the user's effective rights, NOT the policy's
  full set. If `extract` is absent, `tool.invoke` calls that read document text
  MUST be rejected by the add-in with `error.code = -32403`.

Notifications (no `id` field) per JSON-RPC 2.0.

### 3.2 `session.updated`

Sent when metadata changes (e.g. user pressed Save → `is_dirty: false`).
Same payload shape; only changed fields need be populated, but `session_id` is required.

### 3.3 `session.removed`

```json
{
  "jsonrpc": "2.0",
  "method": "session.removed",
  "params": {
    "session_id": "44444444-4444-4444-4444-444444444444",
    "reason": "closed"
  }
}
```

`reason` ∈ `closed | crashed | replaced | unknown`.

## 4. Tool invocation (server → add-in)

```json
{
  "jsonrpc": "2.0",
  "id": "55555555-5555-5555-5555-555555555555",
  "method": "tool.invoke",
  "params": {
    "session_id": "44444444-4444-4444-4444-444444444444",
    "tool": "word.insert_paragraph",
    "args": {
      "anchor": { "kind": "after_paragraph_index", "index": 12 },
      "text": "This paragraph was added by the agent.",
      "style": "Normal"
    },
    "timeout_ms": 30000,
    "client_meta": {
      "mcp_client": "claude-desktop/0.7.4",
      "mcp_request_id": "msg-789",
      "user_intent": "add summary paragraph after intro"
    }
  }
}
```

- `tool` is the fully-qualified tool name (host-prefixed: `word.*`, `excel.*`).
- `args` schema is the tool-specific JSON Schema, validated by the server before forwarding.
- `timeout_ms` is server-enforced; if the add-in exceeds it, the server replies
  to the MCP client with timeout and tells the add-in to cancel via `tool.cancel`.
- `client_meta.user_intent` is the natural-language string from the client (when
  available) — purely for diagnostic logging in the add-in. It MUST NOT change behavior.

### 4.1 Add-in reply

Success:

```json
{
  "jsonrpc": "2.0",
  "id": "55555555-5555-5555-5555-555555555555",
  "result": {
    "ok": true,
    "data": {
      "inserted_paragraph_index": 13,
      "new_word_count": 2423
    },
    "elapsed_ms": 147
  }
}
```

Error:

```json
{
  "jsonrpc": "2.0",
  "id": "55555555-5555-5555-5555-555555555555",
  "error": {
    "code": -32403,
    "message": "Document IRM policy denies edit; only 'view' and 'extract' are granted.",
    "data": {
      "tool": "word.insert_paragraph",
      "denied_rights": ["edit"]
    }
  }
}
```

Error codes are defined in [06-error-model.md](06-error-model.md).

### 4.2 Cancellation: `tool.cancel`

If the MCP client cancels (or the server times out), the server sends:

```json
{
  "jsonrpc": "2.0",
  "method": "tool.cancel",
  "params": {
    "request_id": "55555555-5555-5555-5555-555555555555",
    "reason": "client_disconnected"
  }
}
```

Notification, no reply expected. The add-in SHOULD abort the operation if
possible (e.g. before next `context.sync()`). If already committed, the add-in
SHOULD attempt an undo and reply to the original `tool.invoke` with
`error.code = -32604` ("Cancelled; undo applied" or "Cancelled; undo failed").

## 5. Streaming partial results (deferred)

Not in v1. Long-running tools complete fully or fail. v2 may add `tool.progress`
notifications for partial output (e.g. table generation row-by-row).

## 6. Heartbeat: `ping` / `pong`

Server → add-in every `heartbeat_interval_sec`:

```json
{ "jsonrpc": "2.0", "id": "...", "method": "ping", "params": { "ts": "2026-06-14T01:30:00Z" } }
```

Add-in MUST reply within 10s:

```json
{ "jsonrpc": "2.0", "id": "...", "result": { "ts": "2026-06-14T01:30:00.143Z" } }
```

Two consecutive missed pongs → server marks session stale and closes WS with
code 4002.

## 7. Add-in → server events (opt-in)

The server tells the add-in which events to emit by passing
`enabled_events: [...]` in the `register.result` (advanced, deferred to v1.1).
For v1, the add-in emits only `session.*` events; selection-change / cursor-move
events are NOT emitted by default.

## 8. Close codes

| WS close code | Meaning | Sender |
|---|---|---|
| 1000 | Normal | both |
| 1001 | Office shutting down | add-in |
| 4001 | Auth failure | server |
| 4002 | Heartbeat timeout | server |
| 4003 | Protocol version mismatch | server |
| 4004 | Server shutting down (graceful) | server |
| 4005 | Add-in replaced (same instance_id reconnected) | server |

## 9. Example full session

```
T+0.000  Server starts, writes handshake.json
T+1.200  User launches Word; pinned add-in loads
T+1.350  Add-in reads handshake.json, opens WS
T+1.360  Add-in → register
T+1.370  Server → register.result
T+1.400  Add-in → session.added (Q3-Report.docx)
T+5.000  MCP client → list_resources
T+5.020  Server returns [Q3-Report.docx session]
T+12.000 MCP client → tools/call word.insert_paragraph
T+12.020 Server → tool.invoke (forwards to add-in)
T+12.180 Add-in performs Office.js batch, context.sync
T+12.200 Add-in → result
T+12.210 Server → MCP client returns tool result
T+30.000 Server → ping
T+30.040 Add-in → pong
T+60.000 User closes Q3-Report; add-in detects DocumentClose
T+60.030 Add-in → session.removed
T+90.000 User quits Word; add-in unload → WS close (1001)
T+90.030 Server marks session stale (none currently; nothing to clean)
```
