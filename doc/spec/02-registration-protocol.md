# 02 — Add-in ↔ Server Registration Protocol

Wire-level spec for the local WebSocket channel between Office add-ins and the
`office-mcp` server. JSON-RPC 2.0, one JSON object per WS text frame, UTF-8.

## 1. Conventions

- All method names are lowercase, dot-separated, namespaced
  (`register`, `session.added`, `tool.invoke`, `event.selection_changed`).
- All IDs are stringly-typed UUIDv4.
- All timestamps are RFC 3339 strings.
- All sizes are bytes; durations are milliseconds.
- Unknown fields MUST be ignored by both parties (forward compat).

## 2. Handshake

The production add-in-to-daemon channel uses WSS, including on loopback,
because the add-in itself is loaded from an HTTPS origin. After the WebSocket
upgrade, the first JSON-RPC message is `register`. The register schema is fixed
to runtime metadata, host metadata, add-in protocol metadata, capability
evidence, and optional validated Office SSO user data. The daemon accepts add-in
connections only on loopback and only after exact `Origin` validation; see
[05-security.md §1](05-security.md).

The handshake is deliberately metadata-only. `register` carries only the schema
fields below, and connection acceptance is determined by the local listener, the
WebSocket upgrade, and exact `Origin` validation. TLS is only the
Office-webview-required transport for the local HTTPS/WSS origin.

### 2.1 Address discovery

Office add-ins run in a browser/webview sandbox and cannot read the daemon's
native config file. The v1 add-in therefore uses the compiled default endpoint:

```toml
[addin_channel]
bind = "localhost"
port = 8765
```

- The daemon defaults to `localhost:8765` for the HTTPS/WSS add-in origin.
- The add-in derives its WSS endpoint from the manifest-loaded task pane origin,
  so the default developer manifest connects to `wss://localhost:8765/addin`.
- A user may override the add-in endpoint in its settings UI. The override is
  stored in partitioned browser storage and must match the daemon config.
- Installer-managed deployments may build a manifest/bundle with a different
  default endpoint.
- The installer provisions a per-install local certificate trusted by the
  current user. Its SANs cover the configured local origin. The developer
  manifest uses `localhost`, so the daemon serves both the static add-in bundle
  and WSS from `https://localhost:8765` by default. This certificate is TLS
  transport material required by Office webview HTTPS/WSS rules and is never
  sent in JSON-RPC messages.
- The daemon accepts a WebSocket upgrade only when its `Origin` header exactly
  matches the configured add-in HTTPS origin. For the default endpoint that is
  `https://localhost:8765`.

If either daemon listener is configured with a non-loopback bind address, the
server REFUSES to start in v1. See [05-security.md §2](05-security.md).

### 2.2 First message from add-in: `register`

```json
{
  "jsonrpc": "2.0",
  "id": "11111111-1111-1111-1111-111111111111",
  "method": "register",
  "params": {
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
      "requirement_sets": {
        "WordApi": "1.6",
        "WordApiDesktop": null
      },
      "supported_features": [
        "doc.read",
        "doc.write",
        "doc.tables",
        "doc.comments",
        "doc.tracked_changes"
      ]
    }
  }
}
```

- `instance_id` identifies this document-scoped add-in runtime, not the whole
  Office process. It MUST be stable across WebSocket reconnects for the
  lifetime of the runtime and unique across concurrent runtimes. Generate it
  on add-in initialization and keep it in memory plus `sessionStorage`.
- `requirement_sets` contains the highest versions confirmed at runtime with
  `Office.context.requirements.isSetSupported`. A missing or `null` set is not
  supported. Compile-time Office.js types are not capability evidence.
- `supported_features` is derived from those checks and any successful
  host-specific probes. The daemon never infers a feature from the Office
  build number.
- `user` is optional in v1. It is populated only by deployments that configure
  and validate Office SSO; the base local manifest does not request identity
  assertions and must not infer a UPN from the OS account.
- `register.params` is limited to runtime identifiers, host metadata, add-in
  version/protocol metadata, capability evidence, and the optional validated
  `user` object described above. Fields outside this schema are ignored for
  forward compatibility and must not affect connection acceptance. The add-in
  channel relies on loopback binding plus exact `Origin` validation.
- Extra fields in `register.params` MUST NOT become a second acceptance path.
  The daemon accepts or refuses the connection before any add-in-provided value
  could influence admission.
- Non-loopback bind is rejected at daemon startup. v1 does not define remote
  add-in communication.

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
    "max_pending_per_session": 4,
    "assigned_instance_id": "22222222-2222-2222-2222-222222222222"
  }
}
```

`session_grace_sec` is the maximum reconnect grace period for a closed or
disconnected Office document session. The default is 60 seconds. The daemon
MUST cap the effective value at 300 seconds even if a config file or environment
override asks for a longer retention window. After the grace period expires, the
session is deleted from the daemon registry and must no longer appear in
`office.list_sessions`, `office.get_session_info`, daemon UI document lists, or
tray document counts.

Error (e.g. protocol version skew, malformed register):

```json
{
  "jsonrpc": "2.0",
  "id": "11111111-1111-1111-1111-111111111111",
  "error": {
    "code": -32000,
    "message": "Protocol version mismatch: server supports 1.x, add-in offered 2.0.",
    "data": {
      "office_mcp_code": "PROTOCOL_VERSION_MISMATCH",
      "server_protocol": "1.0"
    }
  }
}
```

After a fatal error the server MUST close the WS with code 4003.

## 3. Session events (add-in → server)

### 3.1 `session.added`

Emitted after the document-scoped add-in runtime registers and has inspected
its current document.

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
      "is_read_only": null,
      "is_protected": null,
      "protection": {
        "kind": "irm",
        "rights": null,
        "rights_source": "unavailable"
      },
      "opened_at": "2026-06-14T01:25:01Z"
    },
    "available_tools": [
      "word.get_text",
      "word.get_outline",
      "word.get_paragraph",
      "word.find_text",
      "word.resolve_anchor",
      "word.get_selection",
      "word.insert_paragraph",
      "word.insert_table",
      "word.insert_image",
      "word.update_image",
      "word.insert_break",
      "word.list_sections",
      "word.insert_list",
      "word.replace_text",
      "word.update_paragraph",
      "word.delete_range",
      "word.apply_formatting",
      "word.apply_style",
      "word.read_table",
      "word.update_table",
      "word.list_content_controls",
      "word.insert_content_control",
      "word.update_content_control",
      "word.delete_content_control",
      "word.add_comment",
      "word.resolve_comment",
      "word.update_tracked_change",
      "word.update_page_setup",
      "word.save"
    ],
    "is_active": null
  }
}
```

- `protection.kind` ∈ `none | password | restricted_editing | irm | signed`.
- `is_dirty` is derived from `Word.Document.saved` and is available in the
  core tier. `url`, `is_read_only`, `is_protected`, and `is_active` are
  nullable because their portable equivalents are not available on every
  supported host. Unknown MUST remain `null`, not a guessed value.
- `available_tools` is authoritative for this session. After the daemon global
  tool access policy allows a call, the daemon returns
  `TOOL_NOT_ENABLED_FOR_DOCUMENT` before dispatch when a tool is absent from the
  target session. The error includes `refresh_session_info: true` so the client
  can refresh `office.get_session_info` or `office.list_sessions` instead of
  repeatedly calling with stale session capability data.
- `available_tools` is the effective tool set after host capability probing and
  user-controlled per-tool permissions. Tools disabled in the task pane settings
  are omitted from this array. The add-in may also include
  `disabled_tools: [{ "tool": "word.insert_paragraph", "reason": "user" }]`
  for UI diagnostics, but clients must plan from `available_tools`.
- A v1 WebSocket connection owns exactly one document session. The
  `session_id` is generated by the add-in, retained in `sessionStorage`, and
  reused when that runtime reconnects. Reopening the file later creates a new
  runtime and a new session ID.
- The daemon binds `session_id` to its registered `instance_id`. Reconnect may
  replace the prior socket only when both IDs match. A different runtime that
  presents an active or grace-period session ID is rejected; it cannot steal
  or merge that session.
- `protection.rights` is optional. When a stable host API exposes effective
  rights, it contains only the current user's effective rights and
  `rights_source` identifies that API. Otherwise it is `null` with
  `rights_source: "unavailable"`.
- When rights are unavailable, the add-in attempts the Office.js operation and
  maps the host's access-denied failure to `IRM_DENIED`; it MUST NOT infer rights
  that the host did not report.

Notifications (no `id` field) per JSON-RPC 2.0.

### 3.2 `session.updated`

Sent when observable metadata changes (e.g. Word reports `saved: true`, so
`is_dirty` becomes `false`).
Same payload shape; only changed fields need be populated, but `session_id` is required.

The add-in MUST send `session.updated` with a fresh `available_tools` array
after the user changes per-tool permissions in task pane settings. The daemon
uses that effective set for capability preflight. If an already-dispatched call
races with a permission change, the add-in returns `TOOL_DISABLED_BY_USER` from
`tool.invoke`.

Daemon-wide tool access policy is evaluated before this session-level preflight.
When the daemon disables a tool globally, `tools/list` omits it and any stale
`tools/call` receives `TOOL_NOT_AVAILABLE` with `refresh_tools: true`. When the
daemon allows the tool globally but the current session does not expose it,
`tools/call` receives `TOOL_NOT_ENABLED_FOR_DOCUMENT` with
`refresh_session_info: true`.

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

- `tool` is the fully-qualified tool name (app-prefixed: `word.*`, `excel.*`).
- `args` schema is the tool-specific JSON Schema, validated by the server before forwarding.
- `timeout_ms` is server-enforced; if the add-in exceeds it, the server replies
  to the MCP client with timeout and tells the add-in to cancel via
  `tool.cancel`. Only one `tool.invoke` is dispatched to a session at a time.
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
      "inserted_paragraph_index": 13
    },
    "elapsed_ms": 147
  }
}
```

Operational failure:

```json
{
  "jsonrpc": "2.0",
  "id": "55555555-5555-5555-5555-555555555555",
  "result": {
    "ok": false,
    "error": {
      "office_mcp_code": "IRM_DENIED",
      "message": "Word denied the requested edit.",
      "tool": "word.insert_paragraph",
      "retriable": false,
      "partial_effect": "none"
    },
    "elapsed_ms": 18
  }
}
```

Error codes are defined in [06-error-model.md](06-error-model.md).

### 4.2 Cancellation: `tool.cancel`

If the MCP client sends an explicit MCP cancellation notification (or the
server deadline expires), the server sends:

```json
{
  "jsonrpc": "2.0",
  "method": "tool.cancel",
  "params": {
    "request_id": "55555555-5555-5555-5555-555555555555",
    "reason": "cancelled_notification"
  }
}
```

Notification, no reply expected. The add-in SHOULD abort the operation before
the next `context.sync()` when possible. Once a sync has started or completed,
rollback is not guaranteed. The original invocation returns `CANCELLED` with
`partial_effect` set to `none`, `possible`, or `unknown`; a daemon deadline
returns `TIMEOUT` with the same field.

Loss of an HTTP/SSE connection is not cancellation under Streamable HTTP. The
daemon continues the call unless it receives the explicit notification or the
deadline expires.

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
| 4002 | Heartbeat timeout | server |
| 4003 | Protocol version mismatch | server |
| 4004 | Server shutting down (graceful) | server |
| 4005 | Add-in replaced (same instance_id reconnected) | server |

## 9. Example full session

```
T+0.000  Server starts, binds wss://localhost:8765/addin (from daemon config)
T+1.200  User launches Word and activates the add-in
T+1.350  Add-in loads its endpoint setting, opens WS
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
T+60.000 User closes Q3-Report; add-in unloads
T+60.030 Add-in → session.removed when unload permits, otherwise WS closes
T+90.000 User quits Word; add-in unload → WS close (1001)
T+90.030 Server marks session stale (none currently; nothing to clean)
```
