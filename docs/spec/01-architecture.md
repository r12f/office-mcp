# 01 — Architecture

## 1. Component overview

```
┌────────────────────────────────────────────────────────────────────┐
│                            User's machine                          │
│                                                                    │
│  ┌──────────────┐                                                  │
│  │  MCP client  │  ── stdio MCP ──┐                                │
│  │  (Claude /   │                 │                                │
│  │   Cursor /   │                 │                                │
│  │   agent)     │                 ▼                                │
│  └──────────────┘     ┌──────────────────────────────┐             │
│                       │                              │             │
│  ┌──────────────┐     │   office-mcp server          │             │
│  │  MCP client  │ HTTP│   (long-lived process)       │             │
│  │  (remote)    │ ───▶│                              │             │
│  └──────────────┘     │  ┌────────────────────────┐  │             │
│                       │  │ MCP frontend           │  │             │
│                       │  │  - stdio transport     │  │             │
│                       │  │  - Streamable HTTP     │  │             │
│                       │  ├────────────────────────┤  │             │
│                       │  │ Tool router            │  │             │
│                       │  ├────────────────────────┤  │             │
│                       │  │ Session registry       │  │             │
│                       │  ├────────────────────────┤  │             │
│                       │  │ Add-in WS backend      │  │             │
│                       │  │  (ws://127.0.0.1:8765) │  │             │
│                       │  └────────────────────────┘  │             │
│                       └────────────▲─────────────────┘             │
│                                    │                               │
│                  ┌─────────────────┼─────────────────┐             │
│                  │                 │                 │             │
│      WebSocket + JSON-RPC 2.0  ────┴────       ──────┴────         │
│                  │                                                 │
│  ┌───────────────▼────────────┐  ┌────────────────────────────┐    │
│  │  Word.exe (instance A)     │  │  Word.exe (instance B)     │    │
│  │  ┌──────────────────────┐  │  │  ┌──────────────────────┐  │    │
│  │  │ office-mcp add-in    │  │  │  │ office-mcp add-in    │  │    │
│  │  │ (Office.js task-pane)│  │  │  │ (Office.js task-pane)│  │    │
│  │  └──────────────────────┘  │  │  └──────────────────────┘  │    │
│  │  Document: report.docx     │  │  Document: contract.docx   │    │
│  │  Document: notes.docx      │  │  (IRM-protected)           │    │
│  └────────────────────────────┘  └────────────────────────────┘    │
└────────────────────────────────────────────────────────────────────┘
```

## 2. Processes and their roles

### 2.1 The MCP server

- **Lifetime**: starts on first need (auto-launched by MCP client over stdio,
  or manually as a service), continues until killed or idle for `IDLE_SHUTDOWN_SEC`.
- **Listens**: WebSocket on `127.0.0.1:<port>` for add-ins (loopback only).
- **Speaks**: MCP (stdio or Streamable HTTP) for clients.
- **State held**:
  - Registry of connected add-in sessions (instance ID, host app, document list,
    capability bitmap, last heartbeat).
  - Pending request map (MCP call ID → add-in session + add-in request ID).
  - (Only when non-loopback bind is configured) shared secret loaded from the
    shared config (see [05-security.md §2](05-security.md)).
- **State NOT held**:
  - Document content. Always lives in the add-in / Office.
  - Undo history. Word owns it.
  - User credentials. Server never sees them; add-in inherits from Office.

### 2.2 The Office add-in

- **Lifetime**: bound to Office instance. Loads when Office starts (if pinned)
  or when the user opens the task pane.
- **Connects to**: `ws://<bind>:<port>` read from the shared config
  ([07-deployment.md §5](07-deployment.md)).
- **State held**:
  - One persistent WS connection back to the server.
  - Per-document Office.js context handles.
  - Pending operation queue (Office.js requires batched `context.sync()`).
- **Responsibilities**:
  - On load: announce instance + enumerate open documents.
  - On document open/close: send `session.added` / `session.removed` events.
  - On server request: execute Office.js calls, return result.
  - On user save / cursor move / selection change: optionally emit events
    (rate-limited, opt-in via client capability).

### 2.3 The MCP client

Just a normal MCP client. Sees a single MCP server. Does not need to know how
many Office instances are open — that is exposed through MCP resources and the
`session_id` parameter on tool calls.

## 3. Transports

### 3.1 Client ↔ Server: MCP

| Transport | When to use | Notes |
|---|---|---|
| **stdio** | Default for Claude Desktop, Cursor, local agents | Server is spawned as child of client. One client per server process. |
| **Streamable HTTP** | Multiple clients sharing one server; remote agents | Server runs as a service. Binds `127.0.0.1` by default; cloud bind requires explicit flag. |

Per the MCP 2026 stateless protocol revision, both transports support multiple concurrent
sessions. SSE-only transport is **not** supported (deprecated upstream).

### 3.2 Server ↔ Add-in: WebSocket + JSON-RPC 2.0

- Bidirectional: server can call add-in; add-in can call server (for events).
- JSON-RPC 2.0 framing (one JSON object per WS message, no batching in v1).
- Heartbeat: server sends `ping` every 30s; add-in must respond within 10s or
  the session is marked stale.
- Reconnect: add-in must back off with jitter (1s, 2s, 5s, 10s, 30s, then 30s cap).

Why WebSocket and not Named Pipes / Unix Socket?

- **Cross-platform** (Office on Web/Mac/Windows).
- **Office.js add-ins are web apps** — they have `WebSocket` natively, no extra
  permissions. Named pipes require a desktop-only bridge.
- **Loopback-only by default** — security-equivalent to a pipe for v1.

### 3.3 Why not stdio between server and add-in?

Add-ins are launched by Office, not by the server. Their lifecycle is controlled
by Office (and the user's task-pane interactions), so the only viable IPC is one
where the add-in is the connecting party.

## 4. Lifecycle scenarios

### 4.1 Cold start (no Office running)

1. MCP client launches the server (or HTTPs in).
2. Server starts, binds the WS endpoint declared by `[addin_channel]` in the
   shared config (default `127.0.0.1:8765`). Refuses to start if the bind is
   non-loopback and no `shared_secret` is set.
3. Client calls `list_sessions` → server returns `[]` (no add-ins connected).
4. Any tool call against a host returns
   `error: { code: -32001, message: "No Office instances connected" }`.

### 4.2 User opens Word after server is up

1. Word loads. The pinned `office-mcp` task-pane add-in initializes.
2. Add-in reads `[addin_channel]` from the shared config to learn the WS URL.
3. Add-in dials `ws://<bind>:<port>`.
4. Add-in sends `register` (see [02-registration-protocol.md](02-registration-protocol.md))
   with its instance ID, host app, and capability bitmap. No auth field is
   sent unless a non-loopback `shared_secret` is configured.
5. Server accepts (loopback default) or validates the secret, replies
   `register.result`, assigns session ID.
6. Add-in enumerates open documents, sends `session.added` for each.
7. Server now exposes those documents via MCP resources & tool addressing.

### 4.3 User opens a new document in an already-registered Word instance

1. Add-in's `Word.run` document-change handler fires.
2. Add-in sends `session.added` with the new document's metadata.
3. Server updates registry; new document is addressable on next MCP client `list_resources`.

### 4.4 Office crashes mid-call

1. WS connection drops.
2. Server marks all sessions for that instance as `state: stale`.
3. In-flight MCP requests for those sessions reject with
   `error: { code: -32002, message: "Document session lost" }`.
4. Server keeps the session in `stale` state for `STALE_GRACE_SEC` (default 60s)
   to allow add-in reconnect (Office auto-recovery).
5. After grace, session is removed; clients see it disappear from `list_resources`.

### 4.5 MCP client disconnects

- stdio: server exits after `IDLE_SHUTDOWN_SEC` of no client connection.
- HTTP: server keeps running. Add-ins stay connected.
- Document sessions are unaffected by client churn.

### 4.6 Two clients want to edit the same document

- v1: **last-write-wins, no locking**. Each MCP call is serialized within the
  add-in (single Office.js context per document), but no transaction across calls.
- v2 (deferred): per-document advisory lock via `acquire_edit_lock` /
  `release_edit_lock` MCP tools.

## 5. Threading and concurrency

| Boundary | Concurrency model |
|---|---|
| MCP client → server | Per-request async. Server handles N requests in parallel. |
| Server → add-in | At most `MAX_INFLIGHT_PER_SESSION` (default 4) concurrent calls per session. |
| Inside add-in | Office.js requires single-threaded batched access. Add-in queues calls into a worker, calls `context.sync()` per batch. |

Tool calls are stamped with a request ID by the server. The add-in echoes the ID
back so the server can correlate even with out-of-order responses.

## 6. Versioning

- **MCP protocol**: server declares MCP version it supports (currently 2026-XX,
  the stateless revision).
- **Add-in ↔ server protocol**: independent `protocol_version` field exchanged in
  `register`. Semver: major bump = incompatible; minor = backward-compatible add.
- **Tool surface**: each tool carries `since: "X.Y"` in its metadata. Clients
  can filter by version.

A server MUST refuse add-ins whose `protocol_version` major differs from its
own, returning `register.error` with a human-readable upgrade message.

## 7. Failure isolation

Every boundary is designed so failure on one side does not crash the other:

- Add-in crash → server marks session stale, other instances unaffected.
- Server crash → add-ins detect WS close, retry forever; on server restart they
  re-register and resume.
- Office crash → same as add-in crash (Office takes the add-in down with it).
- Client crash → server keeps running; sessions stay alive.

The only fatal scenario is server **and** all add-ins down simultaneously. In
that case the user simply restarts the server; Office add-ins reconnect.
