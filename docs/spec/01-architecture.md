# 01 — Architecture

## 0. Deployment model: single-user, local-first

This section is **architecture**, not philosophy. The deployment model is a
concrete design choice that determines where components live, who runs them,
and what crosses the wire. It is independent of the security model in
[05-security.md](05-security.md), which applies the same trust-boundary
principle regardless of deployment shape.

**Default deployment shape**:

- One user identity on one machine.
- Server, add-ins, MCP client, and Office instances all run as the same OS user.
- Server binds loopback only.
- No network exposure outside the user's box.

**Why this shape**:

- It is the deployment the project's killer feature (IRM-aware live editing)
  is built for — the agent operates as the signed-in Office user, which
  requires sharing that user's OS session.
- It eliminates whole categories of design complexity (no remote service
  discovery and no cross-user identity mapping).
- It composes cleanly with the MCP client ecosystem, which today is
  overwhelmingly single-user desktop tooling (Claude Desktop, Cursor, local
  agents).

**Other deployment shapes are future work, not v1 defaults**:

- Non-loopback bind for remote MCP clients or shared screens. Requires
  explicit auth configuration; see [05-security.md §2](05-security.md).
- Multi-user terminal server requires per-user network isolation or a pairing
  protocol; plain TCP loopback does not identify the connecting OS user.

**What this section does NOT define**:

- What "trusted" means between components. That is the security model.
- What boundary the loopback bind itself represents. Also security.

## 1. Component overview

```
┌───────────────────────────────────────────────────────────────────────┐
│                            User's machine                             │
│                                                                       │
│  ┌──────────────┐                                                     │
│  │  MCP client  │ MCP HTTP (Streamable HTTP), loopback                │
│  │  (Claude /   │──────────────┐                                      │
│  │   Cursor /   │              │                                      │
│  │   agent)     │              ▼                                      │
│  └──────────────┘    ┌──────────────────────────────┐                 │
│                      │  office-mcp daemon           │                 │
│                      │  (singleton, autostart)      │                 │
│                      │                              │                 │
│                      │  - MCP HTTP frontend         │                 │
│                      │  - Tool router               │                 │
│                      │  - Session registry          │                 │
│                      │  - Add-in HTTPS/WSS backend  │                 │
│                      │  - Tray + local UI backend   │                 │
│                      └────────────▲─────────────────┘                 │
│                                   │                                   │
│                 WSS + JSON-RPC 2.0  (add-ins dial in)                 │
│                                   │                                   │
│       ┌───────────────────────────┼──────────────────────┐            │
│       │                           │                      │            │
│  ┌────┴────────────────────────┐  │  ┌──────────────────┴────────┐   │
│  │  Word.exe (instance A)      │  │  │  Word.exe (instance B)    │   │
│  │   office-mcp add-in         │  │  │   office-mcp add-in       │   │
│  │   report.docx runtime       │  │  │   contract.docx runtime   │   │
│  └─────────────────────────────┘  │  └───────────────────────────┘   │
└───────────────────────────────────────────────────────────────────────┘
```

## 2. Processes and their roles

### 2.1 The daemon

The daemon is the **single, long-lived office-mcp process** on the machine.
There is exactly one. Everything below assumes it is already running; how it
gets started is install-time concern, not runtime concern (see §2.3).

- **Lifetime**: starts at user login (Windows Scheduled Task / macOS launchd
  agent / Linux systemd `--user` unit). Runs until the user logs out or the
  service is stopped administratively. No idle shutdown.
- **Binds two loopback ports**, both from the daemon config:
  - MCP HTTP (Streamable HTTP) at a single `/mcp` endpoint for MCP clients.
  - Local HTTPS/WSS for the add-in bundle and add-in channel.
- **State**: add-in session registry and in-flight request map. No document
  content or undo history. TLS keys and optional transport secrets are loaded
  from ACL-restricted configuration.
- **User interface**: a desktop tray icon plus main diagnostic window. The UI
  observes daemon state, connected MCP clients, document sessions, in-flight
  commands, and bounded redacted command history. It does not expose document
  body content. Full UI requirements are in [09-ui.md](09-ui.md).

The canonical client transport is Streamable HTTP. MCP clients that only speak
stdio use the bundled `office-mcp stdio` shim, which is a thin protocol bridge
to the already-running daemon. The shim does not own Office sessions or spawn
add-ins; it only adapts process-spawned clients to the long-lived daemon.

### 2.2 The Office add-in

- **Lifetime**: bound to its current document runtime. Loads when the user
  opens the task pane or when a supported deployment activates it.
- **Connects to**: the daemon's WSS endpoint, using the add-in's compiled
  default or browser-storage override.
- **Scope**: one add-in runtime is attached to one Office document. v1 does not
  enumerate or control other documents from that runtime.
- **State**: WS connection, the current document's Office.js context, and a
  pending operation queue. The task pane also keeps the current command and a
  bounded redacted local command history for display.
- **Responsibilities**: register the runtime, announce its one document with
  `session.added`, emit `session.updated` / `session.removed`, and execute tool
  calls forwarded by the daemon.
- **User interface**: a modern task pane showing daemon connection state,
  current document session metadata, the current command, and recent command
  history. See [09-ui.md §6](09-ui.md).

### 2.3 Why the daemon is a singleton, and why nothing auto-spawns it

The add-in is the connecting party. It dials a fixed default loopback port,
optionally overridden in add-in settings. That model only works if exactly one
process is listening on that port at a time. Consequences:

- The daemon must be started by something with deterministic lifetime — the
  OS session start — not by ad-hoc client spawning.
- Multiple installed versions of office-mcp on the same machine resolve the
  same way: whichever the user configured to autostart is the one that runs.
  The other simply isn't running.
- If the daemon is not running, MCP clients get a connect refused. There is
  no "auto-spawn on first request" path. Failing loud is the correct
  behavior; the fix is `office-mcp daemon start` (or fix the autostart).

This is the standard daemon model used by dockerd, pulseaudio, ssh-agent,
language servers in long-running mode. Nothing exotic.

### 2.4 The MCP client

Sees one MCP server, talks to it over HTTP. Doesn't know how many Office
instances are open. Document addressing is via `session_id` returned from
`office.list_sessions`.

## 3. Transports

### 3.1 Client ↔ daemon: MCP Streamable HTTP

| Aspect | Value |
|---|---|
| Transport | MCP Streamable HTTP |
| Endpoint | `http://127.0.0.1:<port>/mcp` by default |
| Bind | loopback by default; non-loopback requires `api_key` |
| Multiplexing | Multiple concurrent clients supported |
| SSE-only | Not supported (deprecated upstream) |
| stdio | Supported only as `office-mcp stdio`, a thin bridge to the daemon |

The implementation follows the negotiated MCP protocol version, including
`Origin` validation, `Accept` handling, `MCP-Protocol-Version`, optional
`MCP-Session-Id`, and POST/GET/DELETE semantics. Invalid browser origins are
rejected with HTTP 403 to prevent DNS rebinding.

### 3.2 Daemon ↔ add-in: WSS + JSON-RPC 2.0

- Bidirectional: daemon can call add-in (`tool.invoke`); add-in can call
  daemon (`session.added` / `session.removed`).
- JSON-RPC 2.0 framing, one JSON object per WS message.
- Production manifests load the local add-in bundle from the daemon over
  HTTPS, and the channel uses `wss://` on the same trusted local origin.
  Plain `ws://` is allowed only in explicit development configurations.
- The WSS upgrade requires an exact configured `Origin` match to the local
  add-in HTTPS origin. Missing or foreign browser origins are rejected before
  JSON-RPC registration.
- Heartbeat: daemon pings at `addin_channel.heartbeat_interval_sec` (config
  default 30s). Add-in must respond within `heartbeat_timeout_sec` (config
  default 10s).
- Reconnect: add-in backs off with jitter on disconnect.

Why WebSocket and not Named Pipes / Unix Socket:

- Available in the browser/webview runtime used by Office add-ins.
- Office.js add-ins are web apps and have `WebSocket` natively.
- Loopback limits network exposure but does not provide peer identity.

### 3.3 Why not stdio between daemon and add-in

Add-ins are launched by Office, not by the daemon. The only IPC where the
add-in can be the connecting party is a network socket.

### 3.4 Daemon UI backend

The daemon UI backend is local-only and not part of the MCP tool surface. It
serves redacted status snapshots to the tray controller and main window. It MAY
share the add-in HTTPS origin or use an internal desktop bridge, but it MUST
preserve the same privacy boundary: no document body text, no unredacted tool
arguments, and no secrets in UI state.

## 4. Lifecycle scenarios

### 4.1 Cold start

1. User logs in. OS-level autostart launches the daemon.
2. Daemon reads its config, binds MCP HTTP plus local HTTPS/WSS, and idles.
3. Daemon tray icon appears and reports `Up` once listeners are ready.
4. No add-ins connected → `office.list_sessions` returns `[]`.

### 4.2 MCP client starts before any Office

1. MCP client connects to `http://127.0.0.1:<mcp_http.port>`.
2. Tool calls against any session return an MCP tool error with
   `office_mcp_code: "NO_SESSIONS"` until the user activates the add-in.

### 4.3 User opens Word

1. Word loads and the user or deployment policy activates the add-in.
2. Add-in loads its compiled endpoint or browser-storage override.
3. Add-in dials the configured WSS endpoint, sends `register` (no auth field on
   loopback; `shared_secret` when configured).
4. Daemon accepts the runtime and replies `register.result`.
5. Add-in sends one `session.added` for its current document. The add-in
   generates the session ID and retains it across WebSocket reconnects for
   the lifetime of that runtime.
6. Documents are now addressable by MCP clients.

### 4.4 User opens a new document in an already-registered Word

Each document has its own add-in runtime. Activating the add-in in another
document starts another runtime, which independently registers and sends its
own `session.added`.

### 4.5 Office crashes mid-call

1. WS drops.
2. Daemon marks affected sessions `stale`; in-flight calls reject with
   `SESSION_LOST`.
3. Session is held in `stale` for `session_grace_sec` (config) to allow a
   transient channel failure in the same runtime to reconnect.
4. On grace timeout, session is removed.

If Office restarts or reopens the document in a new runtime, it receives a new
session ID; it does not reclaim the stale session.

### 4.6 MCP client disconnects

- HTTP client just disconnects. Daemon keeps running.
- Document sessions are independent of client churn.

### 4.7 Daemon crashes

1. All add-ins lose WS, retry on backoff.
2. All MCP clients see their HTTP connection drop and must reconnect.
3. OS-level autostart relaunches the daemon (Scheduled Task / launchd /
   systemd policy). Add-ins reconnect on the next backoff tick.

### 4.8 Two clients edit the same document

v1: last-write-wins, no locking. The daemon dispatches only one call at a
time to a session and keeps a bounded FIFO queue. There is no transaction
across calls. v2 may introduce advisory `acquire_edit_lock`.

## 5. Threading and concurrency

| Boundary | Concurrency model |
|---|---|
| MCP client → server | Per-request async. Server handles N requests in parallel. |
| Server → add-in | Exactly one dispatched call per session; at most `MAX_PENDING_PER_SESSION` (default 4) additional queued calls. |
| Inside add-in | Add-in executes the dispatched call through `Word.run` and one or more required `context.sync()` operations. |

Tool calls are stamped with a request ID by the server. The add-in echoes the ID
back so the server can correlate even with out-of-order responses.

## 6. Versioning

- **MCP protocol**: server negotiates a published MCP protocol version. The
  first implementation targets `2025-11-25`; upgrades require conformance
  tests against the selected version.
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
