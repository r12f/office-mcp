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
- It eliminates whole categories of design complexity (no service discovery,
  no transport encryption, no cross-user identity mapping).
- It composes cleanly with the MCP client ecosystem, which today is
  overwhelmingly single-user desktop tooling (Claude Desktop, Cursor, local
  agents).

**Other deployment shapes are supported but not the default**:

- Non-loopback bind for remote MCP clients or shared screens. Requires
  explicit auth configuration; see [05-security.md §2](05-security.md).
- Multi-user terminal server. Same binary, each user runs their own server
  in their own session, addressing isolation comes from the OS.

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
│                      │  - Add-in WS backend         │                 │
│                      └────────────▲─────────────────┘                 │
│                                   │                                   │
│              WebSocket + JSON-RPC 2.0  (add-ins dial in)              │
│                                   │                                   │
│       ┌───────────────────────────┼──────────────────────┐            │
│       │                           │                      │            │
│  ┌────┴────────────────────────┐  │  ┌──────────────────┴────────┐   │
│  │  Word.exe (instance A)      │  │  │  Word.exe (instance B)    │   │
│  │   office-mcp add-in         │  │  │   office-mcp add-in       │   │
│  │   report.docx, notes.docx   │  │  │   contract.docx (IRM)     │   │
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
- **Binds two loopback ports**, both from the shared config:
  - MCP HTTP (Streamable HTTP) for MCP clients.
  - WebSocket for add-ins.
- **State**: add-in session registry, in-flight request map. No document
  content, no undo history, no credentials.

MCP clients that only speak stdio are out of scope. MCP 2026 promoted
Streamable HTTP to the standard remote transport; if a client needs a
stdio shim it is a generic MCP-stdio-to-HTTP proxy, decoupled from
office-mcp, and not part of this project.

### 2.2 The Office add-in

- **Lifetime**: bound to its Office instance. Loads when Office starts (if
  pinned) or when the user opens the task pane.
- **Connects to**: the daemon's WS endpoint, read from the shared config.
- **State**: WS connection, per-document Office.js context handles, pending
  operation queue (Office.js requires batched `context.sync()`).
- **Responsibilities**: announce instance on load, emit `session.added` /
  `session.removed` as documents open/close, execute tool calls forwarded by
  the daemon.

### 2.3 Why the daemon is a singleton, and why nothing auto-spawns it

The add-in is the connecting party. It dials a fixed loopback port from a
fixed config. That model only works if exactly one process is listening on
that port at a time. Consequences:

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
| Transport | MCP Streamable HTTP (per MCP 2026 stateless revision) |
| Bind | loopback by default; non-loopback requires `api_key` |
| Multiplexing | Multiple concurrent clients supported |
| SSE-only | Not supported (deprecated upstream) |
| stdio | Not supported in-tree; use a generic proxy if needed |

### 3.2 Daemon ↔ add-in: WebSocket + JSON-RPC 2.0

- Bidirectional: daemon can call add-in (`tool.invoke`); add-in can call
  daemon (`session.added` / `session.removed`).
- JSON-RPC 2.0 framing, one JSON object per WS message.
- Heartbeat: daemon pings at `addin_channel.heartbeat_interval_sec` (config
  default 30s). Add-in must respond within `heartbeat_timeout_sec` (config
  default 10s).
- Reconnect: add-in backs off with jitter on disconnect.

Why WebSocket and not Named Pipes / Unix Socket:

- Cross-platform (Office on Web/Mac/Windows).
- Office.js add-ins are web apps and have `WebSocket` natively.
- Loopback bind makes it security-equivalent to a pipe.

### 3.3 Why not stdio between daemon and add-in

Add-ins are launched by Office, not by the daemon. The only IPC where the
add-in can be the connecting party is a network socket.

## 4. Lifecycle scenarios

### 4.1 Cold start

1. User logs in. OS-level autostart launches the daemon.
2. Daemon reads shared config, binds both loopback ports, idles waiting.
3. No add-ins connected → `office.list_sessions` returns `[]`.

### 4.2 MCP client starts before any Office

1. MCP client connects to `http://127.0.0.1:<mcp_http.port>`.
2. Tool calls against any session return
   `error: { code: -32001, message: "No Office instances connected" }`
   until the user opens Office.

### 4.3 User opens Word

1. Word loads; pinned add-in initializes.
2. Add-in reads `[addin_channel]` from shared config.
3. Add-in dials `ws://bind:port`, sends `register` (no auth field on
   loopback; `shared_secret` when configured).
4. Daemon assigns session ID, replies `register.result`.
5. Add-in enumerates open documents, sends `session.added` for each.
6. Documents are now addressable by MCP clients.

### 4.4 User opens a new document in an already-registered Word

1. Add-in's document-change handler fires.
2. Add-in sends `session.added`.

### 4.5 Office crashes mid-call

1. WS drops.
2. Daemon marks affected sessions `stale`; in-flight calls reject with
   `-32002 SESSION_LOST`.
3. Session is held in `stale` for `session_grace_sec` (config) to allow
   Office auto-recovery → add-in reconnect.
4. On grace timeout, session is removed.

### 4.6 MCP client disconnects

- HTTP client just disconnects. Daemon keeps running.
- Document sessions are independent of client churn.

### 4.7 Daemon crashes

1. All add-ins lose WS, retry on backoff.
2. All MCP clients see their HTTP connection drop and must reconnect.
3. OS-level autostart relaunches the daemon (Scheduled Task / launchd /
   systemd policy). Add-ins reconnect on the next backoff tick.

### 4.8 Two clients edit the same document

v1: last-write-wins, no locking. The daemon serializes calls within one
add-in (Office.js requires it), but no transaction across calls. v2 may
introduce advisory `acquire_edit_lock`.

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
