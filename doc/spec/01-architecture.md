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
  a future remote-access security design; v1 refuses to start instead of
  exposing a remote listener. See [05-security.md §2](05-security.md).
- Multi-user terminal server requires per-user network isolation or a future
  transport design; plain TCP loopback does not identify the connecting OS
  user.

**What this section does NOT define**:

- What "trusted" means between components. That is the security model.
- What boundary the loopback bind itself represents. Also security.

## 1. Component overview

Target repository layout:

```text
doc/
  spec/                         # Product, protocol, security, deployment, UI specs.
src/
  office-ctl/                   # Office add-ins, written in TypeScript.
    common/                     # Shared config, logging, channel, protocol, UI helpers.
    word/                       # Word entry point, initialization, and Word commands.
    excel/                      # Excel entry point, initialization, and Excel commands.
  office-mcp/                   # Native office-mcp product runtime.
    daemon/                     # Rust daemon service, daemon-owned state, and daemon UI.
      src/
        common/                 # Shared config, logging, redaction, limits, audit, errors.
        ui/                     # Daemon web UI assets and UI runtime file helpers.
        api/                    # Local daemon UI/control API: status, config, sessions, tasks.
        mcp/                    # MCP Streamable HTTP frontend and stdio bridge.
        addin_mgr/              # Add-in HTTPS/WSS channel, sessions, routing, command queue.
        tray/                   # Native tray/menu-bar host, menu model, UI open/quit actions.
        main.rs                 # CLI entry point and top-level command dispatch.
        lib.rs                  # Public module wiring only.
packaging/                      # Installers, release packaging, catalog/bootstrap scripts.
```

`src/office-ctl` owns code that runs inside Office webview runtimes. Host-specific
folders such as `word` and `excel` may depend on `common`, but `common` MUST NOT
depend on a specific Office host. `src/office-mcp/daemon` owns the local daemon
process, all server-side session management, and the daemon UI it serves. The
daemon UI source/assets, runtime-file handling, and UI-serving code live under
`src/office-mcp/daemon/src/ui` so daemon UI ownership stays inside the daemon
boundary.
The UI may consume only redacted daemon status/control APIs from `api`; it must
not own protocol routing, session mutation, or Office command execution.
`packaging` owns installation and release assembly only; it must not become a
runtime code owner.

Legacy top-level `mcp-server/`, `addin/`, `docs/`, `rust-daemon/`, and sibling
daemon UI source paths MUST NOT remain after the source-layout migration.
Temporary reference code must not remain in the target tree after Rust parity is
proven; protocol and runtime evidence lives in Rust tests and
`src/office-mcp/daemon/evidence`.

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

The production daemon implementation is the native Rust application in
`src/office-mcp/daemon`. Its web console assets and UI runtime helpers live
under the daemon's `ui` module and are served or bridged by the daemon. Existing
MCP transport behavior, add-in JSON-RPC registration, Office tool routing, UI
state redaction, and evidence harnesses are contract tests for the Rust daemon;
they MUST NOT be deleted or weakened while implementation continues.

Rust code is expected to use explicit domain objects and ownership boundaries,
not C-style procedural modules with global state. Rust source files MUST be
organized by functional module, not as a flat list in `daemon/src`. The required
module boundaries are:

- `common`: shared infrastructure with no product-server ownership. Contains
  config loading/redaction, logging, audit logging, limits, shared error types,
  time helpers, and cross-module utilities. It must not depend on `api`, `mcp`,
  `addin_mgr`, `tray`, or host-specific Office logic.
- `ui`: daemon web UI source/assets, UI runtime file helpers, and UI asset
  serving helpers. It depends on `api` view models only through stable DTOs and
  must not dispatch MCP or add-in commands directly.
- `api`: local daemon UI/control API used by the frontend UI and tray. Owns
  redacted status snapshots, sessions/tasks/client view models, config/status
  control endpoints, and UI event streams. It may read daemon state through
  service interfaces but must not contain MCP protocol handling or Office.js
  command execution.
- `mcp`: MCP-facing interfaces only: Streamable HTTP frontend, MCP session and
  client tracking, stdio bridge, MCP JSON-RPC/request validation, MCP resource
  and prompt catalogs, and translation from MCP calls into add-in session
  invocations.
- `addin_mgr`: add-in-facing runtime management: HTTPS/WSS channel, exact
  `Origin` checks, add-in JSON-RPC framing, registration, heartbeat,
  `SessionRegistry`, per-session `CommandRouter`, stale-session handling, and
  forwarding tool invocations to connected Office add-ins.
- `tray`: native tray/menu-bar integration, tray menu model, platform adapters,
  `Show Office MCP`, and graceful quit confirmation. It consumes `api`/`ui`
  surfaces rather than reading add-in or MCP internals directly.
- Root files such as `main.rs`, `lib.rs`, and top-level daemon composition own
  CLI dispatch and dependency wiring only.

The minimum daemon object model is:

- `OfficeMcpDaemon`: owns process lifetime, startup/shutdown, and component
  wiring.
- `DaemonConfigService` (`common`): loads, validates, redacts, and watches daemon config.
- `McpHttpFrontend` (`mcp`): owns the Streamable HTTP server, request validation,
  MCP session/client tracking, and MCP error translation.
- `AddinChannelServer` (`addin_mgr`): owns local HTTPS/WSS serving, origin checks,
  JSON-RPC framing, heartbeat, and registration.
- `SessionRegistry` (`addin_mgr`): owns Office runtime/session identity, stale-session grace,
  capability metadata, and queue depth.
- `CommandRouter` (`addin_mgr`): owns tool dispatch, per-session serialization, timeouts,
  cancellation, and result/error mapping.
- `DaemonApiServer` (`api`): owns local UI/control API routing and event streams.
- `UiStateStore` (`api`): owns redacted daemon UI snapshots, current tasks, bounded
  command history, and UI subscriptions.
- `TrayController` (`tray`): owns native tray/menu-bar integration, status menu updates,
  show-UI action, and graceful quit confirmation.
- `AuditLog` and `Logger` (`common`): own optional audit records and operational logs with
  redaction at the boundary.

These objects may use traits for platform-specific adapters, such as tray
integration and certificate stores, but protocol/domain state should remain in
portable core types.

#### Rust daemon code style

Rust daemon implementation MUST follow these style rules:

- Keep files small and cohesive. A Rust source file should own one primary
  concept or a tight set of helper types for that concept. Do not place many
  unrelated structs, enums, traits, and service implementations in one `.rs`
  file.
- Prefer explicit object/domain types with clear ownership and injected
  dependencies. Avoid procedural catch-all modules, global mutable state, and
  service code hidden in utility files.
- Unit tests for `foo.rs` live in a sibling `foo_tests.rs` file in the same
  module directory. Production files should not contain large inline
  `#[cfg(test)] mod tests` blocks except for tiny compile-only checks where a
  sibling file would add no value.
- Each module directory may have a `mod.rs` only for public exports and module
  wiring; implementation should live in named files.
- New Rust work must add logging with the `tracing` ecosystem. Use spans for
  request/session/tool lifecycles and structured fields for session ID, request
  ID, tool name, client ID, endpoint, and error code where available.
- Log levels must be intentional: `error` for failed operations requiring
  operator attention, `warn` for degraded/retriable or suspicious conditions,
  `info` for lifecycle and user-visible state transitions, `debug` for normal
  diagnostic details, and `trace` only for high-volume protocol internals.
- Logs must be written to a daemon log file configured through the daemon config
  and visible in `daemon status`/UI diagnostics. Console logging may remain for
  developer runs, but file logging is required for production diagnosis.
- Error investigations should use structured logs first. Do not rely on guessing
  from symptoms when the daemon can record the failing boundary, request ID,
  session ID, and error code.
- Each completed Rust implementation task must be committed and pushed before
  starting the next task. Keep commits scoped to one verified task or reviewable
  slice so daemon changes remain auditable and reversible.

- **Lifetime**: starts at user login (Windows Scheduled Task / macOS launchd
  agent / Linux systemd `--user` unit). Runs until the user logs out or the
  service is stopped administratively. No idle shutdown.
- **Binds two loopback ports**, both from the daemon config:
  - MCP HTTP (Streamable HTTP) at a single `/mcp` endpoint for MCP clients.
  - Local HTTPS/WSS for the add-in bundle and add-in channel.
- **State**: add-in session registry and in-flight request map. No document
  content or undo history. TLS keys are loaded from ACL-restricted
  configuration. The add-in channel carries only WebSocket transport metadata,
  JSON-RPC protocol fields, session routing IDs, and tool payloads.
- **User interface**: a desktop tray icon plus main diagnostic window. The UI
  observes daemon state, connected MCP clients, document sessions, in-flight
  commands, and bounded redacted command history. It does not expose document
  body content. Full UI requirements are in [09-ui.md](09-ui.md).

The canonical client transport is Streamable HTTP. MCP clients that only speak
stdio use the bundled `office-mcp stdio` shim, which is a thin protocol bridge
to the already-running daemon. The shim does not own Office sessions or spawn
add-ins; it only adapts process-spawned clients to the long-lived daemon.

### 2.2 The Office add-in

The Office add-in implementation lives under `src/office-ctl`:

- `common`: shared endpoint configuration, channel client, JSON-RPC types,
  logging, redacted task history, and reusable UI primitives.
- `word`: Word task pane entry point, Word-specific startup, capability probing,
  and `word.*` command implementations.
- `excel`: Excel task pane entry point, Excel-specific startup, capability
  probing, and `excel.*` command implementations.

Host-specific folders may expose a consistent registration envelope to the
daemon, but Office.js calls stay inside the host folder that owns them.

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
| Bind | loopback only in v1; non-loopback refuses to start |
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
serves redacted status snapshots to the tray controller and main window. The
web UI source, UI runtime helpers, backend state, and UI/control API live under
`src/office-mcp/daemon`, separated into the `ui` and `api` modules. The UI MAY
share the add-in HTTPS origin or use an internal desktop bridge, but it MUST
preserve the same privacy boundary: no document body text, no unredacted tool
arguments, and no sensitive local configuration values in UI state.

The normal production daemon path must expose the UI server, not only a test
fixture. `daemon run` owns the UI runtime file and keeps `/ui/`, `/ui/state`,
and `/ui/events` available for as long as the daemon is running. The tray host
and `office-mcp-daemon ui` command both open the URL recorded in that runtime
file.

## 4. Lifecycle scenarios

### 4.1 Cold start

1. User logs in. OS-level autostart launches the daemon.
2. Daemon reads its config, binds MCP HTTP plus local HTTPS/WSS, and idles.
3. Daemon UI server is available on the configured local HTTPS origin.
4. Daemon tray icon appears and reports `Up` once listeners are ready.
5. No add-ins connected → `office.list_sessions` returns `[]`.

### 4.2 MCP client starts before any Office

1. MCP client connects to `http://127.0.0.1:<mcp_http.port>`.
2. Tool calls against any session return an MCP tool error with
   `office_mcp_code: "NO_SESSIONS"` until the user activates the add-in.

### 4.3 User opens Word

1. Word loads and the user or deployment policy activates the add-in.
2. Add-in loads its compiled endpoint or browser-storage override.
3. Add-in dials the configured WSS endpoint and sends `register`. The register
   message contains runtime ID, host metadata, add-in protocol metadata, and
   capability evidence.
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
