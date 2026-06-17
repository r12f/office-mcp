# 08 — Roadmap

## Milestones

## Implementation workflow requirements

- Each checked roadmap item or coherent implementation task must be completed as
  a separate commit after its relevant verification passes.
- Push after each completed task commit. Do not batch multiple completed tasks
  into one large commit.
- The next task MUST NOT start until the current completed task has been
  committed and pushed, unless the current task is explicitly blocked and the
  working tree remains isolated from the next slice.
- Implementation agents must keep each commit scoped to exactly one completed
  task or reviewable slice. Drive-by refactors, unrelated formatting, and
  speculative cleanup must be deferred to their own task commits.
- Large migrations, especially M6.5.0, M6.5.1, and M6.6.1, must be decomposed
  into reviewable slices with one commit and push per slice.
- Commit messages should name the task and, when useful, include the local check
  or evidence command that proves the slice.

### M0 — Spec & scaffolding (this PR)

- [x] Architecture & protocol specs
- [x] Automated feasibility harness: protocol, MCP transport, WSS, manifest,
      stable API types, requirement-set boundaries, and save behavior
- [x] Consented Word runtime smoke test plus representative IRM rights matrix
      (`npm run evidence:runtime -- --include-full-word-smoke
      --include-com-tracked-changes` generates structured full Word runtime
      evidence from `src/office-mcp/daemon/evidence`; `npm run evidence:irm`
      plus `--require-irm-preflight` and `--require-irm` validates the
      representative protected document)
- [x] Repository scaffolding (Rust daemon package + Word `src/office-ctl/word/` scaffold)
- [x] CI matrix (Win x64, macOS arm64, Linux x64 for the server; manifest + task pane syntax check for the add-in)

### M1 — Walking skeleton

End-to-end "agent reads paragraph 0 of an open Word doc":

- [x] Daemon: MCP Streamable HTTP frontend, single tool `office.list_sessions`,
      one resource `office://word/<id>/paragraph/0`
- [x] Server: local HTTPS/WSS listener, register, session.added, ping/pong
- [x] Add-in: skeleton task pane with WSS reverse-connect
- [x] Add-in: implements ONE method — `tool.invoke` for reading paragraph 0
- [x] Stdio bridge and client config helper for stdio-only MCP clients
- [x] Runtime evidence that an agent-style stdio MCP client can call the daemon
      and read the active Word document session.

**Exit criterion**: An MCP-capable agent client can ask
"what does paragraph 1 of my open Word doc say" and get the right answer from the open Word document.

### M2 — Read & insert

- [x] `word.get_text`, `word.get_outline`, `word.get_paragraph`, `word.find_text`,
      `word.get_selection`
- [x] `word.insert_paragraph`, `word.insert_heading`, `word.insert_table`,
      `word.insert_page_break`
- [x] IRM/protection metadata surfacing in `session.added`
- [x] Access-denied mapping and optional rights pre-check when host APIs expose it
- [x] Selection-anchored operations
- [x] Reconnect + grace period

**Exit criterion**: A user can have Claude draft a section into their
open document, including in an IRM-protected file where their policy allows the requested operation.

### M3 — Edit & review

- [x] `word.replace_text` (with `dry_run`)
- [x] `word.update_paragraph`, `word.delete_range`, `word.apply_formatting`
- [x] `word.add_comment`, `word.resolve_comment`
- [x] `word.accept_change`, `word.reject_change`
- [x] Track Changes interaction: when the user has Track Changes enabled, Word
      records edits as revisions; v1 exposes tracked-change resources plus
      accept/reject, but does not toggle Track Changes mode.

**Exit criterion**: Agent can do a real editing pass — find, propose,
replace, comment — on a tracked-changes-on document and the user can
review the result like a normal collaborator's edits.

### M4 — Tables & images

- [x] `word.read_table`, `word.update_cell`, `word.add_row`, `word.add_column`,
      `word.format_cell`
- [x] `word.insert_image` (base64 + URL)
- [x] `word.insert_list` (numbered, bulleted)

### M5 — Document IO

- [x] `word.save`
- [x] Re-evaluate `word.save_as` / `word.export_pdf` only if a stable Office.js
      API can produce bytes or a user-approved destination without native
      filesystem access from the add-in. Current Office.js typings do not
      expose portable Word save-as or PDF export APIs, so these remain reserved
      and are not v1 tools.
- [x] Large-result and long-running-operation hardening for Streamable HTTP

### M6 — Distribution

- [x] Windows developer bootstrap script: validates build/manifest, registers
      trusted add-in catalog, exports an already trusted localhost PFX, and
      creates/removes a user logon Scheduled Task.
- [x] Windows MSI build: packages the native Rust daemon executable, daemon UI
      assets, add-in bundle, catalog manifest, and user autostart/catalog
      registration. Native `office-mcp.exe` packaging and Scheduled Task
      replacement remain production hardening items.
- [x] macOS Homebrew formula template and renderer for release tarballs. Actual
      tap publication waits for a signed GitHub Release artifact.
- [x] Linux systemd user unit template and renderer for release packages. Actual
      Linux package publication waits for signed release artifacts and platform
      smoke evidence.
- [x] Hosted manifest renderer for `office-mcp.dev` release artifacts. Actual
      DNS/hosting publication remains a release operations gate.
- [x] AppSource pre-submission package generator: hosted manifest, add-in
      bundle, checksums, and checklist. Partner Center submission and Microsoft
      validation review remain external gates.

### M6.5 — Product UI

- [x] Daemon UI backend and static web console assets: redacted `/ui/state`,
      `/ui/events`, and `/ui/` are served by the daemon evidence fixture and
      validated by `npm run check:ui`.
- [x] User-visible daemon UI server entry point. Running `office-mcp-daemon
      daemon run` must make the daemon web console discoverable and openable by
      a normal user without manually guessing `https://localhost:<port>/ui/`.
      Required entry points: `office-mcp-daemon ui`, tray `Show Office MCP`, and
      a documented URL in `daemon status` output. Current Rust evidence covers
      the `office-mcp-daemon ui` runtime-file path, tray `Show Office MCP`
      dispatch, `daemon status` output including `uiUrl`, `stateUrl`, `logPath`,
      and `uiCommand`, and `ui.production_daemon_tray` proving production
      `daemon run` publishes a reachable UI runtime URL.
- [x] Real desktop tray icon. The current Rust tray model/probe is not enough;
      the product must create a visible Windows notification-area icon during a
      normal daemon/tray launch, with right-click menu items for Up/Down, client
      count, document count, Show Office MCP, and Quit Office MCP. Current
      automated evidence covers production `daemon run` starting the native
      Windows tray host, logging `created native tray icon` for
      `windows-notification-area`, and `tray --probe` reading the live daemon UI
      state with `Status: Up`, client count, document count, Show Office MCP,
      and Quit Office MCP menu items. Manual right-click visibility remains
      tracked under M6.5.0.
- [x] Daemon main window with status, endpoints, connected MCP clients, grouped
      document sessions, current tasks, and recent command history must be
      reachable from the user-visible UI server/tray path, not only from the UI
      evidence fixture. Current evidence validates rendering behavior through
      `ui.browser_smoke`; production reachability is covered by
      `ui.production_daemon_tray`, `office-mcp-daemon ui`, `daemon status`, and
      tray `Show Office MCP` dispatch.
- [x] Per-document detail expansion showing the most recent 10 commands and
      success/failure details. Current evidence: `npm run check:ui` verifies
      collapsed/expanded document detail panels and per-document command
      history rendering in the daemon main window.
- [x] Add-in task pane product UI showing daemon connection, current document
      session, current task, and latest 20 task history entries for Word and
      Excel. Current evidence: `npm run check` in `src/office-ctl/word` and
      `src/office-ctl/excel` validates the shared browser UI helpers,
      current-task/history store usage, endpoint settings, command history
      limits, cancellation state, and failure details.
- [x] Redacted UI state API and tests proving document body content, inserted
      text, image data, and sensitive local configuration values do not appear
      in UI snapshots. Current evidence: `npm run evidence:ui` passes
      `ui.state_api_origin_redaction` and `ui.events_stream`.
- [x] Automated UI coverage for empty, connected, degraded, in-flight, success,
      failure, timeout, and reconnect states.
      Current evidence covers degraded daemon state, connected clients,
      grouped Word/Excel documents, stale/reconnecting document state, empty
      daemon state, in-flight tasks,
      success/failure/timeout/cancelled history filtering, document detail
      inspection, inspector clearing, endpoint copy announcement, dark mode,
      high contrast, reduced motion, desktop layout, and 320 px task pane/main
      window layout.

**Exit criterion**: A non-technical user can see the daemon tray icon, open the
tray menu, open the daemon main window from that menu or `office-mcp-daemon ui`,
and understand whether office-mcp is running, which clients and documents are
connected, what task is currently executing, and why the last relevant command
failed, without inspecting terminal output, logs, or hidden localhost URLs.

### M6.5.0 — Daemon UI server and tray visibility follow-up

User-reported follow-up from live daemon testing:

- [x] Add a production `office-mcp-daemon ui` command that reads the current UI
      runtime file, opens/focuses the daemon web console, and prints a clear
      error when no daemon UI server is running. Covered by `open_ui`,
      `ui_url_from_runtime_path`, and
      `ui_command_reads_runtime_file_url_instead_of_config_defaults`.
- [x] Ensure `daemon status` reports the actual UI URL and state URL for the
      currently running daemon, including non-default ports such as `8766`.
      Covered by `DaemonController::status_json` and
      `status_reports_runtime_details_without_auth_material`.
- [x] Ensure `daemon run` writes a fresh UI runtime file and keeps `/ui/`,
      `/ui/state`, and `/ui/events` available for the lifetime of the daemon.
      Covered by `serve_forever_with_runtime_file`, runtime file tests, and
      `production_bound_daemon_exposes_ui_state_and_events`.
- [x] Add a real Windows tray startup path that creates a visible notification
      icon in normal interactive runs, not just `tray --probe` evidence. Current
      automated evidence covers `office-mcp-daemon daemon run`, which starts the
      native tray host by default on a background thread before running the
      daemon. `--no-tray` is retained for headless/service runs, and
      `--with-tray` remains accepted for compatibility. `ui.production_daemon_tray`
      now proves the production path creates the native Windows tray host,
      exposes live UI state, and logs native notification-area icon creation.
      Manual Windows interaction evidence remains required for right-click menu
      visibility.
- [x] Wire tray `Show Office MCP` to the same UI-opening path as
      `office-mcp-daemon ui`. Covered by `open_ui_from_runtime` and the native
      tray menu action dispatch.
- [x] Wire tray `Quit Office MCP` to graceful daemon shutdown and visible menu
      confirmation. Covered by `stop_daemon`, `confirm_quit`, and
      `native_tray_quit_uses_platform_confirmation_dialogs`.
- [ ] Add manual/e2e evidence that verifies a visible tray icon exists on
      Windows, the right-click menu appears, and `Show Office MCP` opens the
      daemon UI. Automated evidence now covers native tray icon creation logs
      and the live menu model through `ui.production_daemon_tray`; this item
      remains open until a visible desktop interaction run verifies the icon and
      right-click menu in the notification area.
- [x] Add automated coverage that fails when production `daemon run` does not
      expose `/ui/` and when `daemon status` omits the UI URL. Covered by
      `production_bound_daemon_exposes_ui_state_and_events`,
      `status_reports_runtime_details_without_auth_material`, and
      `ui_command_reads_runtime_file_url_instead_of_config_defaults`.

**Exit criterion**: Starting office-mcp in its normal local mode creates a
running UI server and a visible tray icon. A user can open the daemon UI from
the tray or CLI without knowing the localhost URL.

### M6.5.1 — Add-in task pane layout and tool permissions

User-reported follow-up from live task pane testing:

- [x] Replace oversized task pane blocks with content-based section heights.
      The connection/document summary, current task empty state, and recent task
      empty state must be compact and must not reserve large blank panels.
      Covered by Word/Excel compact narrow-width task pane contract tests.
- [x] Merge connection status and document metadata into one compact top block.
      The top block contains connection status, document metadata, settings,
      and tool summary, with Word/Excel contract tests enforcing the order.
- [x] Move settings into an inline child panel of the top block. Clicking the
      gear must not append a detached settings block at the bottom of the task
      pane. Covered by Word/Excel HTML order contract tests.
- [x] Replace bare tool-count badges such as `27 Tools` with an inspectable
      grouped tool list that shows actual tool names, categories, side effects,
      and descriptions. Word and Excel now show `Enabled X of Y`, grouped
      tool names, and per-group enabled counts.
- [x] Add per-tool enable/disable controls in task pane settings for Word and
      Excel. The persisted state is per full tool name; category toggles are
      optional shortcuts only.
- [x] Wire tool permissions into the add-in registration/session update flow:
      disabled tools are omitted from effective `available_tools`, the daemon
      preflight rejects absent tools, and an in-flight race returns
      `TOOL_DISABLED_BY_USER`.
- [x] Add automated UI evidence for the compact layout, inline settings panel,
      grouped tool inspection, and per-tool permission toggles at 320 px width.
      Current evidence is static HTML/CSS/JS contract coverage in Word and Excel
      task pane tests; future pixel-level visual regression can strengthen it.
- [x] Add protocol/unit coverage proving `session.updated.available_tools`
      changes after a permission toggle and that disabled tools cannot be
      invoked. Coverage exists in the shared add-in channel tests, Word/Excel
      task pane tests, and Rust session preflight tests.

**Exit criterion**: In a 320 px Office task pane, the first viewport shows a
compact connection/document summary, current task, and recent task status
without large empty blocks; settings open inside the top block; users can see
which tools are available and disable any individual tool with daemon-visible
effect.

### M6.6 — Rust native daemon migration

- [x] Normalize the source tree away from legacy top-level packages into
      `doc/`, `src/office-ctl/{common,word,excel}`, `src/office-mcp/daemon`,
      and `packaging/`. Daemon UI ownership now lives in
      `src/office-mcp/daemon/src/ui`, including static web console assets.
- [x] Add the Rust daemon under `src/office-mcp/daemon` and preserve protocol
      parity through Rust unit tests plus evidence harnesses.
- [x] Move the daemon web console out of legacy top-level runtime paths. It may
      only consume redacted daemon status APIs and must not own protocol
      routing, session mutation, or Office command execution. Final ownership
      moves under the daemon `ui` module in M6.6.1.
- [x] Port the daemon behind explicit domain objects: `OfficeMcpDaemon`,
      `DaemonConfigService`, `McpHttpFrontend`, `AddinChannelServer`,
      `SessionRegistry`, `CommandRouter`, `UiStateStore`, `TrayController`,
      `AuditLog`, and `Logger`.
- [x] Reuse or mirror the existing protocol and runtime evidence tests as
      Rust parity gates. The rewrite MUST preserve MCP transport semantics,
      add-in JSON-RPC registration, Word tool behavior, error shapes, UI
      redaction, and tray/menu UI evidence.
- [x] Move Windows tray, macOS menu-bar, and Linux tray/status-notifier support
      into native Rust platform adapters behind `TrayController` traits. The
      Rust domain model, shared Windows/macOS/Linux native tray host entry,
      Windows tray probe evidence, and installer autostart wiring are
      implemented.
- [x] Add macOS and Linux packaging smoke evidence for the native tray/menu-bar
      launch path. CI runs the Homebrew and Linux systemd packaging smoke
      checks on macOS and Ubuntu hosts, then runs
      `office-mcp-daemon tray --probe` against a seeded UI state snapshot to
      prove the native tray command entry, platform adapter selection, menu
      model, and UI-state read path on those operating systems. Manual visual
      confirmation that the menu-bar/status-notifier icon is visible remains a
      release-host validation step, not an implementation blocker.
- [x] Finish replacing Node packaging after Rust passes the parity suite on
      Windows and the supported non-Windows packaging smoke gates. Windows MSI
      and developer bootstrap stage the Rust daemon as the runtime; legacy Node
      daemon content has been removed. CLI config loading honors
      `OFFICE_MCP_CONFIG_PATH`, which is required by installed launchers.
- [x] Split Office add-in code into `src/office-ctl/common`,
      `src/office-ctl/word`, and `src/office-ctl/excel` before adding Excel
      host behavior, so shared channel/config/logging code is not duplicated.
      Host-neutral browser UI helpers, endpoint configuration, JSON-RPC channel
      helpers, protocol message helpers, scoped logging, and redacted task
      history now live in `src/office-ctl/common` and are served to Word and
      Excel task panes. The Excel manifest, task pane scaffold, registration
      flow, static daemon routing, and v1 `excel.*` command handlers are
      implemented. Real Excel runtime smoke evidence is covered by
      `artifacts/runtime-evidence-excel.json` and validated with
      `npm run evidence:validate -- --require-excel-smoke`.

**Exit criterion**: The Rust daemon can replace the Node daemon without changing
the add-in protocol, MCP client behavior, evidence report schema, UI state
redaction guarantees, user-visible tray/main-window behavior, or the host add-in
contract shared by Word and Excel.

### M6.6.1 — Rust daemon source organization

User-reported architecture follow-up: the Rust daemon source is currently too
flat. Files under `src/office-mcp/daemon/src` must be reorganized by functional
module so ownership is visible from the directory tree.

- [x] Create `src/office-mcp/daemon/src/common/` for shared config, logger,
      audit log, redaction, limits, shared errors, and utility code. `common`
      must not depend on product-facing modules.
- [x] Create `src/office-mcp/daemon/src/ui/` and merge the former sibling
      daemon UI source/assets into it. Daemon web console assets now
      live in `src/office-mcp/daemon/src/ui/assets`, and the old sibling path is
      removed.
- [x] Create `src/office-mcp/daemon/src/api/` for daemon UI/control APIs:
      status, sessions, current tasks, recent history, config display/control,
      UI runtime file lookup, and UI event streams.
- [x] Create `src/office-mcp/daemon/src/mcp/` for MCP-only code: Streamable HTTP
      frontend, stdio bridge, MCP management client, resources, prompts, tool
      catalog, MCP request validation, and MCP error translation.
- [x] Create `src/office-mcp/daemon/src/addin_mgr/` for add-in-facing code:
      local HTTPS/WSS channel, exact `Origin` validation, add-in JSON-RPC,
      registration, heartbeat, session registry, command router, image fetch
      preprocessing for add-in calls, and stale-session handling.
- [x] Create `src/office-mcp/daemon/src/tray/` for tray/menu-bar code: tray
      controller, native tray host, menu model, `Show Office MCP`, and graceful
      quit confirmation.
- [x] Keep root files minimal: `main.rs` owns CLI dispatch, `lib.rs` exposes
      module wiring, and any top-level daemon composition file only wires
      service objects together. Runtime server and UI evidence fixture code now
      live under `src/office-mcp/daemon/src/runtime/`; the root layout guard
      `daemon_src_root_only_contains_composition_and_transitional_files` proves
      no service modules are added directly under `src`.
- [x] Rename Rust modules where needed to match their service boundary. For
      example, UI state snapshot code belongs under `api`, tray host/controller
      under `tray`, MCP HTTP/frontend/stdio under `mcp`, and session/command
      routing under `addin_mgr`. Current daemon source separates shared daemon
      config model/error types into `common::config_model` and
      `common::config_error`, UI state DTOs into `api::state_model`, MCP HTTP
      request/decision/session DTOs into `mcp::http_frontend_model`, add-in
      registry DTOs into `addin_mgr::session_registry_model`, add-in tool
      payload helpers into `addin_mgr::addin_tool_payload`, add-in protocol and
      clock helpers into dedicated `addin_mgr` modules, and the TLS WebSocket
      session loop into `runtime::websocket_session`.
- [ ] Split oversized Rust files so each `.rs` file owns one primary concept or
      a tight helper set. Avoid catch-all files with many unrelated domain
      objects. Progress so far split config, UI state, MCP HTTP frontend,
      session registry, add-in tool payload, add-in protocol/clock helpers, and
      runtime WebSocket session objects into separate files with sibling
      `*_tests.rs` coverage. Recent slices also extracted add-in registration
      policy and add-in session event handling from `addin_channel.rs`, plus
      command completion helpers from `command_router.rs`, and MCP forwarded
      tool invocation from `runtime/mcp_rpc.rs`, and runtime connection
      handling from `runtime/server.rs`. The MCP prompt catalog now lives in
      `mcp/prompt_catalog.rs` instead of the shared catalog file, and logger
      record serialization now lives in `common/logger_record.rs`. UI evidence
      fixture options, path resolution, and fixture daemon config now live in
      `runtime/evidence_fixture_config.rs`. Tray status DTOs, menu model, and
      quit confirmation text now live in `tray/model.rs`, leaving
      `tray/controller.rs` focused on controller behavior and platform adapter
      contracts. Configuration environment override parsing now lives in
      `common/config_env.rs`, leaving `common/config_service.rs` focused on
      building and validating the daemon config object. Add-in JSON-RPC params
      parsing and register reply serialization now live in
      `runtime/addin_rpc_message.rs`, leaving `runtime/addin_rpc.rs` focused on
      dispatching parsed add-in methods into channel and registry services. The item remains open until the
      remaining production files above the review threshold are audited for
      single-concept ownership.
- [x] Move inline Rust unit tests into sibling files named after the production
      file, such as `logger_tests.rs` for `logger.rs`. Keep only minimal
      compile-only inline test modules when a sibling file is not useful.
- [x] Add `tracing`-based structured logging across daemon boundaries: startup,
      config load/reload, MCP request handling, add-in registration/session
      updates, tool dispatch, command completion/failure, tray actions, UI API
      calls, and shutdown.
      Covered by structured tracing events and tests for daemon startup/config/
      shutdown (`daemon_run_has_startup_config_and_shutdown_tracing`), MCP HTTP
      (`writes_structured_tracing_events_for_session_lifecycle_and_rejections`),
      add-in lifecycle (`writes_structured_tracing_events_for_addin_session_lifecycle`),
      command routing (`writes_structured_tracing_events_for_command_lifecycle`),
      tray host actions (`tray_probe_writes_structured_tracing_event` and
      `native_tray_actions_emit_tracing_events`), UI snapshot rendering
      (`writes_structured_tracing_event_for_ui_api_snapshot`), and file logging
      (`tracing_file_subscriber_writes_json_events_with_level_filter`).
- [x] Configure daemon file logging with intentional levels and structured
      fields. `daemon status` and the daemon UI must expose the current log path
      so errors can be diagnosed from logs instead of guessed from symptoms.
      The daemon initializes a tracing JSON file subscriber from
      `logging.level`/`logging.file`, and daemon UI state exposes `log_path`.
- [x] Add logging tests or evidence that verifies log records are written to the
      configured file, redact sensitive values, include request/session/tool
      context when available, and preserve useful levels. Rust unit tests cover
      file output, level filtering, redaction, and UI log-path exposure.
- [x] Update packaging, evidence harnesses, static asset serving, docs, and tests
      so no production code references the removed sibling daemon UI
      path. Static serving now resolves daemon UI assets from the daemon `ui`
      module ownership boundary.
- [x] Add a source-layout test or lint script that fails if new daemon service
      modules are added directly under `src/office-mcp/daemon/src` instead of
      one of `common`, `ui`, `api`, `mcp`, `addin_mgr`, or `tray`.

**Exit criterion**: A new contributor can identify daemon ownership boundaries
from the directory tree alone; the daemon has no large flat module dump under
`src`, tests live in sibling `_tests.rs` files, tracing-backed file logs are
available for diagnosis, the UI source is owned by the daemon `ui` module, and
all existing Rust, evidence, packaging, and add-in checks still pass.

### M7 — Excel

A separate add-in (Excel.js) connects to the same server. Server gains a
new tool namespace `excel.*`. Catalog (rough): `excel.read_range`,
`excel.write_range`, `excel.add_sheet`, `excel.set_formula`,
`excel.format_range`, `excel.create_table`, `excel.create_chart`.

- [x] Excel add-in scaffold under `src/office-ctl/excel` with Workbook
      manifest validation, modern task pane shell, shared channel/logger/task
      history usage, Excel runtime registration, session announcement, and
      daemon static serving under `/excel/*`.
- [x] Implement representative workbook read/write behavior for
      `excel.read_range` and `excel.write_range`, including daemon tool catalog
      entries, capability-gated forwarding, Excel task pane handlers, and unit
      evidence for MCP-to-add-in dispatch.
- [x] Implement the remaining `excel.*` catalog: `excel.add_sheet`,
      `excel.set_formula`, `excel.format_range`, `excel.create_table`, and
      `excel.create_chart`, including daemon catalog entries, capability-gated
      forwarding, Excel task pane handlers, and static/unit evidence.
- [x] Add real Excel runtime smoke evidence against a live workbook.
      Current evidence: `npm run evidence:excel` writes
      `artifacts/runtime-evidence-excel.json`; validation passes with
      `npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-excel.json --require-excel-smoke`
      against a connected live Excel workbook.

### M8 — PowerPoint

Similar pattern: `powerpoint.*` namespace, second add-in.
Catalog: `add_slide`, `replace_text`, `insert_image`, `apply_layout`,
`export_pdf`.

### M9 — Outlook (cautious)

Outlook add-ins have a stricter sandbox (no full mailbox enumeration without
extra perms). Likely scoped to "active item": read/draft an email, summarize
thread, suggest reply.

## Deferred / under consideration

| Item | Why deferred | Triggers reconsidering |
|---|---|---|
| `tool.progress` streaming partials | v1 keeps protocol simple | First user request for >10s tool calls where partial results matter |
| Per-document `acquire_edit_lock` | Single-client usage is dominant | First report of two-client clobbering |
| Server-issued events back to MCP client (selection-change, etc.) | Most clients don't consume them | MCP gains widely-implemented subscription primitive |
| Publicly trusted WSS endpoint | Office on Web cannot rely on a user-installed localhost certificate | First supported Office on Web deployment |
| Macro execution | Security risk too large | Never, except via explicit allowlist + signed origin |
| Server-side LLM summary tools (`word.summarize`) | Out of scope; that's the client's job | Never |
| Multi-tenant SaaS hosting | Not the project's model | Never |

## Non-goals (restated)

- Replacing `python-docx` for batch/headless work.
- Rendering / formatting fidelity that exceeds what Office.js can produce.
- Authoring or modifying IRM/AIP policies.
- Bypassing any Office security control.

## Open questions

1. **Office on Web**: browser private-network policy and certificate trust make
   localhost access materially different from desktop Office. It is not a v1
   target; a future deployment may require a publicly trusted per-user relay.
2. **Identity binding**: in multi-account Word (one user signed into both work
   and personal Microsoft accounts), how do we tag sessions by identity?
   Probably via `user.upn` + `user.tenant_id` in `session.added`. Test M2.
3. **Multiple installed versions on one machine**: if a user has both
   office-mcp 0.5 and 0.6 installed, both using the default endpoint and
   both starting at logon, they'll race for `127.0.0.1:8765`. Probably
   acceptable (whichever starts second fails loudly with "address already in
   use"); but a better answer might be: registered servers post their version
   to the WS handshake, second one defers. Investigate M2.

## Resolved decisions

1. **Add-in activation UX**: the fallback ribbon command remains **Home >
   office-mcp > Open**. Once the task pane connects, it writes
   `Office.AutoShowTaskpaneWithDocument=true` into the document settings so
   supported sideloaded or centrally deployed documents reopen the pane on
   future opens. Marketplace behavior remains subject to Office host policy.
