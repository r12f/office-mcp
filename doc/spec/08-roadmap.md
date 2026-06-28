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
- [x] Consented Word runtime smoke test plus representative IRM rights matrix.
      `npm run evidence:word` now owns a self-contained live Word document
      lifecycle: it starts or reuses the daemon, creates a driver-owned
      document, activates the add-in, uses the Office tool E2E report path and runs
      Word read/mutation/full-smoke/COM tracked-change evidence, cleans up the
      document, and writes `artifacts/runtime-evidence-word.json`. `npm run
      evidence:runtime` aliases that self-contained Word run. `npm run
      evidence:irm` plus `--require-irm-preflight` and `--require-irm`
      validates the representative protected document.
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
- [x] `word.resize_image` (in-place inline image resizing by paragraph index)
- [x] `word.insert_list` (numbered, bulleted)

### M4.1 — Word Core Tool Surface Refinement

Research basis: Microsoft Word add-in docs identify the core object model as
`Document` -> sections/body -> `Range`, with paragraphs, lists, tables, content
controls, comments, and tracked changes as the high-value user-facing objects.
The refined catalog should be about 25 tools and must apply Occam's razor: do
not keep a separate tool when it only repeats another tool's object owner and
user intent.

Target catalog: `word.get_text`, `word.get_outline`, `word.get_paragraph`,
`word.find_text`, `word.get_selection`, `word.insert_paragraph`,
`word.insert_table`, `word.insert_image`, `word.resize_image`, `word.insert_page_break`,
`word.insert_list`, `word.replace_text`, `word.update_paragraph`,
`word.delete_range`, `word.apply_formatting`, `word.apply_style`,
`word.read_table`, `word.update_table`, `word.list_content_controls`,
`word.insert_content_control`, `word.update_content_control`,
`word.delete_content_control`, `word.add_comment`, `word.resolve_comment`,
`word.update_tracked_change`, and `word.save`.

Superseded compatibility tools: `word.insert_heading`, `word.set_heading_level`,
`word.update_cell`, `word.add_row`, `word.add_column`, `word.format_cell`,
`word.accept_change`, and `word.reject_change`.

- [x] Verify planned Word content-control and consolidated mutation APIs against
      `@types/office-js` and Microsoft API docs before implementation. Current
      spec and tests pin content controls to generic `Word.ContentControl` APIs
      from `WordApi 1.5`, tracked-change mutation to `WordApi 1.6`, and table /
      style consolidation to the existing portable Word API surface.
- [x] Extend `word.insert_paragraph` so heading insertion is represented by
      style or heading-level arguments, then retire `word.insert_heading` from
      the advertised catalog after compatibility evidence is captured.
- [x] Extend `word.apply_style` so heading-level changes are represented as
      semantic style changes, then retire `word.set_heading_level` from the
      advertised catalog.
- [x] Implement `word.update_table` as the single table mutation owner for cell
      values, row/column insertion, table/cell formatting, and table deletion;
      current implementation adds the target owner to the daemon catalog, MCP
      `tools/list`, Word task pane available-tools metadata, permission grouping,
      dispatch path, and spec contract. It reuses the existing tested cell,
      row, column, and format behavior and adds explicit whole-table deletion.
      The compatibility tools were removed from the advertised catalog in the
      target-surface retirement slice; historical handlers remain only as local
      implementation helpers where needed.
- [x] Add content-control CRUD tools: `word.list_content_controls`,
      `word.insert_content_control`, `word.update_content_control`, and
      `word.delete_content_control`, keeping generic text edits owned by
      paragraph/range tools unless the caller targets a content control. Current
      implementation uses generic `Word.ContentControl` APIs from `WordApi 1.5`,
      supports rich text and plain text controls, lists metadata without
      duplicating contained document text, updates tag/title/lock/text metadata,
      and deletes controls with explicit keep/delete content handling.
- [x] Implement `word.update_tracked_change` with an explicit `action` argument
      for accept/reject, then retire `word.accept_change` and
      `word.reject_change` from the advertised catalog. Current implementation
      adds the target owner to the daemon catalog, MCP `tools/list`, Word task
      pane available-tools metadata, permission grouping, dispatch path, and
      spec contract while reusing the existing fingerprint/stale-index safety
      check. The compatibility tools were removed from the advertised catalog in
      the target-surface retirement slice.
- [x] Update daemon MCP catalog entries, JSON schemas, permission categories,
      task pane tool grouping, runtime evidence scripts, and README text from
      the current 27-tool compatibility surface to the refined 25-tool surface.
      Current evidence: daemon `WORD_V1_TOOLS` and Word task pane
      `AVAILABLE_TOOLS` both expose exactly the refined 25-tool catalog;
      compatibility/deprecation tests reject superseded tool names; and runtime
      smoke scripts call the refined owner tools instead of retired compatibility
      tools.
- [x] Add compatibility/deprecation tests proving superseded tools are not
      advertised after migration while their target-owner replacements cover the
      same user workflows without duplicate writes.
- [x] Define the shared typed Word `Anchor` schema in the daemon catalog and
      narrow each anchored tool's advertised `oneOf` to the anchor kinds it
      supports. The daemon must reject unsupported anchor kinds before
      forwarding mutating calls to Office.js, with `INVALID_ARGUMENTS` and
      `partial_effect: none`.

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
- [x] Windows portable zip build: packages the native Rust daemon executable,
      daemon UI assets, add-in bundles, catalog manifests, explicit user install
      scripts, and product assets. MSI packaging was removed from the release
      path because the opaque installer UX was not acceptable for first-contact
      trust.
- [x] Simplify the Windows portable zip root so the user-facing entries are
      `office-mcp-daemon.exe`, `install.ps1`, and `uninstall.ps1`. Remove
      duplicate daemon launcher scripts such as `office-mcp.ps1`,
      `office-mcp-daemon.ps1`, and `office-mcp-tray.ps1`; `install.ps1` should
      set process-scoped portable environment and directly launch
      `office-mcp-daemon.exe daemon run`.
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

### M6.1 — GitHub Release portable package pipeline and user install docs

User-reported distribution gap: the project had packaging scripts but did not
publish a transparent GitHub Release package, and the README did not give
non-developers a clear install path. A program that requires source checkout
commands is not installable software.

- [x] Add `.github/workflows/release.yml` for version-tagged releases. The
      workflow must trigger on tags like `v0.1.0`, support a maintainer dry-run
      `workflow_dispatch`, build the Windows portable zip on `windows-latest`, upload
      artifacts for every run, and attach release assets only for tag builds.
      Current implementation uses the checked release workflow on tag pushes
      and `workflow_dispatch`, uploads the portable zip and `SHA256SUMS` on every run,
      and creates a draft pre-release only when `github.ref_type == 'tag'`.
- [x] Make the release workflow reuse the existing packaging path instead of
      inventing another installer flow: install Rust, Node, add-in dependencies,
      evidence dependencies, and packaging dependencies; run packaging smoke
      checks; then call
      `packaging/windows/build-windows-portable.ps1 -SkipNpmInstall`. Current
      workflow installs Rust, Node, evidence dependencies, all host add-in
      dependencies, and packaging dependencies, runs add-in checks, daemon
      tests, daemon evidence checks, packaging checks, then invokes the existing
      Windows portable builder with `-SkipNpmInstall`.
- [x] Add release-gate tests/checks that fail before publishing when the portable zip is
      missing, unexpectedly small, incorrectly named, or missing required
      payloads: native Rust daemon, daemon UI assets, shared Office add-in
      assets, Word/Excel/PowerPoint task pane bundles, catalog manifests,
      install/uninstall scripts, and product icons. Current workflow stages artifacts
      only after validating the versioned portable zip name, minimum size, and required
      payload paths, and fails if duplicate root launcher scripts re-enter the
      package; `packaging/test/windows.test.mjs` statically locks these gates.
- [x] Generate and publish `SHA256SUMS` for every release artifact. If signing
      is not implemented yet, unsigned pre-releases must be explicitly labeled
      as unsigned rather than silently looking production-signed. Current
      workflow generates `SHA256SUMS`, uploads it with the portable zip, and uses
      `RELEASE_NOTES.md`, which labels 0.1.0 artifacts as unsigned.
- [x] Keep releases draft or pre-release by default until release notes, manual
      tray evidence, portable package smoke evidence, and required live Office evidence
      are attached or explicitly waived for that pre-release. Current workflow
      sets `draft: true`, `prerelease: true`, requires `RELEASE_NOTES.md` for
      tag releases, and publishes that file as the release body.
- [x] Add a README user installation section before the source/developer setup.
      It must explain how to download `office-mcp-windows-portable-<ver>-x64.zip` from
      GitHub Releases, verify checksums, extract it, run `install.ps1`, restart/open Office,
      locate `Office MCP Control` in the Shared Folder catalog when needed,
      open the daemon UI from the tray or CLI, configure MCP clients to
      `http://127.0.0.1:8800/mcp`, find logs, and uninstall. Current README
      includes the user install section before developer setup.
- [x] Add documentation tests or static checks that fail when README no longer
      mentions GitHub Releases, the portable zip asset name, checksum verification,
      daemon UI/tray first-run verification, Office Shared Folder activation,
      MCP endpoint configuration, log collection, and uninstall steps. Current
      packaging static tests cover the release workflow, README install guide,
      and release notes gate.

Current local verification for the release-pipeline implementation:

```powershell
cd C:\Code\office-mcp\packaging
npm run check
cd C:\Code\office-mcp\src\office-mcp\daemon\evidence
npm run check
cd C:\Code\office-mcp
git diff --check
```

Tag-run evidence remains a release-operations gate: the implementation is in
place, but a maintainer still needs to push a version tag or dispatch the
workflow in GitHub Actions to capture the actual hosted Release artifact URL.

**Exit criterion**: Pushing a version tag creates a GitHub Release draft or
pre-release with a downloadable Windows portable zip and checksums. A Windows
desktop user can install from README instructions without cloning the repository
or running developer build commands.

### M6.2 — Fixed-root Windows installer and safe upgrades

User-reported installer gap: the first one-line installer put every release in
`%LOCALAPPDATA%\office-mcp\<release-tag>`, which leaked old versions, made the
installed location hard to reason about, and left Office trusted catalogs at
stale package paths. The portable zip must remain transparent, but installation
must behave like normal software.

- [x] Keep `scripts/install.ps1` as a thin latest-release bootstrap. It resolves
      a published Windows portable zip, extracts it to a temporary staging
      folder, invokes the package-local `install.ps1`, and does not install
      product files into a version-tagged directory.
- [x] Make the package-local `install.ps1` install or upgrade into a stable
      default root, `%LOCALAPPDATA%\office-mcp`, with custom roots supported by
      `OFFICE_MCP_INSTALL_ROOT` for one-line installs and `-InstallRoot` for
      manual package installs.
- [x] Stop existing Office MCP daemon processes before replacing installed
      runtime payloads, preserve user-owned config/certificate/log files, and
      copy the new package payload without mixing old and new UI/add-in assets.
- [x] Clean safe-to-identify stale versioned install roots created by old
      bootstraps and stale Office MCP trusted catalog entries that point at
      previous versioned `addin-catalog` paths.
- [x] Handle running Word, Excel, and PowerPoint during install by listing the
      running hosts, prompting before closing them in interactive mode, and
      failing with an actionable message in non-interactive mode unless
      `-CloseOfficeHosts` is supplied.
- [x] Update README and deployment spec so one-line install, manual install,
      custom install root, upgrade cleanup, Office-host prompting, trusted
      catalog ownership, logs, and uninstall guidance match the actual package.
- [ ] Before the next release, locally verify a clean install from no install
      root and no trusted catalog registry key, an upgrade from an older
      versioned install root, the daemon UI at `https://localhost:8765/ui/`, and
      the trusted catalog URL pointing at the fixed install root's
      `addin-catalog`.

**Exit criterion**: Installing or reinstalling from the one-line command uses a
stable install root, does not accumulate versioned install directories or stale
Office catalog paths, restarts the daemon/tray from the fixed root, and leaves
Word, Excel, and PowerPoint able to discover Office MCP Control from the trusted
Shared Folder catalog after Office restarts.

### M6.5 — Product UI

- [x] Daemon UI backend and static web console assets: redacted `/ui/state`,
      `/ui/events`, and `/ui/` are served by the daemon evidence fixture and
      validated by `npm run check:ui`.
- [x] User-visible daemon UI server entry point. Running `office-mcp-daemon
      daemon run` must make the daemon web console discoverable and openable by
      a normal user without manually guessing `https://localhost:<port>/ui/`.
      Required entry points: `office-mcp-daemon ui`, tray `Show Office MCP Control`, and
      a documented URL in `daemon status` output. Current Rust evidence covers
      the `office-mcp-daemon ui` runtime-file path, tray `Show Office MCP Control`
      dispatch, `daemon status` output including `uiUrl`, `stateUrl`, `logPath`,
      and `uiCommand`, and `ui.production_daemon_tray` proving production
      `daemon run` publishes a reachable UI runtime URL.
- [x] Real desktop tray icon. The current Rust tray model/probe is not enough;
      the product must create a visible Windows notification-area icon during a
      normal daemon/tray launch, with right-click menu items for Up/Down, client
      count, document count, Show Office MCP Control, and Quit Office MCP Control. Current
      automated evidence covers production `daemon run` starting the native
      Windows tray host, logging `created native tray icon` for
      `windows-notification-area`, and `tray --probe` reading the live daemon UI
      state with `Status: Up`, client count, document count, Show Office MCP Control,
      and Quit Office MCP Control menu items. Manual right-click visibility remains
      tracked under M6.5.0.
- [x] Daemon main window with status, endpoints, connected MCP clients, grouped
      document sessions, current tasks, and recent command history must be
      reachable from the user-visible UI server/tray path, not only from the UI
      evidence fixture. Current evidence validates rendering behavior through
      `ui.browser_smoke`; production reachability is covered by
      `ui.production_daemon_tray`, `office-mcp-daemon ui`, `daemon status`, and
      tray `Show Office MCP Control` dispatch.
- [x] Per-document detail expansion showing the most recent 10 commands and
      success/failure details. Current evidence: `npm run check:ui` verifies
      collapsed/expanded document detail panels and per-document command
      history rendering in the daemon main window.
- [x] Rework the daemon main-window top details rail from the latest visible
      build feedback. The current `Config: Not configured` field is ambiguous:
      the UI must show the effective config file path, just like the `Log` file
      path, as selectable text with copy/open affordances. Long config/log paths
      may wrap or use a bounded visual abbreviation, but must not be reduced to
      an unselectable ellipsis-only fragment. A normal default configuration
      must not display `Not configured`; show the default config file path that
      would be read or written, or omit the field from the first row and explain
      the deployment-specific reason in diagnostics/details.
      The details rail must also reduce the width allocated to compact metadata
      such as `Version` and `Uptime`, keep diagnostic values selectable and
      copyable, and reserve the widest flexible diagnostic area for `Last
      Error` so degraded/down states can show useful error text without being
      squeezed by low-priority metadata. `Last error` must be selectable,
      copyable, and able to wrap; an ellipsis-only or copy-only error display is
      not acceptable. The
      top metrics must also rename the ambiguous `Running` counter to an
      explicit task-execution label such as `Active Tasks`, `Tasks Running`, or
      `Running Tasks`; bare `Running` is not acceptable because it can be
      confused with daemon health. Current evidence: `npm run smoke:ui` proves
      effective config/log paths render without `Not configured`, path text is
      selectable and not wrapped by a button, separate copy buttons are present,
      Config/Log paths and `Last error` values are selectable and copyable,
      `Last error` can wrap instead of being an ellipsis-only fragment,
      `Last error` receives the widest details column, the details rail remains
      attached to the compact status strip, and the metric label is `Active
      Tasks` rather than bare `Running`.
- [x] Rework the daemon main-window connected-document list so each document is
      a compact summary block, not a verbose session/activity panel. The block
      should primarily show the document name plus a middle-truncated session
      ID, a user-facing state of only `active` or `dead`, host/add-in version,
      available tool count, queue depth, finished task count, and failed task
      count. The document title has first layout priority and must not be
      squeezed by status badges or right-side metadata. Do not label a stale or
      disconnected session as `reconnecting` in the compact document card unless
      the daemon owns explicit evidence that a live reconnect attempt is
      currently in progress; otherwise stale/disconnected/unavailable sessions
      are `dead`. Remove inline recent-command history and expanded host
      metadata from the default document block because the right-side
      Activity/Inspector pane already owns command history and diagnostics. The
      compact card must show a bounded session ID, while full session IDs remain
      available through copy, tooltip, or Inspector affordances. Current
      evidence: `npm run smoke:ui` validates compact document cards with title,
      `active`/`dead` labels without reconnect wording, session ID on dedicated
      copyable lines, version, tool count, queue depth, finished/failed counts,
      no inline activity history, and preserved title width.
- [x] Fix add-in runtime identity correlation so connected Word, Excel, and
      PowerPoint document cards preserve registered host metadata. A real Word
      session must show app `word` and the Office host version when the add-in
      sent those fields during `register`; it must not fall back to `other` or
      `Version -` because `session.added` was correlated to the wrong runtime
      instance. Add regression coverage for daemon register -> session.added ->
      UI snapshot flow, including daemon-assigned instance IDs returned by the
      register response. Current evidence:
      `register_assigned_instance_id_preserves_host_metadata_in_ui_snapshot`
      parses the daemon register response, sends `session.added` with the
      assigned instance ID, and verifies the UI snapshot preserves `word` and
      host version `16.0`.
- [x] Add scan-friendly document state indicators in the daemon main window.
      Compact document cards must use only `active` and `dead` user-facing
      labels with a stable icon, status dot, or host-accent-compatible color
      marker next to the text label so users can identify dead sessions quickly
      when many documents are listed. Color alone is not enough; keep the text
      label for accessibility and tests. Raw internal states such as `stale`,
      `disconnected`, or future reconnect diagnostics may appear in the
      Inspector, but must not replace the compact card's `dead` label. Current
      evidence: `npm run smoke:ui` validates a visible state marker next to the
      compact card state label and proves stale sessions render as `Dead`
      without `stale` or `reconnecting` copy.
- [x] Enforce closed-session cleanup across daemon runtime state. A closed,
      disconnected, stale, or otherwise dead Office document session may remain
      visible only during the reconnect grace window. The default grace is 60
      seconds and the effective value must be capped at 300 seconds even when
      config or environment variables request more. After the grace expires,
      remove the session from the registry so it disappears from
      `office.list_sessions`, `office.get_session_info`, daemon document cards,
      and tray document counts. Do not keep historical session cards; command
      history is the history surface. Current evidence: registry pruning is
      covered by `stale_sessions_are_pruned_after_grace_period`, config capping
      by `session_grace_is_capped_at_five_minutes`, shared-state pruning by
      `shared_state_prunes_stale_sessions_after_configured_grace`, and the
      request-independent runtime cleanup loop by
      `run_once_prunes_expired_stale_sessions_without_client_request`.
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
      daemon. `--no-tray` is not an accepted local/manual startup path because
      users need the tray icon to know the daemon is running; `--with-tray`
      remains accepted for compatibility. `ui.production_daemon_tray`
      now proves the production path creates the native Windows tray host,
      exposes live UI state, and logs native notification-area icon creation.
      Manual Windows interaction evidence remains required for right-click menu
      visibility.
- [x] Wire tray `Show Office MCP Control` to the same UI-opening path as
      `office-mcp-daemon ui`. Covered by `open_ui_from_runtime` and the native
      tray menu action dispatch.
- [x] Wire tray `Quit Office MCP Control` to graceful daemon shutdown and visible menu
      confirmation. Covered by `stop_daemon`, `confirm_quit`, and
      `native_tray_quit_uses_platform_confirmation_dialogs`.
- [ ] Add manual/e2e evidence that verifies a visible tray icon exists on
      Windows, the right-click menu appears, and `Show Office MCP Control` opens the
      daemon UI. Automated evidence now covers native tray icon creation logs
      and the live menu model through `ui.production_daemon_tray`; this item
      remains open until a visible desktop interaction run verifies the icon and
      right-click menu in the notification area. The manual evidence recorder
      `npm run evidence:record-tray-manual` and validator gate
      `npm run evidence:validate-ui -- --require-manual-tray` now exist so the
      final desktop run can be captured as structured release evidence instead
      of an untracked note. The final run must pass `--daemon-bin` to the
      recorder so the artifact also captures `daemon status` and live
      `tray --probe` output for the same daemon instance that was visually
      inspected; the probe must read live UI state from that daemon and expose
      the expected product tooltip and tray menu snapshot. The recorder marks
      evidence failed and the validator rejects required manual tray evidence
      when it is not bound to that daemon context. Screenshot artifacts must be
      real, complete image files; truncated image headers are rejected by both
      the recorder and validator. Manual tray evidence records freshness metadata
      for every tray screenshot surface; stale screenshots are rejected
      by the recorder, the manual tray validator, and the embedded product visual
      evidence gate so an old tray/menu capture cannot satisfy a new release run.
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
- [x] Arrange task pane top-block metadata into separate Document and
      Daemon/Session groups. Document fields stay together, while daemon
      endpoint, server/protocol version, session ID, and last connection error
      stay together near the `MCP Control` title column instead of being mixed
      into the document rows. Covered by Word/Excel/PowerPoint task pane
      contract tests for `.document-metadata` and `.daemon-metadata` order.
- [x] Remove the separate task pane endpoint error paragraph. Endpoint
      validation and reconnect failures now use the existing `Last Error` row
      instead of reserving a duplicate `endpointError` element that disrupts the
      compact daemon metadata layout. Covered by Word/Excel/PowerPoint task
      pane contract tests rejecting `endpointError` and `.field-error`.
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

### M6.5.2 — Add-in task pane density and state accuracy follow-up

User-reported follow-up from live Excel task pane testing:

- [x] Merge the `Available Tools` and `Tool Permissions` surfaces for Word and
      Excel into one grouped tools surface, unless implementation evidence shows
      a separate read-only summary is materially clearer. The merged surface
      must avoid repeating the same tool names, descriptions, and counts in two
      blocks. Word and Excel now use one `Tools` surface that exposes the same
      categorized rows for inspection and per-tool permission changes.
- [x] Group tool permission rows by the same capability categories as available
      tools. Each category must support independent expand/collapse, show
      `Enabled X of Y` while collapsed, and preserve per-tool saved permission
      state. Contract tests cover categorized collapsible tool controls and
      persisted per-tool permission updates.
- [x] Replace the current settings-card interaction with compact inline editing
      inside the top summary block. Clicking the gear should reveal editable
      rows for daemon endpoint and tool permissions without creating a large
      separate framed settings block that consumes the task pane viewport. The
      settings panel now contains only the endpoint edit row; tool toggles stay
      in the nearby `Tools` surface and become edit-enabled when settings opens.
- [x] Combine server version and protocol version into one metadata row in Word
      and Excel task panes, for example `Server 0.1.0 / Protocol 1.0`.
- [x] Fix document state reporting so normal editable workbooks/documents do not
      show `Dirty: unknown / Read-only: unknown`. When Office APIs can determine
      editability, show `Editable`, `Read-only`, `Protected`, or a more specific
      host state. When a dirty/saved signal is not reliably exposed by the host,
      omit it from the primary summary or label it as diagnostic-only instead
      of presenting `unknown` as the user-facing state. Word and Excel now show
      `Editable`, `Editable, unsaved changes`, `Read-only`, or `Protected...`
      instead of primary `unknown` dirty/read-only text.
- [x] Add Word and Excel task pane contract tests covering grouped collapsible
      permissions, no duplicated tools block, inline settings editing, merged
      server/protocol metadata, and non-unknown editable document state for the
      normal mocked editable host case.
- [ ] Capture live Excel evidence after implementation showing a connected
      workbook with the compact top block, merged tools/permissions surface,
      inline settings edit rows, combined server/protocol row, and a concrete
      editable/read-only/protected state. The visual portion must be recorded
      with `npm run evidence:record-product-visual` using the Excel task pane
      density flags and then validated with `--require-product-visual`.
      Product visual evidence now also has to bind to a passed
      `artifacts/runtime-evidence-excel.json` smoke run, proving the screenshot
      is tied to an active Excel session, workbook title, available tool count,
      marker readback, and the read/write/formula/format/table/chart/sheet smoke
      operations. This item remains open until the runtime-bound evidence also
      includes the live Excel task pane screenshot.

**Exit criterion**: In a 320 px Office task pane, a connected editable Excel
workbook does not show duplicated tool lists or `unknown` document state; tools
and permissions share one categorized collapsible surface; settings edit inline;
and server/protocol metadata fits on one row.

### M6.5.4 — Task pane tools permission polish follow-up

User-reported follow-up from the current Word task pane build:

- [x] Add a top-level capability mode control to the `Tools` surface with three
      states: `Read`, `Write`, and `All`. `Read` enables only read-only tools;
      `Write` enables read tools plus mutating non-destructive tools; `All`
      enables every supported tool, including delete/destructive operations.
      The selected mode updates the effective `available_tools` for the current
      session and still allows individual per-tool toggles inside the selected
      ceiling.
- [x] Simplify every tools count from `Enabled X of Y` to `X/Y`. Category rows
      must right-align the count and place it immediately before the category
      toggle. The aggregate tools header may also use `X/Y`.
- [x] Fix the `Daemon` metadata row height in Word, Excel, and PowerPoint task
      panes. The endpoint row must vertically center with the other document
      metadata rows; the reconnect control is an icon-only button and must not
      make the row card-height.
- [x] Shorten the constrained task-pane title/category text from
      `Office MCP Control` to `MCP Control` where it is used as the panel title
      or local category label. Keep the full product name for catalog,
      installer, tray tooltip, and other wider identity surfaces. Add the
      generated product mark to visible `Open Control Panel` buttons.
- [x] Replace the current Word tool UI grouping by action buckets
      (`Read`/`Insert`/`Edit`) with object-owner categories from
      [04-word-capabilities.md](04-word-capabilities.md): Document & structure,
      Range & selection, Paragraphs & lists, Tables, Media, Content controls,
      and Review. Side-effect level remains separate metadata used by the
      `Read`/`Write`/`All` permission mode.
- [x] Prevent category collapse when clicking category toggles or per-tool
      toggles. Only the explicit category disclosure hit target changes expanded
      state; toggle clicks update permissions only.
- [x] Add task pane contract tests for the new count placement, `Read`/`Write`/
      `All` top-level mode, non-collapsing toggles, compact daemon row, `MCP
      Control` local title, `Open Control Panel` icon, and object-owner Word
      categories across Word, Excel, and PowerPoint where applicable.

**Exit criterion**: At 320 px width, the tools surface shows object-owner
categories, compact right-aligned counts, a top-level permission mode, and
independent non-collapsing toggles; the daemon row is vertically centered and no
long title or text-only control-panel button wastes scarce task-pane space.

Implementation evidence: `feat: polish task pane tool controls` added the
shared `Read`/`Write`/`All` mode UI, compact `X/Y` counts, object-owner Word
groups, compact PowerPoint daemon row, and Word/Excel/PowerPoint task pane
contract tests. Verification passed with `npm test` in each Office add-in and
`cargo test -p office-mcp-daemon static_response`.

### M6.5.5 — Daemon-wide tool access control

User-reported access-control follow-up: tool permissions must not live only in
each Office task pane session. The daemon needs one canonical global tool access
policy that controls discovery and dispatch for every MCP client and every
Word, Excel, and PowerPoint session.

- [x] Add a daemon-owned tool access policy store. It must support a global all
      tools switch, app-level switches for Word/Excel/PowerPoint, category-level
      switches, individual tool switches, and a `Read` / `Write` / `All` access
      ceiling. The policy must persist through daemon restart and be visible in
      the daemon UI. Current implementation loads `[tool_access]` at daemon
      startup, keeps an immediately mutable runtime policy, exposes it in
      `/ui/state`, updates it via `PUT /ui/tool-access-policy`, and writes that
      UI update back to the daemon config file so restart applies the same
      policy.
- [x] Classify every Word, Excel, and PowerPoint tool by app, object/category,
      and side-effect level so the daemon UI can display grouped controls. The
      UI must allow expanding/collapsing categories and independently toggling
      every app/category/tool without collapsing the group accidentally.
      Current evidence: `tool_metadata_covers_every_forwarded_office_tool`
      proves every forwarded Word, Excel, and PowerPoint tool has daemon UI
      metadata, while `npm run smoke:ui` validates grouped app/category/tool
      controls and toggle clicks that do not collapse groups.
- [x] Filter MCP `tools/list` through the daemon global access policy. Disabled
      tools must not be returned to clients, reducing context and preventing
      stale planning from tools the daemon will reject. Management tools such as
      `office.list_sessions` and `office.get_session_info` remain available
      unless a future explicit management-policy section says otherwise.
      Current evidence: `tools_list_filters_daemon_disabled_tools`,
      `tools_list_filters_by_daemon_access_mode`, and
      `mcp_json_rpc_daemon_policy_filters_discovery_and_rejects_before_session_preflight`.
- [x] Enforce daemon policy before session preflight on every `tools/call`.
      When a stale client calls a globally disabled tool, return
      `TOOL_NOT_AVAILABLE` with text beginning `Tool <name> is disabled by
      daemon access policy.`, telling the client to refresh `tools/list`, and
      structured `refresh_tools: true`. Do not forward the request to any add-in
      session. Current evidence:
      `mcp_json_rpc_daemon_policy_filters_discovery_and_rejects_before_session_preflight`
      verifies the daemon-policy rejection and refresh hint before any session
      preflight.
- [x] Keep session `available_tools` as the second access-control layer. After
      daemon policy passes, if the target session does not expose the tool,
      return `TOOL_NOT_ENABLED_FOR_DOCUMENT` with the target `session_id`, text
      beginning `Tool <name> is disabled for this document session.`, a refresh
      instruction for `office.get_session_info` or `office.list_sessions`, and
      structured `refresh_session_info: true`. Current evidence:
      `mcp_json_rpc_disabled_session_tool_uses_unified_error_wording` and
      `invocation_preflight_returns_protocol_errors` verify the second-layer
      session preflight code, wording, and refresh hint.
- [x] Add daemon UI controls for the global policy. The UI must show app,
      category, and tool controls with `Read` / `Write` / `All` mode, and it
      must make clear that daemon policy is global while task-pane
      `available_tools` remains document/session-specific. Current control
      panel implementation renders the daemon-owned tool metadata from
      `/ui/state.daemon.tool_catalog`, groups controls by app and category,
      exposes `Read` / `Write` / `All`, and updates the live daemon policy via
      `PUT /ui/tool-access-policy`.
- [x] Keep Global Tool Access out of the periodic auto-refresh render path so
      user-owned editing state is stable while daemon telemetry updates.
      Automatic polling must not rebuild the access-control tree or reset
      expanded/collapsed app and category groups; the panel reloads only on
      initial load, explicit access-policy updates, or a real tool-catalog shape
      change.
- [x] Add tests covering discovery filtering, daemon-policy dispatch rejection,
      session preflight rejection, refresh hints, persistence, and UI rendering.
      Tests must prove a disabled tool disappears from `tools/list`, a stale
      call gets `TOOL_NOT_AVAILABLE`, and a session-missing tool gets
      `TOOL_NOT_ENABLED_FOR_DOCUMENT` rather than the older generic capability
      failure. Tests must also assert both access-control messages use the
      unified `Tool <name> is disabled...` wording rather than `not enabled`.
      Current evidence covers daemon catalog filtering, daemon-policy dispatch
      rejection, session preflight rejection wording and refresh hints,
      config-file startup and UI write-back persistence, and browser-level UI
      rendering/toggle interaction.

**Exit criterion**: A user can manage tool access once in the daemon UI across
all sessions, MCP clients see only daemon-allowed tools from `tools/list`, and
every call is enforced first by daemon policy and then by the target session's
`available_tools`, with actionable refresh hints when either layer rejects the
call.

Implementation progress: the first daemon-policy runtime slices add an
in-memory `ToolAccessPolicy` owned by the daemon runtime state. MCP
`tools/list` is now filtered through that policy, and `tools/call` rejects a
globally disabled forwarded tool before session lookup with
`TOOL_NOT_AVAILABLE`, `refresh_tools: true`, and the unified `Tool <name> is
disabled...` wording. The policy now classifies Word, Excel, and PowerPoint
tools by app, object/category, and read/mutating/destructive side effect, and
supports `Read` / `Write` / `All` ceilings plus app, category, and individual
tool switches for discovery filtering and dispatch enforcement. The daemon
config file now persists startup policy under `[tool_access]`, runtime state
applies that policy after restart, and UI updates write the same policy back to
the config file. `/ui/state` now exposes the daemon policy and daemon-owned tool
metadata so the control panel can render the global access state. `PUT
/ui/tool-access-policy` updates both the in-memory daemon policy for immediate
UI and MCP discovery changes and the durable config section for restart
consistency.

### M6.5.3 — Product identity, add-in metadata, and native tray polish

User-reported follow-up from live Word add-in and tray testing. This milestone
owns the product identity layer, not protocol behavior: the add-in and tray must
feel like a finished local desktop utility rather than an experimental scaffold.

Latest UI feedback to preserve in implementation planning:

- The product mark must not use Microsoft Office logos or near-logo variants,
  but it must still read as office productivity under local user control. The
  preferred direction is a mature, slightly futuristic control-surface mark:
  document/window panes, routing geometry, layered surfaces, and an explicit
  operator/control affordance that remains legible at tray and ribbon sizes. It
  must not read as a generic settings/file/debug/AI icon. The latest visual
  direction is future office control: precise geometry, layered operational
  surfaces, and restrained control affordances without Office-owned app marks,
  Office tile silhouettes, terminal/debug motifs, neon effects, or decorative
  AI styling.
- The add-in first-run surface must be treated as one product impression. Title,
  icon, provider, description, ribbon command, catalog type/category, and task
  pane chrome must all change together; a polished title paired with a missing
  icon or experimental type/category is still a release failure. The installed
  Word, Excel, and PowerPoint catalogs must all show the same mature identity
  after a fresh install; disappearing catalog entries or stale generated
  manifests are installer bugs. The add-in must not look like a sideloaded
  experiment before the task pane opens: catalog title, catalog type/category,
  icon, ribbon command, provider, and task pane chrome must read as one finished
  local control product.
- The tray must look like native desktop software. A missing/default tray icon,
  non-native-looking right-click surface, webview/HTML imitation menu, or debug
  wording makes the product feel unfinished and must remain blocked until live
  Windows visual evidence proves a real native notification-area icon and native
  context menu. The normal Windows launch must show a deliberate tray glyph,
  native tooltip, right-click menu anchored to the actual notification-area
  icon, OS-native spacing/hover/keyboard behavior, and native quit confirmation.

Latest design feedback from the current visible build:

- The logo needs a deliberate redesign/re-approval pass. It must not use Office
  logos, but it must visibly communicate office productivity and user control
  with a mature, slightly futuristic feel. A generated asset that only reads as
  a generic document, settings mark, debug badge, or placeholder app icon is not
  acceptable.
- The add-in's title, icon, and type/category must be corrected together. The
  first installed Office surface must look like mature software rather than an
  experimental or half-finished sideload package.
- The tray still needs product-quality native presentation. The normal daemon
  launch must show the product icon, and the right-click menu must be OS-native,
  not a custom or web-styled menu surface.

Latest visible-build feedback to carry into the next implementation goal:

- Design and review a product logo that avoids Microsoft Office marks while
  still communicating office productivity plus deliberate local control. The
  direction should feel futuristic through precise control geometry, layered
  document/workspace surfaces, and restrained depth, not through generic neon,
  AI, settings, debug, or placeholder icon language.
- Rework add-in identity as a first-contact software surface, not a collection
  of manifest fields. Title, icon, catalog-visible type/category, provider,
  description, ribbon command, and task pane chrome must all read as one mature
  product for Word, Excel, and PowerPoint after a clean install.
- Rework tray first impression as native desktop software. A normal daemon
  launch must show the accepted product glyph in the Windows notification area,
  and right-click must open a native menu from that icon. A missing/default
  icon, non-native menu, webview/CSS popup, toolkit-demo surface, or prototype
  wording keeps the tray task open.

Current reported issues to carry as TODO requirements:

- The logo needs another design-quality pass, not only an asset-generation pass.
  It must avoid Office logos and near-logo shapes, but it still needs a clear
  office-and-control feeling with a mature, futuristic product tone.
- The add-in title, icon, and visible type/category must be corrected together.
  These fields are one first-contact surface and must make Word, Excel, and
  PowerPoint look like finished software rather than a test add-in or prototype.
- The tray must be corrected as a native Windows product surface. It needs the
  accepted product icon, product tooltip, OS-native right-click menu, native menu
  behavior, and native quit confirmation; custom/web-styled menus are blockers.

Latest current-build feedback to carry as TODO requirements:

- Design or materially refine the logo so it clearly says future office control:
  office/document work under explicit user control, using precise layered panes,
  routing/command geometry, and an operator affordance. It must not use Office
  logos, Office app tiles, Microsoft 365 color ownership, host-app silhouettes,
  gear-only symbolism, debug/terminal motifs, or decorative AI/neon styling.
- Re-audit the add-in first-run surface from the user's point of view. The
  title, icon, provider, description, ribbon command, task pane chrome, and
  catalog-visible type/category must look like mature software, not an Office
  scaffold or experimental sideload.
- Re-audit the tray from a normal Windows launch. The tray must have the final
  product icon and an OS-native right-click menu. A missing/default icon or a
  custom/web-styled menu keeps the tray work open even if automated probes pass.

Current screenshot feedback to preserve for the next implementation goal:

- Redesign or materially refine the logo as a product-design deliverable. The
  mark must not use Office logos or Office-owned visual language, but it must
  still make office productivity and local user control obvious. The target
  impression is mature, slightly futuristic desktop control software, not a
  generic add-in, settings gear, debug badge, or experimental placeholder.
- Correct the add-in first-contact identity in one pass: title, icon, provider,
  description, ribbon command, task pane chrome, catalog-visible type/category,
  and generated catalog metadata. The Office catalog and ribbon must look like a
  finished product before the task pane opens.
- Correct the tray as a native product surface. A visible product icon is
  required in the Windows notification area, and right-click must open a real
  OS-native menu anchored to that icon. A blank/default icon, missing icon,
  webview/CSS/custom-drawn menu, or non-native-looking menu keeps the tray item
  open.

- [ ] Redesign or re-approve the logo from the current visible user surfaces.
      The review must answer the user's current complaint directly: the logo
      cannot use Office logos, but it must still visibly communicate office
      productivity and control with a mature, slightly futuristic feel. Passing
      generated-asset tests is not enough. If the icon reads as a generic add-in,
      settings mark, debug badge, placeholder tile, AI decoration, or legally
      distinct but weak office-control metaphor at tray/ribbon/catalog sizes,
      redesign the source mark and regenerate the asset set before closing this
      item. This item can close only after the final mark is reviewed in the
      real add-in catalog, ribbon, task pane chrome, tray, installer, and daemon
      title-bar surfaces or in release-equivalent rendered evidence tied to the
      current asset fingerprints.
- [ ] Produce and review a new future-office-control logo direction if the
      current mark still feels generic in the latest visible surfaces. The
      design must communicate office productivity plus deliberate control
      without using Office logos, Office-like tiles, Microsoft app-color
      ownership, host-app silhouettes, gear-only settings symbols, terminal or
      debug motifs, or decorative AI/neon effects. The accepted mark must be
      judged at tray, ribbon, catalog, task-pane, daemon title-bar, and
      installer sizes before implementation closes this item.
- [ ] Rework the add-in title, icon, and visible type/category as one mature
      installed-software surface for Word, Excel, and PowerPoint. The catalog
      card, ribbon command, task pane chrome, provider, description, and
      generated catalog metadata must all present `Office MCP Control` as a
      finished local productivity/control product. Generic `Add-in`/`Task Pane`
      type text, default/missing icons, raw MCP/protocol wording, host package
      names, or experimental sideload presentation keep this item open. The
      install/catalog scripts must be fixed as part of this item if any host
      catalog disappears, points to stale assets, or renders the wrong title,
      icon, provider, description, or type/category after reinstall.
- [ ] Re-audit and correct the add-in first-contact package after a clean
      install. `DisplayName`, visible catalog title, icon, provider, short
      description, ribbon group, primary command, task pane title/chrome, and
      catalog-visible type/category must be reviewed as one surface. The result
      must look like a finished local productivity/control product, not a
      default Office task-pane add-in, generic `Add-in`, experimental sideload,
      raw MCP/protocol bridge, or half-finished package.
- [ ] Rework the tray first impression until it looks native and product-grade
      in a normal Windows daemon launch. The tray must show the accepted product
      glyph, product tooltip, OS-native right-click menu anchored to the visible
      notification-area icon, native separators/disabled rows, keyboardable menu
      behavior, and native quit confirmation. A custom-drawn, webview, CSS,
      frameless, or toolkit-demo-looking menu fails this item even when actions
      work. This item cannot close from `tray --probe` alone; it requires live
      interactive Windows evidence showing the icon and native menu opened from
      that exact notification-area icon.
- [ ] Rework the tray icon/menu implementation until a normal Windows launch
      looks native and installed. The notification area must show the final
      product glyph, and right-click must open a real OS-native menu from that
      icon. A missing icon, default/framework icon, blank square, custom-drawn
      menu, webview/HTML/CSS popup, frameless utility panel, toolkit demo menu,
      or non-native-looking menu keeps this item open regardless of functional
      menu actions.
- [x] Run a full product logo concept pass before accepting the final mark. The
      pass must produce multiple original concept directions that communicate
      office productivity plus local control without Office-owned marks, Office
      app-color ownership, generic document thumbnails, gear-only symbols, or
      developer/debug motifs. The accepted design must feel like mature,
      slightly futuristic desktop software and must be reviewed at real tray,
      ribbon, catalog, task pane, installer, and title-bar sizes before this
      item can close. Current evidence: `brand-design.md` documents the
      selected `Command Console Panes` direction plus rejected `Orbiting
      Document Hub` and `Shielded Automation Badge` directions;
      `record-rendered-logo-review.mjs` emits the concept-pass review into the
      rendered-size artifact; `record-product-visual-evidence.ts` and
      `validate-runtime-evidence.ts` reject product visual evidence that omits
      the concept pass. Verified with `npm run check` in
      `src/office-ctl/word` and `src/office-mcp/daemon/evidence`.
- [x] Rework and verify all add-in first-impression metadata as one product
      surface for Word, Excel, and PowerPoint after a clean install. The visible
      add-in title, icon, provider, description, ribbon command, task pane
      title/chrome, catalog title, and catalog type/category must read as one
      mature local productivity automation/control product. A host still showing
      a missing/default icon, raw package name, generic `Add-in`/`Task Pane`
      type, sample/debug wording, or stale catalog entry keeps this item open.
      Current static clean-install evidence: `register-office-catalog.ps1`
      stages fresh Word, Excel, and PowerPoint catalog manifests with synced
      daemon origin and generated icon URLs, and
      `record-catalog-identity-review.mjs` emits a structured
      `catalog_identity_review` artifact proving display name, provider,
      description, command label, task pane URL, generated icon URLs, shared
      origin, and product type/category are coherent. The GitHub Release
      workflow now generates `catalog-identity-review.json` from the staged
      clean catalog, rejects non-ready reviews, publishes the report as a
      release artifact, and includes it in `SHA256SUMS`. Word tests cover both
      the passing clean catalog and prototype-metadata rejection. Final live
      Office catalog/ribbon/task pane visual re-audit remains tracked by the
      live first-contact evidence items below.
- [ ] Rework and verify the Windows tray first-impression surface as native
      desktop software. The normal daemon launch must show the generated product
      glyph in the notification area, product-consistent tooltip text, a native
      right-click menu opened from that exact icon, native separators and
      disabled status rows, keyboardable native menu behavior, and a native quit
      confirmation. Any webview/HTML/CSS/custom-drawn or toolkit-demo-looking
      menu surface keeps this item open even when the daemon probe passes.
- [ ] Re-review the tray visual implementation from a normal Windows launch,
      not only from model/probe output. The tray must have a real product icon,
      native tooltip, native Windows context menu, platform spacing/hover/keyboard
      behavior, and native quit confirmation. Missing icon, framework sample icon,
      blank placeholder, HTML/webview/floating custom menu, or non-native-looking
      right-click surface is a release blocker because it makes the product feel
      unfinished.
- [x] Redesign or formally re-approve the product logo against the latest
      feedback. The accepted mark must not use Office logos or near-logo
      variants, but it must visibly communicate office productivity plus local
      user control with a mature, slightly futuristic control-surface feel. The
      task must produce source artwork, generated icon assets, a short design
      rationale, rendered-size review images, and explicit rejection of generic
      file/settings/debug/AI-only marks. The review must judge the icon from the
      sizes users actually see in Office catalogs, ribbon commands, task pane
      chrome, Windows tray, installer metadata, and the daemon title bar. The
      design review must explicitly answer why the icon feels like future office
      control without borrowing Office logos, Office tile language, Microsoft
      app-color identity, terminal/debug metaphors, or decorative AI/neon
      styling. Re-approval is documented in
      `src/office-ctl/common/assets/brand-design.md` and the rendered review
      generator. Current evidence: `npm test -- --test-name-pattern "brand
      design|rendered-size logo|product identity"` in `src/office-ctl/word`,
      `npm test -- --test-name-pattern "rendered logo|product visual evidence"`
      in `src/office-mcp/daemon/evidence`, and
      `node src/office-ctl/common/scripts/record-rendered-logo-review.mjs
      --output artifacts/logo-rendered-size-review.json --sheet
      artifacts/logo-rendered-size-review.png`. The review now explicitly
      rejects Word document silhouettes, Excel grid marks, PowerPoint slide
      silhouettes, Outlook envelope marks, Office tile language, Microsoft 365
      gradients, Windows logo conventions, placeholder initials, gear-only
      artwork, terminal/debug motifs, and decorative AI/neon styling.
- [ ] Conduct a final product-design review of the logo in the actual user
      surfaces before release. The mark must feel like future office control
      without using Office logos: office/document work must be recognizable,
      control/routing must be visible, and the style must read as mature software
      rather than a placeholder, sample icon, settings gear, generic document,
      developer/debug badge, or loud decorative graphic. This review must inspect
      the add-in catalog, ribbon, task pane chrome, tray, installer, and daemon
      title-bar renderings.
- [ ] Re-open the logo acceptance review against the latest visible-build
      feedback. The reviewer must judge whether the current mark actually feels
      like futuristic office control in real product surfaces, not merely whether
      it is legally distinct from Microsoft marks. If the icon reads as a
      generic add-in, file, settings, debug, AI, or placeholder product mark at
      tray/ribbon/catalog sizes, redesign it before closing this item.
- [x] Redesign or materially refine the current logo if the live rendered
      surfaces still fail the future-office-control brief. The implemented
      refinement keeps the selected `Command Console Panes` direction but
      redraws the mark as three layered control-console document/app panes, a
      command rail, operator nodes, and a yellow control reticle. It avoids
      Office-owned logos and tile language while more clearly suggesting office
      work, command routing, and local user control at 16 px tray size, 32 px
      ribbon size, catalog thumbnail size, title-bar size, and installer size.
      Current evidence: `brand: refine office control mark`, regenerated
      `brand-mark.svg` and icon PNGs, `brand-design.md`,
      `generate-brand-assets.mjs`, `record-rendered-logo-review.mjs`,
      `npm run check` in Word/Excel/PowerPoint add-ins,
      `npm test -- --test-name-pattern "rendered logo|product visual evidence"`
      in `src/office-mcp/daemon/evidence`, and Rust
      `cargo test -p office-mcp-daemon product_tray_icon static_response` run
      as separate filtered commands. Final real Office/tray user-surface review
      remains tracked by the live visual evidence items.
- [ ] Re-audit every add-in first-contact surface after install: title, icon,
      provider, description, ribbon command, task pane chrome, and catalog
      type/category for Word, Excel, and PowerPoint. The installed catalog must
      look like mature product software, not an experimental add-in, sample,
      debug panel, protocol bridge, or raw package. The install/catalog scripts
      must be fixed whenever a host entry disappears, loses its icon, points at
      a stale daemon origin, or shows mismatched title/type metadata. This item
      includes the generated Word catalog, which must remain present and show
      the correct product title, icon, description, provider, and local
      productivity automation/control type after reinstall. The review must be
      performed from Office's real installed surfaces, not only manifest XML, so
      a valid manifest can still fail when Office renders the title, icon, or
      category as a half-finished sideloaded add-in.
- [ ] Rework add-in title, icon, and catalog-visible type/category as one
      coherent product identity update for Word, Excel, and PowerPoint. The
      final installed Office catalog, ribbon command, and task pane chrome must
      expose the same mature product name, generated product icon, provider,
      description, and local productivity automation/control category. Generic
      Office container labels, missing/default icons, raw host package names,
      sideload/debug wording, or mismatched per-host metadata keep this item
      open even if all manifests validate.
- [ ] Re-audit the add-in title, icon, and visible type/category from the user's
      first installed view, not only from source manifests. Word, Excel, and
      PowerPoint must present a finished software identity before the task pane
      opens. The Office catalog and ribbon must not show generic `Add-in` or
      `Task Pane` categories, raw MCP/protocol names, host scaffold names,
      missing/default icons, or any copy that looks like an experiment.
- [ ] Fix add-in product identity as one installable-software surface wherever
      Office exposes it. `DisplayName`, Office catalog title, icon, provider,
      short description, ribbon group, primary command, task pane title/chrome,
      and catalog-visible type/category must be changed and verified together for
      Word, Excel, and PowerPoint. Generic visible types such as `Add-in` or
      `Task Pane`, default/missing icons, raw package names, debug/protocol labels,
      or host-specific scaffold wording keep this item open even when the manifest
      validates.
- [x] Replace prototype add-in identity copy everywhere it can appear in Office:
      manifest `DisplayName`, provider, description, support metadata, ribbon
      group, primary command label, task pane title/chrome, catalog title, and
      catalog type/category. The release identity should read as a finished
      local desktop control product, not a lowercase package slug, raw MCP
      protocol adapter, internal scaffold, or generic `Task Pane` add-in. The
      add-in title, icon, and type/category must be changed and validated
      together; updating only one of those fields is not enough. Current static
      implementation evidence covers Word, Excel, and PowerPoint manifests,
      task pane chrome, generated product icons, catalog metadata, hosted
      manifest/AppSource metadata, and Windows catalog staging. Evidence:
      `npm run check` in `src/office-ctl/word`, `src/office-ctl/excel`, and
      `src/office-ctl/powerpoint`, plus `npm test` in `packaging`. Final live
      Office catalog/ribbon screenshots remain tracked by the first-run visual
      evidence items below.
- [ ] Replace any tray placeholder/default icon and non-native-looking menu
      surface in the normal daemon launch path. Right-click must open a real
      OS-native notification-area context menu with native separators,
      disabled/read-only status rows, hover/selection behavior, keyboard access,
      tooltip text, `Show Office MCP Control`, `Quit Office MCP Control`, and native quit
      confirmation. A webview, frameless HTML panel, CSS popup, or toolkit-demo
      menu fails this item even if automated probes pass. The visual evidence
      must prove the menu is opened from the visible tray icon and looks native
      beside normal Windows tray applications, not like an embedded web UI.
- [ ] Implement and verify the tray as a polished native Windows first-contact
      surface. The visible tray icon must be the accepted product glyph, not a
      blank/default/framework icon, and right-click must open a native Windows
      context menu anchored to that icon. Any menu drawn by HTML/CSS, a webview,
      a frameless floating panel, or a toolkit-demo-looking surface keeps this
      item open. The evidence must come from a normal interactive daemon launch,
      not only from model/probe output.
- [ ] Replace any tray implementation path that draws its own menu surface. The
      Windows build must use native notification-area and menu primitives for
      the visible user interaction, including the right-click menu opened from
      the actual tray icon and the quit confirmation dialog. Automated probe
      output is not enough; this stays open until live Windows evidence proves
      the menu is visually native.
- [ ] Re-audit the normal Windows tray launch from the user's first impression.
      The tray must show the accepted product icon, native tooltip, and
      OS-native right-click menu anchored to the notification-area icon. Any
      blank/default/framework icon, non-native-looking menu, webview/CSS popup,
      toolkit-demo surface, or prototype wording must be fixed before release.
- [x] Design an original office-control product logo and icon system. It must
      avoid Microsoft Office product marks while still communicating office
      productivity, local automation, and user control with a mature, slightly
      futuristic desktop utility style. The first implemented identity is the
      Office MCP Control mark: abstract document panes plus a control node,
      generated from `src/office-ctl/common/assets/brand-mark.svg`.
- [x] Review and refine the implemented logo against the product brief before
      final release. The final mark communicates office productivity plus user
      control through abstract document panes, routing, and a control node; it
      avoids Office logos, Office tile language, Microsoft 365 gradients, Word
      document silhouettes, Excel grid marks, generic gear icons, and text-only
      initials. `src/office-ctl/common/assets/brand-design.md` documents the
      metaphor, palette, minimum sizes, and non-Microsoft distinction, with
      product identity tests guarding the brief.
- [x] Add an explicit brand guard so generated/source assets cannot regress to
      Microsoft Office-owned visual marks, placeholder initials, generic gear
      icons, or single-color debug glyphs. Covered by
      `product-identity.test.mjs`, which verifies the source logo exists, PNG
      app icons are non-empty and multi-color, and Word/Excel manifests do not
      reference Office product logos, placeholder labels, generic `Open`, or
      debug/prototype wording as the add-in identity.
- [x] Add reproducible brand assets: source artwork, generated 16/20/24/32/48/
      64/128/256 px app icons, Office add-in command icons, and a monochrome
      tray/menu-bar glyph that remains legible at 16 px in light, dark, and
      high-contrast themes. Current implementation adds generated PNG assets
      under `src/office-ctl/common/assets/` and a deterministic generator in
      `src/office-ctl/common/scripts/generate-brand-assets.mjs`; Rust tray
      glyph coverage verifies the native 32 px mark is not blank or single-color.
- [x] Re-review and redesign the logo/icon direction against the stricter
      product-quality brief. The current mark uses abstract control-console
      document/app panes, a visible command rail, operator nodes, and a yellow
      control reticle so it communicates office productivity plus user control
      with a mature, slightly futuristic desktop utility feel. It avoids
      Microsoft Office logos, Office tile language, Microsoft 365 gradients,
      Word/Excel/PowerPoint silhouettes, Windows app-icon conventions,
      placeholder initials, generic document icons, and a gear-only settings
      mark. Covered by updated source artwork, reproducible generated raster
      assets, `brand-design.md`, Word/Excel/PowerPoint add-in identity tests,
      rendered-logo review tests, daemon product visual evidence tests, and
      Rust tray icon tests. Final real rendered screenshot evidence remains
      tracked by the visual evidence item below.
- [ ] Capture a rendered-size logo review for the final asset set. The review
      must inspect the mark at 16 px tray size, 32 px ribbon size, catalog
      thumbnail size, daemon title-bar size, and installer/package metadata
      size. The result must prove the icon still reads as office work under
      user control with a mature, slightly futuristic product feel, not as a
      Microsoft Office clone, generic document, gear/settings icon, debug glyph,
      placeholder tile, or abstract mark unrelated to office control. Store the
      review screenshots with the product visual evidence artifact. Automated
      support now exists: `record-rendered-logo-review.mjs` renders a repeatable
      contact-sheet screenshot and JSON report from the generated icon assets,
      records SHA-256 fingerprints for the source SVG and each rendered icon,
      and product identity tests verify the required rendered sizes, non-empty
      output, multi-color product palette, and current-asset fingerprints. The
      product visual evidence recorder and final validator reject rendered logo
      reviews that do not match the current `brand-mark.svg` and generated icon
      files. The GitHub Release workflow now generates
      `logo-rendered-size-review.json` and `logo-rendered-size-review.png`,
      checks that the review is ready, publishes both as release artifacts, and
      includes them in `SHA256SUMS`. This item remains open until that rendered
      review is included in the final product visual evidence artifact from the
      release run.
- [x] Update Word and Excel add-in manifests/catalog rendering so release and
      sideloaded builds use mature product metadata: stable add-in title,
      provider, description, support URL, ribbon group label, action label, and
      product icons. Placeholder labels such as `office-mcp`, generic `Open`,
      blank icons, or debug/experimental naming fail this item. Covered by Word
      and Excel task pane tests plus catalog renderer assertions.
- [x] Tighten add-in product metadata and chrome so the title, icon, provider,
      ribbon group, action label, catalog entry, and task pane header read as a
      stable product. Word and Excel manifests now use `Office MCP Control` as
      provider/display identity, `Open Control Panel` as the action label, and
      host-specific product tooltips. Static tests reject raw implementation
      slugs, vague commands such as `Open`, and prototype/debug wording.
- [x] Add explicit manifest/catalog coverage for add-in type/category metadata.
      Word and Excel catalog entries, hosted manifest output, AppSource
      pre-submission metadata, and installer-generated catalog manifests now
      describe the product as a local productivity automation/control utility,
      not as a sample, debug add-in, raw protocol bridge, or experimental task
      pane. Coverage exists in Word/Excel task pane tests, product identity
      tests, hosted manifest tests, AppSource package tests, and catalog
      registration tests.
- [x] Update task pane visible title/chrome and any in-app product references
      to match the new identity while keeping host-specific Word/Excel accents
      restrained and secondary. Word and Excel task panes now use `Office MCP
      Control` as the document-local chrome title.
- [x] Re-audit the add-in first-run identity across Word and Excel: Office
      catalog entry, ribbon command, task pane title/chrome, manifest metadata,
      provider, description, support metadata, and type/category. The combined
      title/icon/type presentation now reads as mature product software rather
      than an experimental add-in, internal protocol bridge, sample, debug
      panel, or raw implementation package. Automated coverage verifies product
      name/provider/description/type, product task pane icon/title, hosted and
      local catalog icon URLs, support URL, AppSource metadata icons, and
      catalog registration origin sync for Word and Excel. Final real Office
      catalog/ribbon screenshots remain tracked by the visual evidence item
      below.
- [ ] Capture live Word and Excel first-run add-in identity evidence. The
      screenshots must show each host's catalog entry, ribbon command, task pane
      title/chrome, icon, provider, concise description, and type/category as a
      coherent product surface. A passing run must show mature software naming
      and iconography on all visible surfaces at once; partial polish, such as a
      good title with a missing icon or a good icon with experimental
      type/category text, does not satisfy this item. Automated recorder support
      now binds product visual evidence to the current Word and Excel manifests:
      display name, provider, description, standard/high-resolution icon URLs,
      and type/category metadata are validated before product visual evidence can
      pass. This item remains open until those manifest-derived fields are also
      matched by live Office catalog/ribbon/task pane screenshots.
- [x] Redesign the product logo as a deliberate future-office-control identity,
      not as an Office logo variant, framework icon, placeholder mark, gear-only
      settings symbol, or developer/debug badge. The current mark is the
      selected `Command Console Panes` concept after three documented concept
      directions, and uses abstract control-console document/app panes, a
      command rail, operator nodes, and a control reticle to communicate office
      productivity plus user control. `brand-design.md` documents the
      selected/rejected concepts, palette, minimum sizes, and non-Microsoft
      distinction. Generated icon assets cover 16, 20, 24, 32, 48, 64, 80, 128,
      and 256 px, and `record-rendered-logo-review.mjs` plus product identity
      tests validate the rendered tray, ribbon, catalog, daemon title-bar, and
      installer sizes. Final live Office/tray visual evidence remains tracked by
      the separate product visual evidence items below.
- [x] Rework every add-in first-contact identity field so Word, Excel, and
      PowerPoint look like mature installable software after a clean install.
      Title, icon, provider, command label, short description, type/category,
      catalog card, ribbon command, task pane title/chrome, installer metadata,
      and generated catalog output are updated together for all three hosts.
      Static coverage verifies `Office MCP Control` display/provider/ribbon
      group/task pane chrome, `Open Control Panel`, generated icon URLs,
      support metadata, local productivity automation/control descriptions, and
      generated catalog manifests for Word, Excel, and PowerPoint. The catalog
      identity review now rejects generic `Add-in`/`Task Pane`, debug/prototype,
      protocol-bridge, missing-icon, raw package, and shortened ribbon-label
      regressions. Clean install packaging paths stage/run the shared catalog
      generator for all three hosts. Real Office installed-surface screenshots
      remain tracked by the live first-contact evidence items.
- [x] Replace the current tray placeholder/missing icon with the generated
      tray glyph in normal daemon and installed launch paths. The icon must be
      visible in the Windows notification area and visually deliberate beside
      native system icons. Automated coverage verifies the Rust tray icon uses
      the product glyph instead of a blank or single-color placeholder; final
      visible notification-area evidence remains manual.
- [x] Rework tray product polish so the normal Windows notification-area surface
      looks native and finished: visible original icon, product tooltip, no
      missing/default icon, no debug labels, and menu/confirmation text aligned
      with the add-in product name. Automated coverage verifies the generated
      tray glyph, product tooltip, read-only status rows, action labels,
      confirmation text, structured menu roles, and absence of scaffold/debug
      labels. Real Windows native right-click interaction evidence remains
      tracked by the next item and the final visual evidence item.
- [ ] Ensure the tray right-click interaction uses a real platform-native menu
      on Windows, with native separators, disabled/read-only status rows,
      hover/selection behavior, keyboard access, and theme/high-contrast
      adaptation. Custom web-styled or custom-drawn menu panels are not
      acceptable for release. Existing automated tests verify the Rust tray host
      uses `tray_icon::menu`; manual Windows evidence is still required for the
      native look and interaction. A browser window, webview, frameless HTML
      menu, CSS-styled popup, or manually positioned floating panel must fail
      acceptance even if it resembles a Windows menu.
- [ ] Replace any remaining tray prototype presentation with a finished native
      Windows tray surface. A normal daemon launch must show the accepted product
      tray glyph, product tooltip, and OS-native right-click menu. Missing tray
      icon, blank/default/toolkit sample icon, non-product tooltip text,
      webview/custom-drawn menu, or a menu that feels visually non-native is a
      release blocker. Evidence must come from a real interactive Windows launch,
      not only `tray --probe` model output.
- [ ] Capture tray product polish evidence from a normal interactive Windows
      launch. The artifact must show the notification-area icon, product
      tooltip, right-click native menu opened from the actual tray icon, native
      separator/read-only row behavior, `Show Office MCP Control`, `Quit Office MCP Control`,
      and the native quit confirmation dialog. The tray must use the generated
      product glyph and product-consistent labels, with no missing/default icon,
      debug wording, browser/webview menu, frameless HTML panel, or custom-drawn
      menu imitation. Product visual evidence now embeds and revalidates the
      manual tray evidence artifact, including visible icon, right-click menu,
      menu-opened-from-tray-icon proof, native appearance review, product
      tooltip, required menu items, screenshot image, and live daemon/tray probe
      context. The manual tray recorder and final product visual validator now
      also require explicit review flags proving the menu is anchored to the
      visible notification-area icon, uses OS-native menu spacing/hover/theme
      behavior, supports keyboard menu actions, and shows a native quit
      confirmation. The embedded manual tray evidence must include freshness
      metadata for the tray icon, native menu, tooltip, and quit confirmation
      screenshots; stale screenshots keep the product visual evidence from
      passing. This item remains open until those fields are backed by a real
      interactive Windows tray capture.
- [ ] Fix and verify the native tray `Show Office MCP Control` action. Clicking
      the menu item in a normal installed Windows run from
      `%LOCALAPPDATA%\office-mcp` must open or focus the exact daemon UI URL
      reported by runtime metadata / `daemon status`, normally
      `https://localhost:8765/ui/`. The implementation must exercise the real
      menu action dispatch path, not just the tray probe/menu-label model, and
      failures must be logged with the action name, URL, runtime metadata
      source, process ID, and platform launcher error. Automated tests must
      cover runtime URL resolution and action dispatch; release evidence must
      include the visible tray menu action opening the UI.
- [ ] Fix and verify the native tray `Quit Office MCP Control` action. Clicking
      the menu item in a normal installed Windows run from
      `%LOCALAPPDATA%\office-mcp` must show native confirmation, then request
      shutdown of the daemon process that owns the tray icon, MCP listener,
      add-in listener, and UI server. After confirmation, the daemon must stop
      accepting new traffic, close add-in sockets with close code `4004`, close
      listener ports such as `8765` and `8800`, and remove the tray icon as the
      process exits. The implementation must exercise the real quit action
      dispatch path, not just tray probe/menu-label output or confirmation text,
      and failures must be logged with action name, shutdown source/target,
      process ID, and controller/platform error. Automated tests must cover quit
      dispatch and failure messages; release evidence must include the visible
      tray quit action causing daemon process and port shutdown.
- [x] Polish the automated tray product surface model so normal Windows users
      are expected to see a deliberate app icon, native tooltip/title, native
      context menu text, disabled status rows, and confirmation dialogs that
      match the add-in product name. The Rust tray probe now exposes structured
      menu roles (`read_only`, `separator`, `action`), enabled states, action
      IDs, product tooltip text, and quit confirmation details. UI runtime
      evidence and manual evidence validation reject missing tooltip text,
      missing structured menu roles, web-rendered/custom menu substitutions in
      the evidence model, and debug-only primary menu commands. Final visual
      Windows proof remains tracked by the manual evidence items.
- [x] Update package asset installation and manifest renderer tests so the
      generated logo/icon files and product metadata are packaged and referenced
      from the installed add-in catalog without loopback or missing-icon paths.
      Current packaging tests assert portable staging includes the generated assets,
      and AppSource packaging includes `assets/*` in the add-in bundle.
- [x] Add automated tests for manifest metadata/icon URL substitution and asset
      presence, plus manual Windows evidence showing the ribbon command icon,
      add-in title, visible tray glyph, and native right-click menu. Automated
      tests now cover metadata/icon substitution, generated asset dimensions,
      PNG palette checks, source logo guardrails, static asset serving,
      packaging presence, tray glyph generation, tray tooltip format, structured
      native menu roles, disabled status rows, separator position, tray action
      IDs, product visual evidence recording, and product visual evidence
      validation. Live Windows ribbon/tray screenshots remain tracked by the
      final visual evidence item. Product visual evidence now also has to bind
      the screenshots to the same local daemon binary under test by recording
      `daemon status` and `tray --probe` output; the validator rejects product
      visual evidence without that daemon context, and the recorder marks the
      artifact failed before validation when the daemon context is missing, not
      ready, cannot read live UI state, or does not expose the expected tray
      snapshot. The rendered logo review artifact must include the future
      office control design brief in addition to rendered-size palette and
      non-Microsoft distinction checks.
- [ ] Capture visual evidence for the finished identity on Windows: Word ribbon
      command, Word catalog entry including type/category and icon, Word task
      pane title/icon, Excel equivalents, PowerPoint ribbon command,
      PowerPoint catalog entry including type/category and icon, PowerPoint task
      pane title/icon, visible notification-area tray icon, native right-click
      menu opened from that icon, tray tooltip/confirmation dialog, and daemon
      main window with product title/icon, compact status/details rail, and the
      three-column operations layout. The
      evidence must be recorded with
      `npm run evidence:record-product-visual -- --daemon-bin <path>`, tied to
      the same local daemon build under test, validated with
      `--require-product-visual`, and stored as release-checkable artifacts.
      The product visual artifact must embed a passed
      `catalog_identity_review` generated from the clean installed catalog
      manifests, so ribbon/catalog/task pane screenshots are tied back to the
      exact generated title, provider, description, command label, icon URLs,
      shared daemon origin, and local productivity automation/control type that
      Office should render.
      The product visual artifact must include a PowerPoint runtime evidence
      artifact proving an active presentation session, available tool count,
      add-slide, replace-text, layout mutation, and PDF success or explicit host
      rejection, so PowerPoint visual polish cannot pass from static screenshots
      alone.
      The product visual artifact must include a daemon main-window screenshot
      and explicit review flags for product identity, compact status/details
      grouping, and the three-column console layout so release evidence cannot
      pass while the daemon UI has detached oversized blocks or unrelated
      spacing.
      Screenshot artifacts must be real, complete image files; truncated image
      headers are rejected by both the recorder and validator.

**Exit criterion**: Word, Excel, and PowerPoint show a mature product add-in
name and icon in the ribbon/catalog once each host is packaged; the add-in
title, icon, provider, command label, description, and type/category read as one
finished product; the task pane title and chrome match the product identity; the
logo communicates office productivity and user control with a mature, slightly
futuristic feel without using Office owned marks; the tray has a visible
original glyph, native tooltip, native right-click menu opened from the actual
notification-area icon, native quit confirmation, and product-consistent labels;
packaged builds carry the same assets without blank placeholders, generic debug
names, or Microsoft-owned marks.

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
      Windows and the supported non-Windows packaging smoke gates. Windows portable
      package and developer bootstrap stage the Rust daemon as the runtime; legacy Node
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
      `npm run evidence:validate -- --require-office-tool-e2e`.

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
      controller, native tray host, menu model, `Show Office MCP Control`, and graceful
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
- [x] Split oversized Rust files so each `.rs` file owns one primary concept or
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
      dispatching parsed add-in methods into channel and registry services.
      Daemon runtime status-file parsing, process liveness checks, and status
      JSON rendering now live in `api/daemon_status.rs`, leaving
      `api/daemon_control.rs` focused on installed daemon start/stop control.
      The remaining production files above the review threshold were audited:
      `CommandRouter`, `RuntimeServer`, `ImageFetcher`, `DaemonConfigService`,
      `AddinChannelServer`, `McpHttpFrontend`, `NativeTraySurface`,
      `SessionRegistry`, `StdioBridge`, `TrayHost`, `UiRuntimeFile`,
      `UiStateStore`, `RuntimeServerConfig`, and wire/parser modules each now
      own one primary domain object or a tight protocol/platform helper set.
      Further splits should be driven by a new responsibility boundary, not by
      line count alone.
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

### M6.7 — Office tool E2E coverage

The evidence harness proves runtime diagnostics and selected smoke paths, but it
is not enough for tool-level correctness. Every MCP tool needs a real end-to-end
test that drives the production daemon, a real Office document, the add-in
channel, and MCP `tools/call` dispatch.

- [x] Add a non-evidence E2E test harness for Word tools. The harness starts the
      MCP daemon, creates a fresh empty Word document, opens/connects the Office
      add-in, waits until the daemon reports an active Word session, and then
      iterates through every advertised `word.*` tool. The shared lifecycle now
      includes an explicit add-in activation step before session waiting.
- [x] Add equivalent non-evidence E2E harnesses for Excel and PowerPoint tools.
      Each host harness creates a fresh empty workbook/presentation, waits for
      the corresponding add-in session, runs all advertised tools for that host,
      and closes/deletes the test file at the end. The default driver exposes a
      configurable activation step via `OFFICE_MCP_E2E_ACTIVATOR`.
- [x] Keep each host E2E run to one Office application/document lifecycle by
      default: open the host program and driver-owned test file once, connect
      one add-in session, loop over every advertised tool in that session, then
      close and delete the test file once during cleanup. Per-tool Office
      process restarts are allowed only as an explicit diagnostic fallback, not
      as the normal gate.
- [x] Convert the Office E2E driver from step-local automation to a
      lifecycle-safe driver that can keep the driver-owned Word, Excel, and
      PowerPoint files open long enough for the add-in to connect. The driver
      must preserve state across daemon startup, document creation, session
      waiting, tool calls, readback verification, and cleanup without closing or
      modifying any pre-existing user Office windows.
- [x] Add PowerPoint-specific Office automation support that uses a visible
      window path when COM rejects hidden PowerPoint automation. The test must
      prove cleanup closes only the presentation created by the E2E driver.
- [x] Fill every host tool table with one real case per advertised tool. Each
      case follows the common pattern from [03-mcp-tool-surface.md](03-mcp-tool-surface.md):
      create fixed setup content, invoke the MCP tool against that content, and
      verify the result directly for reads or through readback for mutations and
      deletes. Current coverage uses a shared `assertConcreteE2eCases` gate so
      Word, Excel, and PowerPoint must all define setup actions, matching MCP
      tool calls, explicit verifier expectations, and readback tools/resources
      where mutation or delete validation requires rereading host state. Current
      evidence: `node --test src\office-ctl\common\test\tool-e2e-contract.test.mjs`
      and `npm run e2e:tools` in `src/office-ctl/word`,
      `src/office-ctl/excel`, and `src/office-ctl/powerpoint`.
- [x] Add readback verifier helpers for common Office objects so tool cases do
      not duplicate boilerplate: Word document text/paragraph/table/content
      control/comment state, Excel workbook/sheet/range/table/chart/pivot state,
      and PowerPoint presentation/slide/shape/text/table state. Current helper
      coverage is enforced by common E2E contract tests and used by the Word
      tool case table for document text, tables, content controls, comments,
      and tracked changes; `npm run e2e:tools` passes for Word, Excel, and
      PowerPoint after the helper slice. PowerPoint chart helpers remain out of
      v1 scope because the refined PowerPoint catalog explicitly excludes
      charts.
- [x] Make exact advertised-tool coverage part of the Office E2E gate, not only
      the synthetic driver contract. A host-specific `npm run e2e:tools` command
      must fail when `tools/list` or session `available_tools` exposes a tool
      without a real setup/action/verifier row. Current coverage requires every
      Office E2E driver to expose a daemon `tools/list` step, filters that list
      to the current host namespace, compares it exactly with the host case
      table, and still compares session `available_tools` after add-in
      connection. Common contract tests include explicit failure coverage for a
      daemon-only uncovered host tool, and `npm run e2e:tools` passes for Word,
      Excel, and PowerPoint with the strengthened gate.
- [x] Structure each host E2E as one table-driven loop. Every tool case supplies
      only the tool-specific setup content, tool arguments, and verifier. The
      shared loop owns the common steps: create fixed baseline content, call the
      MCP tool, and verify the result. Read tools verify their direct result;
      mutating/comment/destructive tools verify by reading the document back
      through the most appropriate read tool or resource.
- [x] The shared loop must reset or recreate deterministic content before each
      tool case so cases are independent and can be retried without relying on
      previous side effects. The loop records the tool name, setup content,
      MCP request id, and verifier result, but it must not write document body
      content to default logs.
- [x] The harness must fail if any advertised tool lacks a case. Catalog/tool
      list discovery is authoritative: `tools/list` or session
      `available_tools` defines the expected set, and the test asserts exact
      coverage for the host's current catalog.
- [x] The E2E suite must be runnable separately from release evidence commands.
      Evidence recorders may consume its output later, but passing E2E is its
      own gate and must not be implemented as `npm run evidence:*`.
- [x] Fix live Office E2E window/process cleanup across Word, Excel, and
      PowerPoint. The current Windows sideload activator can open both the
      original driver-created file and an Office-generated `* add-in ...` copy,
      and failures can leave driver-owned Office windows/processes behind. The
      harness must keep only one visible driver-owned test file per host run,
      close the original after the sideload copy becomes active, close all
      driver-owned sideload copies in `finally` even when activation/session/tool
      verification fails, quit the Office app when no user documents remain, and
      delete all generated temp files. This is a release blocker for Word as
      well as PowerPoint; the developer must never have to close E2E-created
      Office windows by hand. Current evidence: live `npm run e2e:tools` with
      `OFFICE_MCP_RUN_E2E=1` passed for Word, Excel, and PowerPoint; the runtime
      validator passed `--require-office-tool-e2e` against
      `artifacts/office-tool-e2e-word.json`,
      `artifacts/office-tool-e2e-excel.json`, and
      `artifacts/office-tool-e2e-powerpoint.json`; final process and temp-file
      checks found no E2E-created Word, Excel, or PowerPoint residue. The E2E
      reports now include `deleted_paths` so release validation and product
      visual evidence can inspect the concrete cleanup paths for the
      driver-owned original files and Office sideload copies, not only a count.
- [x] Split the Office E2E timeout budget by lifecycle phase and fail fast on
      add-in/daemon connection failures. A single 120s `waitForSession` timeout
      is too coarse because it hides fast JavaScript/runtime failures behind a
      generic "daemon not connected" symptom. Keep longer budgets only for
      genuinely slow phases such as Office process startup and document open;
      once the task pane is activated, WebSocket `register` plus
      `session.added` should have a short 5-10s budget by default. Timeouts
      must report the exact failed phase and include useful diagnostics: the
      activation log path/tail, daemon log tail, daemon session snapshot,
      latest UIA visible control sample, document path being matched, add-in
      origin/endpoint, and any task pane ready-state evidence available from
      the activator. The harness must expose per-phase environment overrides
      for slow machines, but release defaults should catch connection bugs
      quickly instead of waiting two minutes. Current implementation keeps
      daemon startup and Office startup on separate long phase budgets,
      exposes `OFFICE_MCP_E2E_DAEMON_START_TIMEOUT_MS`,
      `OFFICE_MCP_E2E_OFFICE_START_TIMEOUT_MS`,
      `OFFICE_MCP_E2E_ACTIVATOR_TIMEOUT_MS`,
      `OFFICE_MCP_E2E_ACTIVATION_SESSION_TIMEOUT_MS`, and
      `OFFICE_MCP_E2E_SESSION_TIMEOUT_MS`, and defaults add-in session wait to
      10 seconds with `phase=addin-session-register`, endpoint, activation log,
      daemon log, latest-session, and document diagnostics.

Current implementation evidence: `src/office-ctl/common/test/tool-e2e-contract.mjs`
owns the shared table-driven loop and exact coverage checks; Word, Excel, and
PowerPoint each provide host case tables. The shared loop contract is covered by
`src/office-ctl/common/test/tool-e2e-contract.test.mjs`. The Office host
automation driver now provides daemon startup, Word/Excel COM document creation
and cleanup, a first-class add-in activation step, MCP `tools/call`, and
explicit session waiting through the `OFFICE_MCP_E2E_DRIVER` JSON step
protocol. The default driver can invoke an external UI activator command through
`OFFICE_MCP_E2E_ACTIVATOR`, passing host, document path, add-in origin, and
add-in endpoint environment variables. When no custom activator is configured
on Windows, the driver uses
`src/office-ctl/common/scripts/activate-office-mcp-addin.ps1` to focus the
driver-owned Office document and open the `Open Control Panel` ribbon command;
`OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR=0` disables that fallback for manual
diagnostics. The driver also supports generic
setup/reset action metadata, setup-result bindings such as `${table.shape_id}`,
resource-read bindings such as `${trackChanges.changes.0.fingerprint}`, generic
direct-result assertions, and generic readback verification metadata. A case can
declare MCP setup actions, read MCP resources, call the tool under test, then
assert direct result fields or call a follow-up read tool and assert expected
`contains`/`notContains`/ordered/path markers. The host case tables now cover
all currently advertised Word, Excel, and PowerPoint tools with deterministic
setup and a direct-result or readback verifier. The shared coverage gate also
parses the daemon Rust runtime catalog and fails if a host E2E table covers the
task pane `AVAILABLE_TOOLS` but misses a tool exposed by MCP `tools/list`.
Fully automated Windows ribbon/task-pane activation and live Office execution
evidence now exists for Word, Excel, and PowerPoint. Each host exposes `npm run e2e:tools` as the non-evidence
tool E2E command; live Office execution is still opt-in through
`OFFICE_MCP_RUN_E2E=1`, and release-ready live reports must use a concrete
activator rather than the `no-activator-configured` skipped path. The report
must also include the add-in activator identity and a non-empty `activation_path`;
weak activation proof such as only `activated: true` is not release-ready. When
enabled, the Word, Excel, and PowerPoint commands write
`artifacts/office-tool-e2e-word.json`,
`artifacts/office-tool-e2e-excel.json`, and
`artifacts/office-tool-e2e-powerpoint.json`. Those reports record the ordered
tool loop, per-tool verifier results, and lifecycle counters so the release
evidence can prove each host opened one driver-owned file/session and tested all
tools in that single lifecycle. The runtime evidence validator now exposes
`--require-office-tool-e2e` with explicit Word, Excel, and PowerPoint report
paths, and rejects missing reports, non-passing reports, non-single lifecycle
counts, skipped add-in activation, weak activation proof, incomplete executed
tool lists, missing concrete cleanup paths in `deleted_paths`, or failed
per-tool verifiers.
Product visual evidence now embeds the same Word, Excel, and PowerPoint tool
E2E reports and fails release validation unless those embedded reports are
ready, passing, exact-coverage, single-lifecycle, and include concrete cleanup
paths. This ties final UI/product screenshots to the same tool execution proof
instead of allowing visual polish to pass without full advertised-tool coverage.

**Exit criterion**: A developer can run one command per Office host and see the
daemon start, a blank document connect, every MCP tool execute against fixed
test content, every result verified by direct result or readback, and the test
document closed and deleted afterward. Adding a new advertised tool fails E2E
until a corresponding table row is added.

### M7 — Excel

A separate add-in (Excel.js) connects to the same server. Server gains a
new tool namespace `excel.*`. The implemented v1 catalog covers the initial
seven tools; the target core surface in [04-excel-capabilities.md](04-excel-capabilities.md)
refines the Excel backlog to about 20 task-oriented tools based on the
Microsoft workbook -> worksheet -> range -> table/chart/pivot object workflow.

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
      Current evidence: `npm run evidence:excel` now owns a self-contained live
      Excel workbook lifecycle: it starts or reuses the daemon, creates a
      driver-owned workbook, activates the add-in, writes `artifacts/office-tool-e2e-excel.json` from one driver-owned workbook session,
      and writes `artifacts/runtime-evidence-excel.json`. Validation passes with
      `npm run evidence:validate -- --require-office-tool-e2e`
      against that self-contained live Excel workbook.

#### M7.1 — Excel Core Tool Surface Refinement

Research basis: Microsoft Excel add-in docs identify the most common workflow as
`Workbook` -> `Worksheet` -> `Range`, then higher-level `Table` and `Chart`
objects created from existing range data. The starting source is Microsoft
Learn's core Excel object model concepts page:
https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts.
That page also calls out range `values`, `formulas`, and `format` as the
immediately useful cell-data operations, and states that Excel JavaScript API has
no separate `Cell` object: cells are one-cell `Range` objects. PivotTables are
not part of that page's first-level walkthrough, but they are included in v1
because they are a high-value Excel analysis workflow and are present in the
stable Excel.js object model. The refined catalog is fixed at 20 core tools by
grouping common object lifecycle/configuration operations under `update_*`
commands instead of exposing every Excel.js method. Each tool must own one
distinct user intent; overlapping behavior must be removed or assigned to one
owner before implementation.

Selection criteria: prioritize the operations users actually ask Excel to do in
an agent workflow: workbook orientation, sheet inventory/lifecycle, range/cell
value CRUD, formula authoring, formatting, sorting/filtering, table management,
chart creation/customization, and PivotTable analysis. The API budget is 15-20
tools; v1 intentionally lands at 20 because that is the smallest catalog that
covers workbook, sheet, range/cell, formula, format, data, table, chart, and
PivotTable workflows without method-level sprawl. Do not add tools for every
Office.js object, property, event, shape, comment, slicer, external connection,
or preview feature unless a later user workflow proves that the 20-tool surface
cannot express it safely and the spec either retires or merges another tool.

Research conclusion from the Microsoft Learn core concepts page: Excel v1 must
be a small workflow API, not an Excel.js mirror. The core path is `Workbook` ->
`Worksheet` -> `Range` -> `Table` / `Chart`; single cells are represented by
one-cell ranges; range `values`, `formulas`, and `format` remain separate
intents because they carry different validation and permission profiles.
PivotTables are the only extra analysis object in v1 because they represent a
high-value summarized-analysis workflow. Shapes, comments, slicers, bindings,
events, named items, external data, Power Query, import/export, save-as/close,
and method-level table/chart/PivotTable wrappers remain out of scope.

Target catalog: `excel.get_workbook_info`, `excel.list_sheets`,
`excel.add_sheet`, `excel.update_sheet`, `excel.delete_sheet`,
`excel.get_used_range`, `excel.read_range`, `excel.write_range`,
`excel.clear_range`, `excel.find_replace_cells`, `excel.set_formula`,
`excel.format_range`, `excel.sort_range`, `excel.apply_filter`,
`excel.create_table`, `excel.update_table`, `excel.create_chart`,
`excel.update_chart`, `excel.create_pivot_table`, and
`excel.update_pivot_table`.

The 20 tools are grouped as: Workbook 1, Worksheet 4, Range/cell data 5,
Formula 1, Format 1, Data operations 2, Table 2, Chart 2, and PivotTable 2.
This is the v1 upper bound. Rejected v1 expansions include separate cell CRUD,
worksheet formatting, freeze panes, protection, comments, shapes, images,
slicers, event subscriptions, bindings, named items, custom XML, external data,
Power Query, workbook import/export, save-as/close, and method-level table,
chart, or PivotTable tools that duplicate an existing owner tool.

- [x] Implement workbook and worksheet discovery/lifecycle slice:
      `excel.get_workbook_info`, `excel.list_sheets`, `excel.update_sheet`,
      `excel.delete_sheet`, and `excel.get_used_range`, including daemon catalog
      entries, task pane handlers, and Rust/JS tests. Committed as
      `excel: add workbook and worksheet tools`.
- [x] Record Excel tool-selection research in
      [04-excel-capabilities.md](04-excel-capabilities.md), starting from the
      Microsoft Learn Excel core object model and related range/table/chart/
      PivotTable docs. The spec now explicitly maps workspace-level sheet CRUD,
      sheet/range/cell CRUD, formula, format, sort/filter, table, chart, and
      PivotTable workflows into 20 task-oriented APIs.
- [x] Tighten the Excel v1 API surface to the final 20-tool maximum: document
      the category counts, explain why cells are represented by range tools,
      keep PivotTables as the only analysis object beyond the core concepts
      page's table/chart walkthrough, and record rejected method-level tool
      families so implementation does not drift into an Office.js mirror.
- [x] Record the refined Excel tool budget as an explicit 15-20 API decision:
      start from the Microsoft Learn core object path, keep single-cell work as
      range work, split values/formulas/formatting because they have different
      user intent and permission profiles, merge object lifecycle/configuration
      under `update_*` owners, and cap v1 at 20 tools.
- [x] Verify each Excel tool's minimum ExcelApi requirement set against
      `src/office-ctl/excel/node_modules/@types/office-js/index.d.ts` and the
      relevant Microsoft API docs before changing implementation. Record the
      verified minimum requirement set in
      [04-excel-capabilities.md](04-excel-capabilities.md) and avoid unresolved
      `verify during implementation` language in the final contract. Current
      evidence records the minimum requirement set for all 20 tools in
      [04-excel-capabilities.md](04-excel-capabilities.md), gates host-specific
      paths in the Excel task pane with `requireRequirementSet`, and keeps the
      runtime evidence harness aligned with the final catalog.
- [x] Keep every Excel implementation slice aligned with the final 20-tool
      catalog in [04-excel-capabilities.md](04-excel-capabilities.md): daemon
      catalog, MCP `tools/list`, Excel task pane available-tools metadata,
      permission grouping, and documentation must all expose the same tool names
      and categories. Current evidence: daemon catalog tests, MCP `tools/list`
      tests, Excel task pane contract tests, and runtime evidence tests all name
      the same 20 tools.
- [x] Keep the runtime Excel catalog capped at the refined 20-tool target unless
      a future spec update identifies a distinct object owner, permission
      profile, or user-visible workflow that cannot be represented by the
      existing tools. Any proposed expansion must update the selection matrix
      and retire or merge an existing tool before implementation. Current
      evidence: `ExcelToolCatalog`, daemon catalog/list-tools tests, and Excel
      task pane `AVAILABLE_TOOLS` expose exactly the 20-tool v1 catalog.
- [x] Add or keep contract tests that fail if the daemon catalog, Excel task
      pane available-tools metadata, or UI permission grouping advertises more
      than the v1 20-tool target without a spec update. The tests should also
      prove the required categories remain Workbook, Worksheet, Range, Formula,
      Format, Data, Table, Chart, and PivotTable. Current Excel task pane tests
      validate the grouped categories, metadata, category toggles, per-tool
      toggles, and `session.updated.available_tools`; daemon tests validate the
      catalog and MCP `tools/list` surface.
- [x] Implement range cleanup/search slice: `excel.clear_range` and
      `excel.find_replace_cells`. Tests first: daemon catalog/preflight,
      task pane dispatch/handler tests, argument validation tests, and smoke
      coverage for clear contents/formats/all plus read-only find and replace.
      Current implementation covers range clear/delete through `ExcelApi 1.1`
      and find/replace through `ExcelApi 1.9` host gating.
- [x] Extend formula/format slice: update `excel.set_formula` to accept formula
      matrices, and update `excel.format_range` to cover borders, horizontal and
      vertical alignment, wrap text, autofit rows/columns, and number-format
      matrices. Tests first: shape validation, generated Office.js calls, and
      no regression for existing scalar formula/basic format behavior. Current
      implementation validates formula and number-format matrix shape in the
      Excel task pane, maps alignment and border options through stable Office.js
      enums, gates autofit behind `ExcelApi 1.2`, and passed `npm run check` in
      `src/office-ctl/excel`.
- [x] Implement data operations slice: `excel.sort_range` and
      `excel.apply_filter`. These are the only owners for sorting/filtering both
      plain ranges and table bodies; `excel.update_table` must not duplicate
      this behavior. Tests first: plain range target, table target, clear filter,
      multi-key sort, visible/filtered rows behavior where Office.js exposes it,
      and metadata/permission categories. Current implementation adds both tools
      to the daemon catalog and Excel task pane, groups them under Data,
      capability-gates range/table sort behind `ExcelApi 1.2`, range AutoFilter
      behind `ExcelApi 1.9`, and passed Excel task pane plus targeted daemon
      catalog/list-tools tests.
- [x] Implement table object-owner slice: `excel.update_table` with explicit
      actions for metadata read, add rows, add columns, resize, rename, table
      style/options, and delete. Table cell contents stay owned by
      `excel.read_range`; generic sort/filter stays owned by data tools; generic
      cell formatting stays owned by `excel.format_range`. Current
      implementation adds daemon catalog/list-tools coverage, Excel task pane
      metadata/permission grouping, dispatch coverage, Office.js table rows,
      columns, resize, rename, style/options, delete actions, and gates visual
      table options behind `ExcelApi 1.3` and resize behind `ExcelApi 1.13`.
- [x] Implement chart object-owner slice: `excel.update_chart` with explicit
      actions for metadata read, title, axes, legend, series source, position,
      size, delete, and image export where the host supports it. Unsupported
      export must return a host-capability error, not a silent partial result.
      Current implementation adds daemon catalog/list-tools coverage, Excel
      task pane metadata/permission grouping, dispatch coverage, chart metadata,
      title, legend, axis title/visibility, source range, position, size,
      `getImage` export, and delete actions. Image export is gated behind
      `ExcelApi 1.2`; axis selection and chart type/id metadata are gated behind
      `ExcelApi 1.7`.
- [x] Implement PivotTable slice: `excel.create_pivot_table` and
      `excel.update_pivot_table` for normal range/table sources, row/column/data
      fields, filters, aggregation/calculation, refresh, metadata read, and
      delete. OLAP, Power Pivot, slicers, and preview-only APIs remain deferred.
      Current implementation adds daemon catalog/list-tools coverage, Excel
      task pane metadata/permission grouping, dispatch coverage, PivotTable
      creation from range/table sources, metadata, hierarchy add/remove for row,
      column, filter, and data axes, data aggregation/number format, layout
      options, refresh, manual filters, clear filters, and delete. Creation and
      hierarchy/layout/delete actions are gated behind `ExcelApi 1.8`; manual
      filters are gated behind `ExcelApi 1.12`.
- [x] Update the Excel task pane tools UI and per-tool permissions so all 20
      tools are grouped by Workbook, Worksheet, Range, Formula, Format, Data,
      Table, Chart, and PivotTable categories, with category toggles and per-tool
      toggles preserving `session.updated.available_tools` behavior. Current
      Word/Excel add-in checks validate grouped tool toggles, category toggles,
      and `session.updated.available_tools` updates from the compact task pane
      UI.
- [x] Add final Excel v1 automated evidence: Rust forwarding/preflight tests, daemon
      catalog tests, Excel task pane contract tests, `npm run check` in
      `src/office-ctl/excel`, targeted daemon cargo tests, `git diff --check`,
      and representative live Excel E2E evidence covering all implemented
      categories. Current automated evidence covers the full daemon catalog,
      MCP forwarding/listing, Excel task pane contract, and a 20-tool Office
      tool E2E loop. Current live evidence is self-contained: `npm run
      evidence:excel` delegates to `src/office-ctl/excel` and writes
      `artifacts/office-tool-e2e-excel.json`, and `npm run evidence:validate
      -- --require-office-tool-e2e` validates the regenerated report.

### M8 — PowerPoint

Similar pattern: `powerpoint.*` namespace, second add-in.
Initial catalog: `add_slide`, `replace_text`, `insert_image`, `apply_layout`,
`export_pdf`.

- [x] Add the daemon-side PowerPoint forwarding contract and catalog entries
      for `powerpoint.add_slide`, `powerpoint.replace_text`,
      `powerpoint.insert_image`, `powerpoint.apply_layout`, and
      `powerpoint.export_pdf`. Rust tests cover MCP `tools/list`, catalog
      membership, capability-gated forwarding, and a PowerPoint session dispatch
      path through the add-in connection hub.
- [x] Add a PowerPoint add-in scaffold under `src/office-ctl/powerpoint` with
      Presentation manifest metadata, mature product identity, compact task pane
      UI, shared channel/logger/task-history usage, PowerPoint runtime
      registration, session announcement, Windows catalog staging, packaging
      inclusion, and daemon static serving under `/powerpoint/*`. The initial
      scaffold announced no tools until the handler slice below added tested
      Office.js implementations.
- [x] Finish PowerPoint add-in first-run identity and catalog validation. The
      PowerPoint `DisplayName`, ribbon group, command label, task pane title,
      icon URLs, provider, description, support metadata, and type/category now
      match `Office MCP Control` quality bars and avoid scaffold/sample/protocol
      wording. The shared Windows catalog installer creates
      `office-mcp-powerpoint.xml`, syncs it to the daemon origin and generated
      asset URLs, removes stale legacy host subfolders, and is covered by the
      Office catalog registration tests.
- [x] Extend product identity automated coverage to include PowerPoint wherever
      Word and Excel are already checked: manifest metadata, catalog type/
      category, generated icon references, packaged asset presence, static
      serving under `/powerpoint/*`, and product visual evidence recording
      fields for PowerPoint ribbon/catalog/task pane screenshots. Coverage now
      includes the shared product identity test, PowerPoint task pane tests,
      Windows packaging/catalog tests, daemon static asset tests, and product
      visual evidence recorder/validator gates for PowerPoint first-run identity
      plus ribbon/catalog/task pane screenshot surfaces.
- [x] Implement the PowerPoint task pane handlers for `powerpoint.add_slide`,
      `powerpoint.replace_text`, `powerpoint.insert_image`,
      `powerpoint.apply_layout`, and `powerpoint.export_pdf`, with typed
      argument validation, useful error mapping, cancellation/deadline behavior,
      and unit/static tests for each supported Office.js API path. Current
      implementation uses PowerPoint Office.js APIs for slide creation,
      text-box insertion/replacement, selected-slide image insertion, slide
      layout application, and PDF export through `getFileAsync(Pdf)`, while
      preserving per-tool permission toggles and `session.updated` available
      tool updates.
- [x] Add live PowerPoint runtime smoke evidence against a real presentation.
      The evidence must prove session registration, tool visibility after
      handler implementation, at least one read/write mutation path, and PDF
      export behavior or an explicit host-capability rejection when the current
      Office.js host cannot support export. Automated support now exists via
      `npm run evidence:powerpoint`, which writes
      `artifacts/runtime-evidence-powerpoint.json`, owns a self-contained live
      PowerPoint presentation lifecycle, and validator gate
      `npm run evidence:validate -- --require-office-tool-e2e`.
      Current evidence: `artifacts/office-tool-e2e-powerpoint.json` records a
      passed `office_tool_e2e_report` against that self-contained live
      PowerPoint presentation, and the validator passes
      `--require-office-tool-e2e` with the refined 25-tool catalog,
      per-tool setup/readback proof, cleanup proof, and activation proof.

#### M8.1 — PowerPoint Core Tool Surface Refinement

Research basis: Microsoft Learn's PowerPoint add-in core concepts page
identifies the primary object model as `Presentation` -> `Slide` -> content
objects such as `Shape`, `TextRange`, and `Table`, with `Layout` controlling how
slide content is organized. It also calls out the Common API owners for runtime
context, requirement-set probing, active view, and document file download. The
refined PowerPoint catalog is fixed at 25 tools in
[04-powerpoint-capabilities.md](04-powerpoint-capabilities.md). It applies
Occam's razor: each tool must own one distinct user workflow; method-level or
overlapping tools must be merged into the relevant object-owner update tool.

Target catalog: `powerpoint.get_presentation_info`,
`powerpoint.get_active_view`, `powerpoint.export_file`,
`powerpoint.update_tags`, `powerpoint.list_slides`, `powerpoint.add_slide`,
`powerpoint.update_slide`, `powerpoint.delete_slide`,
`powerpoint.move_slide`, `powerpoint.export_slide`,
`powerpoint.list_layouts`, `powerpoint.apply_layout`,
`powerpoint.get_selection`, `powerpoint.set_selection`,
`powerpoint.list_shapes`, `powerpoint.add_text_box`,
`powerpoint.add_shape`, `powerpoint.insert_image`,
`powerpoint.update_shape`, `powerpoint.read_text`,
`powerpoint.replace_text`, `powerpoint.format_text`,
`powerpoint.add_table`, `powerpoint.read_table`, and
`powerpoint.update_table`.

The 25 tools are grouped as: Presentation 3, Metadata 1, Slides 6, Layout 2,
Selection 2, Shapes 5, Text 3, and Tables 3. Rejected v1 expansions include a
PowerPoint deck save tool, separate `powerpoint.export_pdf`, slide duplication
without a proven stable host API, custom XML mutation, document-property writes,
separate shape move/resize/delete/fill tools, separate table row/column/cell
tools, separate title/body/placeholder text tools, charts, SmartArt, media,
animations, transitions, comments, speaker notes, slide show control,
PowerPoint Designer, and macros.

- [x] Research the PowerPoint v1 tool surface from Microsoft Learn's
      PowerPoint core concepts page and `@types/office-js`, then record the
      refined 25-tool target in
      [04-powerpoint-capabilities.md](04-powerpoint-capabilities.md). Current
      spec keeps export under `powerpoint.export_file`, rejects a PowerPoint
      deck save tool because stable typings do not expose one, keeps slide
      move/export at `PowerPointApi 1.8`, applies layout at `PowerPointApi 1.8`,
      and reserves table operations for `PowerPointApi 1.8/1.9`.
- [x] Verify every planned PowerPoint tool's minimum `PowerPointApi` or Common
      API requirement set against the local `@types/office-js` baseline and the
      relevant Microsoft API docs before implementation. The implementation now
      gates PowerPointApi 1.3/1.4/1.5/1.8/1.9/1.10 and Common API export/view
      paths explicitly, returning `HOST_CAPABILITY_UNAVAILABLE` where the live
      host lacks the required object model.
- [x] Update daemon catalog and MCP `tools/list` from the current 5-tool
      PowerPoint runtime surface to the refined 25-tool surface. Remove
      `powerpoint.export_pdf` from the advertised catalog after
      `powerpoint.export_file` is implemented and covered.
- [x] Update the PowerPoint task pane `AVAILABLE_TOOLS`, permission grouping,
      metadata descriptions, category toggles, per-tool toggles, and
      `session.updated.available_tools` behavior so the UI exposes the same
      Presentation, Metadata, Slides, Layout, Selection, Shapes, Text, and
      Tables categories as the spec.
- [x] Implement Presentation and Metadata slice:
      `powerpoint.get_presentation_info`, `powerpoint.get_active_view`,
      `powerpoint.export_file`, and `powerpoint.update_tags`. Do not implement
      a deck save tool unless stable PowerPoint save support is proven first.
- [x] Implement Slide and Layout slice: `powerpoint.list_slides`, extend
      `powerpoint.add_slide` as needed, add `powerpoint.update_slide`,
      `powerpoint.delete_slide`, `powerpoint.move_slide`,
      `powerpoint.export_slide`, `powerpoint.list_layouts`, and keep
      `powerpoint.apply_layout` aligned with its verified `PowerPointApi 1.8`
      requirement.
- [x] Implement Selection and Shape slice: `powerpoint.get_selection`,
      `powerpoint.set_selection`, `powerpoint.list_shapes`,
      `powerpoint.add_text_box`, `powerpoint.add_shape`, improve
      `powerpoint.insert_image` to a shape-owned implementation where host APIs
      support it, and add `powerpoint.update_shape` as the single owner for
      shape position, size, rotation, fill, line, z-order, grouping, alt text,
      and deletion.
- [x] Implement Text slice: `powerpoint.read_text`, refine
      `powerpoint.replace_text`, and add `powerpoint.format_text`. Do not add
      separate title, subtitle, body, placeholder, or selected-text mutation
      tools; text inside shapes belongs to the Text tools.
- [x] Implement Table slice: `powerpoint.add_table`, `powerpoint.read_table`,
      and `powerpoint.update_table` for cell values, row/column dimensions,
      add/delete rows and columns, merge cells, clear values/formatting, style
      settings, and table-shape deletion. Do not add separate row, column, or
      cell tools.
- [x] Add PowerPoint refinement contract tests: daemon catalog count and names,
      MCP `tools/list`, task pane available-tools metadata, grouped permission
      categories, no duplicate owner tools, no `powerpoint.save`, and no
      `powerpoint.export_pdf` after `powerpoint.export_file` lands.
- [x] Add final PowerPoint v1 evidence: `npm run check` in
      `src/office-ctl/powerpoint`, targeted daemon cargo tests, `git diff
      --check`, and a live PowerPoint smoke covering Presentation, Slides,
      Layout, Shapes, Text, Tables, and export success or explicit
      host-capability rejection. Current evidence includes `npm run check` in
      `src/office-ctl/powerpoint`, `npm run evidence:powerpoint` and
      `npm run check` in `src/office-mcp/daemon/evidence`,
      `cargo test -p office-mcp-daemon`, `cargo fmt --check`, and
      `git diff --check`; the generated PowerPoint runtime evidence records a
      self-contained `office_tool_e2e_report` with all 25 tools advertised and per-tool
      proofs for Presentation, Slides, Layout, Shapes, Text, and Tables.

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
