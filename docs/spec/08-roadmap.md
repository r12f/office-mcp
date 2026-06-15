# 08 — Roadmap

## Milestones

### M0 — Spec & scaffolding (this PR)

- [x] Architecture & protocol specs
- [x] Automated feasibility harness: protocol, MCP transport, WSS, manifest,
      stable API types, requirement-set boundaries, and save behavior
- [x] Consented Word runtime smoke test plus representative IRM rights matrix
      (`npm run evidence:runtime -- --include-full-word-smoke
      --include-com-tracked-changes` generates structured full Word runtime
      evidence; `npm run evidence:irm` plus `--require-irm-preflight` and
      `--require-irm` validates the representative protected document)
- [x] Repository scaffolding (long-running `mcp-server/` package + Word `addin/` scaffold)
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
- [x] Windows MSI build: packages `node.exe`, the compiled daemon, production
      Node dependencies, add-in bundle, catalog manifest, and user
      autostart/catalog registration. Native `office-mcp.exe` packaging and
      Scheduled Task replacement remain production hardening items.
- [x] macOS Homebrew formula template and renderer for release tarballs. Actual
      tap publication waits for a signed GitHub Release artifact.
- [x] Hosted manifest renderer for `office-mcp.dev` release artifacts. Actual
      DNS/hosting publication remains a release operations gate.
- [x] AppSource pre-submission package generator: hosted manifest, add-in
      bundle, checksums, and checklist. Partner Center submission and Microsoft
      validation review remain external gates.

### M6.5 — Product UI

- [ ] Daemon tray icon with right-click status menu: Up/Down, client count,
      document count, separator, Show Office MCP, Quit Office MCP.
- [ ] Daemon main window with status, endpoints, connected MCP clients, grouped
      document sessions, current tasks, and recent command history.
- [ ] Per-document detail expansion showing the most recent 10 commands and
      success/failure details.
- [ ] Add-in task pane rebuilt as a modern TypeScript frontend showing daemon
      connection, current document session, current task, and latest 20 task
      history entries.
- [ ] Redacted UI state API and tests proving document body content, inserted
      text, image data, and secrets do not appear in UI snapshots.
- [ ] Automated UI coverage for empty, connected, degraded, in-flight, success,
      failure, timeout, and reconnect states.

**Exit criterion**: A non-technical user can open the tray menu and main window
to understand whether office-mcp is running, which clients and documents are
connected, what task is currently executing, and why the last relevant command
failed, without inspecting terminal output or logs.

### M7 — Excel

A separate add-in (Excel.js) connects to the same server. Server gains a
new tool namespace `excel.*`. Catalog (rough): `excel.read_range`,
`excel.write_range`, `excel.add_sheet`, `excel.set_formula`,
`excel.format_range`, `excel.create_table`, `excel.create_chart`.

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




