# 08 — Roadmap

## Milestones

### M0 — Spec & scaffolding (this PR)

- [x] Architecture & protocol specs
- [ ] Repository scaffolding (server crate + add-in scaffold)
- [ ] CI matrix (Win x64, macOS arm64, Linux x64 for the server; lint+typecheck for the add-in)

### M1 — Walking skeleton

End-to-end "agent reads paragraph 0 of an open Word doc":

- [ ] Daemon: MCP Streamable HTTP frontend, single tool `office.list_sessions`,
      one resource `office://word/<id>/paragraph/0`
- [ ] Server: WS listener, register, session.added, ping/pong
- [ ] Add-in: skeleton task pane with WS reverse-connect
- [ ] Add-in: implements ONE method — `tool.invoke` for reading paragraph 0
- [ ] Manual test against Claude Desktop on Windows

**Exit criterion**: From a fresh Claude Desktop install, the user can ask
"what does paragraph 1 of my open Word doc say" and get the right answer.

### M2 — Read & insert

- [ ] `word.get_text`, `word.get_outline`, `word.get_paragraph`, `word.find_text`,
      `word.get_selection`
- [ ] `word.insert_paragraph`, `word.insert_heading`, `word.insert_table`,
      `word.insert_page_break`
- [ ] IRM rights surfacing in `session.added`
- [ ] IRM pre-check in add-in
- [ ] Selection-anchored operations
- [ ] Reconnect + grace period

**Exit criterion**: A user can have Claude draft a section into their
open document, including in an IRM-protected file where they have edit rights.

### M3 — Edit & review

- [ ] `word.replace_text` (with `dry_run`)
- [ ] `word.update_paragraph`, `word.delete_range`, `word.apply_formatting`
- [ ] `word.add_comment`, `word.resolve_comment`
- [ ] `word.accept_change`, `word.reject_change`
- [ ] Track Changes mode interaction (auto-routes edits as revisions)

**Exit criterion**: Agent can do a real editing pass — find, propose,
replace, comment — on a tracked-changes-on document and the user can
review the result like a normal collaborator's edits.

### M4 — Tables & images

- [ ] `word.read_table`, `word.update_cell`, `word.add_row`, `word.add_column`,
      `word.format_cell`
- [ ] `word.insert_image` (base64 + URL)
- [ ] `word.insert_list` (numbered, bulleted)

### M5 — Document IO

- [ ] `word.save`, `word.save_as`, `word.export_pdf`
- [ ] HTTP transport for the MCP frontend
- [ ] Streamable HTTP per 2026 MCP revision

### M6 — Distribution

- [ ] Windows MSI
- [ ] macOS Homebrew tap
- [ ] Manifest hosting on `office-mcp.dev`
- [ ] AppSource submission

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
| WSS for non-loopback add-ins | Office on Web behind corp proxy needs it | Production user report from enterprise |
| Macro execution | Security risk too large | Never, except via explicit allowlist + signed origin |
| Server-side LLM summary tools (`word.summarize`) | Out of scope; that's the client's job | Never |
| Multi-tenant SaaS hosting | Not the project's model | Never |

## Non-goals (restated)

- Replacing `python-docx` for batch/headless work.
- Rendering / formatting fidelity that exceeds what Office.js can produce.
- Authoring or modifying IRM/AIP policies.
- Bypassing any Office security control.

## Open questions

1. **Add-in pinning UX**: how do we get the task pane to auto-load without
   the user clicking the ribbon every time? Office requires an `IncludePersistent`
   declaration in the manifest, but support varies by Office channel. Investigate
   M3.
2. **Office on Web**: WS to localhost from a browser-hosted Office add-in
   requires the browser to allow it. Defer to M5; may require WSS + a
   per-user reverse tunnel for the loopback.
3. **Identity binding**: in multi-account Word (one user signed into both work
   and personal Microsoft accounts), how do we tag sessions by identity?
   Probably via `user.upn` + `user.tenant_id` in `session.added`. Test M2.
4. **Multiple installed versions on one machine**: if a user has both
   office-mcp 0.5 and 0.6 installed, both pointing at the shared config and
   both starting at logon, they'll race for `127.0.0.1:8765`. Probably
   acceptable (whichever starts second fails loudly with "address already in
   use"); but a better answer might be: registered servers post their version
   to the WS handshake, second one defers. Investigate M2.
5. **Idle add-in**: if the user opens Word but never opens the task pane, the
   add-in doesn't connect. Is that fine, or should we use Office's
   "[automatically open task pane on document open]" feature so it auto-connects?
   Trade-off: convenience vs. user surprise. Default to NOT auto-open; provide
   a setting.
