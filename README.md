# office-mcp

**A bridge between AI assistants and live Microsoft Office applications via in-process add-ins.**

`office-mcp` exposes Word, Excel, and PowerPoint (with Outlook planned) as MCP tools by running an
add-in inside each Office instance. The add-in reverse-connects to a single long-lived MCP server
process, which multiplexes MCP clients across all running Office windows.

## Why not just use `python-docx` / `docx2pdf` / COM?

| Problem | python-docx / COM | office-mcp |
|---|---|---|
| IRM / RMS protected documents | ❌ Cannot open | Design target: host-enforced access; M0 validation required |
| Live editing in user's open document | ❌ File must be closed | ✅ Operates on the live document |
| Office instance exclusive-access errors | ❌ Common with COM | ✅ Each Office instance has its own add-in |
| Add-in install / discovery | n/a | ✅ Add-in self-registers on Office start |
| MCP client config churn | ❌ Per-doc subprocess | ✅ One persistent server endpoint |
| Platform path | ❌ Windows only | Windows desktop v1; Mac planned; Web requires a different deployment |

## Architecture (one diagram)

```
┌─────────────┐       ┌───────────────────────┐       ┌────────────────────────────┐
│ MCP Client  │◀─────▶│  office-mcp server    │◀─────▶│ Word instance A (add-in)   │
│ (Claude,    │ HTTP  │  (long-lived process) │  WS   ├────────────────────────────┤
│  Cursor,    │       │                       │       │ Word instance B (add-in)   │
│  agent)     │ HTTP  │  - tool router        │       ├────────────────────────────┤
└─────────────┘       │  - session registry   │       │ Excel instance C (add-in)  │
                      │  - capability negotiation │   └────────────────────────────┘
                      └───────────────────────┘
```

- **MCP server** is a single long-running process. It speaks MCP Streamable HTTP
  to clients and JSON-RPC over a local secure WebSocket to add-ins.
- **Office add-ins** are Office.js task-pane add-ins (one per Office app type). Each loaded
  runtime dials out to the server and registers the current host document it can drive.
- **Clients** see a uniform MCP tool surface; the server routes each call to the add-in that
  owns the target document.

## Status

The Windows desktop implementation is in place for Word, Excel, and PowerPoint:

- Local MCP Streamable HTTP endpoint at `http://127.0.0.1:8800/mcp`.
- Local HTTPS task pane and WSS add-in channel at `https://localhost:8765`.
- Word, Excel, and PowerPoint task pane add-ins that reverse-register one live document/workbook/presentation session.
- MCP server catalog covers the full Word v1 tool surface from
  `doc/spec/04-word-capabilities.md`, the Excel v1 tool surface from
  `doc/spec/04-excel-capabilities.md`, and the PowerPoint v1 tools in the roadmap.
- The current Word add-in runtime advertises and executes all 27 Word v1 tools: discovery, read, insert, edit, table, structure, review, tracked-change, and save operations.
- The current Excel add-in runtime advertises and executes the Excel v1 workbook tools: range read/write, sheet creation, formula setting, formatting, table creation, and chart creation.
- The current PowerPoint add-in runtime advertises and executes the PowerPoint v1 presentation tools: add slide, replace text, insert image, apply layout, and PDF export where the host supports it.
- The daemon also exposes the Word v1 resource surface, including document text, structure, paragraph, comments, tracked changes, and selection.

The design docs remain the source of truth for the broader v1 surface.

## Verification

CI runs on Windows, macOS, and Linux for the server package, plus Windows add-in
manifest and task pane syntax validation. The equivalent local checks are:

```powershell
cargo fmt --all --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
cd .\src\office-mcp\daemon\evidence
npm install
npm run check
npm run check:ui
cd ..\..\..\office-ctl\word
npm run check
cd ..\excel
npm run check
cd ..\powerpoint
npm run check
cd ..\..\..\packaging
npm run check
```

## Repository layout

| Path | Owner |
|---|---|
| `src/office-mcp/daemon/` | Native Rust daemon service, daemon-owned state/API, and daemon UI source/assets. |
| `src/office-mcp/daemon/evidence/` | Runtime, UI, smoke, and validation evidence harnesses for daemon parity and release gates. |
| `src/office-ctl/common/` | Shared TypeScript add-in utilities: config, logging, channel/protocol helpers, redaction, and reusable UI primitives. |
| `src/office-ctl/word/` | Word add-in package: XML manifest, task pane static bundle, add-in validation scripts, and Word catalog registration script. |
| `src/office-ctl/excel/` | Excel add-in entry point and host-specific command surface. |
| `src/office-ctl/powerpoint/` | PowerPoint add-in package: XML manifest, compact task pane static bundle, presentation-session registration, add-in validation scripts, and PowerPoint command handlers. |
| `packaging/` | Cross-component packaging assets: Windows bootstrap scripts and WiX MSI source. |
| `doc/spec/` | Product and protocol design. |

The repository root intentionally does not contain a top-level Node package.
Run Rust daemon commands from the repository root, evidence/smoke commands from
`src/office-mcp/daemon/evidence/`, add-in commands from `src/office-ctl/word/`,
and packaging checks from `packaging/` so the daemon, evidence harnesses,
Office add-in, and installers stay independently buildable and packageable.
Historical feasibility spikes and generated installer/runtime evidence belong
under `artifacts/`, not as parallel source packages at the root.

## Design docs

| Doc | Purpose |
|---|---|
| [00-overview.md](doc/spec/00-overview.md) | Goals, non-goals, glossary |
| [01-architecture.md](doc/spec/01-architecture.md) | Process model, transports, lifecycle |
| [02-registration-protocol.md](doc/spec/02-registration-protocol.md) | Add-in ↔ server JSON-RPC wire protocol |
| [03-mcp-tool-surface.md](doc/spec/03-mcp-tool-surface.md) | MCP tools, resources, prompts exposed to clients |
| [04-word-capabilities.md](doc/spec/04-word-capabilities.md) | Word-specific tool catalog (v1) |
| [04-excel-capabilities.md](doc/spec/04-excel-capabilities.md) | Excel-specific tool catalog (v1) |
| [05-security.md](doc/spec/05-security.md) | Security model, sandbox, IRM, trust boundaries |
| [06-error-model.md](doc/spec/06-error-model.md) | Error codes, retry, partial-failure semantics |
| [07-deployment.md](doc/spec/07-deployment.md) | Installation, sideloading, distribution |
| [08-roadmap.md](doc/spec/08-roadmap.md) | Milestones, Excel/PPT/Outlook follow-on |
| [09-ui.md](doc/spec/09-ui.md) | Daemon tray/main window and add-in task pane UI |

## Run the MVP locally

Prerequisites:

- Windows with Word desktop installed.
- Rust toolchain matching `rust-toolchain.toml`.
- Node.js 22 or newer.
- A trusted current-user `CN=localhost` HTTPS certificate. The daemon reads a
  PFX file; it does not import root certificates during startup.

Verify the Rust daemon, evidence harnesses, Office add-ins, and packaging checks:

```powershell
cargo fmt --all --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
cd .\src\office-mcp\daemon\evidence
npm install
npm run check
cd ..\..\..\office-ctl\word
npm install
npm run check
cd ..\excel
npm install
npm run check
cd ..\powerpoint
npm install
npm run check
cd ..\..\..\packaging
npm install
npm run check
```

Export the already trusted localhost certificate to the daemon's default PFX
path:

```powershell
cd ..\..\..
powershell -ExecutionPolicy Bypass -File .\packaging\windows\export-localhost-dev-cert.ps1
```

Start the Rust daemon from the repository root:

```powershell
cargo run -p office-mcp-daemon -- daemon run
```

For release/UI validation on Windows, record the visible tray interaction after
confirming the notification-area icon, right-click menu, and `Show Office MCP Control`
command. Attach a screenshot path from the same run:

```powershell
cd .\src\office-mcp\daemon\evidence
npm run evidence:record-tray-manual -- --daemon-bin C:\Code\office-mcp\target\debug\office-mcp-daemon.exe --visible-icon true --right-click-menu true --menu-opened-from-tray-icon true --native-menu-appearance-reviewed true --menu-anchored-to-tray-icon true --os-native-menu-behavior-reviewed true --keyboard-menu-access-reviewed true --native-quit-confirmation-reviewed true --menu-surface-kind native --show-ui-opened true --tooltip "Office MCP Control - Up - 0 clients - 0 documents" --menu-item "Status: Up" --menu-item "Clients: 0" --menu-item "Documents: 0" --menu-item "Show Office MCP Control" --menu-item "Quit Office MCP Control" --screenshot-path C:\path\to\tray-visible.png --tray-icon-screenshot C:\path\to\tray-icon.png --tray-native-menu-screenshot C:\path\to\tray-menu.png --tray-tooltip-screenshot C:\path\to\tray-tooltip.png --tray-quit-confirmation-screenshot C:\path\to\tray-quit.png
node ..\..\..\office-ctl\common\scripts\record-rendered-logo-review.mjs --output ..\..\..\..\artifacts\logo-rendered-size-review.json --sheet ..\..\..\..\artifacts\logo-rendered-size-review.png
node ..\..\..\office-ctl\common\scripts\record-catalog-identity-review.mjs --catalog-path ..\..\..\..\addin-catalog --output ..\..\..\..\artifacts\catalog-identity-review.json
npm run evidence:record-product-visual -- --daemon-bin C:\Code\office-mcp\target\debug\office-mcp-daemon.exe --manual-tray-evidence-path ..\..\..\..\artifacts\tray-manual-evidence.json --rendered-logo-review-path ..\..\..\..\artifacts\logo-rendered-size-review.json --catalog-identity-review-path ..\..\..\..\artifacts\catalog-identity-review.json --excel-runtime-evidence-path ..\..\..\..\artifacts\runtime-evidence-excel.json --powerpoint-runtime-evidence-path ..\..\..\..\artifacts\runtime-evidence-powerpoint.json --word-ribbon-command "Office MCP Control ribbon command visible" --word-ribbon-command-screenshot C:\path\to\word-ribbon.png --word-catalog-entry "Office MCP Control catalog entry visible" --word-catalog-entry-screenshot C:\path\to\word-catalog.png --word-taskpane-title "Office MCP Control task pane visible" --word-taskpane-title-screenshot C:\path\to\word-taskpane.png --excel-ribbon-command "Office MCP Control ribbon command visible" --excel-ribbon-command-screenshot C:\path\to\excel-ribbon.png --excel-catalog-entry "Office MCP Control catalog entry visible" --excel-catalog-entry-screenshot C:\path\to\excel-catalog.png --excel-taskpane-title "Office MCP Control task pane visible" --excel-taskpane-title-screenshot C:\path\to\excel-taskpane.png --powerpoint-ribbon-command "Office MCP Control ribbon command visible" --powerpoint-ribbon-command-screenshot C:\path\to\powerpoint-ribbon.png --powerpoint-catalog-entry "Office MCP Control catalog entry visible" --powerpoint-catalog-entry-screenshot C:\path\to\powerpoint-catalog.png --powerpoint-taskpane-title "Office MCP Control task pane visible" --powerpoint-taskpane-title-screenshot C:\path\to\powerpoint-taskpane.png --logo-tray-size "Office MCP Control tray-size logo reviewed" --logo-tray-size-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-ribbon-size "Office MCP Control ribbon-size logo reviewed" --logo-ribbon-size-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-catalog-thumbnail "Office MCP Control catalog thumbnail logo reviewed" --logo-catalog-thumbnail-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-daemon-titlebar "Office MCP Control daemon title-bar logo reviewed" --logo-daemon-titlebar-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-installer-metadata "Office MCP Control installer metadata logo reviewed" --logo-installer-metadata-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --tray-icon "Office MCP Control tray icon visible" --tray-icon-screenshot C:\path\to\tray-icon.png --tray-native-menu "Office MCP Control native tray menu visible" --tray-native-menu-screenshot C:\path\to\tray-menu.png --tray-tooltip "Office MCP Control - Up - 0 clients - 0 documents" --tray-tooltip-screenshot C:\path\to\tray-tooltip.png --tray-quit-confirmation "Office MCP Control quit confirmation visible" --tray-quit-confirmation-screenshot C:\path\to\tray-quit.png --catalog-icon-visible true --tray-icon-visible true --tray-menu-native true --quit-confirmation-visible true --logo-quality-reviewed true --final-logo-user-surface-reviewed true --rendered-size-logo-reviewed true --addin-identity-reviewed true --addin-installable-surface-reviewed true --word-first-run-identity-reviewed true --excel-first-run-identity-reviewed true --powerpoint-first-run-identity-reviewed true --tray-product-polish-reviewed true --tray-normal-windows-launch-reviewed true --word-compact-top-block true --word-tools-permissions-merged true --word-inline-settings true --word-server-protocol-row "Server 0.1.0 / Protocol 1.0" --word-document-state "Editable" --excel-compact-top-block true --excel-tools-permissions-merged true --excel-inline-settings true --excel-server-protocol-row "Server 0.1.0 / Protocol 1.0" --excel-document-state "Editable" --powerpoint-compact-top-block true --powerpoint-tools-permissions-merged true --powerpoint-inline-settings true --powerpoint-server-protocol-row "Server 0.1.0 / Protocol 1.0" --powerpoint-document-state "Editable"
npm run evidence:validate-ui -- --input ..\..\..\..\artifacts\ui-runtime-evidence.json --require-manual-tray --manual-tray-evidence-path ..\..\..\..\artifacts\tray-manual-evidence.json --require-product-visual --product-visual-evidence-path ..\..\..\..\artifacts\product-visual-evidence.json
```

Launch the add-in in Word, then open the **office-mcp > Open** ribbon command
so the task pane can connect and register the current document session.

Smoke-test a connected Word session from another terminal in `src/office-mcp/daemon/evidence/`:

```powershell
npm run smoke:mcp -- http://127.0.0.1:8800/mcp sessions
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-core <session_id>
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-formatting <session_id>
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-review <session_id>
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-resources <session_id>
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-spec-args <session_id>
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-track-change <session_id> accept
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-track-change-com <session_id> accept
npm run smoke:mcp -- http://127.0.0.1:8800/mcp word-track-change-com <session_id> reject
```

For structured evidence that can be attached to a release or validation issue,
run:

```powershell
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\..\..\..\artifacts\runtime-evidence-full.json --include-mutation --include-full-word-smoke --include-tracked-changes --include-com-tracked-changes
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-full.json --require-mutation --require-full-word-smoke --require-com-tracked-changes
```

For a connected Excel workbook, run:

```powershell
# Optional when the daemon MCP port is not 8800:
# $env:OFFICE_MCP_MCP_ENDPOINT = 'http://127.0.0.1:8801/mcp'
npm run evidence:excel
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-excel.json --require-excel-smoke
npm run evidence:powerpoint
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-powerpoint.json --require-powerpoint-smoke
```

Against a representative protected document, add `--irm-document-path` for a
read-only COM preflight and `--irm-mode protected-read` or `--irm-mode
protected-edit` once that same document has a connected add-in session:

```powershell
npm run evidence:irm
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-irm.json --require-irm-preflight --require-irm
```

For another protected document, use the explicit form:

```powershell
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\..\..\..\artifacts\runtime-evidence-irm.json --include-mutation --irm-document-path "C:\path\to\protected.docx" --irm-mode protected-read --wait-for-session-ms 120000
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-irm.json --require-irm-preflight --require-irm
```

Start the evidence command before opening the task pane if needed. With
`--wait-for-session-ms`, the harness waits for a connected session whose
document metadata matches `--irm-document-path`, then continues automatically.

Use `--require-claude-desktop-installation` only on a machine where Claude
Desktop is actually installed or running; the gate records whether the config,
install directory, or process is present before the manual UI prompt test. After
a successful agent client prompt, record the observed UI result as structured
evidence:

```powershell
npm run evidence:record-agent-client -- --prompt "what does paragraph 1 of my open Word doc say?" --expected-substring "<expected text from Word>" --observed-answer "<agent answer>" --document-title "<open Word document>" --session-id "<session_id>"
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\..\..\..\artifacts\runtime-evidence-agent-client.json --agent-client-evidence-path ..\..\..\..\artifacts\agent-client-evidence.json
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-agent-client.json --require-agent-client-prompt
```

The report records Word runtime, full Word tool smoke, stdio bridge,
tracked-change COM, optional Claude Desktop installation readiness, agent client
prompt evidence, and IRM gates as `passed`, `failed`, `skipped`, or
`blocked_by_runtime`.

The endpoints are:

- MCP: `http://127.0.0.1:8800/mcp`
- Word task pane: `https://localhost:8765/word/taskpane.html`
- Excel task pane: `https://localhost:8765/excel/taskpane.html`
- PowerPoint task pane: `https://localhost:8765/powerpoint/taskpane.html`
- Add-in WSS: `wss://localhost:8765/addin`

For stdio-only MCP clients, keep the daemon running and use the Rust stdio
bridge:

```powershell
cargo run -p office-mcp-daemon -- stdio
```

Claude Desktop-style config can be generated from the repository root with:

```powershell
cargo run -p office-mcp-daemon -- config claude-desktop
```

For an MSI-installed copy, generate a config that points at the installed
PowerShell launcher instead of the source checkout:

```powershell
.\office-mcp.ps1 config claude-desktop --installed
```

The installed launcher sets `OFFICE_MCP_INSTALL_ROOT`, so the generated config
points back to the current install directory. Use `--install-root <path>` only
when generating config for another install location.

Register the shared-folder catalog for source-checkout development:

```powershell
cd ..
powershell -ExecutionPolicy Bypass -File .\src\office-ctl\common\scripts\register-office-catalog.ps1
```

If another local process already owns `https://localhost:8765`, start the
daemon with a different add-in port and render the catalog manifests for that
same origin:

```powershell
$env:OFFICE_MCP_ADDIN_CHANNEL__PORT = '8766'
$env:OFFICE_MCP_MCP_HTTP__PORT = '8801'
$env:OFFICE_MCP_MCP_ENDPOINT = 'http://127.0.0.1:8801/mcp'
powershell -ExecutionPolicy Bypass -File .\src\office-ctl\common\scripts\register-office-catalog.ps1 -BaseUrl https://localhost:8766
cargo run -p office-mcp-daemon -- daemon run
```

The catalog URL shown by the script is the local folder path to use in Office's
trusted add-in catalog settings if Office does not pick up the registry entry.
The Windows bootstrap and MSI install a broader catalog under
`%LOCALAPPDATA%\office-mcp\addin-catalog\` with Word, Excel, and PowerPoint
manifests placed directly at the catalog root as `office-mcp-word.xml`,
`office-mcp-excel.xml`, and `office-mcp-powerpoint.xml`. After restarting
Office, open **Insert > My Add-ins > Shared Folder** in the target host and add
`Office MCP Control`.

For a repeatable Windows developer bootstrap that also registers a logon
Scheduled Task and local add-in catalog, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\packaging\windows\install-windows.ps1
```

Undo that bootstrap with:

```powershell
powershell -ExecutionPolicy Bypass -File .\packaging\windows\uninstall-windows.ps1
```

Build the Windows MSI artifact without mutating local Office or certificate
state:

```powershell
powershell -ExecutionPolicy Bypass -File .\packaging\windows\build-windows-msi.ps1 -SkipNpmInstall
```

The MSI build stages the same split layout under `artifacts\msi-stage\`:
`office-mcp-daemon.exe` for the Rust daemon runtime, daemon-owned UI assets for
the daemon web console, `office-ctl/word/`, `office-ctl/excel/`, and
`office-ctl/powerpoint/` for the Office task pane bundles,
`scripts/` for installer helper scripts, and `addin-catalog/` for the sideload
manifests. It generates the WiX payload fragment and asserts that the Rust
daemon, UI assets, add-in bundles, catalog manifests, and launcher scripts are
present before building the final MSI.

`office-mcp daemon start|stop` works with the developer Scheduled Task bootstrap
and with the MSI launcher layout.

## License

MIT.


