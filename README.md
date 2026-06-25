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
  `doc/spec/04-excel-capabilities.md`, and the PowerPoint v1 presentation tools in
  `doc/spec/04-powerpoint-capabilities.md`.
- The current Word add-in runtime advertises and executes the refined 25-tool Word v1 surface from
  `doc/spec/04-word-capabilities.md`, with duplicate specialized compatibility tools retired from
  the advertised catalog.
- The current Excel add-in runtime advertises and executes the refined 20-tool Excel v1 surface:
  workbook info, sheet list/add/update/delete, used-range discovery, range read/write/clear/find-replace,
  formula setting, formatting, sort/filter, table/chart creation and updates, and PivotTable creation and updates.
- The current PowerPoint add-in runtime advertises and executes the 25-tool
  PowerPoint v1 presentation tools for presentation info/export, tags, slides,
  layouts, selection, shapes, text, and tables, with host-gated operations
  returning explicit capability errors when the current PowerPoint runtime does
  not support them.
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
| `packaging/` | Cross-component packaging assets: Windows portable package scripts plus Homebrew/Linux templates. |
| `doc/spec/` | Product and protocol design. |

The repository root intentionally does not contain a top-level Node package.
Run Rust daemon commands from the repository root, evidence/smoke commands from
`src/office-mcp/daemon/evidence/`, add-in commands from `src/office-ctl/word/`,
and packaging checks from `packaging/` so the daemon, evidence harnesses,
Office add-in, and release packages stay independently buildable and packageable.
Historical feasibility spikes and generated package/runtime evidence belong
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
| [04-powerpoint-capabilities.md](doc/spec/04-powerpoint-capabilities.md) | PowerPoint-specific tool catalog (v1) |
| [05-security.md](doc/spec/05-security.md) | Security model, sandbox, IRM, trust boundaries |
| [06-error-model.md](doc/spec/06-error-model.md) | Error codes, retry, partial-failure semantics |
| [07-deployment.md](doc/spec/07-deployment.md) | Installation, sideloading, distribution |
| [08-roadmap.md](doc/spec/08-roadmap.md) | Milestones, Excel/PPT/Outlook follow-on |
| [09-ui.md](doc/spec/09-ui.md) | Daemon tray/main window and add-in task pane UI |

## Install from GitHub Releases

Windows desktop is the v1 portable package target. macOS, Linux, Office on the web,
and managed AppSource deployment are still tracked as later distribution paths.

One-command install for the latest release:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command '$ErrorActionPreference="Stop"; $release=Invoke-RestMethod "https://api.github.com/repos/r12f/office-mcp/releases" | Where-Object { $_.assets.name -like "office-mcp-windows-portable-*-x64.zip" } | Select-Object -First 1; if (-not $release) { throw "No Windows portable release asset found." }; $asset=$release.assets | Where-Object { $_.name -like "office-mcp-windows-portable-*-x64.zip" } | Select-Object -First 1; $installRoot=Join-Path $env:LOCALAPPDATA ("office-mcp\" + $release.tag_name); $zipPath=Join-Path $env:TEMP $asset.name; Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath; New-Item -ItemType Directory -Force -Path $installRoot | Out-Null; Expand-Archive -LiteralPath $zipPath -DestinationPath $installRoot -Force; & (Join-Path $installRoot "install.ps1")'
```

1. Open the latest GitHub Releases page and download
   `office-mcp-windows-portable-<ver>-x64.zip` plus `SHA256SUMS`.
2. Optionally verify the package checksum:

   ```powershell
   Get-FileHash -Algorithm SHA256 .\office-mcp-windows-portable-<ver>-x64.zip
   Get-Content .\SHA256SUMS
   ```

3. Extract the zip to the folder where Office MCP Control should live, read
   `README-install.txt`, close Word/Excel/PowerPoint, and run one command from
   the extracted folder:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
   ```

   The extracted folder is the install directory. `install.ps1` writes the
   current-user Office Trusted Add-in Catalog registry entry under
   `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}`
   creates `.office-mcp-localhost.pfx` in the same folder when needed, and
   starts the tray daemon.
4. Open the daemon UI from the tray menu with **Show Office MCP Control**, or
   run `office-mcp-daemon ui`. Check runtime status with:

   ```powershell
   office-mcp-daemon daemon status
   ```

5. Restart Word, Excel, or PowerPoint if it was already open. If the add-in does
   not appear automatically, open **Insert > My Add-ins > Shared Folder** and
   add **Office MCP Control**.
6. Configure MCP clients to use the local Streamable HTTP endpoint:
   `http://127.0.0.1:8800/mcp`. If the daemon config changes the MCP port, use
   the endpoint reported by `office-mcp-daemon config endpoints` or
   `office-mcp-daemon daemon status`.
7. For debugging, collect the log path reported by `office-mcp-daemon daemon
   status` or shown in the daemon UI. The default log location is under the
   current user's local Office MCP data directory.
8. To uninstall, run `uninstall.ps1` and then delete the extracted folder.
   If Office still lists the add-in, remove **Office MCP Control** from Office's
   add-in manager or remove the installed Shared Folder catalog entry.

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
command. Attach screenshot paths from the same run. The recorder writes
freshness metadata for every tray screenshot surface, and the validator rejects stale screenshots
instead of accepting older files from a previous run:

```powershell
cd .\src\office-mcp\daemon\evidence
npm run evidence:record-tray-manual -- --daemon-bin C:\Code\office-mcp\target\debug\office-mcp-daemon.exe --visible-icon true --right-click-menu true --menu-opened-from-tray-icon true --native-menu-appearance-reviewed true --menu-anchored-to-tray-icon true --os-native-menu-behavior-reviewed true --keyboard-menu-access-reviewed true --native-quit-confirmation-reviewed true --menu-surface-kind native --show-ui-opened true --tooltip "Office MCP Control - Up - 0 clients - 0 documents" --menu-item "Status: Up" --menu-item "Clients: 0" --menu-item "Documents: 0" --menu-item "Show Office MCP Control" --menu-item "Quit Office MCP Control" --screenshot-path C:\path\to\tray-icon.png --tray-icon-screenshot C:\path\to\tray-icon.png --tray-native-menu-screenshot C:\path\to\tray-menu.png --tray-tooltip-screenshot C:\path\to\tray-tooltip.png --tray-quit-confirmation-screenshot C:\path\to\tray-quit.png --screenshot-freshness-window-ms 1800000
node ..\..\..\office-ctl\common\scripts\record-rendered-logo-review.mjs --output ..\..\..\..\artifacts\logo-rendered-size-review.json --sheet ..\..\..\..\artifacts\logo-rendered-size-review.png
node ..\..\..\office-ctl\common\scripts\record-catalog-identity-review.mjs --catalog-path ..\..\..\..\artifacts\portable-stage\addin-catalog --output ..\..\..\..\artifacts\catalog-identity-review.json
npm run evidence:record-product-visual -- --daemon-bin C:\Code\office-mcp\target\debug\office-mcp-daemon.exe --manual-tray-evidence-path ..\..\..\..\artifacts\tray-manual-evidence.json --rendered-logo-review-path ..\..\..\..\artifacts\logo-rendered-size-review.json --catalog-identity-review-path ..\..\..\..\artifacts\catalog-identity-review.json --word-tool-e2e-report-path ..\..\..\..\artifacts\office-tool-e2e-word.json --excel-tool-e2e-report-path ..\..\..\..\artifacts\office-tool-e2e-excel.json --powerpoint-tool-e2e-report-path ..\..\..\..\artifacts\office-tool-e2e-powerpoint.json --word-runtime-evidence-path ..\..\..\..\artifacts\runtime-evidence-word.json --excel-runtime-evidence-path ..\..\..\..\artifacts\runtime-evidence-excel.json --powerpoint-runtime-evidence-path ..\..\..\..\artifacts\runtime-evidence-powerpoint.json --word-ribbon-command "Office MCP Control ribbon command visible" --word-ribbon-command-screenshot C:\path\to\word-ribbon.png --word-catalog-entry "Office MCP Control catalog entry visible" --word-catalog-entry-screenshot C:\path\to\word-catalog.png --word-taskpane-title "Office MCP Control task pane visible" --word-taskpane-title-screenshot C:\path\to\word-taskpane.png --excel-ribbon-command "Office MCP Control ribbon command visible" --excel-ribbon-command-screenshot C:\path\to\excel-ribbon.png --excel-catalog-entry "Office MCP Control catalog entry visible" --excel-catalog-entry-screenshot C:\path\to\excel-catalog.png --excel-taskpane-title "Office MCP Control task pane visible" --excel-taskpane-title-screenshot C:\path\to\excel-taskpane.png --powerpoint-ribbon-command "Office MCP Control ribbon command visible" --powerpoint-ribbon-command-screenshot C:\path\to\powerpoint-ribbon.png --powerpoint-catalog-entry "Office MCP Control catalog entry visible" --powerpoint-catalog-entry-screenshot C:\path\to\powerpoint-catalog.png --powerpoint-taskpane-title "Office MCP Control task pane visible" --powerpoint-taskpane-title-screenshot C:\path\to\powerpoint-taskpane.png --daemon-main-window "Office MCP Control daemon main window visible with compact status details" --daemon-main-window-screenshot C:\path\to\daemon-main-window.png --logo-tray-size "Office MCP Control tray-size logo reviewed" --logo-tray-size-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-ribbon-size "Office MCP Control ribbon-size logo reviewed" --logo-ribbon-size-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-catalog-thumbnail "Office MCP Control catalog thumbnail logo reviewed" --logo-catalog-thumbnail-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-daemon-titlebar "Office MCP Control daemon title-bar logo reviewed" --logo-daemon-titlebar-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --logo-installer-metadata "Office MCP Control installer metadata logo reviewed" --logo-installer-metadata-screenshot ..\..\..\..\artifacts\logo-rendered-size-review.png --tray-icon "Office MCP Control tray icon visible" --tray-icon-screenshot C:\path\to\tray-icon.png --tray-native-menu "Office MCP Control native tray menu visible" --tray-native-menu-screenshot C:\path\to\tray-menu.png --tray-tooltip "Office MCP Control - Up - 0 clients - 0 documents" --tray-tooltip-screenshot C:\path\to\tray-tooltip.png --tray-quit-confirmation "Office MCP Control quit confirmation visible" --tray-quit-confirmation-screenshot C:\path\to\tray-quit.png --catalog-icon-visible true --tray-icon-visible true --tray-menu-native true --quit-confirmation-visible true --logo-quality-reviewed true --logo-future-office-control-reviewed true --final-logo-user-surface-reviewed true --current-logo-screenshot-feedback-reviewed true --rendered-size-logo-reviewed true --addin-identity-reviewed true --addin-title-icon-type-reviewed true --addin-installable-surface-reviewed true --current-addin-screenshot-feedback-reviewed true --word-first-run-identity-reviewed true --excel-first-run-identity-reviewed true --powerpoint-first-run-identity-reviewed true --tray-product-polish-reviewed true --tray-native-first-impression-reviewed true --tray-normal-windows-launch-reviewed true --current-tray-screenshot-feedback-reviewed true --daemon-main-window-reviewed true --daemon-main-window-compact-reviewed true --daemon-main-window-three-column-reviewed true --word-compact-top-block true --word-tools-permissions-merged true --word-inline-settings true --word-server-protocol-row "Server 0.1.0 / Protocol 1.0" --word-document-state "Editable" --excel-compact-top-block true --excel-tools-permissions-merged true --excel-inline-settings true --excel-server-protocol-row "Server 0.1.0 / Protocol 1.0" --excel-document-state "Editable" --powerpoint-compact-top-block true --powerpoint-tools-permissions-merged true --powerpoint-inline-settings true --powerpoint-server-protocol-row "Server 0.1.0 / Protocol 1.0" --powerpoint-document-state "Editable" --screenshot-freshness-window-ms 1800000
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
npm run evidence:word
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-word.json
```

For live tool-level E2E evidence, enable the live driver and run one host loop
per Office app. Each command starts the daemon, opens one driver-owned
file/session for that host, connects the add-in once, loops all advertised
tools in that session, cleans up once, and writes
`artifacts/office-tool-e2e-<host>.json`:

The release gate must not restart Office, recreate the document, or reconnect
the add-in per tool. Per-tool Office launches are diagnostic only; the stable
path is one opened host program, one connected session, and one table-driven
loop across the host's complete tool catalog.

Release-ready tool E2E reports must prove the add-in activation step. On
Windows, the default activator is
`src/office-ctl/common/scripts/activate-office-mcp-addin.ps1`; it activates the
driver-owned document and opens the **Open Control Panel** ribbon command before
the session wait. Set `OFFICE_MCP_E2E_ACTIVATOR` to override it with another
command, or set `OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR=0` for manual debugging.
The driver passes `OFFICE_MCP_E2E_HOST`, `OFFICE_MCP_E2E_DOCUMENT_PATH`,
`OFFICE_MCP_E2E_ADDIN_ORIGIN`, and `OFFICE_MCP_E2E_ADDIN_ENDPOINT` to the
activator. Reports with `no-activator-configured` are useful for local manual
debugging, but `npm run evidence:validate -- --require-office-tool-e2e` rejects
them. Release-ready reports must also include the concrete add-in activator identity
and a non-empty `activation_path`; weak activation proof such as only
`activated: true` is rejected. Cleanup proof must include `deleted_paths` with
the concrete cleanup paths for the driver-owned original file and any Office
sideload copies that were closed and deleted; a count-only cleanup report is not
release-ready.

```powershell
cd ..\..\office-ctl\word
$env:OFFICE_MCP_RUN_E2E = '1'
$env:OFFICE_MCP_E2E_ACTIVATOR = 'powershell -NoProfile -ExecutionPolicy Bypass -File C:\path\to\activate-office-mcp-addin.ps1'
npm run e2e:tools
cd ..\excel
npm run e2e:tools
cd ..\powerpoint
npm run e2e:tools
cd ..\..\office-mcp\daemon\evidence
npm run evidence:validate -- --input ..\..\..\..\artifacts\runtime-evidence-word.json --require-office-tool-e2e --word-tool-e2e-report-path ..\..\..\..\artifacts\office-tool-e2e-word.json --excel-tool-e2e-report-path ..\..\..\..\artifacts\office-tool-e2e-excel.json --powerpoint-tool-e2e-report-path ..\..\..\..\artifacts\office-tool-e2e-powerpoint.json
```

For a connected Excel workbook, run:

```powershell
cd ..\..\office-ctl\excel
$env:OFFICE_MCP_RUN_E2E = '1'
npm run e2e:tools
cd ..\powerpoint
npm run e2e:tools
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
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\..\..\..\artifacts\runtime-evidence-irm.json --irm-document-path "C:\path\to\protected.docx" --irm-mode protected-read --wait-for-session-ms 120000
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

For a portable package copy, generate a config that points at the extracted
daemon executable instead of the source checkout:

```powershell
.\office-mcp-daemon.exe config claude-desktop --installed --install-root <extracted-folder>
```

The generated config points the MCP client at `office-mcp-daemon.exe` and sets
the portable runtime environment for that client process.

Register the shared-folder catalog for source-checkout development:

```powershell
cd ..
powershell -ExecutionPolicy Bypass -File .\src\office-ctl\common\scripts\register-office-catalog.ps1
```

The catalog registration script also clears stale Office WEF add-in caches for
Office MCP Control so versioned task pane URLs refresh after upgrades. Close
Word, Excel, and PowerPoint before running it; the script fails fast if those
hosts are still open instead of leaving Office to reuse an older cached
manifest. Use `-SkipOfficeCache` only for diagnostics where the current Office
cache must be preserved.

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
The Windows bootstrap and portable package install a broader catalog under
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

Build the Windows portable artifact without mutating local Office or certificate
state:

```powershell
powershell -ExecutionPolicy Bypass -File .\packaging\windows\build-windows-portable.ps1 -SkipNpmInstall
```

The portable build stages the same split layout under `artifacts\portable-stage\`:
`office-mcp-daemon.exe` for the Rust daemon runtime, daemon-owned UI assets for
the daemon web console, `office-ctl/word/`, `office-ctl/excel/`, and
`office-ctl/powerpoint/` for the Office task pane bundles,
`scripts/` for certificate helper scripts, and `addin-catalog/` for the sideload
manifests. It asserts that the Rust daemon, UI assets, add-in bundles, catalog
manifests, `install.ps1`, and `uninstall.ps1` are present before building the
final zip.

`office-mcp daemon start|stop` works with the developer Scheduled Task bootstrap.

## License

MIT.


