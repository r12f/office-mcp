# office-mcp

**A bridge between AI assistants and live Microsoft Office applications via in-process add-ins.**

`office-mcp` exposes Word (and eventually Excel, PowerPoint, Outlook) as MCP tools by running an
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

MVP implementation is in place for Word desktop on Windows:

- Local MCP Streamable HTTP endpoint at `http://127.0.0.1:8800/mcp`.
- Local HTTPS task pane and WSS add-in channel at `https://localhost:8765`.
- Word task pane add-in that reverse-registers one document session.
- MCP server catalog covers the full Word v1 tool surface from
  `docs/spec/04-word-capabilities.md`.
- The current Word add-in runtime advertises and executes all 27 Word v1 tools: discovery, read, insert, edit, table, structure, review, tracked-change, and save operations.
- The daemon also exposes the Word v1 resource surface, including document text, structure, paragraph, comments, tracked changes, and selection.

The design docs remain the source of truth for the broader v1 surface.

## Verification

CI runs on Windows, macOS, and Linux for the server package, plus Windows add-in
manifest and task pane syntax validation. The equivalent local checks are:

```powershell
cd .\mcp-server
npm run check
cd ..\addin
npm run check
```

## Repository layout

| Path | Owner |
|---|---|
| `mcp-server/` | Long-running local service: MCP frontend, HTTPS/WSS add-in channel, session registry, service tests, and service-owned scripts. |
| `addin/` | Word add-in package: XML manifest, task pane static bundle, add-in validation scripts, and Word catalog registration script. |
| `packaging/` | Cross-component packaging assets: Windows bootstrap scripts and WiX MSI source. |
| `docs/spec/` | Product and protocol design. |

The repository root intentionally does not contain a top-level Node package.
Run server commands from `mcp-server/` and add-in commands from `addin/` so the
daemon and Office add-in stay independently buildable and packageable.
Historical feasibility spikes and generated installer/runtime evidence belong
under `artifacts/`, not as parallel source packages at the root.

## Design docs

| Doc | Purpose |
|---|---|
| [00-overview.md](docs/spec/00-overview.md) | Goals, non-goals, glossary |
| [01-architecture.md](docs/spec/01-architecture.md) | Process model, transports, lifecycle |
| [02-registration-protocol.md](docs/spec/02-registration-protocol.md) | Add-in ↔ server JSON-RPC wire protocol |
| [03-mcp-tool-surface.md](docs/spec/03-mcp-tool-surface.md) | MCP tools, resources, prompts exposed to clients |
| [04-word-capabilities.md](docs/spec/04-word-capabilities.md) | Word-specific tool catalog (v1) |
| [05-security.md](docs/spec/05-security.md) | Auth, sandbox, IRM, trust boundaries |
| [06-error-model.md](docs/spec/06-error-model.md) | Error codes, retry, partial-failure semantics |
| [07-deployment.md](docs/spec/07-deployment.md) | Installation, sideloading, distribution |
| [08-roadmap.md](docs/spec/08-roadmap.md) | Milestones, Excel/PPT/Outlook follow-on |

## Run the MVP locally

Prerequisites:

- Windows with Word desktop installed.
- Node.js 22 or newer.
- A trusted current-user `CN=localhost` HTTPS certificate. The daemon reads a
  PFX file; it does not import root certificates during startup.

Install service dependencies and verify the build:

```powershell
cd .\mcp-server
npm install
npm run check
cd ..\addin
npm install
npm run check
```

Export the already trusted localhost certificate to the daemon's default PFX
path:

```powershell
cd ..\mcp-server
powershell -ExecutionPolicy Bypass -File .\scripts\export-localhost-dev-cert.ps1
```

Start the daemon from `mcp-server/`:

```powershell
npm run daemon
```

Launch the add-in in Word, then open the **office-mcp > Open** ribbon command
so the task pane can connect and register the current document session.

Smoke-test a connected Word session from another terminal in `mcp-server/`:

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
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\artifacts\runtime-evidence-full.json --include-mutation --include-full-word-smoke --include-tracked-changes --include-com-tracked-changes
npm run evidence:validate -- --input ..\artifacts\runtime-evidence-full.json --require-mutation --require-full-word-smoke --require-com-tracked-changes
```

Against a representative protected document, add `--irm-document-path` for a
read-only COM preflight and `--irm-mode protected-read` or `--irm-mode
protected-edit` once that same document has a connected add-in session:

```powershell
npm run evidence:irm
npm run evidence:validate -- --input ..\artifacts\runtime-evidence-irm.json --require-irm-preflight --require-irm
```

For another protected document, use the explicit form:

```powershell
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\artifacts\runtime-evidence-irm.json --include-mutation --irm-document-path "C:\path\to\protected.docx" --irm-mode protected-read --wait-for-session-ms 120000
npm run evidence:validate -- --input ..\artifacts\runtime-evidence-irm.json --require-irm-preflight --require-irm
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
npm run evidence:runtime -- --endpoint http://127.0.0.1:8800/mcp --output ..\artifacts\runtime-evidence-agent-client.json --agent-client-evidence-path ..\artifacts\agent-client-evidence.json
npm run evidence:validate -- --input ..\artifacts\runtime-evidence-agent-client.json --require-agent-client-prompt
```

The report records Word runtime, full Word tool smoke, stdio bridge,
tracked-change COM, optional Claude Desktop installation readiness, agent client
prompt evidence, and IRM gates as `passed`, `failed`, `skipped`, or
`blocked_by_runtime`.

The endpoints are:

- MCP: `http://127.0.0.1:8800/mcp`
- Add-in task pane: `https://localhost:8765/taskpane.html`
- Add-in WSS: `wss://localhost:8765/addin`

For stdio-only MCP clients, keep the daemon running and use the stdio bridge:

```powershell
npm run build
node .\dist\src\cli.js stdio
```

Claude Desktop-style config can be generated from `mcp-server/` with:

```powershell
npx tsx src\cli.ts config claude-desktop
```

For an MSI-installed copy, generate a config that points at the installed
PowerShell launcher instead of the source checkout:

```powershell
.\office-mcp.ps1 config claude-desktop --installed
```

The installed launcher sets `OFFICE_MCP_INSTALL_ROOT`, so the generated config
points back to the current install directory. Use `--install-root <path>` only
when generating config for another install location.

Register the Word shared-folder catalog:

```powershell
cd ..
powershell -ExecutionPolicy Bypass -File .\addin\scripts\register-word-catalog.ps1
```

The catalog URL shown by the script is the local folder path to use in Word's
trusted add-in catalog settings if Word does not pick up the registry entry.
After restarting Word, open **Insert > My Add-ins > Shared Folder** and add
`office-mcp`.

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
`mcp-server/` for the long-running daemon, `addin/` for the Word task pane
bundle, and `addin-catalog/` for the sideload manifest. It installs production
server dependencies there, copies a local `node.exe`, generates the WiX payload
fragment, and asserts that the daemon, add-in bundle, catalog manifest,
launcher scripts, and runtime dependencies are present before building the
final MSI.

`office-mcp daemon start|stop` works with the developer Scheduled Task bootstrap
and with the MSI launcher layout.

## License

TBD (likely MIT, matching the surrounding Office add-in ecosystem).




