# Release Notes

## 0.1.10

Office MCP Control 0.1.10 tightens PowerPoint and MCP discovery behavior after
the 0.1.9 tool-contract release. PowerPoint presentation resources now report
real slide counts, the resource fallback includes collection-level slide text,
and callable PowerPoint action tools are exposed through MCP `tools/list` so
session metadata and the client-visible tool surface stay aligned. MCP HTTP
rate limiting now separates read-only discovery traffic from document operation
traffic and returns structured retry metadata for 429 responses.

Expected assets:

- `office-mcp-windows-portable-0.1.10-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- `powerpoint.get_presentation_info` and the
  `office://powerpoint/<session_id>/presentation` resource report the actual
  slide count instead of `null`.
- PowerPoint collection text resources can read slide ranges through
  `office://powerpoint/<session_id>/slides/text{?start,end}` while preserving
  single-slide text resources.
- PowerPoint tools advertised in session metadata, including mutating action
  tools such as `powerpoint.add_slide`, `powerpoint.add_text_box`, and
  `powerpoint.update_shape`, are visible through MCP `tools/list` when Global
  Tool Access allows them and still route through session capability checks.
- Read-only MCP discovery requests such as `initialize`, `tools/list`,
  `resources/list`, `resources/templates/list`, `office://sessions`, and
  `office.describe_tools` use an independent discovery rate-limit budget;
  operation calls remain protected by the normal operation budget.
- HTTP 429 responses include `Retry-After`, and JSON-RPC-shaped requests return
  structured `RATE_LIMITED` error metadata.
- Release CI builds the Windows portable artifact after Rust daemon, Word,
  Excel, PowerPoint, evidence, packaging, rendered-logo, and catalog-identity
  gates pass.

## 0.1.9

Office MCP Control 0.1.9 improves MCP client planning and resource discovery
confidence. The daemon now exposes batch-only tool contract discovery through
`office.describe_tools`, including client-friendly top-level parameter
summaries, and adds regression coverage that keeps `resources/read` routing
aligned with `resources/list` and `resources/templates/list` for Word, Excel,
and PowerPoint.

Expected assets:

- `office-mcp-windows-portable-0.1.9-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- MCP clients can request multiple detailed tool contracts with one
  `office.describe_tools` call; the previous `office.describe_tool` helper is
  no longer exposed as a public tool.
- Returned tool contracts include `input_schema`, top-level `parameters`,
  examples, side-effect classification, app/category metadata, and common error
  hints.
- Unknown requested tool names return structured per-entry `UNKNOWN_TOOL`
  results without failing the whole batch.
- Resource discovery coverage proves every accepted Word, Excel, and PowerPoint
  read-only resource route is represented in either `resources/list` or
  `resources/templates/list`.
- Release CI builds the Windows portable artifact after Rust daemon, Word,
  Excel, PowerPoint, evidence, packaging, rendered-logo, and catalog-identity
  gates pass.

## 0.1.8

Office MCP Control 0.1.8 improves agent-facing Word tool planning. MCP
`tools/list` now exposes rich schemas, examples, side-effect metadata, and
common error hints directly, and Word mutating tools now support validation-only
planning calls before writes are queued.

Expected assets:

- `office-mcp-windows-portable-0.1.8-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- MCP clients can construct Office tool calls from `tools/list` metadata without
  inspecting add-in JavaScript or making a secondary discovery call.
- `office.describe_tools` remains aligned with the richer `tools/list` contract
  for schemas, top-level parameter summaries, examples, side-effect
  classification, app/category metadata, and common error hints.
- `word.insert_image`, `word.replace_text`, `word.update_paragraph`, and
  `word.delete_range` support `validate_only: true` or the existing dry-run
  alias where documented, returning no-mutation planning results or structured
  no-effect validation failures.
- Release CI builds the Windows portable artifact after Rust daemon, Word,
  Excel, PowerPoint, evidence, packaging, rendered-logo, and catalog-identity
  gates pass.

## 0.1.7

Office MCP Control 0.1.7 fixes the native Windows tray menu actions for normal
portable installs. The tray **Show Office MCP Control** action now opens the
daemon UI through the runtime UI URL, and **Quit Office MCP Control** now routes
through a tested daemon shutdown dispatcher with actionable logging.

Expected assets:

- `office-mcp-windows-portable-0.1.7-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- The native tray **Show Office MCP Control** action resolves the daemon runtime
  UI URL, normally `https://localhost:8765/ui/`, and logs action, source, URL,
  process ID, and launcher errors when opening fails.
- The native tray **Quit Office MCP Control** action shows native confirmation,
  dispatches daemon shutdown through the tray shutdown controller, and logs
  action, source, process ID, and controller errors when shutdown fails.
- The tray actions work from the fixed Windows install root
  `%LOCALAPPDATA%\office-mcp` used by the portable installer.
- Rust tray tests cover Show UI dispatch, Quit dispatch, and actionable failure
  messages before the release artifact is built.

## 0.1.6

Office MCP Control 0.1.6 redesigns the Windows portable installer so upgrades
use a stable install root instead of creating a new versioned directory for each
release.

Expected assets:

- `office-mcp-windows-portable-0.1.6-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- The one-line installer downloads the latest portable zip to a temporary
  staging directory and invokes the package-local `install.ps1`.
- The package-local installer installs or upgrades `%LOCALAPPDATA%\office-mcp`
  by default, with custom roots supported by `OFFICE_MCP_INSTALL_ROOT` and
  `-InstallRoot`.
- Upgrade stops existing Office MCP daemon processes before replacing runtime
  files, preserves config/certificate/log files, removes safe-to-identify stale
  versioned install roots, and starts `office-mcp-daemon.exe daemon run` from the
  fixed install root.
- The trusted Office catalog registry entry points at the fixed install root's
  `addin-catalog` folder and stale Office MCP catalog paths are removed.
- If Word, Excel, or PowerPoint is running, interactive installs ask before
  closing them; non-interactive installs fail unless `-CloseOfficeHosts` is
  supplied.

## 0.1.5

Office MCP Control 0.1.5 fixes the daemon control panel in the Windows portable
one-line install path. The portable package already included `office-mcp\ui`,
but the daemon only discovered the source-tree UI assets, so installed daemons
returned `404 Not found` for `https://localhost:8765/ui/`.

Expected assets:

- `office-mcp-windows-portable-0.1.5-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- The one-line installer downloads this release and expands a portable package
  whose daemon can discover `office-mcp\ui` from the install root.
- `https://localhost:8765/ui/` returns `HTTP/1.1 200 OK` with the Office MCP
  Control HTML after install.
- The installer writes the GUID-based trusted catalog registry entry and clears
  stale Office WEF add-in caches.
- The portable zip root contains `office-mcp-daemon.exe`, `install.ps1`,
  `uninstall.ps1`, `config.toml`, and `README-install.txt`.
- The CI portable zip verification uses `packaging/package.json` instead of a
  hard-coded historical version.

## 0.1.4

Office MCP Control 0.1.4 fixes Office trusted catalog registration so the
installer no longer reports success while Word, Excel, or PowerPoint still has
the old catalog state loaded.

Expected assets:

- `office-mcp-windows-portable-0.1.4-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- The package-local `install.ps1` refuses to continue while Word, Excel, or
  PowerPoint is running. The installer error includes the text
  `Word, Excel, or PowerPoint is running` because Office only reloads trusted
  add-in catalogs on startup.
- The installer writes the GUID-based trusted catalog registry entry, clears
  stale Office WEF add-in caches, and tells the user to reopen Office.
- The portable zip root contains `office-mcp-daemon.exe`, `install.ps1`,
  `uninstall.ps1`, `config.toml`, and `README-install.txt`.
- The one-line installer downloads this release and starts
  `office-mcp-daemon.exe daemon run` for `http://127.0.0.1:8800/mcp`.

## 0.1.3

Office MCP Control 0.1.3 fixes the Windows one-line installer so installation
starts the actual daemon runtime, not only the standalone tray process.

Expected assets:

- `office-mcp-windows-portable-0.1.3-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- The one-line installer downloads the latest published portable zip and runs
  the package-local `install.ps1` successfully.
- The package-local `install.ps1` starts `office-mcp-daemon.exe daemon run`,
  which serves the MCP endpoint, add-in endpoint, UI runtime file, and tray.
- The portable zip root contains `office-mcp-daemon.exe`, `install.ps1`,
  `uninstall.ps1`, `config.toml`, and `README-install.txt`.
- The portable zip root does not contain duplicate launcher wrappers such as
  `office-mcp.ps1`, `office-mcp-daemon.ps1`, `office-mcp-tray.ps1`, or
  `office-mcp-env.ps1`.
- MCP clients can connect to `http://127.0.0.1:8800/mcp`.

## 0.1.2

Office MCP Control 0.1.2 republished the Windows portable package from the
current portable-only layout, but its installer started only the tray process.
Use 0.1.3 or newer for the one-line installer.

## 0.1.1

Office MCP Control 0.1.1 was intended to be the first Windows desktop
portable-only pre-release, but its release artifact was built from the older
portable layout. Use 0.1.3 or newer for the one-line installer.

## 0.1.0

Initial prerelease artifacts. This release is superseded by portable releases.
