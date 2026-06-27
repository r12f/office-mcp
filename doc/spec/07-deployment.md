# 07 — Deployment

How the server and the add-in get onto a user's machine and start talking to
each other.

## 1. Components to ship

| Component | What it is | Where it lives | Updated how |
|---|---|---|---|
| `office-mcp` | Native Rust long-running daemon from `src/office-mcp/daemon`. | Stable user-selected install root (Win) / `/usr/local/bin/office-mcp` wrapper target (Mac, Linux) | Portable zip / Homebrew tap |
| `office-mcp-ui` | Web UI assets owned by `src/office-mcp/daemon/src/ui`, opened from the tray and served or bridged by the daemon. | Installed beside the daemon | Portable zip / Homebrew tap |
| `office-ctl` | Office add-in bundles from `src/office-ctl`: shared `common` code plus host entries such as `word` and `excel`. | Installed beside the daemon and served from its trusted local HTTPS origin | Installer / atomic local replacement |
| Manifest | XML / JSON describing the add-in | Sideloaded via the trusted catalog by the installer; AppSource / M365 admin push for managed deployments | See §3 |
| Bootstrap package | Portable zip / .pkg / shell script | Downloaded from GitHub Releases | Per-release |

The target source tree is:

```text
doc/                 # Specifications and design documentation.
src/office-ctl/      # TypeScript Office add-ins: common, word, excel, powerpoint.
src/office-mcp/
  daemon/            # Rust daemon service, daemon-owned state/API, and daemon UI.
packaging/           # Installers and release assembly.
```

The actual installation procedure — including daemon autostart and add-in
catalog registration — is in §6, not here. §1 is just the artifact list.

The production daemon is the native Rust executable built from
`src/office-mcp/daemon`. Its main-window web assets are owned by the daemon's
`ui` module under `src/office-mcp/daemon/src/ui/assets` and packaged beside the
executable. Installers MUST stage the Rust daemon as the runtime and keep using
the protocol, runtime, UI, tray,
redaction, and packaging evidence gates as release checks.

## 3. Add-in distribution

Office add-ins can be deployed three ways. They share one add-in ID and
capability model, but deployment-specific manifest variants may use different
source URLs and activation settings.

### 3.1 Sideload (developer / individual user)

```
# Word desktop on Windows
1. Open Word → File → Options → Trust Center → Trusted Add-in Catalogs.
2. Add a trusted catalog path (a network share or local folder).
3. Drop the Word, Excel, and PowerPoint manifest XML files in that folder.
4. Restart the target Office host.
5. Insert → My Add-ins → Shared Folder → office-mcp → Add.
```

The Windows portable package includes `addin-catalog\` and `install.ps1`
pre-registers it as a trusted catalog when the user runs the script. Host
manifests live directly under the catalog root with stable file names such as
`addin-catalog\office-mcp-word.xml`,
`addin-catalog\office-mcp-excel.xml`, and
`addin-catalog\office-mcp-powerpoint.xml` so Word, Excel, and PowerPoint can
appear in Office's Shared Folder add-in picker without relying on recursive
catalog scanning.

The shared-folder catalog must present the same product identity users see in
the ribbon and task pane: polished product title, product icon, provider,
description, and local automation/control category or type metadata where the
manifest or catalog format supports it. The user-facing catalog entry must not
show raw package names, blank icons, generic Office icons, debug wording, or a
host-only name that hides the product family.

### 3.2 AppSource (eventual)

Published via [partner.microsoft.com](https://partner.microsoft.com). Users
install with one click from Office's Add-ins picker. Requires:

- Microsoft Partner Center account
- AppSource validation review
- Hosted manifest URL (HTTPS, public)

The repository can generate the pre-submission artifact set, but it cannot
complete the external Partner Center submission or Microsoft's validation
review:

```
powershell -ExecutionPolicy Bypass -File .\src\office-ctl\word\scripts\build-appsource-package.ps1 `
  -Version <package-version> `
  -BaseUrl https://office-mcp.dev `
  -AddinId <release-guid> `
  -AddinVersion <office-four-part-version>
```

The generated package contains the hosted manifest, add-in static bundle,
checksums, and a submission checklist. The checklist explicitly records the
remaining external gates: public hosting, Office webview validation from the
hosted origin, Partner Center listing metadata, and AppSource review.

### 3.3 Centralized deployment (enterprise)

M365 admins deploy via Microsoft 365 admin center → Settings → Integrated apps.
Users get the add-in automatically through a centrally managed manifest
variant.

## 4. Manifest

v1 ships an **add-in-only XML manifest**. Microsoft's unified manifest support
for Excel, PowerPoint, and Word remains preview and is not a production basis
for this project.

The production XML manifest:

- Declares a Word task-pane add-in with `ReadWriteDocument`.
- Requires `WordApi 1.3` for activation.
- Uses `VersionOverridesV1_0` and `AddinCommands 1.1` for the ribbon command.
- Loads `https://localhost:8765/taskpane.html` in local sideloaded builds.
- Loads production Office.js from Microsoft's CDN from within that page.
- Probes `WordApi 1.4`, `WordApi 1.6`, and desktop-only sets at runtime.

The checked developer manifests live under `src/office-ctl/<host>/`, such as
`src/office-ctl/word/manifest.xml`. Running the host add-in check validates the
manifest against the current Office add-in schemas and checks the task pane
TypeScript bundle. The release build substitutes the real add-in ID, version,
icon assets, and support URL. The XML manifest's four-part version starts at
`1.0.0.0` because Office rejects values below 1.0; it is mapped to, but not
textually identical with, the add-in package semver.

Release manifests MUST also substitute product-quality add-in identity fields:
`DisplayName`, ribbon group labels, command labels, description, provider name,
support URL, add-in type/category metadata where supported, and all icon URLs.
These values must match the product identity in [09-ui.md](09-ui.md), avoid
developer placeholders such as `office-mcp` and generic command labels such as
`Open`, and must not imply Microsoft ownership. Manifest and catalog renderers
must fail if required identity fields are missing, still point at placeholder
assets, or expose inconsistent names between the catalog, ribbon command, and
task pane title.

The hosted manifest is rendered from the checked developer manifest rather than
maintained as a second XML file:

```
powershell -ExecutionPolicy Bypass -File .\src\office-ctl\word\scripts\render-hosted-manifest.ps1 `
  -BaseUrl https://office-mcp.dev `
  -AddinId <release-guid> `
  -AddinVersion <office-four-part-version> `
  -AssetVersion <package-version>
```

The canonical home for add-in build and manifest scripts is
`src/office-ctl/<host>/scripts/`; legacy top-level wrapper paths must not remain.

The renderer refuses non-HTTPS and loopback origins, replaces task-pane and icon
URLs with the public base URL, and fails if any loopback URL remains in the
output. Publishing `https://office-mcp.dev/manifest.xml` is still a release
hosting task; the checked renderer provides the reproducible artifact.

Although the XML schema permits other Word platforms, the v1 task pane checks
`Office.onReady()` and displays an unsupported-platform message unless the host
is Word desktop on Windows. Store submission is deferred until every platform
implied by its manifest requirements is supported.

An AppSource build cannot use a loopback `SourceLocation`. It requires a
publicly hosted HTTPS task pane and separate validation that the Office webview
may connect from that origin to the local WSS daemon. AppSource is therefore a
future deployment track, not a packaging variant promised by v1.

## 5. Configuration

`office-mcp` uses a native TOML config file for the daemon. The Office add-in
cannot read this file because it runs in a browser/webview sandbox.

Location:

| OS | Path |
|---|---|
| Windows | `%APPDATA%\office-mcp\config.toml` |
| macOS | `~/Library/Application Support/office-mcp/config.toml` |
| Linux | `~/.config/office-mcp/config.toml` |

The add-in derives its WSS endpoint from the HTTPS origin that loaded the task
pane. The checked developer manifest loads from `https://localhost:8765`, so it
connects to `wss://localhost:8765/addin`. An endpoint override may be stored in
the add-in's partitioned browser storage through its settings UI.
Installer-managed builds may compile a different manifest origin. The daemon
CLI prints the endpoint values that should be entered:

```
office-mcp config endpoints
```

```toml
# ─── daemon configuration; endpoint values are mirrored in add-in settings ──
[addin_channel]
# HTTPS/WSS origin the daemon serves and the add-in dials.
bind = "localhost"
port = 8765
heartbeat_interval_sec = 30
heartbeat_timeout_sec  = 10
session_grace_sec = 60
max_pending_per_session = 4
certificate_path = ""            # installer-managed default
private_key_path = ""            # ACL restricted to the current user
# These certificate fields are HTTPS/WSS transport settings.

# session_grace_sec defaults to 60 and is capped by the daemon at 300 seconds.
# Closed document sessions are deleted after this grace period; they are not
# retained as historical sessions.

# ─── daemon only ─────────────────────────────────────────────────────────
[mcp_http]
# MCP Streamable HTTP frontend the daemon exposes for MCP clients.
bind = "127.0.0.1"
port = 8800
# v1 is loopback-only. Non-loopback bind = startup refusal.

[limits]
max_response_bytes = 1048576     # 1 MiB
max_request_bytes = 16777216      # 16 MiB; supports a 10 MiB base64 image
max_ws_frame_bytes = 16777216
default_tool_timeout_ms = 30000
requests_per_minute = 120

[audit]
enabled = false
path = ""                        # default: %LOCALAPPDATA%\office-mcp\audit.jsonl

[logging]
level = "info"                   # trace | debug | info | warn | error
file  = ""                       # default: platform log dir
```

The Rust daemon uses the `tracing` ecosystem for structured logs. Production
runs MUST write logs to a file, either the configured `logging.file` path or the
platform default log location when that field is empty. The effective log path
is part of daemon diagnostics and must be visible from `daemon status` and the
daemon UI. Logs must redact document body content, inserted/replacement text,
base64 image data, certificate passphrases, and other sensitive configuration
values while preserving request/session/tool/error context for debugging.

Environment variables override daemon config keys, prefixed `OFFICE_MCP_`
(e.g. `OFFICE_MCP_ADDIN_CHANNEL__PORT=9000`). They are not visible to the
web add-in. The section-style names matching the TOML structure are canonical;
legacy flat names such as `OFFICE_MCP_ADDIN_PORT` remain accepted for existing
developer scripts. These variables configure endpoints and limits only; they
are daemon process configuration, not web add-in storage.

## 6. Installation and first-run flow

Installation has two prerequisites: the daemon must be installed and
autostarted by the OS, and the add-in must be registered with Office. Both
are done by the platform package (portable zip / .pkg / brew formula); the user
should never have to hand-edit autostart entries or sideload from a developer
console for a production install.

### 6.1 Windows

The repository includes a developer bootstrap script that performs the same
user-scoped installation shape from a source checkout:

```
powershell -ExecutionPolicy Bypass -File .\packaging\windows\install-windows.ps1
```

It validates the protocol/evidence checks and manifest, builds the native Rust
daemon, installs the daemon web assets, exports an already trusted localhost
certificate to `%LOCALAPPDATA%\office-mcp\`, registers the Office trusted catalog,
and creates a logon Scheduled Task that starts the tray. The tray starts the
installed Rust daemon executable. It does not import root certificates. It can
be removed with:

```
powershell -ExecutionPolicy Bypass -File .\packaging\windows\uninstall-windows.ps1
```

The current Windows release package is a transparent portable zip installer
source for the native Rust daemon. It contains `office-mcp-daemon.exe`, daemon-owned UI assets,
the Word/Excel/PowerPoint add-in bundles, catalog manifests, default
`config.toml`, product assets, and auditable `install.ps1` / `uninstall.ps1`
scripts. The package MUST NOT expose multiple user-facing PowerShell launchers
for the same native daemon. `office-mcp-daemon.exe` is the product binary; it
owns the tray, daemon, UI, status, and MCP server subcommands. The user-run
`install.ps1` installs or upgrades the package into a stable install root,
registers the Office trusted catalog, creates or exports the localhost
certificate, sets any portable runtime environment needed for the child process,
and directly starts `office-mcp-daemon.exe daemon run`. MSI packaging is not a
release target.

The root bootstrap script at `scripts/install.ps1` is only a downloader. It MUST
resolve the latest published Windows portable release asset, extract it to a
temporary staging directory, and invoke the package-local `install.ps1` with the
selected install root. It MUST NOT install product files into a version-tagged
directory or duplicate package-local install semantics. The default Windows
install root is stable across releases, `%LOCALAPPDATA%\office-mcp` unless the
user supplies another path through `OFFICE_MCP_INSTALL_ROOT` for the one-line
bootstrap flow or `-InstallRoot` for the package-local installer.

Upgrade installs MUST stop the running Office MCP daemon/tray before replacing
runtime files, replace the installed payload from the new package without
mixing old and new UI/add-in assets, and preserve user config, certificates,
logs, and audit files unless an uninstall or purge path explicitly removes
them. Installers MUST remove stale versioned install roots created by earlier
bootstrap releases when they can be identified as Office MCP roots, and MUST
avoid creating new version-tagged roots by default.

Office trusted catalog registration is installer-owned. Users must not need to
add the catalog manually in Trust Center. Reinstall and upgrade MUST use a
single stable current-user Office MCP trusted catalog entry, update it to the
fixed install root's `addin-catalog` folder, and remove stale Office MCP catalog
entries that point at previous versioned install roots. A reinstall that leaves
multiple Office MCP catalog paths behind is an installer bug.

If Word, Excel, or PowerPoint is running, the installer MUST list the running
Office hosts and request explicit confirmation before closing them. After
confirmation it should close those hosts for the user so Office can reload the
trusted catalog on next start. In non-interactive mode the installer MUST fail
with an actionable message unless an explicit close/force option is supplied.

The production portable zip remains the Windows release packaging target. A
checked-in build script is not enough for user distribution: release builds MUST
be produced by GitHub Actions from an immutable version tag and published as
GitHub Release assets so a non-developer can download and inspect the product
without cloning the repository.

The GitHub Release pipeline MUST:

- Trigger on version tags such as `v0.1.0` and support an explicit
  `workflow_dispatch` dry run for maintainers.
- Build the Windows portable zip on `windows-latest` from the tagged source using
  `packaging/windows/build-windows-portable.ps1 -SkipNpmInstall` after installing
  Rust, Node, and all add-in dependencies.
- Run the packaging smoke checks, PowerShell parser checks, manifest checks,
  daemon tests needed for packaging confidence, and `git diff --check` before
  publishing any artifact.
- Produce versioned artifacts under `artifacts/`, including
  `office-mcp-windows-portable-<ver>-x64.zip`, `SHA256SUMS`, and later any tarballs,
  hosted manifests, or add-in bundles that become supported release outputs.
- Verify the portable zip artifact is non-empty, has the expected versioned file name,
  contains the native Rust daemon, daemon UI assets, shared Office add-in
  assets, Word/Excel/PowerPoint task pane bundles, generated catalog manifests,
  `install.ps1`, `uninstall.ps1`, and installer icon assets. It must fail if
  duplicate user-facing daemon launcher scripts such as `office-mcp.ps1`,
  `office-mcp-daemon.ps1`, or `office-mcp-tray.ps1` are present in the portable
  package root.
- Upload artifacts with `actions/upload-artifact` for every run and attach them
  to a GitHub Release with `softprops/action-gh-release` or `gh release` only
  when the run is for a version tag.
- Publish draft releases by default until signing, release notes, manual tray
  evidence, and any required live Office evidence are attached or explicitly
  waived for a pre-release.
- Keep signing as a separate gate: if signing secrets are unavailable, the
  pipeline may publish unsigned pre-release artifacts but MUST label them as
  unsigned and MUST still publish `SHA256SUMS`.
- Fail rather than publish when the version in the tag, portable zip file name, package
  metadata, and release notes disagree.

The repository README MUST include a user-facing installation guide before the
developer setup section. The guide MUST explain how to install from GitHub
Releases, not only how to build from source. At minimum it must cover:

- Supported platform status, with Windows desktop as the v1 portable package target.
- Running the one-line latest-release bootstrap command
  `irm https://raw.githubusercontent.com/r12f/office-mcp/main/scripts/install.ps1 | iex`.
- Explaining that the bootstrap downloads the latest portable zip to a temporary
  staging directory, invokes the package-local `install.ps1`, and installs into
  the stable `%LOCALAPPDATA%\office-mcp` root by default instead of a
  version-tagged directory.
- Explaining how to choose a custom install root with `OFFICE_MCP_INSTALL_ROOT`
  for one-line installs and `-InstallRoot` for manual package installs.
- Manual fallback: downloading `office-mcp-windows-portable-<ver>-x64.zip` from
  the latest GitHub Release, verifying `SHA256SUMS` when desired, extracting the
  portable zip, reading `README-install.txt`, and running
  `powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1`.
- Opening the daemon UI from the tray or `office-mcp-daemon ui` and checking
  `daemon status`.
- Restarting Office if needed, opening Word/Excel/PowerPoint, and adding
  `Office MCP Control` from the Shared Folder catalog when Office does not
  auto-show it.
- Configuring MCP clients to use `http://127.0.0.1:8800/mcp` or the configured
  endpoint.
- Where logs live and how to collect them for debugging.
- How to uninstall with `uninstall.ps1` and delete the fixed install root when
  appropriate, including the separate Office add-in removal caveat.

The Windows user flow remains:

1. User runs the one-line bootstrap or extracts
   `office-mcp-windows-portable-<ver>-x64.zip` and runs `install.ps1`. The
   package:
   - Places the native Rust daemon executable in the stable selected install root.
   - Installs the generated product logo, tray icon, add-in command icons, and
     main-window UI assets.
   - Verifies the installed add-in catalog resolves product metadata and icon
     paths for Word, Excel, and PowerPoint before declaring installation
     complete once each host is packaged.
   - Installs the static add-in bundle beside the daemon.
   - Exports or creates a current-user trusted localhost certificate when the
     user runs `install.ps1`; it does not require an opaque installer step.
   - Writes or preserves a default `config.toml` and points the daemon at it from
     `install.ps1` with `OFFICE_MCP_CONFIG_PATH` or equivalent process-scoped
     configuration.
   - Starts the daemon runtime by launching `office-mcp-daemon.exe daemon run` from
     `install.ps1`, so the user does not need a separate daemon startup script
     during install.
   - Registers an add-in trusted catalog folder under the fixed install root's
     `addin-catalog\` directory and drops the Word, Excel, and PowerPoint
     manifests directly under that catalog root once each host is packaged.
   - Stops any existing Office MCP daemon, replaces installed package payloads,
     removes safe-to-identify stale versioned install roots and stale Office MCP
     trusted catalog paths, then starts the daemon from the fixed install root.
   - Future production builds should start the daemon once so the user does not
     have to log out / log in for the daemon to come up.
   - Future production builds should make the tray icon visible immediately
     after install and expose the main window from a native tray menu.

2. User configures their MCP client to connect to
   `http://127.0.0.1:8800` (or whatever `mcp_http.port` is in their config).

3. User opens Word and launches the add-in from the ribbon. Auto-open is an
   opt-in document setting for sideloaded or centrally deployed builds; it is
   not assumed to work for Marketplace distribution.

4. User asks the MCP client to do something with the doc.

### 6.2 macOS / Linux

Same shape, different mechanism: launchd agent (macOS) or systemd `--user`
unit (Linux) replaces the Scheduled Task. brew formula / Linux package
performs the equivalent setup.

The checked Homebrew packaging assets live under `packaging/homebrew/`:

- `Formula/office-mcp.rb.in` is the formula template for a release tarball.
- `render-formula.ps1` renders the formula with the release tarball URL and
  SHA-256 digest.

The formula installs the release tarball under Homebrew `libexec`, writes an
`office-mcp` wrapper that sets `OFFICE_MCP_INSTALL_ROOT` and
`OFFICE_MCP_CONFIG_PATH`, and defines a `brew services` daemon running
`office-mcp daemon run`. macOS Office trusted-catalog registration and
localhost certificate trust remain explicit first-run or enterprise deployment
steps; the formula does not silently mutate those trust stores.

### 6.3 Verifying

```
office-mcp daemon status     # are the MCP and add-in listener ports up?
office-mcp daemon stop/start # start or stop the Windows autostart integration
office-mcp ui                # open or focus the daemon main window
office-mcp config show       # show effective config
office-mcp sessions          # list documents with a connected add-in runtime
```

The tray icon and main window are part of production verification, not optional
developer diagnostics. UI behavior and state redaction are specified in
[09-ui.md](09-ui.md).

Production verification MUST include the normal daemon path, not only fixture
or probe commands:

- `office-mcp-daemon daemon run` writes the UI runtime file and serves `/ui/`,
  `/ui/state`, and `/ui/events` on the configured local HTTPS origin.
- `office-mcp-daemon daemon status` reports the current UI URL and state URL.
- `office-mcp-daemon ui` opens or prints the current daemon UI URL and fails
  clearly when no daemon UI server is running.
- The Windows tray launch path creates a visible notification-area icon in an
  interactive user session. `tray --probe` is useful automated coverage, but it
  is not sufficient evidence that the user can see or use the tray.
- The tray `Show Office MCP Control` action opens or focuses the same UI URL reported by
  `daemon status`.

## 7. Versioning & upgrades

- **Server version** is reported in `register.result.server_version`.
- **Add-in version** is reported in `register.params.add_in.version`.
- The server and add-in protocols are versioned independently (see
  [01-architecture.md §6](01-architecture.md)).
- Auto-update:
  - Server: opt-in update check at startup (GitHub Releases API; can be disabled).
  - Sideloaded add-in: updated atomically by the platform installer.
  - Marketplace / centrally deployed add-in: updated through that deployment
    channel.

## 8. Uninstall

`uninstall.ps1` removes:

- Binaries
- Scheduled task / launchd plist
- Local certificate and private key created by the installer
- `%LOCALAPPDATA%\office-mcp\` (audit log, sideload catalog)
- `%APPDATA%\office-mcp\` (config) — only with `--purge`

The add-in must be removed separately via Office's add-in manager
(or by removing the manifest from the trusted catalog).

## 9. Release artifacts (per version tag)

| File | Purpose |
|---|---|
| `office-mcp-<ver>-x64.exe` | Standalone Windows binary |
| `office-mcp-windows-portable-<ver>-x64.zip` | Windows portable package |
| `office-mcp-<ver>-x64.tar.gz` | Linux/macOS tarball |
| `office-mcp-<ver>-aarch64-darwin.tar.gz` | Apple Silicon |
| `manifest-<ver>.xml` | XML manifest for sideload |
| `office-mcp-addin-<ver>.zip` | Add-in static bundle (for self-hosters) |
| `SHA256SUMS` | Checksums |
| `SHA256SUMS.asc` | GPG signature (release key) |
