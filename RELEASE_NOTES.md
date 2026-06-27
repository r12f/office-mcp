# Release Notes

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
