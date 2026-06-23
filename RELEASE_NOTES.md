# Release Notes

## 0.1.1

Office MCP Control 0.1.1 is the first Windows desktop portable-only pre-release.
It packages the native Rust daemon, daemon UI assets, Word/Excel/PowerPoint
add-in bundles, trusted catalog manifests, launcher scripts, product icons, and
the Windows tray entry point into the auditable portable package
`office-mcp-windows-portable-0.1.1-x64.zip`.

This release supersedes the earlier 0.1.0 prerelease artifacts and removes the
opaque Windows installer packaging path. Users should install from the portable zip so the
program location, add-in catalog, scripts, certificate helper, and uninstall
steps are visible before anything runs.

This release is published as a draft pre-release until signing status, tray
evidence, portable package smoke evidence, and required live Office evidence are
attached to the GitHub Release or explicitly waived for the pre-release.
Artifacts are unsigned unless the release page states otherwise.

Expected assets:

- `office-mcp-windows-portable-0.1.1-x64.zip`
- `SHA256SUMS`

Validation gates before promoting the release:

- Windows desktop portable install completes without requiring a source checkout.
- The portable zip contains `README-install.txt`, `install-user.ps1`,
  `start-daemon.ps1`, and `uninstall-user.ps1` so users can inspect the install
  location and user-level registry/certificate changes before running them.
- Office MCP Control appears from the tray and opens the daemon UI.
- Word, Excel, and PowerPoint can load Office MCP Control from the Shared Folder
  catalog when Office does not auto-show the add-in.
- MCP clients can connect to `http://127.0.0.1:8800/mcp`.
- The release page includes tray evidence, portable package smoke evidence, and live
  Office evidence for the supported hosts.