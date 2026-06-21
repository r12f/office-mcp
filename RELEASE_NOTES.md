# Release Notes

## 0.1.0

Office MCP Control 0.1.0 is the first Windows desktop installer pre-release.
It packages the native Rust daemon, daemon UI assets, Word/Excel/PowerPoint
add-in bundles, trusted catalog manifests, launcher scripts, product icons, and
the Windows tray entry point into `office-mcp-setup-0.1.0-x64.msi`.

This release is published as a draft pre-release until signing status, tray
evidence, installer smoke evidence, and required live Office evidence are
attached to the GitHub Release or explicitly waived for the pre-release.
Artifacts are unsigned unless the release page states otherwise.

Expected assets:

- `office-mcp-setup-0.1.0-x64.msi`
- `SHA256SUMS`

Validation gates before promoting the release:

- Windows desktop install completes without requiring a source checkout.
- Office MCP Control appears from the tray and opens the daemon UI.
- Word, Excel, and PowerPoint can load Office MCP Control from the Shared Folder
  catalog when Office does not auto-show the add-in.
- MCP clients can connect to `http://127.0.0.1:8800/mcp`.
- The release page includes tray evidence, installer smoke evidence, and live
  Office evidence for the supported hosts.
