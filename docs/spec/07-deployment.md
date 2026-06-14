# 07 — Deployment

How the server and the add-in get onto a user's machine and start talking to
each other.

## 1. Components to ship

| Component | What it is | Where it lives | Updated how |
|---|---|---|---|
| `office-mcp-server` | Single binary (~15 MB) | `%LOCALAPPDATA%\office-mcp\office-mcp-server.exe` (Win) / `/usr/local/bin` (Mac) | MSI / Homebrew tap |
| `office-mcp-addin` | Static web bundle (~2 MB) + `manifest.xml` | Hosted at `https://office-mcp.dev/addin/v1/` (CDN) | Atomic versioned URLs |
| Manifest | XML / JSON describing the add-in | Sideloaded, AppSource, or M365 centralized deployment | See §3 |
| Bootstrap installer | MSI / .pkg / shell script | Downloaded from GitHub Releases | Per-release |

## 2. Server installation

### 2.1 Windows (primary target)

`office-mcp-setup-x64.msi` performs:

1. Copies `office-mcp-server.exe` to `%LOCALAPPDATA%\office-mcp\`.
2. Creates a Start Menu shortcut.
3. Registers a Scheduled Task `office-mcp` set to run at logon (optional, off by default).
4. Writes the default config to `%APPDATA%\office-mcp\config.toml`.
5. Does NOT install the add-in (that happens via manifest sideload, see §3).

### 2.2 macOS

`brew install r12f/tap/office-mcp` installs the binary; the user runs
`office-mcp install` to set up a launchd plist for autostart.

### 2.3 Linux (developer-only)

`cargo install office-mcp` or pre-built tarball. Mostly for protocol testing;
no Office host is normally present.

## 3. Add-in distribution

Office add-ins can be deployed three ways. office-mcp publishes the same
manifest for all three.

### 3.1 Sideload (developer / individual user)

```
# Word desktop on Windows
1. Open Word → File → Options → Trust Center → Trusted Add-in Catalogs.
2. Add a trusted catalog path (a network share or local folder).
3. Drop manifest.xml in that folder.
4. Restart Word.
5. Insert → My Add-ins → Shared Folder → office-mcp → Add.
```

The MSI optionally creates `%LOCALAPPDATA%\office-mcp\addin-catalog\` and
pre-registers it as a trusted catalog (with the user's explicit consent at
install time).

### 3.2 AppSource (eventual)

Published via [partner.microsoft.com](https://partner.microsoft.com). Users
install with one click from Office's Add-ins picker. Requires:

- Microsoft Partner Center account
- AppSource validation review
- Hosted manifest URL (HTTPS, public)

Tracked in [08-roadmap.md](08-roadmap.md); not blocking v1.

### 3.3 Centralized deployment (enterprise)

M365 admins deploy via Microsoft 365 admin center → Settings → Integrated apps.
Users get the add-in automatically. Same manifest, different distribution.

## 4. Manifest

Unified manifest (JSON, the recommended format for 2026+):

```jsonc
// manifest.json — abridged
{
  "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
  "manifestVersion": "1.17",
  "version": "0.1.0",
  "id": "11111111-aaaa-bbbb-cccc-222222222222",
  "developer": {
    "name": "office-mcp project",
    "websiteUrl": "https://github.com/r12f/office-mcp",
    "privacyUrl": "https://github.com/r12f/office-mcp/blob/main/PRIVACY.md",
    "termsOfUseUrl": "https://github.com/r12f/office-mcp/blob/main/LICENSE"
  },
  "name": { "short": "office-mcp", "full": "office-mcp — AI bridge for Word" },
  "description": {
    "short": "Lets AI agents read and edit your Word documents",
    "full": "Exposes Word as MCP tools so AI assistants (Claude, Cursor, agents) can read and edit the document you're working on, including IRM-protected files."
  },
  "icons": { "outline": "icon-32.png", "color": "icon-192.png" },
  "accentColor": "#0078d4",
  "extensions": [
    {
      "requirements": {
        "scopes": ["document"],
        "capabilities": [{ "name": "WordApi", "minVersion": "1.3" }]
      },
      "ribbons": [
        {
          "contexts": ["default"],
          "tabs": [
            {
              "builtInTabId": "TabHome",
              "groups": [
                {
                  "id": "officeMcpGroup",
                  "label": "AI",
                  "controls": [
                    {
                      "id": "openOfficeMcp",
                      "type": "button",
                      "label": "office-mcp",
                      "actionId": "showTaskpane"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    }
  ],
  "validDomains": [],
  "webApplicationInfo": { "id": "00000000-...", "resource": "..." }
}
```

XML-format equivalent is shipped alongside for compatibility with Office
versions that don't yet support the unified manifest (still ~30% of installed
base as of 2026).

## 5. Configuration

`office-mcp` uses a **single config file shared by the server and the add-in**.
Changing the port (or any other shared value) in one place updates both ends.

Location:

| OS | Path |
|---|---|
| Windows | `%APPDATA%\office-mcp\config.toml` |
| macOS | `~/Library/Application Support/office-mcp/config.toml` |
| Linux | `~/.config/office-mcp/config.toml` |

The add-in resolves the same path via `Office.context.requirements` /
`fetch('app://office-mcp/config')` shim provided by the add-in package; users
do not maintain two copies.

```toml
# ─── shared by server and add-in ─────────────────────────────────────────
[addin_channel]
# WebSocket the server listens on and the add-in dials.
bind = "127.0.0.1"
port = 8765
heartbeat_interval_sec = 30
session_grace_sec = 60
max_inflight_per_session = 4
# Only required if `bind` is non-loopback. Empty (default) = no auth, which
# is safe and correct on a loopback bind. See docs/spec/05-security.md.
shared_secret = ""

# ─── server only ─────────────────────────────────────────────────────────
[mcp_stdio]
# stdio is always available; nothing to configure here. Listed for clarity.
enabled = true

[mcp_http]
# Optional HTTP frontend for MCP clients (Streamable HTTP transport).
enabled = false
bind = "127.0.0.1:8800"
# Only required if `bind` is non-loopback. Empty + non-loopback = startup refusal.
api_key = ""

[limits]
max_response_bytes = 1048576     # 1 MiB
default_tool_timeout_ms = 30000

[lifecycle]
idle_shutdown_sec = 0            # 0 = never auto-shutdown

[audit]
enabled = false
path = ""                        # default: %LOCALAPPDATA%\office-mcp\audit.jsonl

[logging]
level = "info"                   # trace | debug | info | warn | error
file  = ""                       # default: stderr
```

Environment variables override config keys, prefixed `OFFICE_MCP_` (e.g.
`OFFICE_MCP_ADDIN_CHANNEL__PORT=9000`). Add-in honors only the keys under
`[addin_channel]`; everything else is server-side and silently ignored if
present in the add-in's view of the config.

## 6. First-run flow (Windows desktop)

1. User installs `office-mcp-setup-x64.msi`. Reboots not required.
2. User installs the add-in manifest (sideload or via admin push).
3. User configures their MCP client to launch office-mcp:
   ```json
   {
     "mcpServers": {
       "office": {
         "command": "%LOCALAPPDATA%\\office-mcp\\office-mcp-server.exe",
         "args": ["--transport", "stdio"]
       }
     }
   }
   ```
4. User opens Word. The pinned ribbon button "office-mcp" appears.
5. User clicks it to open the task pane; the add-in connects to the server.
   (For pinned-on-load behavior, the add-in declares it in the manifest;
   then no click is required.)
6. User asks Claude (or whatever) to do something with the doc.

## 7. Versioning & upgrades

- **Server version** is reported in `register.result.server_version`.
- **Add-in version** is reported in `register.params.add_in.version`.
- The server and add-in protocols are versioned independently (see
  [01-architecture.md §6](01-architecture.md)).
- Auto-update:
  - Server: opt-in update check at startup (GitHub Releases API; can be disabled).
  - Add-in: Office handles add-in updates automatically when manifest changes
    on the trusted catalog or AppSource.

## 8. Uninstall

`office-mcp uninstall` (or the MSI's uninstall) removes:

- Binaries
- Scheduled task / launchd plist
- `%LOCALAPPDATA%\office-mcp\` (audit log, sideload catalog)
- `%APPDATA%\office-mcp\` (config) — only with `--purge`

The add-in must be removed separately via Office's add-in manager
(or by removing the manifest from the trusted catalog).

## 9. Release artifacts (per version tag)

| File | Purpose |
|---|---|
| `office-mcp-server-<ver>-x64.exe` | Standalone Windows binary |
| `office-mcp-setup-<ver>-x64.msi` | Windows MSI installer |
| `office-mcp-server-<ver>-x64.tar.gz` | Linux/macOS tarball |
| `office-mcp-server-<ver>-aarch64-darwin.tar.gz` | Apple Silicon |
| `manifest-<ver>.xml` | XML manifest for sideload |
| `manifest-<ver>.json` | Unified manifest for sideload |
| `office-mcp-addin-<ver>.zip` | Add-in static bundle (for self-hosters) |
| `SHA256SUMS` | Checksums |
| `SHA256SUMS.asc` | GPG signature (release key) |
