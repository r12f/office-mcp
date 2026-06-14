# 07 — Deployment

How the server and the add-in get onto a user's machine and start talking to
each other.

## 1. Components to ship

| Component | What it is | Where it lives | Updated how |
|---|---|---|---|
| `office-mcp` | Single binary (~15 MB) | `%LOCALAPPDATA%\office-mcp\office-mcp.exe` (Win) / `/usr/local/bin/office-mcp` (Mac, Linux) | MSI / Homebrew tap |
| `office-mcp-addin` | Static web bundle (~2 MB) + manifest | Hosted at `https://office-mcp.dev/addin/v1/` (CDN); a copy is staged into the per-user trusted catalog by the installer | Atomic versioned URLs |
| Manifest | XML / JSON describing the add-in | Sideloaded via the trusted catalog by the installer; AppSource / M365 admin push for managed deployments | See §3 |
| Bootstrap installer | MSI / .pkg / shell script | Downloaded from GitHub Releases | Per-release |

The actual installation procedure — including daemon autostart and add-in
catalog registration — is in §6, not here. §1 is just the artifact list.

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

`office-mcp` uses a **single config file shared by the daemon and the add-in**.
Changing the port (or any other shared value) in one place updates both ends.

Location:

| OS | Path |
|---|---|
| Windows | `%APPDATA%\office-mcp\config.toml` |
| macOS | `~/Library/Application Support/office-mcp/config.toml` |
| Linux | `~/.config/office-mcp/config.toml` |

The add-in resolves the same path via a small loader in the add-in package;
users do not maintain two copies.

```toml
# ─── shared by daemon and add-in ─────────────────────────────────────────
[addin_channel]
# WebSocket the daemon listens on and the add-in dials.
bind = "127.0.0.1"
port = 8765
heartbeat_interval_sec = 30
heartbeat_timeout_sec  = 10
session_grace_sec = 60
max_inflight_per_session = 4
# Only required if `bind` is non-loopback. See docs/spec/05-security.md.
shared_secret = ""

# ─── daemon only ─────────────────────────────────────────────────────────
[mcp_http]
# MCP Streamable HTTP frontend the daemon exposes for MCP clients.
bind = "127.0.0.1"
port = 8800
# Only required if `bind` is non-loopback. Empty + non-loopback = startup refusal.
api_key = ""

[limits]
max_response_bytes = 1048576     # 1 MiB
default_tool_timeout_ms = 30000

[audit]
enabled = false
path = ""                        # default: %LOCALAPPDATA%\office-mcp\audit.jsonl

[logging]
level = "info"                   # trace | debug | info | warn | error
file  = ""                       # default: platform log dir
```

Environment variables override config keys, prefixed `OFFICE_MCP_` (e.g.
`OFFICE_MCP_ADDIN_CHANNEL__PORT=9000`). The add-in honors only the keys under
`[addin_channel]`. The daemon honors all keys.

## 6. Installation and first-run flow

Installation has two prerequisites: the daemon must be installed and
autostarted by the OS, and the add-in must be registered with Office. Both
are done by the platform installer (MSI / .pkg / brew formula); the user
should never have to hand-edit autostart entries or sideload from a developer
console for a production install.

### 6.1 Windows

1. User installs `office-mcp-setup-x64.msi`. The installer:
   - Drops `office-mcp.exe` to `%LOCALAPPDATA%\office-mcp\`.
   - Writes a default `config.toml` to `%APPDATA%\office-mcp\`.
   - Registers a Scheduled Task `office-mcp` set to run at logon
     (`office-mcp daemon run`).
   - Registers an add-in trusted catalog folder under
     `%LOCALAPPDATA%\office-mcp\addin-catalog\` and drops the manifest
     there.
   - Runs `office-mcp daemon start` once so the user doesn't have to log
     out / log in for the daemon to come up.

2. User configures their MCP client to connect to
   `http://127.0.0.1:8800` (or whatever `mcp_http.port` is in their config).

3. User opens Word. The pinned add-in appears in the ribbon. Pinned-on-load
   is set in the manifest, so the add-in connects without the user clicking.

4. User asks the MCP client to do something with the doc.

### 6.2 macOS / Linux

Same shape, different mechanism: launchd agent (macOS) or systemd `--user`
unit (Linux) replaces the Scheduled Task. brew formula / Linux package
performs the equivalent setup.

### 6.3 Verifying

```
office-mcp daemon status     # is the daemon up?
office-mcp daemon stop/start # restart it
office-mcp config show       # show effective config
office-mcp sessions          # list connected add-ins (= open Office instances)
```

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
| `office-mcp-<ver>-x64.exe` | Standalone Windows binary |
| `office-mcp-setup-<ver>-x64.msi` | Windows MSI installer |
| `office-mcp-<ver>-x64.tar.gz` | Linux/macOS tarball |
| `office-mcp-<ver>-aarch64-darwin.tar.gz` | Apple Silicon |
| `manifest-<ver>.xml` | XML manifest for sideload |
| `manifest-<ver>.json` | Unified manifest for sideload |
| `office-mcp-addin-<ver>.zip` | Add-in static bundle (for self-hosters) |
| `SHA256SUMS` | Checksums |
| `SHA256SUMS.asc` | GPG signature (release key) |
