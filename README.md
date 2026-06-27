# office-mcp

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/r12f/office-mcp?include_prereleases)](https://github.com/r12f/office-mcp/releases)
[![Release workflow](https://github.com/r12f/office-mcp/actions/workflows/release.yml/badge.svg)](https://github.com/r12f/office-mcp/actions/workflows/release.yml)
[![Platform](https://img.shields.io/badge/platform-Windows%20desktop-lightgrey.svg)](doc/spec/07-deployment.md)

`office-mcp` connects MCP clients to live Microsoft Office documents through Office add-ins and a local MCP daemon.

## Install and MCP config

Install the latest Windows portable release from PowerShell:

```powershell
irm https://raw.githubusercontent.com/r12f/office-mcp/main/scripts/install.ps1 | iex
```

Configure MCP clients to use the local Streamable HTTP endpoint:

```text
http://127.0.0.1:8800/mcp
```

See the [deployment spec](doc/spec/07-deployment.md) for manual install, upgrade, log, catalog, and uninstall details.

## Why office-mcp

AI agents need access to the document the user is actually editing, not only the file that happens to be on disk.

Office documents may be protected, open, locked, or host-managed in ways file parsers cannot handle correctly. A local add-in can run inside the Office host, inherit the user's current Office context, and preserve the live editing state the user sees.

## Key idea

One local long-lived MCP daemon exposes a stable MCP endpoint.

Office add-ins run inside Word, Excel, and PowerPoint. Each add-in reverse-connects to the daemon and registers the live document session it owns.

MCP clients call tools through the daemon, and the daemon routes calls to the correct Office host session.

## Difference vs Python / COM-based MCP servers or skills

`python-docx`, `openpyxl`, `python-pptx`, and similar libraries work on file formats, not the live Office host. They are useful for batch file processing, but they cannot see every part of an already-open, protected, or host-managed editing session.

COM automation is Windows-only, can fight with already-open Office instances, and often depends on process or desktop-session assumptions that do not fit agent workflows.

`office-mcp` is designed around Office.js add-ins and live host sessions, giving agents a route to user-open documents through the Office application itself while keeping capability claims aligned with the current implementation and spec.

## License

[MIT](LICENSE).
