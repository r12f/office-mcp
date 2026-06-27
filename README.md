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

## Why not just use `python-docx` / `docx2pdf` / COM?

`python-docx`, `openpyxl`, `python-pptx`, `docx2pdf`, and similar tools work on files or automation surfaces, not the live Office host session. They are useful for batch processing, but they cannot reliably see every part of an already-open, protected, locked, or host-managed editing session.

| Problem | File libraries / COM | office-mcp |
|---|---|---|
| IRM / RMS protected documents | Cannot reliably access host-enforced state | Design target: run inside the Office host; protected-document behavior remains host-validated |
| Live editing in the user's open document | File-oriented tools need the file on disk; COM can fight the active desktop session | Operates through the add-in loaded in the live document |
| Office instance exclusive-access errors | Common when automation opens or locks the same file | Each Office window owns its own add-in session |
| Add-in install / discovery | Not applicable | Installer registers one trusted catalog for Word, Excel, and PowerPoint |
| MCP client config churn | Often one subprocess or config path per workflow | One persistent local Streamable HTTP endpoint |
| Platform path | COM is Windows-only; file libraries miss host behavior | Windows desktop v1; other Office platforms need separate deployment validation |

## Architecture (one diagram)

```text
+-------------+       +-----------------------+       +----------------------------+
| MCP Client  |<----->|  office-mcp daemon    |<----->| Word instance A (add-in)   |
| Claude,     | HTTP  |  long-lived process   | WSS   +----------------------------+
| Cursor,     |       |                       |       | Word instance B (add-in)   |
| agent       |       |  - tool router        |       +----------------------------+
+-------------+       |  - session registry   |       | Excel instance C (add-in)  |
                      |  - access control     |       +----------------------------+
                      +-----------------------+
```

- **MCP daemon** is a single long-running process. It speaks MCP Streamable HTTP to clients and JSON-RPC over a local secure WebSocket to add-ins.
- **Office add-ins** are Office.js task-pane add-ins for Word, Excel, and PowerPoint. Each loaded runtime dials out to the daemon and registers the current host document it can drive.
- **Clients** see a uniform MCP tool surface. The daemon filters disabled tools, checks session capabilities, and routes each call to the add-in that owns the target document.

## License

[MIT](LICENSE).
