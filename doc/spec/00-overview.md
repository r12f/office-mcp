# 00 — Overview

## 1. Goals

`office-mcp` exposes live Microsoft Office applications (Word and Excel first,
PowerPoint / Outlook later) as Model Context Protocol (MCP) tools, with three concrete properties that
existing MCP implementations do not provide together:

1. **Operates on the user's live, open document.** No "save, close, re-open" round trips.
   The AI sees what the user sees and edits in place.
2. **Targets Information Rights Management (IRM / AIP / Purview) protected content.**
   Because the add-in runs inside the Office process, it inherits the user's already-validated
   Office context — no MIP SDK integration or separate document authentication.
   Exact supported operations remain host-dependent and are an M0 validation gate.
3. **Survives the "every Office window is a different process" problem.** A single
   long-lived MCP server holds the client-facing endpoint; Office add-ins reverse-connect
   to it when each instance loads, and disconnect cleanly when the instance dies.

## 2. Non-goals

| Out of scope (v1) | Why |
|---|---|
| Replacing python-docx for batch / headless `.docx` processing | Different problem. Use python-docx when no user is at the keyboard. |
| Authoring IRM policies (encrypting unprotected documents) | Requires MIP SDK; see [05-security.md](05-security.md). v1 only consumes existing protections. |
| Server-side document rendering (PDF, images) | Office.js can request these, but quality and fidelity depend on the host; out of v1. |
| Macro execution | VBA / OOXML macros are out — too easy to weaponize via prompt injection. |
| Cross-tenant document sharing | Each add-in instance is scoped to one user identity. |
| Remote add-in (no local Office) | The whole point is to leverage local Office. If you don't have Office, use python-docx. |

## 3. Glossary

| Term | Meaning |
|---|---|
| **MCP** | Model Context Protocol. Open standard for AI ↔ tools (Anthropic-led, multi-vendor). |
| **MCP client** | The thing that wants tools — Claude Desktop, Cursor, an agent, etc. |
| **MCP server** | What we build. Speaks MCP outward, JSON-RPC over WSS inward. |
| **Office add-in** | A web app (HTML/JS) hosted inside an Office app via Office.js. v1 targets Word and Excel desktop on Windows; macOS and Office on the web are later validation targets. |
| **Daemon UI** | The desktop tray icon and main window for monitoring the long-running `office-mcp` process. |
| **Task pane UI** | The Office add-in's document-local UI, hosted in the Office webview. |
| **App** | Office application: Word, Excel, PowerPoint, Outlook. URI namespace and tool prefix. |
| **Document session** | One open document (`.docx`, `.xlsx`, etc.) inside one app instance. Unit of addressing. |
| **Add-in runtime** | One document-scoped web runtime. It owns one WebSocket connection and one document session in v1. |
| **IRM** | Information Rights Management — Microsoft's per-document access control. Also AIP / MIP / Purview. |
| **Reverse-registration** | Add-in dials out to the server on load (the server does not connect to add-ins). |
| **Selection-anchored operation** | Tool call whose target is "wherever the user's cursor is now" rather than an explicit range. |

## 4. Design priorities (in order)

1. **Reliability over feature surface.** Better to do 20 things that never fail than 200
   things that sometimes hang. COM-based MCPs fail this test today; the project exists
   because the current options are unreliable.
2. **Stateless clients.** A client should be able to disconnect and reconnect without
   losing track of which documents are addressable. State lives in the server.
3. **Stateful documents.** Per-document state (last selection and pending operation
   queue) lives in the document-scoped add-in runtime, addressed by stable session IDs.
4. **Boring transports.** Streamable HTTP for MCP; WSS + JSON-RPC 2.0 for
   add-in ↔ server. No bespoke binary frames, no SSE-only paths. Stdio-only
   clients use the bundled stdio-to-daemon bridge.
5. **Local-first install.** Transparent Windows portable package and no project-operated
   cloud service. Marketplace builds load the production Office.js library
   from Microsoft's CDN, so fully offline operation is not a v1 promise. See
   [01-architecture.md §0](01-architecture.md) for the full deployment model.
6. **Observable by default.** The daemon exposes a tray icon and main window,
   and the add-in task pane exposes document-local connection and task status.
   Users should not need a terminal to answer basic support questions. See
   [09-ui.md](09-ui.md).
7. **Graceful degradation.** When no Office instance is running, the server still works
   — clients see "no document sessions available" instead of a crash.

## 5. Comparison with existing MCPs

| Project | Approach | IRM | Live edit | Reliability under load |
|---|---|---|---|---|
| `GongRzhe/Office-Word-MCP-Server` | python-docx + msoffcrypto | ❌ Password only | ❌ File-based | ✅ but no Office |
| Various COM-based MCPs | Win32 COM automation | ⚠️ Only if Office UI handles it | ✅ | ❌ Exclusive-access errors, hung processes |
| `office-mcp` (this) | Office.js add-in + reverse-connect | ✅ Inherits user rights | ✅ | ✅ Per-instance isolation |

## 6. Open questions tracked elsewhere

| Q | Where |
|---|---|
| How exactly does a client target one document among many? | [03-mcp-tool-surface.md §3](03-mcp-tool-surface.md) |
| What does the add-in send when it first dials in? | [02-registration-protocol.md §2](02-registration-protocol.md) |
| What happens when Office is killed mid-operation? | [06-error-model.md §4](06-error-model.md) |
| Distribution: sideload vs AppSource vs centralized deployment? | [07-deployment.md §3](07-deployment.md) |
| What should users see in the tray, daemon window, and add-in task pane? | [09-ui.md](09-ui.md) |
