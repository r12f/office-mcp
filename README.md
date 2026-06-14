# office-mcp

**A bridge between AI assistants and live Microsoft Office applications via in-process add-ins.**

`office-mcp` exposes Word (and eventually Excel, PowerPoint, Outlook) as MCP tools by running an
add-in inside each Office instance. The add-in reverse-connects to a single long-lived MCP server
process, which multiplexes MCP clients across all running Office windows.

## Why not just use `python-docx` / `docx2pdf` / COM?

| Problem | python-docx / COM | office-mcp |
|---|---|---|
| IRM / RMS protected documents | ❌ Cannot open | Design target: host-enforced access; M0 validation required |
| Live editing in user's open document | ❌ File must be closed | ✅ Operates on the live document |
| Office instance exclusive-access errors | ❌ Common with COM | ✅ Each Office instance has its own add-in |
| Add-in install / discovery | n/a | ✅ Add-in self-registers on Office start |
| MCP client config churn | ❌ Per-doc subprocess | ✅ One persistent server endpoint |
| Platform path | ❌ Windows only | Windows desktop v1; Mac planned; Web requires a different deployment |

## Architecture (one diagram)

```
┌─────────────┐       ┌───────────────────────┐       ┌────────────────────────────┐
│ MCP Client  │◀─────▶│  office-mcp server    │◀─────▶│ Word instance A (add-in)   │
│ (Claude,    │ HTTP  │  (long-lived process) │  WS   ├────────────────────────────┤
│  Cursor,    │       │                       │       │ Word instance B (add-in)   │
│  agent)     │ HTTP  │  - tool router        │       ├────────────────────────────┤
└─────────────┘       │  - session registry   │       │ Excel instance C (add-in)  │
                      │  - capability negotiation │   └────────────────────────────┘
                      └───────────────────────┘
```

- **MCP server** is a single long-running process. It speaks MCP Streamable HTTP
  to clients and JSON-RPC over a local secure WebSocket to add-ins.
- **Office add-ins** are Office.js task-pane add-ins (one per Office app type). Each loaded
  runtime dials out to the server and registers the current host document it can drive.
- **Clients** see a uniform MCP tool surface; the server routes each call to the add-in that
  owns the target document.

## Status

Spec phase. See [docs/spec/](docs/spec/) for the full design.

| Doc | Purpose |
|---|---|
| [00-overview.md](docs/spec/00-overview.md) | Goals, non-goals, glossary |
| [01-architecture.md](docs/spec/01-architecture.md) | Process model, transports, lifecycle |
| [02-registration-protocol.md](docs/spec/02-registration-protocol.md) | Add-in ↔ server JSON-RPC wire protocol |
| [03-mcp-tool-surface.md](docs/spec/03-mcp-tool-surface.md) | MCP tools, resources, prompts exposed to clients |
| [04-word-capabilities.md](docs/spec/04-word-capabilities.md) | Word-specific tool catalog (v1) |
| [05-security.md](docs/spec/05-security.md) | Auth, sandbox, IRM, trust boundaries |
| [06-error-model.md](docs/spec/06-error-model.md) | Error codes, retry, partial-failure semantics |
| [07-deployment.md](docs/spec/07-deployment.md) | Installation, sideloading, distribution |
| [08-roadmap.md](docs/spec/08-roadmap.md) | Milestones, Excel/PPT/Outlook follow-on |
| [09-feasibility.md](docs/spec/09-feasibility.md) | Evidence, probes, and remaining host validation gates |

## License

TBD (likely MIT, matching the surrounding Office add-in ecosystem).
