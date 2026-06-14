# 09 - Feasibility Record

Checked on June 14, 2026 against Microsoft Word for Windows
`16.0.20131.20044`, current production Office.js typings, MCP protocol
`2025-11-25`, and the current Microsoft manifest validator.

## 1. Conclusion

The reduced Windows v1 design is implementable with stable public APIs. The
automated checks found no remaining protocol, schema, transport, or compile-time
API contradiction.

This is not a claim that every IRM policy works. Two host-only gates remain and
must pass before M0 is closed:

1. A user-consented trusted localhost certificate plus a sideloaded task pane
   must execute the reversible runtime smoke test in Word.
2. Representative Purview/AIP documents must verify read, edit, comment, save,
   and denied-operation behavior for each rights profile.

These are validation gates, not assumed capabilities. A failed gate reduces the
advertised feature set; it does not permit a bypass or inferred right.

## 2. Feasibility matrix

| Area | Result | Evidence |
|---|---|---|
| MCP frontend | Pass | Official SDK client/server initialization, tool list/call, structured success, `isError` execution failure |
| MCP HTTP security | Pass | Invalid `Origin` receives HTTP 403 |
| Add-in channel | Pass | HTTPS/WSS loopback handshake with SAN certificate; foreign WSS origin rejected |
| Session ordering | Pass | One active call plus bounded FIFO and reconnect/new-runtime ID assertions |
| Manifest | Pass | Add-in-only XML passes Microsoft's current manifest validator |
| Core Word APIs | Pass at type/API-set level | Paragraphs, search, insert/edit, tables, lists, selection, formatting, saved state |
| Review APIs | Pass at type/API-set level | Comments/bookmarks require `WordApi 1.4`; tracked changes require `WordApi 1.6` |
| Native Save As/PDF | Removed from v1 | No stable portable arbitrary-path or PDF-byte API |
| Character offsets | Removed from v1 | Stable Word API does not expose portable document character offsets |
| Active/protection metadata | Capability-gated | Desktop-only APIs; nullable when unavailable |
| IRM behavior | Manual gate | Office enforces policy; stable API does not enumerate every effective right |

## 3. Reproducible checks

Run from `feasibility/`:

```powershell
npm ci
npm run check
npm audit --audit-level=low
```

The suite contains:

- `word-api-probe.ts`: compiles representative stable Word API usage and
  runtime requirement checks.
- `protocol-probe.ts`: verifies queue capacity, FIFO order, and session ID
  lifecycle.
- `wss-probe.ts`: verifies TLS WebSocket registration and origin rejection.
- `mcp-probe.ts`: uses the official MCP SDK over Streamable HTTP.
- `manifest.xml`: validates the production-format XML manifest shape.

## 4. Primary references

- [MCP Streamable HTTP transport](https://modelcontextprotocol.io/specification/2025-11-25/basic/transports)
- [MCP tools and execution errors](https://modelcontextprotocol.io/specification/2025-11-25/server/tools)
- [Word API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- [Runtime API requirement checks](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-api-requirements-runtime)
- [Word Document API](https://learn.microsoft.com/en-us/javascript/api/word/word.document)
- [Office add-in-only XML manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview)
- [Unified manifest status](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview)
- [Office.js CDN requirements](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn)
- [Persisting add-in state](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings)
- [Automatically open a task pane](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)
