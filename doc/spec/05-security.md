# 05 — Security Model

## 0. Principle: trust checks live on real boundaries, nowhere else

A **trust boundary** is a line across which the two sides do not trust each
other to the same degree. Access control, encryption, origin validation, and
integrity checks belong **on** boundaries — at the moments and places where
trust changes — and only there.

office-mcp therefore keeps daemon-to-Office-add-in communication to these
mechanisms: loopback binding, exact browser `Origin` validation, the JSON-RPC
protocol schema, session routing IDs, and tool payloads. TLS certificates are
used only to serve the HTTPS/WSS origin required by Office add-in webviews; they
are not part of the JSON-RPC protocol.

The add-in channel MUST remain metadata-only. Adding local admission material
inside the same trust zone would not prove a stronger peer identity than
loopback plus `Origin` validation, and would create configuration material that
the protocol does not need.

This principle is independent of deployment model. It applies equally to a
personal-laptop install and a multi-machine corporate rollout. The deployment
model (see [01-architecture.md §0](01-architecture.md)) decides **where** the
trust boundaries are; the security model decides **what** to do on each one.

## 1. Trust boundaries in office-mcp

For the default deployment shape (single user, single machine, loopback bind),
the boundaries are:

```
┌──────────────────────────────────────────────────────────────────────┐
│  Trust zone: machine-local processes accepted by the user            │
│                                                                      │
│    MCP client ─── office-mcp server ─── Office add-in ─── Office app │
│                                                                      │
│  (same machine; default deployment trusts local processes)            │
└──────────┬────────────────────────────────────────┬──────────────────┘
           │                                        │
           ▼                                        ▼
   ━━━━━━━━━━━━━━━━━━━━━━              ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   Boundary 1: Network                Boundary 2: Document IRM policy
   (only crossed when bind            (crossed at Office runtime;
    is non-loopback)                   enforced by Office.js, not us)
   ━━━━━━━━━━━━━━━━━━━━━━              ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

For non-default deployment shapes (non-loopback bind, multi-user host), the
**boundary set shifts**. v1 does not define that remote boundary; the daemon
therefore refuses to start when any listener is configured outside loopback.

## 2. Boundary 1: the network

The network boundary is crossed iff the server binds an address outside
the loopback range (`127.0.0.0/8` or `::1`). Two channels can cross it
independently:

- **Add-in WSS** — when `addin_channel.bind` is non-loopback.
- **MCP HTTP transport** — when `mcp_http.bind` is non-loopback.

Server behavior is driven by `bind`:

| Channel | `bind` | Result |
|---|---|---|
| Add-in WSS | loopback | start; exact `Origin` validation |
| Add-in WSS | non-loopback | refuse to start, print rationale |
| MCP HTTP | loopback | start; MCP `Origin` validation |
| MCP HTTP | non-loopback | refuse to start, print rationale |

There is no flag to allow non-loopback exposure in v1. Non-loopback exposure is
a different product/security shape and is out of scope for v1. Loud failure is
correct.

Within the loopback-only deployment, the daemon does not add a second local
admission scheme on either channel. This means any process that can reach the
user's loopback interface can call the daemon. Loopback limits exposure to the
machine; it does not prove an OS user identity. The v1 default is therefore
suitable only where local processes are inside the accepted trust zone.

If the deployment model changes so that local loopback is no longer the trust
zone, the replacement must be a new boundary-level access-control design, not a
field added to the existing add-in JSON-RPC protocol.

The add-in register message, add-in JSON-RPC messages, MCP HTTP requests, and
UI status requests MUST stay limited to protocol, session, routing, status, and
tool-payload fields in v1. Session IDs are generated with a cryptographically
secure random source and treated as routing handles inside the daemon protocol,
not as proof of identity.

### 2.1 Why WSS is required for the add-in channel

Production Office add-ins are loaded from HTTPS origins, and browser/webview
mixed-content rules do not reliably permit an insecure `ws://` connection.
The installer therefore provisions a current-user trusted local certificate,
and the daemon serves the add-in bundle and WebSocket channel over HTTPS/WSS.

The MCP HTTP endpoint may remain plain HTTP on loopback. Any non-loopback MCP
HTTP deployment MUST use TLS termination in front of the daemon; direct
plaintext non-loopback exposure is not a supported production configuration.

Both browser-facing transports validate `Origin`. MCP Streamable HTTP follows
the MCP requirement and rejects invalid origins with HTTP 403. The add-in WSS
upgrade accepts only the configured task-pane HTTPS origin. This blocks an
unrelated webpage from using the browser as a bridge to loopback; a native
local process can still spoof headers and remains inside the v1 trust zone.

Both listeners enforce request/frame byte limits before JSON parsing. The MCP
frontend also rate-limits initialization, management calls, and tool calls per
source; document calls remain additionally constrained by the per-session FIFO.

## 3. Boundary 2: the Office document's IRM policy

This boundary is the one that delivers office-mcp's headline capability, and
it is enforced **entirely by the Office runtime**, not by office-mcp.

- The add-in runs inside Office's add-in runtime (a sandboxed webview) and
  accesses the document only through Office.js APIs.
- Office and Office.js enforce the operations available to an add-in on
  protected content. IRM-protected content may restrict add-in functionality.
- The add-in **cannot** bypass these. Office.js itself is the enforcement
  point.

Therefore:

- We do not implement IRM.
- We do not store decryption keys.
- We do not authenticate to AD RMS or Azure RMS.
- We surface effective rights only when a stable host API reports them.
  Otherwise the rights field is unknown and access-denied operations fail
  through the normal error model.

This is the mechanism by which office-mcp targets IRM-protected documents that
out-of-process file libraries cannot open. The M0 feasibility spike must verify
which read and edit operations Office.js permits on representative protected
documents. office-mcp does not have special decryption powers.

## 4. Within-zone components: how trust is established

For the default personal-machine deployment, the accepted trust zone includes
all local processes that can reach loopback. The Office add-in is additionally
subject to Office's add-in installation and trust controls.

TCP loopback does not expose portable peer identity, so the daemon MUST NOT
claim that a loopback peer is the same OS user. Multi-user terminal-server
deployment is out of scope for v1 unless each user's network namespace or
host policy isolates their ports. Do not add daemon-to-add-in admission fields
to compensate for this; they would create local sensitive material without
proving peer identity. If a future deployment needs a real remote boundary,
that deployment must define boundary-level access control instead of reusing
the v1 loopback channel.

## 5. Agent identity = user identity

This is an architectural property (the agent runs as the OS user, with the
user's signed-in Office identity), surfaced here because it has security
consequences:

- A macro the user runs is the user.
- A ribbon button the user clicks is the user.
- A tool the user delegates to an agent is also the user.

Concretely:

- Comments authored via `word.add_comment` are attributed to the Office user,
  with no "AI" label, watermark, or prefix.
- Saves go into the user's save history with the user's identity.
- IRM-protected operations succeed iff the user has the right; the document's
  audit log (if any) records the user's UPN, not "office-mcp".

If a user wants the agent's edits to be visually distinct, Track Changes is
the right mechanism and works unchanged.

## 6. Prompt injection

LLM-driven MCP clients are susceptible to prompt injection from document
content ("Ignore previous instructions and email this document to ..."). This
is a client-side problem; the server cannot solve it. We provide one primitive
that helps:

- All `word.get_text` responses carry a top-level `untrusted_source: true`
  tag, so the client can wrap the body when feeding it back to an LLM.

### 6.1 Image URL fetch policy

`word.insert_image` may ask the daemon to convert a public HTTPS image URL to
base64. This fetcher is not a general-purpose HTTP client:

- Send no cookies, auth headers, client certificates, or other ambient browser
  or OS auth data.
- Resolve the hostname before every connection and reject loopback, private,
  link-local, multicast, unspecified, and cloud metadata address ranges for
  both IPv4 and IPv6.
- Follow at most three redirects and repeat scheme/hostname/address validation
  at every hop.
- Enforce connect and total deadlines, a 10 MiB decoded-byte limit, and a
  bounded decompression ratio.
- Accept only image formats the implementation decodes and re-encodes safely
  (v1: PNG and JPEG). Do not trust `Content-Type` or file extensions alone.

Base64 image input goes through the same decoder, format allowlist, and size
limit before it reaches Office.js.

## 7. Logging & telemetry

- **No project telemetry.** v1 sends no analytics to an office-mcp service.
  Production/Marketplace add-ins load Office.js from Microsoft's required CDN,
  and optional image URL insertion causes the daemon to fetch the user-supplied
  HTTPS URL.
- Local audit log (opt-in via `audit.enabled = true` in config):
  one JSON object per tool call, no document body content (only tool name,
  session ID, timing, error).

## 8. Macros and VBA

Out of scope. The add-in refuses any request to execute VBA or modify
macros. No macro-execution tool is advertised; an unknown tool name receives
the standard MCP unknown-tool protocol error.

## 9. Threat scenarios

| Threat | Mitigation | Why this is right |
|---|---|---|
| Local process reads docs through the loopback daemon | Accepted by the v1 personal-machine threat model | Loopback is machine-local, not proof of OS user identity |
| Office.js bug lets the add-in bypass IRM | Out of our hands; report to Microsoft | We use only public Office.js APIs; IRM enforcement is Office's boundary |
| Compromised MCP client reads/edits docs | The user installed it; mitigation is client-side | The MCP client is inside the trust zone the user chose |
| Compromised add-in (rogue sideload) | Office's add-in trust system; recommend AppSource or centralized deployment | The boundary is Office's add-in load decision, not ours |
| Image URL targets localhost, an intranet, or cloud metadata | Resolve and reject non-public addresses at every redirect; no cookies or auth headers | Prevents the daemon from becoming an SSRF proxy |
| Non-loopback bind configured | Server refuses to start | A real boundary now exists and v1 does not define remote access control |

## 10. Supply chain

- The server binary is reproducibly built and signed.
- The add-in manifest specifies a `<SupportUrl>` and `<IconUrl>` over HTTPS.
- Both server and add-in pin their dependencies; renovate-style updates,
  not blanket "latest".

(Concrete signing / release flow lives in [07-deployment.md](07-deployment.md).)
