# office-ctl common

Shared TypeScript utilities for Office add-ins: endpoint configuration,
logging, channel/protocol helpers, redaction, and reusable UI primitives.
Host-specific Office.js calls must stay in the host folders.

`browser-ui.js` is a host-neutral browser helper loaded by task panes before
their host entry point. It may contain formatting, redaction, escaping, and
small view helpers. It must not call `Word.*`, `Excel.*`, or other host APIs.

`addin-channel.js` is a host-neutral channel helper for endpoint storage,
runtime/session IDs, register request tracking, JSON-RPC parsing, send/reply,
and reconnect timing. It must not call Office host APIs or own host-specific
registration metadata.

`logger.js` is a host-neutral scoped logger for high-level task pane events. It
redacts user-controlled strings and document-shaped fields before writing to the
browser console. It must not persist logs or send telemetry.

`task-history.js` is a host-neutral in-memory current-task and bounded-history
store for task panes. It owns cancellation markers and redacted high-level task
history only; host folders still own command execution and rendering.
