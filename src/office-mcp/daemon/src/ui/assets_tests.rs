use super::{UiAssetError, UiAssetStore};
use std::fs;

#[test]
fn reads_assets_from_explicit_root_with_content_type() {
    let dir =
        std::env::temp_dir().join(format!("office-mcp-ui-assets-test-{}", std::process::id()));
    fs::create_dir_all(&dir).expect("asset dir");
    fs::write(dir.join("index.html"), "<main>Office MCP</main>").expect("html");
    fs::write(dir.join("app.css"), "body { margin: 0; }").expect("css");
    fs::write(dir.join("app.js"), "window.ok = true;").expect("js");

    let store = UiAssetStore::with_root(dir.clone());

    let html = store.read("index.html").expect("html asset");
    assert_eq!(html.content_type, "text/html; charset=utf-8");
    assert_eq!(html.content, b"<main>Office MCP</main>");

    let css = store.read("app.css").expect("css asset");
    assert_eq!(css.content_type, "text/css; charset=utf-8");

    let js = store.read("app.js").expect("js asset");
    assert_eq!(js.content_type, "text/javascript; charset=utf-8");

    let _ = fs::remove_dir_all(dir);
}

#[test]
fn rejects_unsafe_asset_names() {
    let store = UiAssetStore::with_root(std::env::temp_dir());

    assert!(matches!(store.read(""), Err(UiAssetError::UnsafeName)));
    assert!(matches!(
        store.read("../index.html"),
        Err(UiAssetError::UnsafeName)
    ));
    assert!(matches!(
        store.read("ui/index.html"),
        Err(UiAssetError::UnsafeName)
    ));
    assert!(matches!(
        store.read("ui\\index.html"),
        Err(UiAssetError::UnsafeName)
    ));
}

#[test]
fn default_store_finds_repo_daemon_ui_assets() {
    let html = UiAssetStore::default()
        .read("index.html")
        .expect("default html asset");

    assert_eq!(html.content_type, "text/html; charset=utf-8");
    assert!(String::from_utf8_lossy(&html.content).contains("Office MCP"));
}

#[test]
fn default_daemon_ui_assets_keep_accessible_dense_operations_layout() {
    let store = UiAssetStore::default();
    let html =
        String::from_utf8(store.read("index.html").expect("html").content).expect("html utf8");
    let html = html.replace("\r\n", "\n");
    let css = String::from_utf8(store.read("app.css").expect("css").content).expect("css utf8");
    let js = String::from_utf8(store.read("app.js").expect("js").content).expect("js utf8");

    assert!(html.contains("<title>Office MCP Control</title>"));
    assert!(html.contains("<h1>Office MCP Control</h1>"));
    assert!(html.contains("<span id=\"healthBadge\" class=\"badge neutral\">Loading…</span>"));
    assert!(!html.contains("<span id=\"healthBadge\" class=\"badge neutral\">Loading</span>"));
    assert!(html.contains("<img class=\"product-mark\" src=\"/assets/icon-32.png\" width=\"32\" height=\"32\" alt=\"\" aria-hidden=\"true\" />"));
    assert!(!html.contains("<title>Office MCP</title>"));
    assert!(!html.contains("<h1>Office MCP</h1>"));
    assert!(html.contains("aria-label=\"Copy MCP endpoint\""));
    assert!(html.contains("aria-label=\"Copy add-in endpoint\""));
    assert!(html.contains("aria-label=\"Copy config path\""));
    assert!(html.contains("aria-label=\"Copy log path\""));
    assert!(html.contains("aria-label=\"Open config file location\""));
    assert!(html.contains("aria-label=\"Open log file location\""));
    assert!(html.contains("data-open-diagnostic=\"config\""));
    assert!(html.contains("data-open-diagnostic=\"log\""));
    assert!(html.contains("id=\"toolAccessPanel\""));
    assert!(html.contains("Global Tool Access"));
    assert!(html.contains("id=\"toolAccessMode\""));
    assert!(html.contains("data-access-mode=\"read\""));
    assert!(html.contains("data-access-mode=\"write\""));
    assert!(html.contains("data-access-mode=\"all\""));
    assert!(html.contains("<div><dt>Active Tasks</dt><dd id=\"taskCount\">0</dd></div>"));
    assert!(!html.contains("<dt>Running</dt><dd id=\"taskCount\">"));
    assert!(html.contains("<option value=\"timeout\">Timed Out</option>"));
    assert!(!html.contains("<option value=\"timeout\">Timed out</option>"));
    assert!(html.contains("<div class=\"details\" aria-label=\"Daemon details\">"));
    assert!(html.contains("</div>\n      </header>\n\n      <section class=\"workspace\">"));
    assert!(!html.contains("<section class=\"details\" aria-label=\"Daemon details\">"));
    assert!(html.contains("class=\"detail-path-value\"><code id=\"configPath\" tabindex=\"0\""));
    assert!(html.contains("class=\"detail-path-value\"><code id=\"logPath\" tabindex=\"0\""));
    assert!(html.contains("class=\"detail-copy\" data-copy=\"configPath\""));
    assert!(html.contains("class=\"detail-copy\" data-copy=\"logPath\""));
    assert!(html.contains("class=\"detail-copy\" data-copy=\"lastError\""));
    assert!(html.contains("<textarea id=\"lastError\" readonly spellcheck=\"false\" aria-label=\"Last daemon error\">None</textarea>"));
    assert!(!html.contains("<code id=\"lastError\""));
    assert!(!html.contains("class=\"detail-copy\" data-copy=\"logPath\" aria-label=\"Copy log path\"><code id=\"logPath\""));
    assert!(html.contains("id=\"appFilter\" name=\"app-filter\""));
    assert!(html.contains("aria-label=\"Filter documents by app\""));
    assert!(html.contains("class=\"activity-filters\" aria-label=\"Activity filters\""));
    assert!(html.contains("id=\"clientFilter\" name=\"client-filter\""));
    assert!(html.contains("aria-label=\"Filter activity by client\""));
    assert!(html.contains("<option value=\"all\">All clients</option>"));
    assert!(html.contains("name=\"session-filter\""));
    assert!(html.contains("Title, app, session&hellip;"));
    assert!(!html.contains("session..."));
    assert!(html.contains("aria-live=\"polite\""));
    assert!(css.contains("grid-template-rows: auto minmax(0, 1fr) auto"));
    assert!(css.contains(".status-strip, .panel"));
    assert!(!css.contains(".status-strip, .details, .panel { background"));
    assert!(css.contains("grid-template-columns: 32px auto minmax(0, 1fr)"));
    assert!(css.contains(".product-mark { width: 32px; height: 32px;"));
    assert!(css.contains(
        ".detail-path-value { display: grid; grid-template-columns: minmax(0, 1fr) auto;"
    ));
    assert!(css.contains(".detail-path-value code { display: block; min-width: 0; white-space: normal; overflow-wrap: anywhere; user-select: text;"));
    assert!(css.contains(".detail-actions { display: inline-flex; gap: 2px; align-items: center;"));
    assert!(css.contains(
        ".detail-log-value { display: grid; grid-template-columns: minmax(0, 1fr) auto;"
    ));
    assert!(css.contains(".detail-log-value textarea { display: block; width: 100%; min-width: 0; min-height: 96px; max-height: 240px; resize: vertical; overflow: auto;"));
    assert!(css.contains("white-space: pre-wrap; overflow-wrap: anywhere; user-select: text;"));
    assert!(!css.contains(".detail-log-value textarea { display: block; width: 100%; min-width: 0; overflow: hidden; text-overflow: ellipsis;"));
    assert!(css.contains(".detail-copy, .detail-open { display: inline-flex; min-height: 24px;"));
    assert!(css.contains(".id-copy { display: inline-flex; max-width: 100%; min-height: 32px;"));
    assert!(css.contains(".id-copy code { display: block; max-width: 18ch; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(css.contains(
        ".document-card-session { display: grid; grid-template-columns: minmax(0, 1fr);"
    ));
    assert!(css.contains(
        ".document-card-session span { overflow: visible; text-overflow: clip; white-space: nowrap;"
    ));
    assert!(css.contains(".document-card-session code { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(css.contains(".tool-access-panel { display: grid;"));
    assert!(css.contains(".segmented { display: inline-grid; grid-template-columns: repeat(3"));
    assert!(css.contains(".tool-access-list details { border: 1px solid var(--border);"));
    assert!(css.contains(".tool-access-toggle { display: inline-flex; width: 34px;"));
    assert!(css.contains(
        ".detail-copy:hover, .detail-copy:focus-visible, .detail-open:hover, .detail-open:focus-visible, .id-copy:hover, .id-copy:focus-visible"
    ));
    assert!(css.contains(".details { grid-column: 1 / -1;"));
    assert!(css.contains("border-top: 1px solid var(--border); padding-top: 8px;"));
    assert!(css.contains(".details dl { display: grid; grid-template-columns: minmax(64px, .35fr) minmax(64px, .35fr) minmax(150px, .9fr) minmax(150px, .9fr) minmax(320px, 2.3fr);"));
    assert!(
        css.contains(".details dd { min-width: 0; overflow-wrap: anywhere; user-select: text;")
    );
    assert!(css.contains(".details .detail-path code { direction: ltr; text-align: left;"));
    assert!(!css.contains(".details dd { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(css.contains("--powerpoint: #b7472a;"));
    assert!(css.contains("--accent:"));
    assert!(css.contains(".accent { color: var(--accent); }"));
    assert!(css.contains(".row.powerpoint { border-left-color: var(--powerpoint); }"));
    assert!(css.contains(".filter-row { display: grid; grid-template-columns: auto minmax(82px, 104px) auto minmax(120px, 1fr);"));
    assert!(css.contains(".activity-filters { display: grid; grid-template-columns: auto minmax(110px, 160px) auto minmax(112px, 150px);"));
    assert!(css.contains("#search, #appFilter, #clientFilter, #resultFilter { min-width: 0;"));
    assert!(css.contains("minmax(0, 1fr)"));
    assert!(css.contains("content-visibility: auto"));
    assert!(css.contains(".empty strong"));
    assert!(
        css.contains(".empty-copy { display: grid; grid-template-columns: auto minmax(0, 1fr);")
    );
    assert!(css.contains(".empty-copy code { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(!css.contains("transition: all"));
    assert!(js.contains("emptyState('No documents connected'"));
    assert!(js.contains("$('configPath').textContent = snapshot.daemon?.config_path || '-'"));
    assert!(js.contains("$('daemonMeta').textContent = 'Daemon state is live'"));
    assert!(js.contains("setTextareaValue($('lastError'), snapshot.daemon?.last_error || 'None')"));
    assert!(js.contains("function setTextareaValue(textarea, value)"));
    assert!(js.contains("if (!textarea || textarea.value === value) return;"));
    assert!(js.contains("textarea.setSelectionRange(Math.min(selectionStart, value.length), Math.min(selectionEnd, value.length))"));
    assert!(
        js.contains("const value = copy.dataset.copyValue || target?.value || target?.textContent")
    );
    assert!(js.contains("event.target.closest('[data-open-diagnostic]')"));
    assert!(js.contains("function handleRowNavigation(event)"));
    assert!(js.contains("renderCommands('currentTasks', filterCommands(snapshot.current_tasks || [], true), true)"));
    assert!(js.contains("const history = filterCommands(snapshot.recent_commands || [], false)"));
    assert!(js.contains("function filterCommands(commands, running)"));
    assert!(js.contains("(running || state.result === 'all' || command.status === state.result) && matches(JSON.stringify(command))"));
    assert!(js.contains("['ArrowDown', 'ArrowUp', 'Home', 'End', 'PageDown', 'PageUp']"));
    assert!(js.contains("row.closest('#documents, #clients, #currentTasks, #history')"));
    assert!(js.contains("filter((item) => !item.disabled)"));
    assert!(!js.contains("offsetParent !== null"));
    assert!(js.contains("function rowNavigationIndex(key, index, count, pageStep)"));
    assert!(js.contains("const focusKey = focusedRowKey()"));
    assert!(js.contains("selectedRowKey: null"));
    assert!(js.contains("restoreSelectedRow()"));
    assert!(js.contains("restoreRowFocus(focusKey)"));
    assert!(js.contains("function focusedRowKey()"));
    assert!(js.contains("function restoreRowFocus(focusKey)"));
    assert!(js.contains("function restoreSelectedRow()"));
    assert!(js.contains("row.dataset.focusKey === state.selectedRowKey ? 'true' : 'false'"));
    assert!(js.contains("state.selectedRowKey = element.dataset.focusKey || null"));
    assert!(js.contains("state.selectedRowKey = null"));
    assert!(js.contains("data-focus-key=\"document:${esc(doc.session_id || label)}\""));
    assert!(js.contains("type=\"button\" data-key-activate data-focus-key=\"client:${esc(client.client_id || client.name)}\""));
    assert!(js.contains("data-focus-key=\"client:${esc(client.client_id || client.name)}\""));
    assert!(js.contains("clientEmptyState(state.snapshot?.daemon?.mcp_endpoint)"));
    assert!(js.contains("function clientEmptyState(mcpEndpoint)"));
    assert!(js.contains("No MCP clients connected</strong>Connect an MCP client using either local transport."));
    assert!(js.contains("emptyCopy(mcpEndpoint, 'Copy MCP endpoint', 'MCP endpoint')"));
    assert!(js.contains("emptyCopy('office-mcp-daemon stdio', 'Copy stdio bridge command', 'Stdio bridge')"));
    assert!(js.contains("function emptyCopy(codeText, copyLabel, label)"));
    assert!(js.contains("data-focus-key=\"command:${esc(command.command_id || command.mcp_request_id || command.tool)}\""));
    assert!(js.contains("fetch('/ui/open-diagnostic'"));
    assert!(js.contains("method: 'POST'"));
    assert!(js.contains("body: JSON.stringify({ target })"));
    assert!(!js.contains("$('lastError').textContent"));
    assert!(!js.contains("$('daemonMeta').textContent = snapshot.daemon?.last_error"));
    assert!(!js.contains("'Not configured'"));
    assert!(js.contains("renderToolAccess(snapshot.daemon?.tool_catalog || [], snapshot.daemon?.tool_access_policy || {})"));
    assert!(js.contains("function renderToolAccess(catalog, policy)"));
    assert!(js.contains("function groupedToolAccessCatalog(catalog)"));
    assert!(js.contains("fetch('/ui/tool-access-policy'"));
    assert!(js.contains("method: 'PUT'"));
    assert!(js.contains("event.stopPropagation()"));
    assert!(js.contains("function renderDocumentCard(doc, app)"));
    assert!(js.contains("class=\"row document-card ${esc(app)}\""));
    assert!(js.contains("const sessionId = doc.session_id || '-'"));
    assert!(js.contains("const hostVersion = `${title(doc.host?.app || app)} ${doc.host?.version || '-'}`"));
    assert!(js.contains(
        "class=\"document-card-session\" data-copy-value=\"${esc(sessionId)}\""
    ));
    assert!(js.contains("<span>Session ID</span><code>${esc(middleTruncate(sessionId, 24))}</code>"));
    assert!(!js.contains("Session ${esc(middleTruncate(doc.session_id, 18))}"));
    assert!(js.contains("title=\"${esc(sessionId)}\""));
    assert!(js.contains("<span>${esc(hostVersion)}</span>"));
    assert!(!js.contains("<span>Version ${esc(doc.host?.version || '-')}</span>"));
    assert!(js.contains("function documentConnectionLabel(status) { return status === 'active' || !status ? 'active' : 'dead'; }"));
    assert!(js.contains("function documentStateTone(status) { return status === 'active' || !status ? 'success' : 'danger'; }"));
    assert!(js.contains("function statusLabel(value)"));
    assert!(js.contains("if (value === 'success') return 'Succeeded';"));
    assert!(js.contains("if (value === 'failure') return 'Failed';"));
    assert!(js.contains("if (value === 'timeout') return 'Timed Out';"));
    assert!(js.contains("if (value === 'cancelled') return 'Cancelled';"));
    assert!(js.contains("value === 'timeout' ? 'warning'"));
    assert!(js.contains("value === 'running' ? 'accent'"));
    assert!(js.contains("Finished ${esc(metrics.finished)}"));
    assert!(js.contains("Failed ${esc(metrics.failed)}"));
    assert!(js.contains("function documentTaskMetrics(sessionId)"));
    assert!(js.contains("command.status === 'success'"));
    assert!(!js.contains("function renderDocumentHistory"));
    assert!(!js.contains("doc-history"));
    assert!(!js.contains("aria-expanded"));
    assert!(!js.contains("Show details"));
    assert!(!js.contains("Hide details"));
    assert!(js.contains("state.app = event.target.value"));
    assert!(js.contains("client: 'all'"));
    assert!(js.contains("$('clientFilter').addEventListener('change'"));
    assert!(js.contains("renderClientFilter(snapshot)"));
    assert!(js.contains("function renderClientFilter(snapshot)"));
    assert!(js.contains("function clientFilterOptions(snapshot)"));
    assert!(js.contains("function clientMatches(command)"));
    assert!(js.contains("clientMatches(command) && (running || state.result === 'all' || command.status === state.result)"));
    assert!(js.contains("if (state.app !== 'all' && app !== state.app) continue;"));
    assert!(js.contains("emptyState('No matching documents'"));
    assert!(js.contains("Open Word, Excel, or PowerPoint, then open Office MCP Control."));
    assert!(js.contains("'Copy add-in endpoint'"));
    assert!(js.contains("'Copy MCP endpoint'"));
    assert!(js.contains("const label = copyLabel.replace(/^Copy\\s+/i, '')"));
    assert!(js.contains("class=\"empty-copy\" data-copy-value=\"${esc(codeText)}\""));
    assert!(js.contains("<span>${esc(label)}</span>"));
    assert!(js.contains("middleTruncate(codeText, 46)"));
    assert!(!js.contains("load the Office MCP add-in"));
    assert!(js.contains("fallbackCopy"));
    assert!(js.contains("announceStatus"));
    assert!(js.contains("event.target.closest('[data-copy], [data-copy-value]')"));
    assert!(js.contains("function copyableId(value, label)"));
    assert!(js.contains("class=\"id-copy\" data-copy-value=\"${esc(text)}\""));
    assert!(js.contains("function middleTruncate(value, maxLength = 30)"));
    assert!(!js.contains("stale | reconnecting"));
    assert!(!js.contains("copyableId(doc.session_id, 'Copy session ID')"));
    assert!(
        js.contains("copyableId(command.command_id || command.mcp_request_id, 'Copy command ID')")
    );
    assert!(js.contains("copyableId(command.session_id, 'Copy session ID')"));
}
