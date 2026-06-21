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
    assert!(html.contains("<img class=\"product-mark\" src=\"/assets/icon-32.png\" width=\"32\" height=\"32\" alt=\"\" aria-hidden=\"true\" />"));
    assert!(!html.contains("<title>Office MCP</title>"));
    assert!(!html.contains("<h1>Office MCP</h1>"));
    assert!(html.contains("aria-label=\"Copy MCP endpoint\""));
    assert!(html.contains("aria-label=\"Copy add-in endpoint\""));
    assert!(html.contains("aria-label=\"Copy config path\""));
    assert!(html.contains("aria-label=\"Copy log path\""));
    assert!(html.contains("<div><dt>Active Tasks</dt><dd id=\"taskCount\">0</dd></div>"));
    assert!(!html.contains("<dt>Running</dt><dd id=\"taskCount\">"));
    assert!(html.contains("<div class=\"details\" aria-label=\"Daemon details\">"));
    assert!(html.contains("</div>\n      </header>\n\n      <section class=\"workspace\">"));
    assert!(!html.contains("<section class=\"details\" aria-label=\"Daemon details\">"));
    assert!(html.contains("class=\"detail-copy\" data-copy=\"configPath\""));
    assert!(html.contains("class=\"detail-copy\" data-copy=\"logPath\""));
    assert!(html.contains("id=\"appFilter\" name=\"app-filter\""));
    assert!(html.contains("aria-label=\"Filter documents by app\""));
    assert!(html.contains("name=\"session-filter\""));
    assert!(html.contains("Title, app, session&hellip;"));
    assert!(!html.contains("session..."));
    assert!(html.contains("aria-live=\"polite\""));
    assert!(css.contains("grid-template-rows: auto minmax(0, 1fr) auto"));
    assert!(css.contains(".status-strip, .panel"));
    assert!(!css.contains(".status-strip, .details, .panel { background"));
    assert!(css.contains("grid-template-columns: 32px auto minmax(0, 1fr)"));
    assert!(css.contains(".product-mark { width: 32px; height: 32px;"));
    assert!(css.contains(".detail-copy { display: inline-flex; width: 100%; min-height: 32px;"));
    assert!(css.contains(".detail-copy code { display: block; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(css.contains(".id-copy { display: inline-flex; max-width: 100%; min-height: 32px;"));
    assert!(css.contains(".id-copy code { display: block; max-width: 18ch; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(css.contains(
        ".detail-copy:hover, .detail-copy:focus-visible, .id-copy:hover, .id-copy:focus-visible"
    ));
    assert!(css.contains(".details { grid-column: 1 / -1;"));
    assert!(css.contains("border-top: 1px solid var(--border); padding-top: 8px;"));
    assert!(css.contains(".details dl { display: grid; grid-template-columns: minmax(64px, .35fr) minmax(64px, .35fr) minmax(150px, .9fr) minmax(150px, .9fr) minmax(240px, 1.7fr);"));
    assert!(css.contains(".details dd { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"));
    assert!(css.contains(".details .detail-copy { min-height: 24px; padding: 0 4px;"));
    assert!(css.contains(".details .detail-path code { direction: rtl; text-align: left;"));
    assert!(css.contains("--powerpoint: #b7472a;"));
    assert!(css.contains(".row.powerpoint { border-left-color: var(--powerpoint); }"));
    assert!(css.contains(".filter-row { display: grid; grid-template-columns: auto minmax(82px, 104px) auto minmax(120px, 1fr);"));
    assert!(css.contains("#search, #appFilter { min-width: 0;"));
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
    assert!(!js.contains("'Not configured'"));
    assert!(js.contains("state.app = event.target.value"));
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
    assert!(js.contains("event.stopPropagation()"));
    assert!(js.contains("function copyableId(value, label)"));
    assert!(js.contains("class=\"id-copy\" data-copy-value=\"${esc(text)}\""));
    assert!(js.contains("function middleTruncate(value, maxLength = 30)"));
    assert!(js.contains("copyableId(doc.session_id, 'Copy session ID')"));
    assert!(
        js.contains("copyableId(command.command_id || command.mcp_request_id, 'Copy command ID')")
    );
    assert!(js.contains("copyableId(command.session_id, 'Copy session ID')"));
}
