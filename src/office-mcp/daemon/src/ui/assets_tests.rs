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
