use super::{find_addin_public_dir_from, static_asset_content_type};
use std::fs;
use std::path::{Path, PathBuf};

#[test]
fn maps_static_asset_content_types() {
    assert_eq!(
        static_asset_content_type(Path::new("taskpane.html")),
        "text/html; charset=utf-8"
    );
    assert_eq!(
        static_asset_content_type(Path::new("taskpane.js")),
        "text/javascript; charset=utf-8"
    );
    assert_eq!(
        static_asset_content_type(Path::new("taskpane.css")),
        "text/css; charset=utf-8"
    );
    assert_eq!(
        static_asset_content_type(Path::new("icon.png")),
        "application/octet-stream"
    );
}

#[test]
fn finds_word_addin_public_dir_from_descendant() {
    let root = temp_root("office-mcp-addin-public-dir");
    let public = root
        .join("src")
        .join("office-ctl")
        .join("word")
        .join("public");
    let nested = root.join("src").join("office-mcp").join("daemon");
    fs::create_dir_all(&public).expect("public dir");
    fs::create_dir_all(&nested).expect("nested dir");
    fs::write(public.join("taskpane.html"), "<main></main>").expect("taskpane");

    let found = find_addin_public_dir_from(&nested).expect("addin public dir");

    assert_eq!(found, public);
    let _ = fs::remove_dir_all(root);
}

#[test]
fn ignores_word_public_dir_without_taskpane() {
    let root = temp_root("office-mcp-addin-public-dir-missing-taskpane");
    let public = root
        .join("src")
        .join("office-ctl")
        .join("word")
        .join("public");
    fs::create_dir_all(&public).expect("public dir");

    assert!(find_addin_public_dir_from(&root).is_none());
    let _ = fs::remove_dir_all(root);
}

fn temp_root(name: &str) -> PathBuf {
    std::env::temp_dir().join(format!("{name}-{}", std::process::id()))
}
