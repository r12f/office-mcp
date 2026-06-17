use super::ui_url_from_runtime_path;

#[test]
fn ui_command_reads_runtime_file_url_instead_of_config_defaults() {
    let dir =
        std::env::temp_dir().join(format!("office-mcp-ui-command-test-{}", std::process::id()));
    let path = dir.join("ui-runtime.json");
    std::fs::create_dir_all(&dir).expect("temp dir");
    std::fs::write(
        &path,
        concat!(
            "{",
            "\"origin\":\"https://localhost:8766\",",
            "\"stateUrl\":\"https://localhost:8766/ui/state\",",
            "\"uiUrl\":\"https://localhost:8766/ui/\",",
            "\"pid\":123,",
            "\"createdAt\":\"1\"",
            "}"
        ),
    )
    .expect("runtime file");

    let url = ui_url_from_runtime_path(&path).expect("ui url");

    assert_eq!(url, "https://localhost:8766/ui/");
    let _ = std::fs::remove_dir_all(dir);
}

#[test]
fn daemon_run_with_tray_is_wired_to_background_tray_start() {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("main.rs");
    let source = std::fs::read_to_string(source_path).expect("read main source");

    assert!(source.contains("daemon\" && subcommand == \"run\" && flag == \"--with-tray"));
    assert!(source.contains("serve_daemon_with_optional_tray(true)"));
    assert!(source.contains("start_tray_background();"));
}
