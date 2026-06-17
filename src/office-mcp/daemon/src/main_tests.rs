#[test]
fn daemon_run_starts_background_tray_by_default() {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("main.rs");
    let source = std::fs::read_to_string(source_path).expect("read main source");

    assert!(source.contains("daemon\" && subcommand == \"run\" =>"));
    assert!(source.contains("serve_daemon_with_optional_tray(true)"));
    assert!(source.contains("flag == \"--no-tray\""));
    assert!(source.contains("serve_daemon_with_optional_tray(false)"));
    assert!(source.contains("flag == \"--with-tray\""));
    assert!(source.contains("start_tray_background();"));
    assert!(source.contains("daemon run [--no-tray|--with-tray]"));
}

#[test]
fn daemon_run_has_startup_config_and_shutdown_tracing() {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("main.rs");
    let source = std::fs::read_to_string(source_path).expect("read main source");

    assert!(source.contains("office-mcp-daemon run requested"));
    assert!(source.contains("loaded daemon configuration"));
    assert!(source.contains("validated daemon boundary configuration"));
    assert!(source.contains("office-mcp-daemon started"));
    assert!(source.contains("office-mcp-daemon stopped with error"));
    assert!(source.contains("office-mcp-daemon stopped cleanly"));
    assert!(source.contains("mcp_endpoint = %endpoints.mcp"));
    assert!(source.contains("log_path = %config.logging.file"));
}
