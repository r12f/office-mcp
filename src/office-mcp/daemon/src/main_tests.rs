#[test]
fn daemon_run_starts_background_tray_by_default() {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("main.rs");
    let source = std::fs::read_to_string(source_path).expect("read main source");

    assert!(source.contains("daemon\" && subcommand == \"run\" =>"));
    assert!(source.contains("serve_daemon_with_optional_tray(true)"));
    assert!(source.contains("flag == \"--with-tray\""));
    assert!(source.contains("start_tray_background();"));
    assert!(source.contains("daemon run [--with-tray]"));
    assert!(!source.contains("flag == \"--no-tray\""));
    assert!(!source.contains("daemon run [--no-tray"));
    assert!(!source.contains("serve_daemon_with_optional_tray(false)"));
}

#[test]
fn daemon_run_starts_tray_after_file_tracing_is_initialized() {
    let source = read_main_source();
    let tracing_index = source
        .find("DaemonLogger::init_tracing_file")
        .expect("tracing init call");
    let tray_index = source
        .find("start_tray_background();")
        .expect("tray background start");

    assert!(
        tracing_index < tray_index,
        "tray startup should be logged to the configured tracing file"
    );
}

#[test]
fn daemon_run_has_startup_config_and_shutdown_tracing() {
    let source = read_main_source();

    assert!(source.contains("office-mcp-daemon run requested"));
    assert!(source.contains("loaded daemon configuration"));
    assert!(source.contains("validated daemon boundary configuration"));
    assert!(source.contains("office-mcp-daemon started"));
    assert!(source.contains("office-mcp-daemon stopped with error"));
    assert!(source.contains("office-mcp-daemon stopped cleanly"));
    assert!(source.contains("mcp_endpoint = %endpoints.mcp"));
    assert!(source.contains("log_path = %config.logging.file"));
}

#[test]
fn default_description_uses_product_language_not_scaffold_wording() {
    let source = read_main_source();

    assert!(source.contains("Office MCP Control daemon"));
    assert!(source.contains("Local Office automation control server for live add-in sessions."));
    assert!(source.contains("Usage: office-mcp-daemon daemon run [--with-tray]"));
    assert!(source.contains("Components:"));
    assert!(!source.to_ascii_lowercase().contains("reference scaffold"));
    assert!(!source.to_ascii_lowercase().contains("debug daemon"));
}

fn read_main_source() -> String {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("main.rs");
    std::fs::read_to_string(source_path).expect("read main source")
}
