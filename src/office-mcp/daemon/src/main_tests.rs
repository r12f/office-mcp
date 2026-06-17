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

fn read_main_source() -> String {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("main.rs");
    std::fs::read_to_string(source_path).expect("read main source")
}
