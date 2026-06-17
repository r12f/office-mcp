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
