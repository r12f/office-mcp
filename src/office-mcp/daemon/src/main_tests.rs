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
