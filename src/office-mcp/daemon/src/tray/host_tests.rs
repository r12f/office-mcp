use super::{TrayHost, TrayHostOptions};
use std::fs;

#[test]
fn probe_snapshot_reads_state_file() {
    let dir =
        std::env::temp_dir().join(format!("office-mcp-tray-host-test-{}", std::process::id()));
    fs::create_dir_all(&dir).expect("create temp dir");
    let state_path = dir.join("state.json");
    fs::write(
        &state_path,
        r#"{"daemon":{"status":"up"},"clients":[{}],"documents":{"word":[{}]},"current_tasks":[]}"#,
    )
    .expect("write state");

    let host = TrayHost::new(TrayHostOptions {
        runtime_path: None,
        probe_state_path: Some(state_path),
        probe: true,
    });
    let probe = host.probe_json();

    assert_eq!(probe["native_host"], true);
    assert_eq!(probe["snapshot"]["menu_items"][0], "Status: Up");
    assert_eq!(probe["snapshot"]["menu_items"][1], "Clients: 1");
    assert_eq!(probe["snapshot"]["menu_items"][2], "Documents: 1");
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn linux_tray_manifest_enables_native_status_notifier_features() {
    let manifest = fs::read_to_string(env!("CARGO_MANIFEST_PATH")).expect("read manifest");

    assert!(manifest.contains("target_os = \"linux\""));
    assert!(manifest.contains("features = [\"dbus\", \"rwh_06\", \"x11\"]"));
    assert!(manifest.contains("features = [\"gtk\", \"libxdo\"]"));
}

#[test]
fn native_tray_quit_uses_platform_confirmation_dialogs() {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("tray")
        .join("host.rs");
    let source = fs::read_to_string(source_path).expect("read source");

    assert!(source.contains("confirm_quit(&surface.snapshot.quit_confirmation)"));
    assert!(source.contains("impl TrayPlatformAdapter for NativeTraySurface"));
    assert!(source.contains("surface.show_ui()"));
    assert!(source.contains("surface.quit()"));
    assert!(source.contains("System.Windows.MessageBox"));
    assert!(source.contains("display dialog"));
    assert!(source.contains("zenity"));
    assert!(source.contains("kdialog"));
}

#[test]
fn background_tray_launcher_owns_native_tray_thread() {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("tray")
        .join("host.rs");
    let source = fs::read_to_string(source_path).expect("read source");

    assert!(source.contains("pub fn start_tray_background()"));
    assert!(source.contains(".name(\"office-mcp-tray\".to_string())"));
    assert!(source.contains("TrayHost::new(TrayHostOptions::default()).run()"));
    assert!(source.contains("office-mcp tray host thread failed to start"));
}
