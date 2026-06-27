use super::{TrayHost, TrayHostOptions};
use crate::common::{Logger, LoggerLogLevel};
use crate::tray::quit_request::{RecordingShutdownController, TrayQuitRequest};
use crate::tray::ui_launch::{RecordingUiLauncher, TrayUiOpenRequest};
use crate::ui::{UiRuntimeFile, UiRuntimeInfo};
use std::fs;
use std::fs::{read_to_string, remove_dir_all};

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
fn tray_probe_writes_structured_tracing_event() {
    let dir = std::env::temp_dir().join(format!("office-mcp-tray-host-log-{}", std::process::id()));
    fs::create_dir_all(&dir).expect("create temp dir");
    let state_path = dir.join("state.json");
    let log_path = dir.join("office-mcp.log");
    fs::write(
        &state_path,
        r#"{"daemon":{"status":"up"},"clients":[{}],"documents":{"word":[{}]},"current_tasks":[]}"#,
    )
    .expect("write state");
    let (subscriber, guard) =
        Logger::tracing_file_default(LoggerLogLevel::Info, &log_path).expect("init tracing");
    let host = TrayHost::new(TrayHostOptions {
        runtime_path: None,
        probe_state_path: Some(state_path),
        probe: true,
    });

    tracing::subscriber::with_default(subscriber, || host.run().expect("run probe"));
    drop(guard);

    let contents = read_to_string(&log_path).expect("read tracing log file");
    assert!(contents.contains("ran tray probe"));
    assert!(contents.contains("\"component\":\"tray_host\""));
    assert!(contents.contains("\"state_fetch_ok\":true"));
    let _ = remove_dir_all(dir);
}

#[test]
fn show_ui_action_opens_runtime_ui_url_through_launcher() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-tray-show-ui-runtime-{}",
        std::process::id()
    ));
    fs::create_dir_all(&dir).expect("create temp dir");
    let runtime_path = dir.join("ui-runtime.json");
    UiRuntimeFile::with_path(
        runtime_path.clone(),
        UiRuntimeInfo::with_origin("https://localhost:9876".to_string()),
    )
    .write()
    .expect("write runtime file");
    let launcher = RecordingUiLauncher::default();
    let options = TrayHostOptions {
        runtime_path: Some(runtime_path),
        probe_state_path: None,
        probe: false,
    };

    let request = TrayUiOpenRequest::from_runtime(&options).expect("resolve runtime UI request");
    request.open_with(&launcher).expect("open UI request");

    assert_eq!(launcher.opened_urls(), vec!["https://localhost:9876/ui/"]);
    assert_eq!(request.action(), "show_ui");
    assert_eq!(request.source(), "runtime_file");
    assert_eq!(request.process_id(), std::process::id());
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn show_ui_action_error_includes_action_source_url_and_pid() {
    let missing_runtime_path = std::env::temp_dir().join(format!(
        "office-mcp-missing-ui-runtime-{}.json",
        std::process::id()
    ));
    let launcher = RecordingUiLauncher::failing("launcher unavailable");
    let options = TrayHostOptions {
        runtime_path: Some(missing_runtime_path),
        probe_state_path: None,
        probe: false,
    };

    let error = TrayUiOpenRequest::from_runtime(&options)
        .expect("resolve fallback UI request")
        .open_with(&launcher)
        .expect_err("launcher failure should be returned");
    let message = error.to_string();

    assert!(message.contains("show_ui"));
    assert!(message.contains("fallback"));
    assert!(message.contains("https://localhost:8765/ui/"));
    assert!(message.contains(&format!("pid={}", std::process::id())));
    assert!(message.contains("launcher unavailable"));
}

#[test]
fn quit_action_requests_daemon_shutdown_through_controller() {
    let controller = RecordingShutdownController::default();

    let request = TrayQuitRequest::new("native_tray_menu");
    request
        .shutdown_with(&controller)
        .expect("quit request should ask controller to stop daemon");

    assert_eq!(controller.stop_count(), 1);
    assert_eq!(request.action(), "quit");
    assert_eq!(request.source(), "native_tray_menu");
    assert_eq!(request.process_id(), std::process::id());
}

#[test]
fn quit_action_error_includes_action_source_pid_and_controller_error() {
    let controller = RecordingShutdownController::failing("scheduled task stop failed");

    let error = TrayQuitRequest::new("native_tray_menu")
        .shutdown_with(&controller)
        .expect_err("shutdown controller failure should be returned");
    let message = error.to_string();

    assert!(message.contains("quit"));
    assert!(message.contains("native_tray_menu"));
    assert!(message.contains(&format!("pid={}", std::process::id())));
    assert!(message.contains("scheduled task stop failed"));
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
    let source = read_tray_source("native_tray.rs");

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
fn native_tray_menu_uses_platform_menu_primitives_not_web_surfaces() {
    let source = read_tray_source("native_tray.rs");
    let normalized = source.to_ascii_lowercase();

    assert!(
        source.contains("tray_icon::menu::{Menu, MenuEvent, MenuId, MenuItem, PredefinedMenuItem}")
    );
    assert!(source.contains("TrayIconBuilder::new()"));
    assert!(source.contains(".with_menu(Box::new(menu))"));
    assert!(source.contains("Menu::new()"));
    assert!(source.contains("MenuItem::new(label, false, None)"));
    assert!(source.contains("MenuItem::with_id(MenuId::new(SHOW_ID), label, true, None)"));
    assert!(source.contains("PredefinedMenuItem::separator()"));

    for forbidden in [
        "webview",
        "web_view",
        "html",
        "css",
        "frameless",
        "popup",
        "browserwindow",
        "windowbuilder",
        "tao::window",
        "wry::",
    ] {
        assert!(
            !normalized.contains(forbidden),
            "native tray menu must not use custom web or window surface: {forbidden}"
        );
    }
}

#[test]
fn native_tray_icon_uses_generated_product_glyph_not_framework_default() {
    let source = read_tray_source("native_tray.rs");
    let product_icon = read_tray_source("product_icon.rs");

    assert!(
        source.contains(
            "use crate::tray::product_icon::{ICON_HEIGHT, ICON_WIDTH, product_icon_rgba};"
        )
    );
    assert!(source.contains(".with_icon(app_icon())"));
    assert!(source.contains("Icon::from_rgba(product_icon_rgba(), ICON_WIDTH, ICON_HEIGHT)"));
    assert!(source.contains("generated tray icon is valid"));
    assert!(product_icon.contains("pub(crate) const ICON_WIDTH: u32 = 32;"));
    assert!(product_icon.contains("pub(crate) const ICON_HEIGHT: u32 = 32;"));
    assert!(product_icon.contains("[248, 216, 74, 255]"));

    for forbidden in [
        "from_path",
        "from_resource",
        "Default::default",
        "include_bytes!",
    ] {
        assert!(
            !source.contains(forbidden),
            "native tray icon must use the generated product glyph, not {forbidden}"
        );
    }
}

#[test]
fn native_tray_allows_background_event_loop_on_windows_and_linux() {
    let source = read_tray_source("native_tray.rs");

    assert!(source.contains("EventLoopBuilderExtWindows"));
    assert!(source.contains("EventLoopBuilderExtUnix"));
    assert!(source.contains("builder.with_any_thread(true);"));
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

#[test]
fn native_tray_actions_emit_tracing_events() {
    let source = format!(
        "{}\n{}\n{}\n{}",
        read_tray_source("host.rs"),
        read_tray_source("native_tray.rs"),
        read_tray_source("ui_launch.rs"),
        read_tray_source("quit_request.rs")
    );

    assert!(source.contains("created native tray icon"));
    assert!(source.contains("refreshing tray snapshot"));
    assert!(source.contains("applied tray snapshot"));
    assert!(source.contains("tray action show UI selected"));
    assert!(source.contains("tray action quit selected"));
    assert!(source.contains("opening daemon UI from tray"));
    assert!(source.contains("stopping daemon from tray"));
}

fn read_tray_source(name: &str) -> String {
    let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("tray")
        .join(name);
    fs::read_to_string(source_path).expect("read tray source")
}
