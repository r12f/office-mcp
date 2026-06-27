use super::{RecordingUiLauncher, TrayUiOpenRequest};
use crate::tray::TrayHostOptions;
use crate::ui::{UiRuntimeFile, UiRuntimeInfo};
use std::fs;

#[test]
fn request_uses_runtime_file_url_when_available() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-tray-ui-launch-runtime-{}",
        std::process::id()
    ));
    fs::create_dir_all(&dir).expect("create temp dir");
    let runtime_path = dir.join("ui-runtime.json");
    UiRuntimeFile::with_path(
        runtime_path.clone(),
        UiRuntimeInfo::with_origin("https://localhost:9988".to_string()),
    )
    .write()
    .expect("write runtime file");

    let request = TrayUiOpenRequest::from_runtime(&TrayHostOptions {
        runtime_path: Some(runtime_path),
        probe_state_path: None,
        probe: false,
    })
    .expect("resolve runtime request");

    assert_eq!(request.url(), "https://localhost:9988/ui/");
    assert_eq!(request.source(), "runtime_file");
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn request_falls_back_when_runtime_file_is_missing() {
    let request = TrayUiOpenRequest::from_runtime(&TrayHostOptions {
        runtime_path: Some(std::env::temp_dir().join(format!(
            "office-mcp-ui-runtime-missing-{}.json",
            std::process::id()
        ))),
        probe_state_path: None,
        probe: false,
    })
    .expect("resolve fallback request");

    assert_eq!(request.url(), "https://localhost:8765/ui/");
    assert_eq!(request.source(), "fallback");
}

#[test]
fn request_records_launcher_url() {
    let launcher = RecordingUiLauncher::default();
    let request = TrayUiOpenRequest::new("https://localhost:8765/ui/".to_string(), "fallback");

    request.open_with(&launcher).expect("open request");

    assert_eq!(launcher.opened_urls(), vec!["https://localhost:8765/ui/"]);
}
