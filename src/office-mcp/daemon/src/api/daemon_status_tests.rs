use super::DaemonStatusReporter;
use std::fs;

#[test]
fn status_reports_runtime_details_without_auth_material() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-daemon-status-test-{}",
        std::process::id()
    ));
    let path = dir.join("ui-runtime.json");
    fs::create_dir_all(&dir).expect("temp dir");
    fs::write(
        &path,
        "{\"pid\":0,\"uiUrl\":\"https://localhost:8765/ui/\",\"stateUrl\":\"https://localhost:8765/ui/state\",\"logPath\":\"C:\\\\logs\\\\office-mcp.log\"}",
    )
    .expect("runtime file");

    let json = DaemonStatusReporter::new(path).status_json();

    assert!(json.contains("\"running\": false"));
    assert!(json.contains("https://localhost:8765/ui/"));
    assert!(json.contains("\"uiCommand\": \"office-mcp-daemon ui\""));
    assert!(json.contains("\"logPath\": \"C:\\\\logs\\\\office-mcp.log\""));
    assert!(!json.contains("token"));
    assert!(!json.contains("secret"));
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn status_handles_missing_runtime_file() {
    let json =
        DaemonStatusReporter::new(std::env::temp_dir().join("missing-runtime.json")).status_json();

    assert!(json.contains("\"running\": false"));
    assert!(json.contains("\"pid\": null"));
    assert!(json.contains("\"uiUrl\": null"));
    assert!(json.contains("\"logPath\": null"));
}
