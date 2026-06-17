use super::{UiRuntimeFile, UiRuntimeInfo, default_path_from_env};
use crate::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, LimitsConfig, LoggingConfig, McpConfig,
};
use std::collections::BTreeMap;
use std::fs;

#[test]
fn runtime_info_uses_ui_urls_without_credentials() {
    let info = UiRuntimeInfo::with_origin("https://localhost:8765".to_string());
    let json = info.to_json();

    assert!(json.contains("\"origin\": \"https://localhost:8765\""));
    assert!(json.contains("\"stateUrl\": \"https://localhost:8765/ui/state\""));
    assert!(json.contains("\"uiUrl\": \"https://localhost:8765/ui/\""));
    assert!(json.contains("\"logPath\": null"));
    assert!(!json.contains("token"));
    assert!(!json.contains("secret"));
}

#[test]
fn runtime_info_can_publish_configured_log_path() {
    let info = UiRuntimeInfo::with_origin_and_log_path(
        "https://localhost:8765".to_string(),
        Some("C:\\logs\\office-mcp.log".to_string()),
    );

    let json = info.to_json();
    let parsed = UiRuntimeInfo::from_json(&json).expect("parse runtime info");

    assert!(json.contains("\"logPath\": \"C:\\\\logs\\\\office-mcp.log\""));
    assert_eq!(parsed.log_path.as_deref(), Some("C:\\logs\\office-mcp.log"));
}

#[test]
fn writes_and_removes_runtime_file() {
    let dir =
        std::env::temp_dir().join(format!("office-mcp-ui-runtime-test-{}", std::process::id()));
    let path = dir.join("ui-runtime.json");
    let file = UiRuntimeFile::with_path(
        path.clone(),
        UiRuntimeInfo::with_origin("https://localhost:8765".to_string()),
    );

    file.write().expect("runtime file writes");
    let body = fs::read_to_string(&path).expect("runtime file readable");
    assert!(body.contains("https://localhost:8765/ui/"));

    file.remove().expect("runtime file removes");
    assert!(!path.exists());
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn reads_runtime_file_body() {
    let info = UiRuntimeInfo::with_origin("https://localhost:8765".to_string());
    let parsed = UiRuntimeInfo::from_json(&info.to_json()).expect("parse runtime info");

    assert_eq!(parsed.origin, info.origin);
    assert_eq!(parsed.state_url, info.state_url);
    assert_eq!(parsed.ui_url, info.ui_url);
    assert_eq!(parsed.log_path, None);
}

#[test]
fn default_runtime_path_uses_local_app_data_on_windows() {
    if !cfg!(windows) {
        return;
    }
    let env = BTreeMap::from([("LOCALAPPDATA".to_string(), "C:\\Local".to_string())]);
    assert_eq!(
        default_path_from_env(&env),
        std::path::PathBuf::from("C:\\Local")
            .join("office-mcp")
            .join("ui-runtime.json")
    );
}

#[test]
fn runtime_file_can_be_built_from_config() {
    let config = DaemonConfig {
        addin: AddinConfig {
            host: "localhost".to_string(),
            port: 8765,
            origin: "https://localhost:8765".to_string(),
            pfx_path: String::new(),
            pfx_passphrase: String::new(),
            heartbeat_interval_sec: 30,
            heartbeat_timeout_sec: 10,
            session_grace_sec: 60,
            max_pending_per_session: 4,
        },
        mcp: McpConfig {
            host: "127.0.0.1".to_string(),
            port: 8800,
        },
        limits: LimitsConfig {
            max_response_bytes: 1,
            max_request_bytes: 1,
            max_ws_frame_bytes: 1,
            default_tool_timeout_ms: 1,
            requests_per_minute: 1,
        },
        audit: AuditConfig {
            enabled: false,
            path: String::new(),
        },
        logging: LoggingConfig {
            level: ConfigLogLevel::Info,
            file: "C:\\logs\\office-mcp.log".to_string(),
        },
    };

    let file = UiRuntimeFile::from_config(&config);
    assert_eq!(file.info().origin, "https://localhost:8765");
    assert_eq!(
        file.info().log_path.as_deref(),
        Some("C:\\logs\\office-mcp.log")
    );
    assert!(file.path().ends_with("ui-runtime.json"));
}
