use super::{json_escape, render_redacted_config, ui_url_from_runtime_path};
use office_mcp_daemon::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, LimitsConfig, LoggingConfig, McpConfig,
};

#[test]
fn redacted_config_json_hides_secrets_and_includes_endpoints() {
    let config = DaemonConfig {
        addin: AddinConfig {
            host: "localhost".to_string(),
            port: 8765,
            origin: "https://localhost:8765".to_string(),
            pfx_path: "C:\\cert.pfx".to_string(),
            pfx_passphrase: "secret".to_string(),
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
            file: String::new(),
        },
    };

    let json = render_redacted_config(&config);

    assert!(json.contains("<redacted>"));
    assert!(!json.contains("secret"));
    assert!(json.contains("http://127.0.0.1:8800/mcp"));
    assert!(json.contains("wss://localhost:8765/addin"));
}

#[test]
fn escapes_json_control_characters() {
    assert_eq!(json_escape("a\\b\"c\n"), "a\\\\b\\\"c\\n");
}

#[test]
fn ui_command_reads_runtime_file_url_instead_of_config_defaults() {
    let dir =
        std::env::temp_dir().join(format!("office-mcp-ui-command-test-{}", std::process::id()));
    let path = dir.join("ui-runtime.json");
    std::fs::create_dir_all(&dir).expect("temp dir");
    std::fs::write(
        &path,
        concat!(
            "{",
            "\"origin\":\"https://localhost:8766\",",
            "\"stateUrl\":\"https://localhost:8766/ui/state\",",
            "\"uiUrl\":\"https://localhost:8766/ui/\",",
            "\"pid\":123,",
            "\"createdAt\":\"1\"",
            "}"
        ),
    )
    .expect("runtime file");

    let url = ui_url_from_runtime_path(&path).expect("ui url");

    assert_eq!(url, "https://localhost:8766/ui/");
    let _ = std::fs::remove_dir_all(dir);
}
