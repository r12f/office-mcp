use super::{json_escape, render_endpoints, render_redacted_config};
use crate::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, EndpointConfig, LimitsConfig,
    LoggingConfig, McpConfig,
};

#[test]
fn redacted_config_json_hides_secrets_and_includes_endpoints() {
    let config = config();

    let json = render_redacted_config(&config);

    assert!(json.contains("<redacted>"));
    assert!(!json.contains("secret"));
    assert!(json.contains("http://127.0.0.1:8800/mcp"));
    assert!(json.contains("wss://localhost:8765/addin"));
}

#[test]
fn endpoints_json_uses_public_endpoint_names() {
    let json = render_endpoints(&EndpointConfig {
        mcp: "http://127.0.0.1:8800/mcp".to_string(),
        addin_origin: "https://localhost:8765".to_string(),
        addin_wss: "wss://localhost:8765/addin".to_string(),
    });

    assert!(json.contains("\"mcp\": \"http://127.0.0.1:8800/mcp\""));
    assert!(json.contains("\"addin_origin\": \"https://localhost:8765\""));
    assert!(json.contains("\"addin_wss\": \"wss://localhost:8765/addin\""));
}

#[test]
fn escapes_json_control_characters() {
    assert_eq!(json_escape("a\\b\"c\n"), "a\\\\b\\\"c\\n");
}

fn config() -> DaemonConfig {
    DaemonConfig {
        config_path: "C:\\office-mcp\\config.toml".to_string(),
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
    }
}
