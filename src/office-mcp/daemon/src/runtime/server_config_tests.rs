use super::{RuntimeServerConfig, RuntimeServerError};
use crate::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, LimitsConfig, LoggingConfig, McpConfig,
};

#[test]
fn runtime_config_converts_daemon_config_into_socket_settings() {
    let config = RuntimeServerConfig::from_daemon_config(&daemon_config()).expect("runtime config");

    assert_eq!(config.mcp_bind_addr(), "127.0.0.1:8801");
    assert_eq!(config.addin_bind_addr(), "localhost:8766");
    assert_eq!(config.addin_origin, "https://localhost:8766");
    assert_eq!(config.max_pending_per_session, 7);
    assert_eq!(config.session_grace, std::time::Duration::from_secs(42));
    assert_eq!(
        config.config_path.as_deref(),
        Some("C:\\office-mcp\\config.toml")
    );
    assert_eq!(config.log_path.as_deref(), Some("C:\\logs\\office-mcp.log"));
    assert_eq!(config.mcp_http_config().requests_per_minute, 99);
    assert_eq!(
        config.addin_channel_config().origin,
        "https://localhost:8766"
    );
    assert_eq!(
        config.addin_channel_config().session_grace,
        std::time::Duration::from_secs(42)
    );
}

#[test]
fn runtime_config_rejects_ports_that_do_not_fit_tcp_port_type() {
    let mut daemon_config = daemon_config();
    daemon_config.mcp.port = u64::from(u16::MAX) + 1;

    let error = RuntimeServerConfig::from_daemon_config(&daemon_config).expect_err("port error");

    assert!(matches!(error, RuntimeServerError::InvalidConfig(_)));
    assert!(error.to_string().contains("mcp.port"));
}

#[test]
fn default_runtime_config_uses_loopback_mcp_endpoint() {
    let config = RuntimeServerConfig::default();

    assert_eq!(config.mcp_bind_addr(), "127.0.0.1:8800");
    assert_eq!(config.addin_bind_addr(), "localhost:8765");
}

fn daemon_config() -> DaemonConfig {
    DaemonConfig {
        config_path: "C:\\office-mcp\\config.toml".to_string(),
        addin: AddinConfig {
            host: "localhost".to_string(),
            port: 8766,
            origin: "https://localhost:8766".to_string(),
            pfx_path: "C:\\cert.pfx".to_string(),
            pfx_passphrase: "secret".to_string(),
            heartbeat_interval_sec: 30,
            heartbeat_timeout_sec: 10,
            session_grace_sec: 42,
            max_pending_per_session: 7,
        },
        mcp: McpConfig {
            host: "127.0.0.1".to_string(),
            port: 8801,
        },
        limits: LimitsConfig {
            max_response_bytes: 1,
            max_request_bytes: 2,
            max_ws_frame_bytes: 3,
            default_tool_timeout_ms: 4,
            requests_per_minute: 99,
        },
        audit: AuditConfig {
            enabled: false,
            path: String::new(),
        },
        logging: LoggingConfig {
            level: ConfigLogLevel::Info,
            file: "C:\\logs\\office-mcp.log".to_string(),
        },
    }
}
