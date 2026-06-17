use super::{ConfigError, DaemonConfigService, LoadConfigOptions, LogLevel};
use std::collections::{BTreeMap, HashMap};
use std::env;
use std::fs;
use std::path::PathBuf;

#[test]
fn loads_config_file_values_and_endpoints() {
    let path = temp_config(
        r#"
[addin_channel]
bind = "localhost"
port = 9443
certificate_path = "C:\\certs\\office-mcp.pfx"
certificate_passphrase = "secret"
heartbeat_interval_sec = 11
heartbeat_timeout_sec = 5
session_grace_sec = 22
max_pending_per_session = 2

[mcp_http]
bind = "127.0.0.1"
port = 9900

[limits]
max_response_bytes = 1234
max_request_bytes = 4567
max_ws_frame_bytes = 6789
default_tool_timeout_ms = 321
requests_per_minute = 42

[audit]
enabled = true
path = "C:\\logs\\office-mcp-audit.jsonl"

[logging]
level = "debug"
file = "C:\\logs\\office-mcp.log"
"#,
    );
    let loaded = DaemonConfigService::with_env(BTreeMap::new())
        .load(LoadConfigOptions {
            config_path: Some(path.clone()),
        })
        .expect("load config");

    assert_eq!(loaded.addin.port, 9443);
    assert_eq!(loaded.addin.origin, "https://localhost:9443");
    assert_eq!(loaded.addin.pfx_path, "C:\\certs\\office-mcp.pfx");
    assert_eq!(loaded.addin.pfx_passphrase, "secret");
    assert_eq!(loaded.addin.heartbeat_interval_sec, 11);
    assert_eq!(loaded.addin.heartbeat_timeout_sec, 5);
    assert_eq!(loaded.addin.session_grace_sec, 22);
    assert_eq!(loaded.addin.max_pending_per_session, 2);
    assert_eq!(loaded.mcp.port, 9900);
    assert_eq!(loaded.limits.max_response_bytes, 1234);
    assert_eq!(loaded.limits.max_request_bytes, 4567);
    assert_eq!(loaded.limits.max_ws_frame_bytes, 6789);
    assert_eq!(loaded.limits.default_tool_timeout_ms, 321);
    assert_eq!(loaded.limits.requests_per_minute, 42);
    assert!(loaded.audit.enabled);
    assert_eq!(loaded.audit.path, "C:\\logs\\office-mcp-audit.jsonl");
    assert_eq!(loaded.logging.level, LogLevel::Debug);
    assert_eq!(loaded.logging.file, "C:\\logs\\office-mcp.log");
    let _ = fs::remove_file(path);
}

#[test]
fn office_mcp_config_path_environment_selects_config_file() {
    let path = temp_config(
        r#"
[addin_channel]
bind = "localhost"
port = 8891

[mcp_http]
bind = "127.0.0.1"
port = 8890
"#,
    );
    let loaded = DaemonConfigService::with_env(env_map(HashMap::from([(
        "OFFICE_MCP_CONFIG_PATH",
        path.to_str().expect("path utf8"),
    )])))
    .load(LoadConfigOptions::default())
    .expect("load config from env path");

    assert_eq!(loaded.addin.port, 8891);
    assert_eq!(loaded.mcp.port, 8890);
    let _ = fs::remove_file(path);
}

#[test]
fn empty_audit_path_falls_back_to_platform_default_audit_path() {
    let path = temp_config(
        r#"
[audit]
enabled = true
path = ""
"#,
    );
    let loaded = DaemonConfigService::with_env(BTreeMap::new())
        .load(LoadConfigOptions {
            config_path: Some(path.clone()),
        })
        .expect("load config");

    assert!(loaded.audit.enabled);
    assert!(
        loaded
            .audit
            .path
            .replace('\\', "/")
            .ends_with("office-mcp/audit.jsonl")
    );
    let _ = fs::remove_file(path);
}

#[test]
fn empty_logging_file_falls_back_to_platform_default_log_file() {
    let path = temp_config(
        r#"
[logging]
level = "warn"
file = ""
"#,
    );
    let loaded = DaemonConfigService::with_env(BTreeMap::new())
        .load(LoadConfigOptions {
            config_path: Some(path.clone()),
        })
        .expect("load config");

    assert_eq!(loaded.logging.level, LogLevel::Warn);
    assert!(
        loaded
            .logging
            .file
            .replace('\\', "/")
            .ends_with("office-mcp/office-mcp.log")
    );
    let _ = fs::remove_file(path);
}

#[test]
fn legacy_environment_variables_override_config_file_values() {
    let path = temp_config(
        r#"
[addin_channel]
bind = "0.0.0.0"
port = 9443
certificate_path = "file.pfx"

[mcp_http]
bind = "127.0.0.2"
port = 9900
"#,
    );
    let loaded = DaemonConfigService::with_env(env_map(HashMap::from([
        ("OFFICE_MCP_ADDIN_HOST", "localhost"),
        ("OFFICE_MCP_ADDIN_PORT", "8765"),
        ("OFFICE_MCP_MCP_HOST", "127.0.0.1"),
        ("OFFICE_MCP_MCP_PORT", "8800"),
    ])))
    .load(LoadConfigOptions {
        config_path: Some(path.clone()),
    })
    .expect("load config");

    assert_eq!(loaded.addin.host, "localhost");
    assert_eq!(loaded.addin.port, 8765);
    assert_eq!(loaded.mcp.host, "127.0.0.1");
    assert_eq!(loaded.mcp.port, 8800);
    DaemonConfigService::assert_boundary_auth_config(&loaded).expect("loopback config");
    let _ = fs::remove_file(path);
}

#[test]
fn section_style_environment_variables_override_legacy_names() {
    let loaded = DaemonConfigService::with_env(env_map(HashMap::from([
        ("OFFICE_MCP_ADDIN_HOST", "127.0.0.1"),
        ("OFFICE_MCP_ADDIN_CHANNEL__BIND", "localhost"),
        ("OFFICE_MCP_ADDIN_PORT", "9999"),
        ("OFFICE_MCP_ADDIN_CHANNEL__PORT", "8766"),
        ("OFFICE_MCP_ADDIN_PFX_PATH", "legacy.pfx"),
        ("OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH", "section.pfx"),
        ("OFFICE_MCP_ADDIN_PFX_PASSPHRASE", "legacy-pass"),
        (
            "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE",
            "section-pass",
        ),
        ("OFFICE_MCP_MCP_HOST", "127.0.0.2"),
        ("OFFICE_MCP_MCP_HTTP__BIND", "127.0.0.1"),
        ("OFFICE_MCP_MCP_PORT", "9998"),
        ("OFFICE_MCP_MCP_HTTP__PORT", "8801"),
        ("OFFICE_MCP_ADDIN_HEARTBEAT_INTERVAL_SEC", "99"),
        ("OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_INTERVAL_SEC", "31"),
        ("OFFICE_MCP_ADDIN_HEARTBEAT_TIMEOUT_SEC", "98"),
        ("OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_TIMEOUT_SEC", "12"),
        ("OFFICE_MCP_ADDIN_SESSION_GRACE_SEC", "97"),
        ("OFFICE_MCP_ADDIN_CHANNEL__SESSION_GRACE_SEC", "61"),
        ("OFFICE_MCP_ADDIN_MAX_PENDING_PER_SESSION", "96"),
        ("OFFICE_MCP_ADDIN_CHANNEL__MAX_PENDING_PER_SESSION", "5"),
        ("OFFICE_MCP_MAX_RESPONSE_BYTES", "9997"),
        ("OFFICE_MCP_LIMITS__MAX_RESPONSE_BYTES", "1001"),
        ("OFFICE_MCP_MAX_REQUEST_BYTES", "9996"),
        ("OFFICE_MCP_LIMITS__MAX_REQUEST_BYTES", "1002"),
        ("OFFICE_MCP_MAX_WS_FRAME_BYTES", "9995"),
        ("OFFICE_MCP_LIMITS__MAX_WS_FRAME_BYTES", "1003"),
        ("OFFICE_MCP_DEFAULT_TOOL_TIMEOUT_MS", "9994"),
        ("OFFICE_MCP_LIMITS__DEFAULT_TOOL_TIMEOUT_MS", "1004"),
        ("OFFICE_MCP_REQUESTS_PER_MINUTE", "9993"),
        ("OFFICE_MCP_LIMITS__REQUESTS_PER_MINUTE", "1005"),
        ("OFFICE_MCP_AUDIT_ENABLED", "false"),
        ("OFFICE_MCP_AUDIT__ENABLED", "true"),
        ("OFFICE_MCP_AUDIT_PATH", "legacy-audit.jsonl"),
        ("OFFICE_MCP_AUDIT__PATH", "section-audit.jsonl"),
        ("OFFICE_MCP_LOG_LEVEL", "error"),
        ("OFFICE_MCP_LOGGING__LEVEL", "trace"),
        ("OFFICE_MCP_LOG_FILE", "legacy.log"),
        ("OFFICE_MCP_LOGGING__FILE", "section.log"),
    ])))
    .load(LoadConfigOptions {
        config_path: Some(PathBuf::from("missing-config.toml")),
    })
    .expect("load config");

    assert_eq!(loaded.addin.host, "localhost");
    assert_eq!(loaded.addin.port, 8766);
    assert_eq!(loaded.addin.pfx_path, "section.pfx");
    assert_eq!(loaded.addin.pfx_passphrase, "section-pass");
    assert_eq!(loaded.addin.heartbeat_interval_sec, 31);
    assert_eq!(loaded.addin.heartbeat_timeout_sec, 12);
    assert_eq!(loaded.addin.session_grace_sec, 61);
    assert_eq!(loaded.addin.max_pending_per_session, 5);
    assert_eq!(loaded.mcp.host, "127.0.0.1");
    assert_eq!(loaded.mcp.port, 8801);
    assert_eq!(loaded.limits.max_response_bytes, 1001);
    assert_eq!(loaded.limits.max_request_bytes, 1002);
    assert_eq!(loaded.limits.max_ws_frame_bytes, 1003);
    assert_eq!(loaded.limits.default_tool_timeout_ms, 1004);
    assert_eq!(loaded.limits.requests_per_minute, 1005);
    assert!(loaded.audit.enabled);
    assert_eq!(loaded.audit.path, "section-audit.jsonl");
    assert_eq!(loaded.logging.level, LogLevel::Trace);
    assert_eq!(loaded.logging.file, "section.log");
}

#[test]
fn refuses_non_loopback_binds() {
    let mut config = DaemonConfigService::with_env(BTreeMap::new())
        .load(LoadConfigOptions::default())
        .expect("load defaults");
    config.addin.host = "0.0.0.0".to_string();
    let error =
        DaemonConfigService::assert_boundary_auth_config(&config).expect_err("reject addin");
    assert!(matches!(
        error,
        ConfigError::Validation(message) if message.contains("loopback-only")
    ));

    config.addin.host = "localhost".to_string();
    config.mcp.host = "0.0.0.0".to_string();
    let error = DaemonConfigService::assert_boundary_auth_config(&config).expect_err("reject mcp");
    assert!(matches!(
        error,
        ConfigError::Validation(message) if message.contains("loopback-only")
    ));
}

#[test]
fn redacts_secrets_without_hiding_public_endpoint_fields() {
    let config = DaemonConfigService::with_env(BTreeMap::new())
        .load(LoadConfigOptions::default())
        .expect("load defaults");
    let redacted = DaemonConfigService::redacted(&config);

    assert_eq!(redacted.addin.host, config.addin.host);
    assert_eq!(redacted.addin.pfx_passphrase, "<redacted>");
}

fn temp_config(contents: &str) -> PathBuf {
    let path = env::temp_dir().join(format!(
        "office-mcp-config-{}-{}.toml",
        std::process::id(),
        unique_suffix()
    ));
    fs::write(&path, contents).expect("write temp config");
    path
}

fn unique_suffix() -> u128 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("time")
        .as_nanos()
}

fn env_map(values: HashMap<&str, &str>) -> BTreeMap<String, String> {
    values
        .into_iter()
        .map(|(key, value)| (key.to_string(), value.to_string()))
        .collect()
}
