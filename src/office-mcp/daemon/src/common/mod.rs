pub mod audit_log;
pub mod client_config;
pub mod config_display;
pub mod config_error;
pub mod config_model;
pub(crate) mod config_paths;
pub mod config_service;
pub mod config_toml;
pub mod logger;
pub mod logger_record;
pub(crate) mod logger_redaction;

pub use audit_log::{AuditLog, AuditLogError, AuditRecord};
pub use client_config::ClaudeDesktopConfigBuilder;
pub use config_display::{json_escape, render_endpoints, render_redacted_config};
pub use config_error::ConfigError;
pub use config_model::{
    AddinConfig, AuditConfig, DaemonConfig, EndpointConfig, LimitsConfig, LoadConfigOptions,
    LogLevel as ConfigLogLevel, LoggingConfig, McpConfig, RedactedAddinConfig,
    RedactedDaemonConfig, RedactedMcpConfig,
};
pub use config_service::DaemonConfigService;
pub use config_toml::{RawTomlValue, parse_toml};
pub use logger::{Logger, LoggerError, TracingLogGuard};
pub use logger_record::{LogLevel as LoggerLogLevel, LogRecord};
