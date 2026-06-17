pub mod audit_log;
pub mod client_config;
pub mod config_display;
pub(crate) mod config_paths;
pub mod config_service;
pub mod config_toml;
pub mod logger;

pub use audit_log::{AuditLog, AuditLogError, AuditRecord};
pub use client_config::ClaudeDesktopConfigBuilder;
pub use config_display::{json_escape, render_endpoints, render_redacted_config};
pub use config_service::{
    AddinConfig, AuditConfig, ConfigError, DaemonConfig, DaemonConfigService, EndpointConfig,
    LimitsConfig, LoadConfigOptions, LogLevel as ConfigLogLevel, LoggingConfig, McpConfig,
};
pub use config_toml::{RawTomlValue, parse_toml};
pub use logger::{LogLevel as LoggerLogLevel, LogRecord, Logger, LoggerError, TracingLogGuard};
