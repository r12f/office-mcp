use crate::common::ConfigError;
use crate::mcp::AccessMode;
use std::path::PathBuf;

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct LoadConfigOptions {
    pub config_path: Option<PathBuf>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DaemonConfig {
    pub config_path: String,
    pub addin: AddinConfig,
    pub mcp: McpConfig,
    pub limits: LimitsConfig,
    pub audit: AuditConfig,
    pub logging: LoggingConfig,
    pub tool_access: ToolAccessConfig,
}

impl DaemonConfig {
    #[must_use]
    pub fn endpoints(&self) -> EndpointConfig {
        EndpointConfig {
            mcp: format!("http://{}:{}/mcp", self.mcp.host, self.mcp.port),
            addin_origin: self.addin.origin.clone(),
            addin_wss: format!(
                "{}/addin",
                self.addin.origin.replacen("https://", "wss://", 1)
            ),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EndpointConfig {
    pub mcp: String,
    pub addin_origin: String,
    pub addin_wss: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddinConfig {
    pub host: String,
    pub port: u64,
    pub origin: String,
    pub pfx_path: String,
    pub pfx_passphrase: String,
    pub heartbeat_interval_sec: u64,
    pub heartbeat_timeout_sec: u64,
    pub session_grace_sec: u64,
    pub max_pending_per_session: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpConfig {
    pub host: String,
    pub port: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LimitsConfig {
    pub max_response_bytes: u64,
    pub max_request_bytes: u64,
    pub max_ws_frame_bytes: u64,
    pub default_tool_timeout_ms: u64,
    pub requests_per_minute: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuditConfig {
    pub enabled: bool,
    pub path: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoggingConfig {
    pub level: LogLevel,
    pub file: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToolAccessConfig {
    pub access_mode: AccessMode,
    pub disabled_apps: Vec<String>,
    pub disabled_categories: Vec<(String, String)>,
    pub disabled_tools: Vec<String>,
}

impl Default for ToolAccessConfig {
    fn default() -> Self {
        Self {
            access_mode: AccessMode::All,
            disabled_apps: Vec::new(),
            disabled_categories: Vec::new(),
            disabled_tools: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LogLevel {
    Trace,
    Debug,
    Info,
    Warn,
    Error,
}

impl LogLevel {
    pub(crate) fn parse(value: &str) -> Result<Self, ConfigError> {
        match value {
            "trace" => Ok(Self::Trace),
            "debug" => Ok(Self::Debug),
            "info" => Ok(Self::Info),
            "warn" => Ok(Self::Warn),
            "error" => Ok(Self::Error),
            _ => Err(ConfigError::Validation(format!(
                "logging.level must be one of trace, debug, info, warn, error; got {value}."
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RedactedDaemonConfig {
    pub addin: RedactedAddinConfig,
    pub mcp: RedactedMcpConfig,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RedactedAddinConfig {
    pub host: String,
    pub port: u64,
    pub origin: String,
    pub pfx_path: String,
    pub pfx_passphrase: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RedactedMcpConfig {
    pub host: String,
    pub port: u64,
}
