use crate::addin_mgr::{
    AddinChannelConfig, ImageFetcher, WebSocketCodecError, WebSocketProtocolError,
};
use crate::common::{AuditLog, DaemonConfig};
use crate::mcp::McpHttpConfig;
use crate::ui::UiRuntimeError;
use native_tls::{Identity, TlsAcceptor};
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::fs;
use std::path::PathBuf;
use std::time::Duration;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeServerConfig {
    pub mcp_host: String,
    pub mcp_port: u16,
    pub addin_host: String,
    pub addin_port: u16,
    pub addin_origin: String,
    pub addin_public_dir: PathBuf,
    pub certificate_path: PathBuf,
    pub certificate_passphrase: String,
    pub max_request_bytes: usize,
    pub max_ws_frame_bytes: usize,
    pub max_pending_per_session: usize,
    pub heartbeat_interval: Duration,
    pub heartbeat_timeout: Duration,
    pub requests_per_minute: u64,
    pub config_path: Option<String>,
    pub log_path: Option<String>,
    pub audit_log: AuditLog,
    pub image_fetcher: ImageFetcher,
}

impl RuntimeServerConfig {
    /// Builds a runtime server config from validated daemon configuration.
    ///
    /// # Errors
    ///
    /// Returns an error when a configured port or byte limit does not fit the
    /// native runtime type used by the socket server.
    pub fn from_daemon_config(config: &DaemonConfig) -> Result<Self, RuntimeServerError> {
        Ok(Self {
            mcp_host: config.mcp.host.clone(),
            mcp_port: to_u16("mcp.port", config.mcp.port)?,
            addin_host: config.addin.host.clone(),
            addin_port: to_u16("addin.port", config.addin.port)?,
            addin_origin: config.addin.origin.clone(),
            addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
            certificate_path: PathBuf::from(&config.addin.pfx_path),
            certificate_passphrase: config.addin.pfx_passphrase.clone(),
            max_request_bytes: to_usize(
                "limits.max_request_bytes",
                config.limits.max_request_bytes,
            )?,
            max_ws_frame_bytes: to_usize(
                "limits.max_ws_frame_bytes",
                config.limits.max_ws_frame_bytes,
            )?,
            max_pending_per_session: to_usize(
                "addin.max_pending_per_session",
                config.addin.max_pending_per_session,
            )?,
            heartbeat_interval: Duration::from_secs(config.addin.heartbeat_interval_sec),
            heartbeat_timeout: Duration::from_secs(config.addin.heartbeat_timeout_sec),
            requests_per_minute: config.limits.requests_per_minute,
            config_path: None,
            log_path: Some(config.logging.file.clone()),
            audit_log: if config.audit.enabled {
                AuditLog::enabled(&config.audit.path)
            } else {
                AuditLog::new()
            },
            image_fetcher: ImageFetcher::new(),
        })
    }

    #[must_use]
    pub(crate) fn mcp_bind_addr(&self) -> String {
        format!("{}:{}", self.mcp_host, self.mcp_port)
    }

    #[must_use]
    pub(crate) fn addin_bind_addr(&self) -> String {
        format!("{}:{}", self.addin_host, self.addin_port)
    }

    #[must_use]
    pub(crate) fn mcp_http_config(&self) -> McpHttpConfig {
        McpHttpConfig {
            host: self.mcp_host.clone(),
            port: self.mcp_port,
            max_request_bytes: self.max_request_bytes,
            requests_per_minute: self.requests_per_minute,
        }
    }

    /// Builds the TLS acceptor used by the add-in listener.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured certificate cannot be read or loaded.
    pub(crate) fn tls_acceptor(&self) -> Result<TlsAcceptor, RuntimeServerError> {
        let pfx = fs::read(&self.certificate_path).map_err(|error| {
            RuntimeServerError::Tls(format!(
                "Failed to read add-in HTTPS certificate {}: {error}",
                self.certificate_path.display()
            ))
        })?;
        let identity = Identity::from_pkcs12(&pfx, &self.certificate_passphrase)
            .map_err(|error| RuntimeServerError::Tls(error.to_string()))?;
        TlsAcceptor::new(identity).map_err(|error| RuntimeServerError::Tls(error.to_string()))
    }

    #[must_use]
    pub(crate) fn addin_channel_config(&self) -> AddinChannelConfig {
        AddinChannelConfig {
            origin: self.addin_origin.clone(),
            heartbeat_interval: self.heartbeat_interval,
            heartbeat_timeout: self.heartbeat_timeout,
            max_pending_per_session: self.max_pending_per_session,
            ..AddinChannelConfig::default()
        }
    }
}

impl Default for RuntimeServerConfig {
    fn default() -> Self {
        Self {
            mcp_host: "127.0.0.1".to_string(),
            mcp_port: 8800,
            addin_host: "localhost".to_string(),
            addin_port: 8765,
            addin_origin: "https://localhost:8765".to_string(),
            addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
            certificate_path: default_pfx_path(),
            certificate_passphrase: "office-mcp-localhost".to_string(),
            max_request_bytes: 16 * 1024 * 1024,
            max_ws_frame_bytes: 16 * 1024 * 1024,
            max_pending_per_session: 4,
            heartbeat_interval: Duration::from_secs(30),
            heartbeat_timeout: Duration::from_secs(10),
            requests_per_minute: 120,
            config_path: None,
            log_path: None,
            audit_log: AuditLog::new(),
            image_fetcher: ImageFetcher::new(),
        }
    }
}

#[derive(Debug)]
pub enum RuntimeServerError {
    Io(std::io::Error),
    Tls(String),
    InvalidConfig(String),
    BadRequest(String),
    WebSocketProtocol(WebSocketProtocolError),
    Internal(String),
}

impl Display for RuntimeServerError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::Tls(message) => formatter.write_str(message),
            Self::InvalidConfig(message) | Self::BadRequest(message) | Self::Internal(message) => {
                formatter.write_str(message)
            }
            Self::WebSocketProtocol(error) => formatter.write_str(&error.reason),
        }
    }
}

impl Error for RuntimeServerError {}

impl From<std::io::Error> for RuntimeServerError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<UiRuntimeError> for RuntimeServerError {
    fn from(error: UiRuntimeError) -> Self {
        Self::Internal(error.to_string())
    }
}

impl From<WebSocketCodecError> for RuntimeServerError {
    fn from(error: WebSocketCodecError) -> Self {
        match error {
            WebSocketCodecError::Io(error) => Self::Io(error),
            WebSocketCodecError::Protocol(error) => Self::WebSocketProtocol(error),
        }
    }
}

#[must_use]
pub fn default_pfx_path() -> PathBuf {
    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let local_candidate = ancestor.join(".office-mcp-localhost.pfx");
        if local_candidate.is_file() {
            return local_candidate;
        }
    }
    current.join(".office-mcp-localhost.pfx")
}

fn to_u16(name: &str, value: u64) -> Result<u16, RuntimeServerError> {
    value
        .try_into()
        .map_err(|_| RuntimeServerError::InvalidConfig(format!("{name} must fit into a TCP port.")))
}

fn to_usize(name: &str, value: u64) -> Result<usize, RuntimeServerError> {
    value.try_into().map_err(|_| {
        RuntimeServerError::InvalidConfig(format!("{name} is too large for this platform."))
    })
}

#[cfg(test)]
#[path = "server_config_tests.rs"]
mod server_config_tests;
