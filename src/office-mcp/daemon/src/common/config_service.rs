use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::fs;
use std::net::IpAddr;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DaemonConfigService {
    description: &'static str,
    env: BTreeMap<String, String>,
}

impl DaemonConfigService {
    #[must_use]
    pub fn new() -> Self {
        Self::with_env(env::vars().collect())
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        self.description
    }

    #[must_use]
    pub fn with_env(env: BTreeMap<String, String>) -> Self {
        Self {
            description: "loads validates redacts and watches daemon configuration",
            env,
        }
    }

    /// Loads daemon configuration from file defaults plus environment overrides.
    ///
    /// # Errors
    ///
    /// Returns an error when the config file cannot be read, the supported TOML
    /// subset cannot be parsed, an environment override has the wrong type, or
    /// a typed config value is invalid.
    #[allow(clippy::too_many_lines)]
    pub fn load(&self, options: LoadConfigOptions) -> Result<DaemonConfig, ConfigError> {
        let config_path = options
            .config_path
            .or_else(|| self.env.get("OFFICE_MCP_CONFIG_PATH").map(PathBuf::from))
            .unwrap_or_else(|| self.default_config_path());
        let file_config = Self::load_config_file(&config_path)?;
        let addin_channel = file_config.section("addin_channel");
        let mcp_http = file_config.section("mcp_http");
        let limits = file_config.section("limits");
        let audit = file_config.section("audit");
        let logging = file_config.section("logging");

        let addin_host = self.string_env_any(
            &["OFFICE_MCP_ADDIN_CHANNEL__BIND", "OFFICE_MCP_ADDIN_HOST"],
            addin_channel.string_value("bind", "localhost")?,
        );
        let addin_port = self.int_env_any(
            &["OFFICE_MCP_ADDIN_CHANNEL__PORT", "OFFICE_MCP_ADDIN_PORT"],
            addin_channel.int_value("port", 8765)?,
        )?;

        Ok(DaemonConfig {
            addin: AddinConfig {
                host: addin_host.clone(),
                port: addin_port,
                origin: self.string_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__ORIGIN",
                        "OFFICE_MCP_ADDIN_ORIGIN",
                    ],
                    format!("https://{addin_host}:{addin_port}"),
                ),
                pfx_path: self.string_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH",
                        "OFFICE_MCP_ADDIN_PFX_PATH",
                    ],
                    addin_channel.string_value(
                        "certificate_path",
                        &Self::default_pfx_path().display().to_string(),
                    )?,
                ),
                pfx_passphrase: self.string_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE",
                        "OFFICE_MCP_ADDIN_PFX_PASSPHRASE",
                    ],
                    addin_channel.string_value("certificate_passphrase", "office-mcp-localhost")?,
                ),
                heartbeat_interval_sec: self.int_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_INTERVAL_SEC",
                        "OFFICE_MCP_ADDIN_HEARTBEAT_INTERVAL_SEC",
                    ],
                    addin_channel.int_value("heartbeat_interval_sec", 30)?,
                )?,
                heartbeat_timeout_sec: self.int_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_TIMEOUT_SEC",
                        "OFFICE_MCP_ADDIN_HEARTBEAT_TIMEOUT_SEC",
                    ],
                    addin_channel.int_value("heartbeat_timeout_sec", 10)?,
                )?,
                session_grace_sec: self.int_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__SESSION_GRACE_SEC",
                        "OFFICE_MCP_ADDIN_SESSION_GRACE_SEC",
                    ],
                    addin_channel.int_value("session_grace_sec", 60)?,
                )?,
                max_pending_per_session: self.int_env_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__MAX_PENDING_PER_SESSION",
                        "OFFICE_MCP_ADDIN_MAX_PENDING_PER_SESSION",
                    ],
                    addin_channel.int_value("max_pending_per_session", 4)?,
                )?,
            },
            mcp: McpConfig {
                host: self.string_env_any(
                    &["OFFICE_MCP_MCP_HTTP__BIND", "OFFICE_MCP_MCP_HOST"],
                    mcp_http.string_value("bind", "127.0.0.1")?,
                ),
                port: self.int_env_any(
                    &["OFFICE_MCP_MCP_HTTP__PORT", "OFFICE_MCP_MCP_PORT"],
                    mcp_http.int_value("port", 8800)?,
                )?,
            },
            limits: LimitsConfig {
                max_response_bytes: self.int_env_any(
                    &[
                        "OFFICE_MCP_LIMITS__MAX_RESPONSE_BYTES",
                        "OFFICE_MCP_MAX_RESPONSE_BYTES",
                    ],
                    limits.int_value("max_response_bytes", 1024 * 1024)?,
                )?,
                max_request_bytes: self.int_env_any(
                    &[
                        "OFFICE_MCP_LIMITS__MAX_REQUEST_BYTES",
                        "OFFICE_MCP_MAX_REQUEST_BYTES",
                    ],
                    limits.int_value("max_request_bytes", 16 * 1024 * 1024)?,
                )?,
                max_ws_frame_bytes: self.int_env_any(
                    &[
                        "OFFICE_MCP_LIMITS__MAX_WS_FRAME_BYTES",
                        "OFFICE_MCP_MAX_WS_FRAME_BYTES",
                    ],
                    limits.int_value("max_ws_frame_bytes", 16 * 1024 * 1024)?,
                )?,
                default_tool_timeout_ms: self.int_env_any(
                    &[
                        "OFFICE_MCP_LIMITS__DEFAULT_TOOL_TIMEOUT_MS",
                        "OFFICE_MCP_DEFAULT_TOOL_TIMEOUT_MS",
                    ],
                    limits.int_value("default_tool_timeout_ms", 30_000)?,
                )?,
                requests_per_minute: self.int_env_any(
                    &[
                        "OFFICE_MCP_LIMITS__REQUESTS_PER_MINUTE",
                        "OFFICE_MCP_REQUESTS_PER_MINUTE",
                    ],
                    limits.int_value("requests_per_minute", 120)?,
                )?,
            },
            audit: AuditConfig {
                enabled: self.bool_env_any(
                    &["OFFICE_MCP_AUDIT__ENABLED", "OFFICE_MCP_AUDIT_ENABLED"],
                    audit.bool_value("enabled", false)?,
                )?,
                path: self.string_env_any(
                    &["OFFICE_MCP_AUDIT__PATH", "OFFICE_MCP_AUDIT_PATH"],
                    optional_path_value(
                        audit.string_value("path", &self.default_audit_path())?,
                        &self.default_audit_path(),
                    ),
                ),
            },
            logging: LoggingConfig {
                level: LogLevel::parse(&self.string_env_any(
                    &["OFFICE_MCP_LOGGING__LEVEL", "OFFICE_MCP_LOG_LEVEL"],
                    logging.string_value("level", "info")?,
                ))?,
                file: self.string_env_any(
                    &["OFFICE_MCP_LOGGING__FILE", "OFFICE_MCP_LOG_FILE"],
                    optional_path_value(
                        logging.string_value("file", &self.default_log_path())?,
                        &self.default_log_path(),
                    ),
                ),
            },
        })
    }

    /// Validates that v1 listeners stay on loopback.
    ///
    /// # Errors
    ///
    /// Returns an error when either listener is exposed outside loopback.
    pub fn assert_boundary_auth_config(config: &DaemonConfig) -> Result<(), ConfigError> {
        if !is_loopback_host(&config.addin.host) {
            return Err(ConfigError::Validation(
                "Refusing to bind add-in WSS to a non-loopback address; v1 is loopback-only."
                    .to_string(),
            ));
        }
        if !is_loopback_host(&config.mcp.host) {
            return Err(ConfigError::Validation(
                "Refusing to bind MCP HTTP to a non-loopback address; v1 is loopback-only."
                    .to_string(),
            ));
        }
        Ok(())
    }

    #[must_use]
    pub fn redacted(config: &DaemonConfig) -> RedactedDaemonConfig {
        RedactedDaemonConfig {
            addin: RedactedAddinConfig {
                host: config.addin.host.clone(),
                port: config.addin.port,
                origin: config.addin.origin.clone(),
                pfx_path: config.addin.pfx_path.clone(),
                pfx_passphrase: "<redacted>".to_string(),
            },
            mcp: RedactedMcpConfig {
                host: config.mcp.host.clone(),
                port: config.mcp.port,
            },
        }
    }

    fn load_config_file(path: &Path) -> Result<RawToml, ConfigError> {
        if !path.exists() {
            return Ok(RawToml::default());
        }
        parse_toml(&fs::read_to_string(path).map_err(ConfigError::Io)?)
    }

    fn string_env_any(&self, names: &[&str], fallback: String) -> String {
        names
            .iter()
            .find_map(|name| self.env.get(*name).cloned())
            .unwrap_or(fallback)
    }

    fn int_env_any(&self, names: &[&str], fallback: u64) -> Result<u64, ConfigError> {
        let Some((name, raw)) = names
            .iter()
            .find_map(|name| self.env.get(*name).map(|raw| (*name, raw)))
        else {
            return Ok(fallback);
        };
        raw.parse::<u64>()
            .ok()
            .filter(|value| *value > 0)
            .ok_or_else(|| ConfigError::Validation(format!("{name} must be a positive integer")))
    }

    fn bool_env_any(&self, names: &[&str], fallback: bool) -> Result<bool, ConfigError> {
        let Some((name, raw)) = names
            .iter()
            .find_map(|name| self.env.get(*name).map(|raw| (*name, raw)))
        else {
            return Ok(fallback);
        };
        match raw.as_str() {
            "true" => Ok(true),
            "false" => Ok(false),
            _ => Err(ConfigError::Validation(format!(
                "{name} must be true or false"
            ))),
        }
    }

    fn default_config_path(&self) -> PathBuf {
        if cfg!(windows) {
            return PathBuf::from(
                self.env
                    .get("APPDATA")
                    .cloned()
                    .or_else(|| {
                        self.env
                            .get("USERPROFILE")
                            .map(|path| format!("{path}\\AppData\\Roaming"))
                    })
                    .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Roaming".to_string()),
            )
            .join("office-mcp")
            .join("config.toml");
        }
        if cfg!(target_os = "macos") {
            return PathBuf::from(
                self.env
                    .get("HOME")
                    .cloned()
                    .unwrap_or_else(|| ".".to_string()),
            )
            .join("Library")
            .join("Application Support")
            .join("office-mcp")
            .join("config.toml");
        }
        PathBuf::from(self.env.get("XDG_CONFIG_HOME").cloned().unwrap_or_else(|| {
            PathBuf::from(
                self.env
                    .get("HOME")
                    .cloned()
                    .unwrap_or_else(|| ".".to_string()),
            )
            .join(".config")
            .display()
            .to_string()
        }))
        .join("office-mcp")
        .join("config.toml")
    }

    fn default_pfx_path() -> PathBuf {
        env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(".office-mcp-localhost.pfx")
    }

    fn default_audit_path(&self) -> String {
        platform_state_path(&self.env, "audit.jsonl")
    }

    fn default_log_path(&self) -> String {
        platform_state_path(&self.env, "office-mcp.log")
    }
}

impl Default for DaemonConfigService {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct LoadConfigOptions {
    pub config_path: Option<PathBuf>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DaemonConfig {
    pub addin: AddinConfig,
    pub mcp: McpConfig,
    pub limits: LimitsConfig,
    pub audit: AuditConfig,
    pub logging: LoggingConfig,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LogLevel {
    Trace,
    Debug,
    Info,
    Warn,
    Error,
}

impl LogLevel {
    fn parse(value: &str) -> Result<Self, ConfigError> {
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

#[derive(Debug)]
pub enum ConfigError {
    Io(std::io::Error),
    Parse(String),
    Validation(String),
}

impl Display for ConfigError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::Parse(message) | Self::Validation(message) => formatter.write_str(message),
        }
    }
}

impl Error for ConfigError {}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RawToml(BTreeMap<String, BTreeMap<String, RawTomlValue>>);

impl RawToml {
    fn section(&self, name: &str) -> RawTomlSection<'_> {
        RawTomlSection(self.0.get(name))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RawTomlValue {
    String(String),
    Integer(u64),
    Boolean(bool),
}

struct RawTomlSection<'a>(Option<&'a BTreeMap<String, RawTomlValue>>);

impl RawTomlSection<'_> {
    fn string_value(&self, key: &str, fallback: &str) -> Result<String, ConfigError> {
        let Some(value) = self.0.and_then(|section| section.get(key)) else {
            return Ok(fallback.to_string());
        };
        match value {
            RawTomlValue::String(value) => Ok(value.clone()),
            other => Err(ConfigError::Parse(format!(
                "Expected string config value for {key}, got {}.",
                other.kind()
            ))),
        }
    }

    fn int_value(&self, key: &str, fallback: u64) -> Result<u64, ConfigError> {
        let Some(value) = self.0.and_then(|section| section.get(key)) else {
            return Ok(fallback);
        };
        match value {
            RawTomlValue::Integer(value) if *value > 0 => Ok(*value),
            _ => Err(ConfigError::Parse(format!(
                "Expected positive integer config value for {key}."
            ))),
        }
    }

    fn bool_value(&self, key: &str, fallback: bool) -> Result<bool, ConfigError> {
        let Some(value) = self.0.and_then(|section| section.get(key)) else {
            return Ok(fallback);
        };
        match value {
            RawTomlValue::Boolean(value) => Ok(*value),
            other => Err(ConfigError::Parse(format!(
                "Expected boolean config value for {key}, got {}.",
                other.kind()
            ))),
        }
    }
}

impl RawTomlValue {
    const fn kind(&self) -> &'static str {
        match self {
            Self::String(_) => "string",
            Self::Integer(_) => "integer",
            Self::Boolean(_) => "boolean",
        }
    }
}

/// Parses the small TOML subset accepted by the daemon config file.
///
/// # Errors
///
/// Returns an error when the input contains unsupported syntax or unsupported
/// scalar values.
pub fn parse_toml(input: &str) -> Result<RawToml, ConfigError> {
    let mut result = BTreeMap::<String, BTreeMap<String, RawTomlValue>>::new();
    let mut current_section = String::new();
    for (index, raw_line) in input.lines().enumerate() {
        let line = strip_toml_comment(raw_line).trim().to_string();
        if line.is_empty() {
            continue;
        }
        if let Some(section) = parse_section(&line) {
            current_section = section;
            result.entry(current_section.clone()).or_default();
            continue;
        }
        let Some((key, value)) = line.split_once('=') else {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML syntax at line {}: {raw_line}",
                index + 1
            )));
        };
        if current_section.is_empty() {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML syntax at line {}: {raw_line}",
                index + 1
            )));
        }
        let key = key.trim();
        if !is_identifier(key) {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML syntax at line {}: {raw_line}",
                index + 1
            )));
        }
        let parsed_value = parse_toml_value(value.trim(), index + 1)?;
        result
            .entry(current_section.clone())
            .or_default()
            .insert(key.to_string(), parsed_value);
    }
    Ok(RawToml(result))
}

fn parse_section(line: &str) -> Option<String> {
    let inner = line.strip_prefix('[')?.strip_suffix(']')?;
    is_identifier(inner).then(|| inner.to_string())
}

fn parse_toml_value(raw: &str, line_number: usize) -> Result<RawTomlValue, ConfigError> {
    if raw.starts_with('"') && raw.ends_with('"') {
        return Ok(RawTomlValue::String(unescape_toml_string(
            raw,
            line_number,
        )?));
    }
    match raw {
        "true" => return Ok(RawTomlValue::Boolean(true)),
        "false" => return Ok(RawTomlValue::Boolean(false)),
        _ => {}
    }
    if raw.chars().all(|char| char.is_ascii_digit()) {
        return raw.parse::<u64>().map(RawTomlValue::Integer).map_err(|_| {
            ConfigError::Parse(format!(
                "Unsupported TOML value at line {line_number}: {raw}"
            ))
        });
    }
    Err(ConfigError::Parse(format!(
        "Unsupported TOML value at line {line_number}: {raw}"
    )))
}

fn unescape_toml_string(raw: &str, line_number: usize) -> Result<String, ConfigError> {
    let mut result = String::new();
    let mut chars = raw[1..raw.len() - 1].chars();
    while let Some(char) = chars.next() {
        if char != '\\' {
            result.push(char);
            continue;
        }
        let Some(escaped) = chars.next() else {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML value at line {line_number}: {raw}"
            )));
        };
        match escaped {
            '"' => result.push('"'),
            '\\' => result.push('\\'),
            'n' => result.push('\n'),
            'r' => result.push('\r'),
            't' => result.push('\t'),
            other => {
                result.push('\\');
                result.push(other);
            }
        }
    }
    Ok(result)
}

fn strip_toml_comment(line: &str) -> String {
    let mut in_string = false;
    let mut escaped = false;
    for (index, char) in line.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        if char == '\\' && in_string {
            escaped = true;
            continue;
        }
        if char == '"' {
            in_string = !in_string;
        }
        if char == '#' && !in_string {
            return line[..index].to_string();
        }
    }
    line.to_string()
}

fn is_identifier(value: &str) -> bool {
    !value.is_empty()
        && value
            .chars()
            .all(|char| char.is_ascii_alphanumeric() || char == '_')
}

fn optional_path_value(value: String, fallback: &str) -> String {
    if value.is_empty() {
        fallback.to_string()
    } else {
        value
    }
}

fn platform_state_path(env: &BTreeMap<String, String>, filename: &str) -> String {
    if cfg!(windows) {
        return PathBuf::from(
            env.get("LOCALAPPDATA")
                .cloned()
                .or_else(|| {
                    env.get("USERPROFILE")
                        .map(|path| format!("{path}\\AppData\\Local"))
                })
                .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Local".to_string()),
        )
        .join("office-mcp")
        .join(filename)
        .display()
        .to_string();
    }
    if cfg!(target_os = "macos") {
        return PathBuf::from(env.get("HOME").cloned().unwrap_or_else(|| ".".to_string()))
            .join("Library")
            .join("Logs")
            .join("office-mcp")
            .join(filename)
            .display()
            .to_string();
    }
    PathBuf::from(env.get("XDG_STATE_HOME").cloned().unwrap_or_else(|| {
        PathBuf::from(env.get("HOME").cloned().unwrap_or_else(|| ".".to_string()))
            .join(".local")
            .join("state")
            .display()
            .to_string()
    }))
    .join("office-mcp")
    .join(filename)
    .display()
    .to_string()
}

fn is_loopback_host(host: &str) -> bool {
    let normalized = host.to_ascii_lowercase();
    if normalized == "localhost" || normalized == "::1" || normalized == "[::1]" {
        return true;
    }
    normalized
        .parse::<IpAddr>()
        .is_ok_and(|address| address.is_loopback())
}

#[cfg(test)]
#[path = "config_service_tests.rs"]
mod config_service_tests;
