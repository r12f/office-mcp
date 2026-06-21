use crate::common::config_env::ConfigEnv;
use crate::common::config_model::LogLevel;
use crate::common::config_paths::ConfigPathResolver;
use crate::common::config_toml::{RawToml, parse_toml};
use crate::common::{
    AddinConfig, AuditConfig, ConfigError, DaemonConfig, LimitsConfig, LoadConfigOptions,
    LoggingConfig, McpConfig, RedactedAddinConfig, RedactedDaemonConfig, RedactedMcpConfig,
    ToolAccessConfig,
};
use crate::mcp::AccessMode;
use std::collections::BTreeMap;
use std::env;
use std::fs;
use std::net::IpAddr;
use std::path::{Path, PathBuf};

const MAX_SESSION_GRACE_SEC: u64 = 300;

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
        let tool_access = file_config.section("tool_access");
        let config_env = ConfigEnv::new(&self.env);

        let addin_host = config_env.string_any(
            &["OFFICE_MCP_ADDIN_CHANNEL__BIND", "OFFICE_MCP_ADDIN_HOST"],
            addin_channel.string_value("bind", "localhost")?,
        );
        let addin_port = config_env.positive_int_any(
            &["OFFICE_MCP_ADDIN_CHANNEL__PORT", "OFFICE_MCP_ADDIN_PORT"],
            addin_channel.int_value("port", 8765)?,
        )?;

        Ok(DaemonConfig {
            config_path: config_path.display().to_string(),
            addin: AddinConfig {
                host: addin_host.clone(),
                port: addin_port,
                origin: config_env.string_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__ORIGIN",
                        "OFFICE_MCP_ADDIN_ORIGIN",
                    ],
                    format!("https://{addin_host}:{addin_port}"),
                ),
                pfx_path: config_env.string_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH",
                        "OFFICE_MCP_ADDIN_PFX_PATH",
                    ],
                    addin_channel.string_value(
                        "certificate_path",
                        &ConfigPathResolver::pfx_path().display().to_string(),
                    )?,
                ),
                pfx_passphrase: config_env.string_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE",
                        "OFFICE_MCP_ADDIN_PFX_PASSPHRASE",
                    ],
                    addin_channel.string_value("certificate_passphrase", "office-mcp-localhost")?,
                ),
                heartbeat_interval_sec: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_INTERVAL_SEC",
                        "OFFICE_MCP_ADDIN_HEARTBEAT_INTERVAL_SEC",
                    ],
                    addin_channel.int_value("heartbeat_interval_sec", 30)?,
                )?,
                heartbeat_timeout_sec: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_TIMEOUT_SEC",
                        "OFFICE_MCP_ADDIN_HEARTBEAT_TIMEOUT_SEC",
                    ],
                    addin_channel.int_value("heartbeat_timeout_sec", 10)?,
                )?,
                session_grace_sec: config_env
                    .positive_int_any(
                        &[
                            "OFFICE_MCP_ADDIN_CHANNEL__SESSION_GRACE_SEC",
                            "OFFICE_MCP_ADDIN_SESSION_GRACE_SEC",
                        ],
                        addin_channel.int_value("session_grace_sec", 60)?,
                    )?
                    .min(MAX_SESSION_GRACE_SEC),
                max_pending_per_session: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_ADDIN_CHANNEL__MAX_PENDING_PER_SESSION",
                        "OFFICE_MCP_ADDIN_MAX_PENDING_PER_SESSION",
                    ],
                    addin_channel.int_value("max_pending_per_session", 4)?,
                )?,
            },
            mcp: McpConfig {
                host: config_env.string_any(
                    &["OFFICE_MCP_MCP_HTTP__BIND", "OFFICE_MCP_MCP_HOST"],
                    mcp_http.string_value("bind", "127.0.0.1")?,
                ),
                port: config_env.positive_int_any(
                    &["OFFICE_MCP_MCP_HTTP__PORT", "OFFICE_MCP_MCP_PORT"],
                    mcp_http.int_value("port", 8800)?,
                )?,
            },
            limits: LimitsConfig {
                max_response_bytes: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_LIMITS__MAX_RESPONSE_BYTES",
                        "OFFICE_MCP_MAX_RESPONSE_BYTES",
                    ],
                    limits.int_value("max_response_bytes", 1024 * 1024)?,
                )?,
                max_request_bytes: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_LIMITS__MAX_REQUEST_BYTES",
                        "OFFICE_MCP_MAX_REQUEST_BYTES",
                    ],
                    limits.int_value("max_request_bytes", 16 * 1024 * 1024)?,
                )?,
                max_ws_frame_bytes: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_LIMITS__MAX_WS_FRAME_BYTES",
                        "OFFICE_MCP_MAX_WS_FRAME_BYTES",
                    ],
                    limits.int_value("max_ws_frame_bytes", 16 * 1024 * 1024)?,
                )?,
                default_tool_timeout_ms: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_LIMITS__DEFAULT_TOOL_TIMEOUT_MS",
                        "OFFICE_MCP_DEFAULT_TOOL_TIMEOUT_MS",
                    ],
                    limits.int_value("default_tool_timeout_ms", 30_000)?,
                )?,
                requests_per_minute: config_env.positive_int_any(
                    &[
                        "OFFICE_MCP_LIMITS__REQUESTS_PER_MINUTE",
                        "OFFICE_MCP_REQUESTS_PER_MINUTE",
                    ],
                    limits.int_value("requests_per_minute", 1000)?,
                )?,
            },
            audit: AuditConfig {
                enabled: config_env.bool_any(
                    &["OFFICE_MCP_AUDIT__ENABLED", "OFFICE_MCP_AUDIT_ENABLED"],
                    audit.bool_value("enabled", false)?,
                )?,
                path: config_env.string_any(
                    &["OFFICE_MCP_AUDIT__PATH", "OFFICE_MCP_AUDIT_PATH"],
                    ConfigEnv::optional_path_value(
                        audit.string_value("path", &self.default_audit_path())?,
                        &self.default_audit_path(),
                    ),
                ),
            },
            logging: LoggingConfig {
                level: LogLevel::parse(&config_env.string_any(
                    &["OFFICE_MCP_LOGGING__LEVEL", "OFFICE_MCP_LOG_LEVEL"],
                    logging.string_value("level", "info")?,
                ))?,
                file: config_env.string_any(
                    &["OFFICE_MCP_LOGGING__FILE", "OFFICE_MCP_LOG_FILE"],
                    ConfigEnv::optional_path_value(
                        logging.string_value("file", &self.default_log_path())?,
                        &self.default_log_path(),
                    ),
                ),
            },
            tool_access: ToolAccessConfig {
                access_mode: parse_access_mode(&config_env.string_any(
                    &["OFFICE_MCP_TOOL_ACCESS__ACCESS_MODE"],
                    tool_access.string_value("access_mode", "all")?,
                ))?,
                disabled_apps: parse_csv_list(&config_env.string_any(
                    &["OFFICE_MCP_TOOL_ACCESS__DISABLED_APPS"],
                    tool_access.string_value("disabled_apps", "")?,
                )),
                disabled_categories: parse_category_list(&config_env.string_any(
                    &["OFFICE_MCP_TOOL_ACCESS__DISABLED_CATEGORIES"],
                    tool_access.string_value("disabled_categories", "")?,
                ))?,
                disabled_tools: parse_csv_list(&config_env.string_any(
                    &["OFFICE_MCP_TOOL_ACCESS__DISABLED_TOOLS"],
                    tool_access.string_value("disabled_tools", "")?,
                )),
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

    /// Persists the daemon-wide tool access section in the config file.
    ///
    /// # Errors
    ///
    /// Returns an error when the config file cannot be read or written.
    pub fn save_tool_access_config(
        path: &Path,
        tool_access: &ToolAccessConfig,
    ) -> Result<(), ConfigError> {
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).map_err(ConfigError::Io)?;
        }
        let current = if path.exists() {
            fs::read_to_string(path).map_err(ConfigError::Io)?
        } else {
            String::new()
        };
        let updated =
            replace_tool_access_section(&current, &render_tool_access_section(tool_access));
        fs::write(path, updated).map_err(ConfigError::Io)
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

    fn default_config_path(&self) -> PathBuf {
        ConfigPathResolver::new(&self.env).config_path()
    }

    fn default_audit_path(&self) -> String {
        ConfigPathResolver::new(&self.env).audit_path()
    }

    fn default_log_path(&self) -> String {
        ConfigPathResolver::new(&self.env).log_path()
    }
}

impl Default for DaemonConfigService {
    fn default() -> Self {
        Self::new()
    }
}

fn replace_tool_access_section(current: &str, section: &str) -> String {
    let Some((start, end)) = tool_access_section_range(current) else {
        let mut output = current.trim_end().to_string();
        if !output.is_empty() {
            output.push_str("\n\n");
        }
        output.push_str(section);
        return output;
    };
    let mut output = String::new();
    output.push_str(current[..start].trim_end());
    if !output.is_empty() {
        output.push_str("\n\n");
    }
    output.push_str(section);
    let suffix = current[end..].trim_start_matches(['\r', '\n']);
    if !suffix.is_empty() {
        output.push_str("\n\n");
        output.push_str(suffix.trim_start());
    }
    output
}

fn tool_access_section_range(contents: &str) -> Option<(usize, usize)> {
    let mut start = None;
    for (index, line) in line_spans(contents) {
        if strip_inline_comment(line).trim() == "[tool_access]" {
            start = Some(index);
            break;
        }
    }
    let start = start?;
    let mut end = contents.len();
    for (index, line) in line_spans(&contents[start..]) {
        if index == 0 {
            continue;
        }
        let trimmed = strip_inline_comment(line).trim();
        if trimmed.starts_with('[') && trimmed.ends_with(']') {
            end = start + index;
            break;
        }
    }
    Some((start, end))
}

fn line_spans(contents: &str) -> impl Iterator<Item = (usize, &str)> {
    contents.split_inclusive('\n').scan(0, |offset, line| {
        let start = *offset;
        *offset += line.len();
        Some((start, line.trim_end_matches(['\r', '\n'])))
    })
}

fn strip_inline_comment(line: &str) -> &str {
    line.split_once('#').map_or(line, |(value, _)| value)
}

fn render_tool_access_section(tool_access: &ToolAccessConfig) -> String {
    format!(
        "[tool_access]\naccess_mode = \"{}\"\ndisabled_apps = \"{}\"\ndisabled_categories = \"{}\"\ndisabled_tools = \"{}\"\n",
        access_mode_config_value(tool_access.access_mode),
        comma_list(&tool_access.disabled_apps),
        category_list(&tool_access.disabled_categories),
        comma_list(&tool_access.disabled_tools),
    )
}

fn access_mode_config_value(value: AccessMode) -> &'static str {
    match value {
        AccessMode::Read => "read",
        AccessMode::Write => "write",
        AccessMode::All => "all",
    }
}

fn comma_list(values: &[String]) -> String {
    values.join(",")
}

fn category_list(values: &[(String, String)]) -> String {
    values
        .iter()
        .map(|(app, category)| format!("{app}:{category}"))
        .collect::<Vec<_>>()
        .join(";")
}

fn parse_access_mode(value: &str) -> Result<AccessMode, ConfigError> {
    match value.to_ascii_lowercase().as_str() {
        "read" => Ok(AccessMode::Read),
        "write" => Ok(AccessMode::Write),
        "all" => Ok(AccessMode::All),
        _ => Err(ConfigError::Validation(format!(
            "tool_access.access_mode must be one of read, write, all; got {value}."
        ))),
    }
}

fn parse_csv_list(value: &str) -> Vec<String> {
    value
        .split(',')
        .map(str::trim)
        .filter(|item| !item.is_empty())
        .map(ToString::to_string)
        .collect()
}

fn parse_category_list(value: &str) -> Result<Vec<(String, String)>, ConfigError> {
    value
        .split(';')
        .map(str::trim)
        .filter(|item| !item.is_empty())
        .map(|item| {
            let Some((app, category)) = item.split_once(':') else {
                return Err(ConfigError::Validation(format!(
                    "tool_access.disabled_categories entries must use app:category; got {item}."
                )));
            };
            let app = app.trim();
            let category = category.trim();
            if app.is_empty() || category.is_empty() {
                return Err(ConfigError::Validation(format!(
                    "tool_access.disabled_categories entries must use app:category; got {item}."
                )));
            }
            Ok((app.to_string(), category.to_string()))
        })
        .collect()
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
