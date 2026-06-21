use crate::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, LimitsConfig, LoggingConfig, McpConfig,
    ToolAccessConfig,
};
use crate::ui::UiRuntimeFile;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiFixtureOptions {
    pub runtime_path: PathBuf,
    pub certificate_path: PathBuf,
    pub certificate_passphrase: String,
    pub state: UiFixtureState,
}

impl UiFixtureOptions {
    #[must_use]
    pub fn from_env() -> Self {
        Self {
            runtime_path: std::env::var_os("OFFICE_MCP_UI_RUNTIME_PATH")
                .map_or_else(UiRuntimeFile::default_path, PathBuf::from),
            certificate_path: std::env::var_os("OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH")
                .or_else(|| std::env::var_os("OFFICE_MCP_ADDIN_PFX_PATH"))
                .map_or_else(default_fixture_pfx_path, PathBuf::from),
            certificate_passphrase: std::env::var(
                "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE",
            )
            .or_else(|_| std::env::var("OFFICE_MCP_ADDIN_PFX_PASSPHRASE"))
            .unwrap_or_else(|_| "office-mcp-localhost".to_string()),
            state: UiFixtureState::from_env(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum UiFixtureState {
    #[default]
    Seeded,
    Empty,
}

impl UiFixtureState {
    fn from_env() -> Self {
        Self::from_value(std::env::var("OFFICE_MCP_UI_FIXTURE_STATE").ok().as_deref())
    }

    pub(crate) fn from_value(value: Option<&str>) -> Self {
        match value {
            Some(value) if value.eq_ignore_ascii_case("empty") => Self::Empty,
            _ => Self::Seeded,
        }
    }
}

pub(crate) fn fixture_config(
    mcp_port: u16,
    addin_port: u16,
    certificate_path: &Path,
    certificate_passphrase: String,
) -> DaemonConfig {
    DaemonConfig {
        config_path: default_fixture_config_path().display().to_string(),
        addin: AddinConfig {
            host: "127.0.0.1".to_string(),
            port: u64::from(addin_port),
            origin: format!("https://localhost:{addin_port}"),
            pfx_path: certificate_path.display().to_string(),
            pfx_passphrase: certificate_passphrase,
            heartbeat_interval_sec: 30,
            heartbeat_timeout_sec: 10,
            session_grace_sec: 60,
            max_pending_per_session: 4,
        },
        mcp: McpConfig {
            host: "127.0.0.1".to_string(),
            port: u64::from(mcp_port),
        },
        limits: LimitsConfig {
            max_response_bytes: 1024 * 1024,
            max_request_bytes: 16 * 1024 * 1024,
            max_ws_frame_bytes: 16 * 1024 * 1024,
            default_tool_timeout_ms: 30_000,
            requests_per_minute: 1000,
        },
        audit: AuditConfig {
            enabled: false,
            path: "audit.jsonl".to_string(),
        },
        logging: LoggingConfig {
            level: ConfigLogLevel::Error,
            file: String::new(),
        },
        tool_access: ToolAccessConfig::default(),
    }
}

fn default_fixture_config_path() -> PathBuf {
    std::env::temp_dir().join("office-mcp").join("config.toml")
}

pub(crate) fn default_addin_public_dir() -> PathBuf {
    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let candidate = ancestor
            .join("src")
            .join("office-ctl")
            .join("word")
            .join("public");
        if candidate.is_dir() {
            return candidate;
        }
    }
    PathBuf::from("src/office-ctl/word/public")
}

fn default_fixture_pfx_path() -> PathBuf {
    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let local_candidate = ancestor.join(".office-mcp-localhost.pfx");
        if local_candidate.is_file() {
            return local_candidate;
        }
    }
    PathBuf::from(".office-mcp-localhost.pfx")
}

#[cfg(test)]
#[path = "evidence_fixture_config_tests.rs"]
mod evidence_fixture_config_tests;
