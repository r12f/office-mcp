use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, PartialEffect, ProtectionInfo, RuntimeInfo,
    SessionRegistry,
};
use crate::common::AuditLog;
use crate::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, LimitsConfig, LoggingConfig, McpConfig,
};
use crate::image_fetcher::ImageFetcher;
use crate::runtime_server::{
    RuntimeSeedState, RuntimeServer, RuntimeServerConfig, RuntimeServerError,
};
use crate::ui::UiRuntimeFile;
use crate::ui::{
    CommandFailure, CommandResult, RegisterClientInput, StartCommandInput, UiClientTransport,
    UiHealth, UiStateStore,
};
use std::net::TcpListener;
use std::path::{Path, PathBuf};
use std::time::{Duration, SystemTime};

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
        match std::env::var("OFFICE_MCP_UI_FIXTURE_STATE") {
            Ok(value) if value.eq_ignore_ascii_case("empty") => Self::Empty,
            _ => Self::Seeded,
        }
    }
}

/// Runs a Rust-owned daemon UI evidence fixture until the process exits.
///
/// # Errors
///
/// Returns an error when loopback ports cannot be allocated, the runtime file
/// cannot be written, TLS cannot be loaded, or the fixture server fails.
pub fn run_ui_fixture(options: UiFixtureOptions) -> Result<(), RuntimeServerError> {
    let mcp_listener = TcpListener::bind("127.0.0.1:0")?;
    let addin_listener = TcpListener::bind("127.0.0.1:0")?;
    let mcp_port = mcp_listener.local_addr()?.port();
    let addin_port = addin_listener.local_addr()?.port();
    let config = fixture_config(
        mcp_port,
        addin_port,
        &options.certificate_path,
        options.certificate_passphrase,
    );
    let runtime_file = UiRuntimeFile::with_path(
        options.runtime_path,
        crate::ui::UiRuntimeInfo::from_config(&config),
    );
    runtime_file.write()?;
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        mcp_host: config.mcp.host,
        mcp_port: u16::try_from(config.mcp.port).unwrap_or(mcp_port),
        addin_host: config.addin.host,
        addin_port: u16::try_from(config.addin.port).unwrap_or(addin_port),
        addin_origin: config.addin.origin,
        addin_public_dir: default_addin_public_dir(),
        certificate_path: PathBuf::from(config.addin.pfx_path),
        certificate_passphrase: config.addin.pfx_passphrase,
        max_request_bytes: usize::try_from(config.limits.max_request_bytes)
            .unwrap_or(16 * 1024 * 1024),
        max_ws_frame_bytes: usize::try_from(config.limits.max_ws_frame_bytes)
            .unwrap_or(16 * 1024 * 1024),
        max_pending_per_session: usize::try_from(config.addin.max_pending_per_session).unwrap_or(4),
        heartbeat_interval: Duration::from_secs(config.addin.heartbeat_interval_sec),
        heartbeat_timeout: Duration::from_secs(config.addin.heartbeat_timeout_sec),
        requests_per_minute: config.limits.requests_per_minute,
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
    });
    let seed = match options.state {
        UiFixtureState::Seeded => seeded_state(SystemTime::now()),
        UiFixtureState::Empty => empty_state(),
    };
    let result = server.serve_bound_with_state_forever(
        &mcp_listener,
        &addin_listener,
        seed.ui_state,
        seed.registry,
    );
    if let Err(error) = runtime_file.remove() {
        eprintln!("{error}");
    }
    result
}

#[must_use]
pub fn seeded_state(now: SystemTime) -> RuntimeSeedState {
    let mut registry = SessionRegistry::new();
    registry.register_runtime(runtime("word-instance", "word", now));
    registry.add_session(word_session(), now);
    registry.register_runtime(runtime("excel-instance", "excel", now));
    registry.add_session(excel_session(), now);
    registry.mark_session_stale("33333333-3333-4333-8333-333333333333", now);

    let mut ui_state = UiStateStore::new();
    ui_state.set_health(
        UiHealth::Degraded,
        Some("Certificate reload failed. Check the configured local PFX path."),
    );
    let client_id = ui_state.register_client(RegisterClientInput {
        client_id: Some("client-1".to_string()),
        transport: UiClientTransport::Http,
        name: Some("copilot-cli/1.0 token=secret-value".to_string()),
    });
    ui_state.start_command(StartCommandInput {
        client_id: Some(client_id.clone()),
        client_name: Some("copilot-cli/1.0".to_string()),
        session_id: Some("11111111-1111-4111-8111-111111111111".to_string()),
        host_app: Some("word".to_string()),
        tool: "word.insert_paragraph".to_string(),
        user_intent: Some("running smoke task".to_string()),
        timeout_ms: Some(30_000),
        ..start_command_default()
    });
    for index in 0..12 {
        let command_id = ui_state.start_command(StartCommandInput {
            client_id: Some(client_id.clone()),
            client_name: Some("copilot-cli/1.0".to_string()),
            session_id: Some("11111111-1111-4111-8111-111111111111".to_string()),
            host_app: Some("word".to_string()),
            tool: if index % 2 == 0 {
                "word.get_text"
            } else {
                "word.insert_paragraph"
            }
            .to_string(),
            user_intent: Some(format!(
                "summarize status token=secret-value base64,QUJDREVGRw== {index}"
            )),
            timeout_ms: Some(30_000),
            started_at: Some(now + Duration::from_millis(index * 10)),
            ..start_command_default()
        });
        let result = match index % 4 {
            0 => CommandResult::Success,
            1 => CommandResult::Failure(CommandFailure {
                office_mcp_code: "IRM_DENIED".to_string(),
                message: "certificate_passphrase=secret-value blocked by document policy."
                    .to_string(),
                tool: Some("word.insert_paragraph".to_string()),
                retriable: false,
                partial_effect: Some(PartialEffect::None),
            }),
            2 => CommandResult::Failure(CommandFailure {
                office_mcp_code: "TIMEOUT".to_string(),
                message: "The add-in did not respond before the deadline.".to_string(),
                tool: Some("word.get_text".to_string()),
                retriable: true,
                partial_effect: Some(PartialEffect::Unknown),
            }),
            _ => CommandResult::Failure(CommandFailure {
                office_mcp_code: "CANCELLED".to_string(),
                message: "The client cancelled the command.".to_string(),
                tool: Some("word.insert_paragraph".to_string()),
                retriable: true,
                partial_effect: Some(PartialEffect::Unknown),
            }),
        };
        ui_state.finish_command(
            &command_id,
            result,
            now + Duration::from_millis(index * 10 + 5),
        );
    }
    RuntimeSeedState { ui_state, registry }
}

#[must_use]
pub fn empty_state() -> RuntimeSeedState {
    let registry = SessionRegistry::new();
    let mut ui_state = UiStateStore::new();
    ui_state.set_health(UiHealth::Up, None);
    RuntimeSeedState { ui_state, registry }
}

fn fixture_config(
    mcp_port: u16,
    addin_port: u16,
    certificate_path: &Path,
    certificate_passphrase: String,
) -> DaemonConfig {
    DaemonConfig {
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
            requests_per_minute: 120,
        },
        audit: AuditConfig {
            enabled: false,
            path: "audit.jsonl".to_string(),
        },
        logging: LoggingConfig {
            level: ConfigLogLevel::Error,
            file: String::new(),
        },
    }
}

fn runtime(instance_id: &str, app: &str, registered_at: SystemTime) -> RuntimeInfo {
    RuntimeInfo {
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: app.to_string(),
            version: Some("16.0".to_string()),
            platform: Some("pc".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.6".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["doc.read".to_string(), "doc.write".to_string()],
        },
        registered_at,
    }
}

fn word_session() -> NewSessionInfo {
    NewSessionInfo {
        session_id: "11111111-1111-4111-8111-111111111111".to_string(),
        instance_id: "word-instance".to_string(),
        document: DocumentInfo {
            title: Some("Runtime Evidence.docx".to_string()),
            filename: Some("Runtime Evidence.docx".to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(true),
            protection: Some(ProtectionInfo {
                kind: Some("irm".to_string()),
                rights: None,
                rights_source: Some("unavailable".to_string()),
            }),
            ..DocumentInfo::default()
        },
        available_tools: vec![
            "word.get_text".to_string(),
            "word.insert_paragraph".to_string(),
        ],
        is_active: Some(true),
    }
}

fn excel_session() -> NewSessionInfo {
    NewSessionInfo {
        session_id: "33333333-3333-4333-8333-333333333333".to_string(),
        instance_id: "excel-instance".to_string(),
        document: DocumentInfo {
            title: Some("Budget.xlsx".to_string()),
            filename: Some("Budget.xlsx".to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(false),
            ..DocumentInfo::default()
        },
        available_tools: vec![
            "excel.read_range".to_string(),
            "excel.write_range".to_string(),
            "excel.add_sheet".to_string(),
            "excel.set_formula".to_string(),
            "excel.format_range".to_string(),
            "excel.create_table".to_string(),
            "excel.create_chart".to_string(),
        ],
        is_active: Some(false),
    }
}

fn start_command_default() -> StartCommandInput {
    StartCommandInput {
        command_id: None,
        mcp_request_id: None,
        client_id: None,
        client_name: None,
        session_id: None,
        host_app: None,
        tool: String::new(),
        user_intent: None,
        timeout_ms: None,
        started_at: None,
    }
}

fn default_addin_public_dir() -> PathBuf {
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
mod tests {
    use super::{empty_state, seeded_state};
    use std::time::SystemTime;

    #[test]
    fn seeded_ui_fixture_is_redacted_and_grouped() {
        let seed = seeded_state(SystemTime::UNIX_EPOCH);
        let sessions = seed.registry.list_sessions();
        let snapshot = seed.ui_state.snapshot(&sessions, SystemTime::UNIX_EPOCH);

        assert_eq!(snapshot.clients.len(), 1);
        assert_eq!(snapshot.documents["word"].len(), 1);
        assert_eq!(snapshot.documents["excel"].len(), 1);
        assert_eq!(snapshot.current_tasks.len(), 1);
        assert_eq!(snapshot.recent_commands.len(), 10);
        let debug = format!("{snapshot:?}");
        assert!(!debug.contains("secret-value"));
        assert!(!debug.contains("base64,QUJDREVGRw"));
        assert!(debug.contains("certificate_passphrase=[redacted]"));
    }

    #[test]
    fn empty_ui_fixture_exercises_zero_state() {
        let seed = empty_state();
        let sessions = seed.registry.list_sessions();
        let snapshot = seed.ui_state.snapshot(&sessions, SystemTime::UNIX_EPOCH);

        assert_eq!(snapshot.clients.len(), 0);
        assert_eq!(snapshot.documents.values().map(Vec::len).sum::<usize>(), 0);
        assert_eq!(snapshot.current_tasks.len(), 0);
        assert_eq!(snapshot.recent_commands.len(), 0);
    }
}
