use crate::addin_mgr::ImageFetcher;
use crate::common::AuditLog;
use crate::runtime::evidence_fixture_config::{
    UiFixtureOptions, UiFixtureState, default_addin_public_dir, fixture_config,
};
use crate::runtime::evidence_fixture_seed::{empty_state, seeded_state};
use crate::runtime::{RuntimeServer, RuntimeServerConfig, RuntimeServerError};
use crate::ui::UiRuntimeFile;
use std::net::TcpListener;
use std::path::PathBuf;
use std::time::Duration;

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
        session_grace: Duration::from_secs(config.addin.session_grace_sec),
        heartbeat_interval: Duration::from_secs(config.addin.heartbeat_interval_sec),
        heartbeat_timeout: Duration::from_secs(config.addin.heartbeat_timeout_sec),
        requests_per_minute: config.limits.requests_per_minute,
        config_path: None,
        log_path: Some(config.logging.file.clone()),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
    });
    let seed = match options.state {
        UiFixtureState::Seeded => seeded_state(std::time::SystemTime::now()),
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

#[cfg(test)]
#[path = "evidence_fixture_tests.rs"]
mod evidence_fixture_tests;
