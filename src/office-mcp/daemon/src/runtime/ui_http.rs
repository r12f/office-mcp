use crate::api::{UiSnapshotEndpoints, UiSnapshotService, UiStateStore, redact_text_with_limit};
use crate::common::{DaemonConfigService, ToolAccessConfig};
use crate::mcp::{AccessMode, HttpMethod, ToolAccessPolicy};
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::mcp_response::RuntimeSharedState;
use crate::runtime::server_config::RuntimeServerConfig;
use crate::runtime::static_response::StaticResponseService;
use serde_json::Value;
use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::{Arc, Mutex};

const LOG_TAIL_MAX_BYTES: u64 = 64 * 1024;
const LOG_TAIL_MAX_CHARS: usize = 64 * 1024;

#[derive(Debug, Clone)]
pub(crate) struct UiHttpService {
    addin_origin: String,
    mcp_endpoint: String,
    addin_endpoint: String,
    assets: StaticResponseService,
    diagnostic_opener: Arc<dyn DiagnosticOpener>,
}

impl PartialEq for UiHttpService {
    fn eq(&self, other: &Self) -> bool {
        self.addin_origin == other.addin_origin
            && self.mcp_endpoint == other.mcp_endpoint
            && self.addin_endpoint == other.addin_endpoint
            && self.assets == other.assets
    }
}

impl Eq for UiHttpService {}

impl UiHttpService {
    #[must_use]
    pub(crate) fn from_config(config: &RuntimeServerConfig) -> Self {
        Self::from_config_with_diagnostic_opener(config, SystemDiagnosticOpener)
    }

    #[must_use]
    pub(crate) fn from_config_with_diagnostic_opener(
        config: &RuntimeServerConfig,
        diagnostic_opener: impl DiagnosticOpener + 'static,
    ) -> Self {
        Self {
            addin_origin: config.addin_origin.clone(),
            mcp_endpoint: format!("http://{}:{}/mcp", config.mcp_host, config.mcp_port),
            addin_endpoint: format!("{}/addin", config.addin_origin),
            assets: StaticResponseService::new(config.addin_public_dir.clone()),
            diagnostic_opener: Arc::new(diagnostic_opener),
        }
    }

    #[must_use]
    pub(crate) fn try_handle(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Option<WireHttpResponse> {
        if matches!(request.path.as_str(), "/ui" | "/ui/" | "/ui/index.html") {
            return Some(StaticResponseService::serve_ui_asset("index.html"));
        }
        if request.path == "/ui/app.css" {
            return Some(StaticResponseService::serve_ui_asset("app.css"));
        }
        if request.path == "/ui/app.js" {
            return Some(StaticResponseService::serve_ui_asset("app.js"));
        }
        if request.path == "/ui/state" {
            return Some(self.ui_state_response(request, ui_state, shared_state));
        }
        if request.path == "/ui/events" {
            return Some(self.ui_events_response(request, ui_state, shared_state));
        }
        if request.path == "/ui/log-tail" {
            return Some(self.log_tail_response(request, ui_state, shared_state));
        }
        if request.path == "/ui/tool-access-policy" && request.method == HttpMethod::Put {
            return Some(self.tool_access_policy_update_response(request, ui_state, shared_state));
        }
        if request.path == "/ui/open-diagnostic" && request.method == HttpMethod::Post {
            return Some(self.open_diagnostic_response(request, ui_state, shared_state));
        }
        None
    }

    fn ui_state_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        WireHttpResponse::json(
            200,
            BTreeMap::new(),
            self.render_snapshot(ui_state, shared_state),
        )
    }

    fn ui_events_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        WireHttpResponse::binary(
            200,
            "text/event-stream; charset=utf-8",
            format!(
                "event: snapshot\ndata: {}\n\n",
                self.render_snapshot(ui_state, shared_state)
            )
            .into_bytes(),
            BTreeMap::from([
                ("Cache-Control".to_string(), "no-store".to_string()),
                ("X-Accel-Buffering".to_string(), "no".to_string()),
            ]),
        )
    }

    fn tool_access_policy_update_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        let policy = match parse_tool_access_policy(&request.body) {
            Ok(policy) => policy,
            Err(message) => return WireHttpResponse::text(400, message),
        };
        if let Some(config_path) = &shared_state.config_path {
            if let Err(error) = DaemonConfigService::save_tool_access_config(
                &PathBuf::from(config_path),
                &tool_access_config_from_policy(&policy),
            ) {
                tracing::error!(%error, config_path, "failed to persist daemon tool access policy");
                return WireHttpResponse::text(
                    500,
                    "Failed to persist tool access policy".to_string(),
                );
            }
        } else {
            tracing::warn!("updated daemon tool access policy without config persistence path");
        }
        if !shared_state.set_tool_access_policy(policy.clone()) {
            return WireHttpResponse::text(500, "Failed to update tool access policy".to_string());
        }
        if let Ok(mut ui_state) = ui_state.lock() {
            ui_state.set_tool_access_policy(policy);
        } else {
            tracing::warn!("failed to lock UI state for tool access policy update");
            return WireHttpResponse::text(500, "Failed to update UI state".to_string());
        }
        WireHttpResponse::json(
            200,
            BTreeMap::new(),
            self.render_snapshot(ui_state, shared_state),
        )
    }

    fn log_tail_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        let path = match diagnostic_path(DiagnosticTarget::Log, ui_state, shared_state) {
            Ok(path) => path,
            Err(message) => return WireHttpResponse::text(404, message),
        };
        let tail = match read_log_tail(&path) {
            Ok(tail) => tail,
            Err(error) => {
                tracing::warn!(%error, path = %path.display(), "failed to read daemon log tail");
                return WireHttpResponse::text(404, "Log file is not readable".to_string());
            }
        };
        let redacted = redact_text_with_limit(&tail.text, LOG_TAIL_MAX_CHARS);
        let body = serde_json::json!({
            "path": path.display().to_string(),
            "text": redacted,
            "truncated": tail.truncated,
            "bytes_read": tail.bytes_read,
        });
        WireHttpResponse::json(200, BTreeMap::new(), body.to_string())
    }

    fn open_diagnostic_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        let target = match parse_open_diagnostic_target(&request.body) {
            Ok(target) => target,
            Err(message) => return WireHttpResponse::text(400, message),
        };
        let path = match diagnostic_path(target, ui_state, shared_state) {
            Ok(path) => path,
            Err(message) => return WireHttpResponse::text(404, message),
        };
        let request = DiagnosticOpenRequest { target, path };
        if let Err(error) = self.diagnostic_opener.open(&request) {
            tracing::error!(%error, target = request.target.as_str(), path = %request.path.display(), "failed to open daemon diagnostic path");
            return WireHttpResponse::text(500, "Failed to open diagnostic path".to_string());
        }
        tracing::info!(target = request.target.as_str(), path = %request.path.display(), "opened daemon diagnostic path");
        WireHttpResponse::json(200, BTreeMap::new(), r#"{"ok":true}"#.to_string())
    }

    fn allows_origin(&self, request: &WireHttpRequest) -> bool {
        request
            .headers
            .get("origin")
            .is_none_or(|origin| origin == &self.addin_origin)
    }

    fn render_snapshot(
        &self,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> String {
        let registry = shared_state.registry.clone();
        UiSnapshotService::new().render_runtime_snapshot(
            ui_state,
            &registry,
            &UiSnapshotEndpoints {
                mcp_endpoint: self.mcp_endpoint.clone(),
                addin_endpoint: self.addin_endpoint.clone(),
            },
        )
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct LogTail {
    text: String,
    truncated: bool,
    bytes_read: usize,
}

fn read_log_tail(path: &Path) -> Result<LogTail, std::io::Error> {
    let mut file = File::open(path)?;
    let length = file.metadata()?.len();
    let start = length.saturating_sub(LOG_TAIL_MAX_BYTES);
    file.seek(SeekFrom::Start(start))?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    if start > 0
        && let Some(index) = bytes.iter().position(|byte| *byte == b'\n')
    {
        bytes.drain(..=index);
    }
    Ok(LogTail {
        bytes_read: bytes.len(),
        text: String::from_utf8_lossy(&bytes).into_owned(),
        truncated: start > 0,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DiagnosticTarget {
    Config,
    Log,
}

impl DiagnosticTarget {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Config => "config",
            Self::Log => "log",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DiagnosticOpenRequest {
    pub(crate) target: DiagnosticTarget,
    pub(crate) path: PathBuf,
}

pub(crate) trait DiagnosticOpener: Send + Sync + std::fmt::Debug {
    fn open(&self, request: &DiagnosticOpenRequest) -> Result<(), String>;
}

#[derive(Debug, Clone, Copy, Default)]
struct SystemDiagnosticOpener;

impl DiagnosticOpener for SystemDiagnosticOpener {
    fn open(&self, request: &DiagnosticOpenRequest) -> Result<(), String> {
        open_system_path(&request.path)
    }
}

fn diagnostic_path(
    target: DiagnosticTarget,
    ui_state: &Arc<Mutex<UiStateStore>>,
    shared_state: &Arc<RuntimeSharedState>,
) -> Result<PathBuf, String> {
    match target {
        DiagnosticTarget::Config => shared_state
            .config_path
            .as_deref()
            .map(PathBuf::from)
            .ok_or_else(|| "Config path is not available".to_string()),
        DiagnosticTarget::Log => ui_state
            .lock()
            .map_err(|_| "Failed to read UI state".to_string())?
            .snapshot(&[], std::time::SystemTime::UNIX_EPOCH)
            .daemon
            .log_path
            .map(PathBuf::from)
            .ok_or_else(|| "Log path is not available".to_string()),
    }
}

fn parse_open_diagnostic_target(body: &[u8]) -> Result<DiagnosticTarget, String> {
    let value = serde_json::from_slice::<Value>(body).map_err(|error| error.to_string())?;
    match value.get("target").and_then(Value::as_str) {
        Some("config") => Ok(DiagnosticTarget::Config),
        Some("log") => Ok(DiagnosticTarget::Log),
        Some(other) => Err(format!("Unsupported diagnostic target: {other}")),
        None => Err("Diagnostic target is required".to_string()),
    }
}

fn open_system_path(path: &Path) -> Result<(), String> {
    #[cfg(windows)]
    {
        Command::new("explorer.exe")
            .arg(path)
            .spawn()
            .map_err(|error| error.to_string())?;
        return Ok(());
    }
    #[cfg(target_os = "macos")]
    {
        Command::new("open")
            .arg(path)
            .spawn()
            .map_err(|error| error.to_string())?;
        return Ok(());
    }
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        Command::new("xdg-open")
            .arg(path)
            .spawn()
            .map_err(|error| error.to_string())?;
        return Ok(());
    }
    #[allow(unreachable_code)]
    Err("opening diagnostic paths is unsupported on this platform".to_string())
}

fn tool_access_config_from_policy(policy: &ToolAccessPolicy) -> ToolAccessConfig {
    let snapshot = policy.snapshot();
    ToolAccessConfig {
        access_mode: snapshot.access_mode,
        disabled_apps: snapshot.disabled_apps,
        disabled_categories: snapshot.disabled_categories,
        disabled_tools: snapshot.disabled_tools,
    }
}

fn parse_tool_access_policy(body: &[u8]) -> Result<ToolAccessPolicy, String> {
    let value = serde_json::from_slice::<Value>(body).map_err(|error| error.to_string())?;
    let access_mode = match value
        .get("access_mode")
        .and_then(Value::as_str)
        .unwrap_or("all")
    {
        "read" => AccessMode::Read,
        "write" => AccessMode::Write,
        "all" => AccessMode::All,
        other => return Err(format!("Unsupported access_mode: {other}")),
    };
    let mut policy = ToolAccessPolicy::default().with_access_mode(access_mode);
    for app in string_array(&value, "disabled_apps")? {
        policy = policy.with_disabled_app(&app);
    }
    for category in value
        .get("disabled_categories")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
    {
        let app = category
            .get("app")
            .and_then(Value::as_str)
            .ok_or_else(|| "disabled_categories entries require app".to_string())?;
        let category_name = category
            .get("category")
            .and_then(Value::as_str)
            .ok_or_else(|| "disabled_categories entries require category".to_string())?;
        policy = policy.with_disabled_category(app, category_name);
    }
    for tool in string_array(&value, "disabled_tools")? {
        policy = policy.with_disabled_tool(&tool);
    }
    Ok(policy)
}

fn string_array(value: &Value, field: &str) -> Result<Vec<String>, String> {
    value
        .get(field)
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .map(|item| {
            item.as_str()
                .map(ToString::to_string)
                .ok_or_else(|| format!("{field} entries must be strings"))
        })
        .collect()
}
