use crate::api::{UiSnapshotEndpoints, UiSnapshotService, UiStateStore};
use crate::mcp::{AccessMode, HttpMethod, ToolAccessPolicy};
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::mcp_response::RuntimeSharedState;
use crate::runtime::server_config::RuntimeServerConfig;
use crate::runtime::static_response::StaticResponseService;
use serde_json::Value;
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct UiHttpService {
    addin_origin: String,
    mcp_endpoint: String,
    addin_endpoint: String,
    assets: StaticResponseService,
}

impl UiHttpService {
    #[must_use]
    pub(crate) fn from_config(config: &RuntimeServerConfig) -> Self {
        Self {
            addin_origin: config.addin_origin.clone(),
            mcp_endpoint: format!("http://{}:{}/mcp", config.mcp_host, config.mcp_port),
            addin_endpoint: format!("{}/addin", config.addin_origin),
            assets: StaticResponseService::new(config.addin_public_dir.clone()),
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
        if request.path == "/ui/tool-access-policy" && request.method == HttpMethod::Put {
            return Some(self.tool_access_policy_update_response(request, ui_state, shared_state));
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
