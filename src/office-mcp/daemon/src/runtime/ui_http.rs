use crate::addin_mgr::SessionRegistry;
use crate::api::{UiSnapshotEndpoints, UiSnapshotService, UiStateStore};
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::server_config::RuntimeServerConfig;
use crate::runtime::static_response::StaticResponseService;
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
        registry: &Arc<Mutex<SessionRegistry>>,
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
            return Some(self.ui_state_response(request, ui_state, registry));
        }
        if request.path == "/ui/events" {
            return Some(self.ui_events_response(request, ui_state, registry));
        }
        None
    }

    fn ui_state_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        registry: &Arc<Mutex<SessionRegistry>>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        WireHttpResponse::json(
            200,
            BTreeMap::new(),
            self.render_snapshot(ui_state, registry),
        )
    }

    fn ui_events_response(
        &self,
        request: &WireHttpRequest,
        ui_state: &Arc<Mutex<UiStateStore>>,
        registry: &Arc<Mutex<SessionRegistry>>,
    ) -> WireHttpResponse {
        if !self.allows_origin(request) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        WireHttpResponse::binary(
            200,
            "text/event-stream; charset=utf-8",
            format!(
                "event: snapshot\ndata: {}\n\n",
                self.render_snapshot(ui_state, registry)
            )
            .into_bytes(),
            BTreeMap::from([
                ("Cache-Control".to_string(), "no-store".to_string()),
                ("X-Accel-Buffering".to_string(), "no".to_string()),
            ]),
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
        registry: &Arc<Mutex<SessionRegistry>>,
    ) -> String {
        UiSnapshotService::new().render_runtime_snapshot(
            ui_state,
            registry,
            &UiSnapshotEndpoints {
                mcp_endpoint: self.mcp_endpoint.clone(),
                addin_endpoint: self.addin_endpoint.clone(),
            },
        )
    }
}
