use crate::addin_mgr::{SessionRegistry, websocket_accept_key};
use crate::api::UiStateStore;
use crate::mcp::HttpMethod;
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::server_config::RuntimeServerConfig;
use crate::runtime::static_response::StaticResponseService;
use crate::runtime::ui_http::UiHttpService;
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct AddinHttpService {
    addin_origin: String,
    assets: StaticResponseService,
    ui: UiHttpService,
}

impl AddinHttpService {
    #[must_use]
    pub(crate) fn from_config(config: &RuntimeServerConfig) -> Self {
        Self {
            addin_origin: config.addin_origin.clone(),
            assets: StaticResponseService::new(config.addin_public_dir.clone()),
            ui: UiHttpService::from_config(config),
        }
    }

    #[must_use]
    pub(crate) fn route_request(
        &self,
        ui_state: &Arc<Mutex<UiStateStore>>,
        registry: &Arc<Mutex<SessionRegistry>>,
        request: &WireHttpRequest,
    ) -> WireHttpResponse {
        if request.path == "/healthz" && request.method == HttpMethod::Get {
            return WireHttpResponse::json(200, BTreeMap::new(), "{\"ok\":true}".to_string());
        }
        if request.path == "/addin/diagnostics" {
            return self.route_addin_diagnostics(request);
        }
        if request.path == "/addin" {
            return self.route_websocket_upgrade(request);
        }
        if request.method == HttpMethod::Get
            && let Some(response) = self.ui.try_handle(request, ui_state, registry)
        {
            return response;
        }
        if request.method != HttpMethod::Get {
            return WireHttpResponse::text(405, "Method not allowed".to_string());
        }
        self.assets.serve_addin_asset(&request.path)
    }

    fn route_addin_diagnostics(&self, request: &WireHttpRequest) -> WireHttpResponse {
        if request.method != HttpMethod::Post {
            return WireHttpResponse::text(405, "Method not allowed".to_string());
        }
        if request
            .headers
            .get("origin")
            .is_some_and(|origin| origin != &self.addin_origin)
        {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        let body = String::from_utf8_lossy(&request.body);
        let clipped = body.chars().take(2000).collect::<String>();
        tracing::info!(
            component = "addin_diagnostics",
            body = %clipped,
            "received add-in diagnostic event"
        );
        WireHttpResponse::json(200, BTreeMap::new(), "{\"ok\":true}".to_string())
    }

    #[must_use]
    pub(crate) fn is_valid_websocket_upgrade(&self, request: &WireHttpRequest) -> bool {
        request.path == "/addin"
            && request.method == HttpMethod::Get
            && request.headers.get("origin") == Some(&self.addin_origin)
            && request
                .headers
                .get("upgrade")
                .is_some_and(|value| value.eq_ignore_ascii_case("websocket"))
            && request
                .headers
                .get("connection")
                .is_some_and(|value| value.to_ascii_lowercase().contains("upgrade"))
            && request.headers.contains_key("sec-websocket-key")
    }

    fn route_websocket_upgrade(&self, request: &WireHttpRequest) -> WireHttpResponse {
        if request.method != HttpMethod::Get {
            return WireHttpResponse::text(405, "Method not allowed".to_string());
        }
        if request.headers.get("origin") != Some(&self.addin_origin) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        let upgrade = request
            .headers
            .get("upgrade")
            .is_some_and(|value| value.eq_ignore_ascii_case("websocket"));
        let connection = request
            .headers
            .get("connection")
            .is_some_and(|value| value.to_ascii_lowercase().contains("upgrade"));
        let Some(key) = request.headers.get("sec-websocket-key") else {
            return WireHttpResponse::text(400, "Missing Sec-WebSocket-Key".to_string());
        };
        if !upgrade || !connection {
            return WireHttpResponse::text(400, "Invalid WebSocket upgrade".to_string());
        }
        WireHttpResponse::switching_protocols(BTreeMap::from([
            ("Upgrade".to_string(), "websocket".to_string()),
            ("Connection".to_string(), "Upgrade".to_string()),
            (
                "Sec-WebSocket-Accept".to_string(),
                websocket_accept_key(key),
            ),
        ]))
    }
}

#[cfg(test)]
#[path = "addin_http_tests.rs"]
mod addin_http_tests;
