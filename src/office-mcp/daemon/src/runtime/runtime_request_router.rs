use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::mcp::{HttpMethod, McpHttpFrontend, McpHttpRequest};
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::mcp_response::{McpHttpResponseService, RuntimeSharedState};
use std::collections::BTreeMap;
use std::sync::Arc;
use std::time::SystemTime;

pub(crate) struct RuntimeRequestRouter;

impl RuntimeRequestRouter {
    pub(crate) fn route(
        frontend: &mut McpHttpFrontend,
        ui_state: &mut UiStateStore,
        registry: &SessionRegistry,
        shared_state: &Arc<RuntimeSharedState>,
        remote_addr: Option<String>,
        request: WireHttpRequest,
    ) -> WireHttpResponse {
        if request.path == "/healthz" && request.method == HttpMethod::Get {
            return WireHttpResponse::json(200, BTreeMap::new(), "{\"ok\":true}".to_string());
        }
        if request.path != "/mcp" {
            return WireHttpResponse::text(404, "Not found".to_string());
        }
        let body_bytes = request.body.len();
        let is_initialize = request.is_initialize();
        let body = request.body;
        let decision = frontend.handle_request(
            ui_state,
            &McpHttpRequest {
                method: request.method,
                headers: request.headers,
                remote_addr,
                body_bytes,
                is_initialize,
            },
            SystemTime::now(),
        );
        McpHttpResponseService::runtime_response(decision, registry, ui_state, shared_state, &body)
    }
}

#[cfg(test)]
fn test_shared_state() -> Arc<RuntimeSharedState> {
    use crate::addin_mgr::{AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher};
    use crate::common::AuditLog;
    use std::sync::Mutex;

    Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: std::time::Duration::from_secs(60),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::new())),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
    })
}

#[cfg(test)]
#[path = "runtime_request_router_tests.rs"]
mod runtime_request_router_tests;
