use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::mcp::{HttpMethod, McpHttpFrontend, McpHttpRequest, McpHttpRequestClass};
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
        let class = mcp_request_class(&request.body, is_initialize);
        let body = request.body;
        let decision = frontend.handle_request(
            ui_state,
            &McpHttpRequest {
                method: request.method,
                headers: request.headers,
                remote_addr,
                body_bytes,
                is_initialize,
                class,
            },
            SystemTime::now(),
        );
        McpHttpResponseService::runtime_response(decision, registry, ui_state, shared_state, &body)
    }
}

fn mcp_request_class(body: &[u8], is_initialize: bool) -> McpHttpRequestClass {
    if is_initialize {
        return McpHttpRequestClass::Discovery;
    }
    let Ok(value) = serde_json::from_slice::<serde_json::Value>(body) else {
        return McpHttpRequestClass::Operation;
    };
    match value.get("method").and_then(serde_json::Value::as_str) {
        Some("tools/list" | "resources/list" | "resources/templates/list") => {
            McpHttpRequestClass::Discovery
        }
        Some("resources/read") if is_sessions_resource_read(&value) => {
            McpHttpRequestClass::Discovery
        }
        Some("tools/call") if is_discovery_tool_call(&value) => McpHttpRequestClass::Discovery,
        _ => McpHttpRequestClass::Operation,
    }
}

fn is_sessions_resource_read(value: &serde_json::Value) -> bool {
    value
        .get("params")
        .and_then(|params| params.get("uri"))
        .and_then(serde_json::Value::as_str)
        == Some("office://sessions")
}

fn is_discovery_tool_call(value: &serde_json::Value) -> bool {
    matches!(
        value
            .get("params")
            .and_then(|params| params.get("name"))
            .and_then(serde_json::Value::as_str),
        Some("office.list_sessions" | "office.get_session_info" | "office.describe_tools")
    )
}

#[cfg(test)]
fn test_shared_state() -> Arc<RuntimeSharedState> {
    use crate::addin_mgr::{AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher};
    use crate::common::AuditLog;
    use crate::mcp::ToolAccessPolicy;
    use std::sync::{Arc, Mutex};

    Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: std::time::Duration::from_mins(1),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::new())),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
        tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
        config_path: None,
    })
}

#[cfg(test)]
#[path = "runtime_request_router_tests.rs"]
mod runtime_request_router_tests;
