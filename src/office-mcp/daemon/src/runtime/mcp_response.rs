use crate::addin_mgr::SessionRegistry;
use crate::addin_mgr::{AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::mcp::{McpHttpDecision, ToolAccessPolicy};
use crate::runtime::http_wire::WireHttpResponse;
use crate::runtime::mcp_rpc::{McpDispatchContext, McpJsonRpcRuntime};
use serde_json::{Value, json};
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone)]
pub(crate) struct RuntimeSharedState {
    pub(crate) registry: Arc<Mutex<SessionRegistry>>,
    pub(crate) session_grace: Duration,
    pub(crate) addin_channel: Arc<Mutex<AddinChannelServer>>,
    pub(crate) connection_hub: Arc<AddinConnectionHub>,
    pub(crate) command_router: Arc<Mutex<CommandRouter>>,
    pub(crate) audit_log: AuditLog,
    pub(crate) image_fetcher: ImageFetcher,
    pub(crate) tool_access_policy: ToolAccessPolicy,
}

impl RuntimeSharedState {
    pub(crate) fn prune_stale_sessions(&self, now: SystemTime) -> usize {
        let Ok(mut registry) = self.registry.lock() else {
            tracing::warn!("failed to lock session registry for stale-session pruning");
            return 0;
        };
        let removed = registry.prune_stale_sessions(now, self.session_grace);
        if removed > 0 {
            tracing::info!(removed, "pruned stale Office document sessions");
        }
        removed
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HeartbeatLoopDecision {
    KeepOpen,
    Close,
}

pub(crate) struct McpHttpResponseService;

impl McpHttpResponseService {
    pub(crate) fn runtime_response(
        decision: McpHttpDecision,
        registry: &SessionRegistry,
        ui_state: &mut UiStateStore,
        shared_state: &Arc<RuntimeSharedState>,
        body: &[u8],
    ) -> WireHttpResponse {
        let request_id = json_rpc_request_id(body);
        match decision {
            McpHttpDecision::InitializeTransport { session_id } => {
                let initialize_body = json!({
                    "jsonrpc": "2.0",
                    "id": request_id.unwrap_or(Value::Null),
                    "result": {
                        "protocolVersion": "2025-06-18",
                        "capabilities": { "tools": {}, "resources": {}, "prompts": {} },
                        "serverInfo": { "name": "office-mcp", "version": "0.1.0" }
                    }
                })
                .to_string();
                WireHttpResponse::json(
                    200,
                    BTreeMap::from([
                        ("MCP-Session-Id".to_string(), session_id),
                        ("MCP-Protocol-Version".to_string(), "2025-06-18".to_string()),
                    ]),
                    initialize_body,
                )
            }
            McpHttpDecision::ForwardToTransport { .. } => {
                let mut context = McpDispatchContext {
                    registry,
                    ui_state,
                    addin_channel: &shared_state.addin_channel,
                    connection_hub: &shared_state.connection_hub,
                    command_router: &shared_state.command_router,
                    audit_log: &shared_state.audit_log,
                    image_fetcher: &shared_state.image_fetcher,
                    tool_access_policy: &shared_state.tool_access_policy,
                };
                let body = McpJsonRpcRuntime::handle_body(&mut context, body);
                WireHttpResponse::json(200, BTreeMap::new(), body)
            }
            McpHttpDecision::Reject {
                status,
                body,
                headers,
            } => {
                let mut response = WireHttpResponse::text(status, body);
                response.headers.extend(headers);
                response
            }
            McpHttpDecision::JsonRpcError {
                status,
                code,
                message,
            } => WireHttpResponse::json(
                status,
                BTreeMap::new(),
                json!({
                    "jsonrpc": "2.0",
                    "id": request_id.unwrap_or(Value::Null),
                    "error": { "code": code, "message": message }
                })
                .to_string(),
            ),
        }
    }
}

fn json_rpc_request_id(body: &[u8]) -> Option<Value> {
    serde_json::from_slice::<Value>(body)
        .ok()
        .and_then(|value| value.get("id").cloned())
}

#[cfg(test)]
#[path = "mcp_response_tests.rs"]
mod mcp_response_tests;
