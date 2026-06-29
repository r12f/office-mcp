use super::{McpHttpResponseService, RuntimeSharedState};
use crate::addin_mgr::SessionRegistry;
use crate::addin_mgr::{AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::mcp::{McpHttpDecision, ToolAccessPolicy};
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};

#[test]
fn initialize_response_includes_session_headers_and_request_id() {
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let shared_state = shared_state();
    let response = McpHttpResponseService::runtime_response(
        McpHttpDecision::InitializeTransport {
            session_id: "mcp-session-1".to_string(),
        },
        &registry,
        &mut ui_state,
        &shared_state,
        br#"{"jsonrpc":"2.0","id":7,"method":"initialize"}"#,
    );
    let text = response_text(&response);

    assert!(text.starts_with("HTTP/1.1 200 OK"));
    assert!(text.contains("MCP-Session-Id: mcp-session-1"));
    assert!(text.contains("MCP-Protocol-Version: 2025-06-18"));
    assert!(text.contains("\"id\":7"));
    assert!(text.contains("\"name\":\"office-mcp\""));
}

#[test]
fn reject_response_preserves_frontend_headers() {
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let shared_state = shared_state();
    let response = McpHttpResponseService::runtime_response(
        McpHttpDecision::Reject {
            status: 429,
            body: "Rate limit exceeded".to_string(),
            headers: BTreeMap::from([("Retry-After".to_string(), "60".to_string())]),
            json_rpc_code: Some(-32000),
            office_mcp_code: Some("RATE_LIMITED".to_string()),
        },
        &registry,
        &mut ui_state,
        &shared_state,
        b"{}",
    );
    let text = response_text(&response);

    assert!(text.starts_with("HTTP/1.1 429 Too Many Requests"));
    assert!(text.contains("Retry-After: 60"));
    assert!(text.contains("Rate limit exceeded"));
    assert!(text.contains("\"office_mcp_code\":\"RATE_LIMITED\""));
}

#[test]
fn json_rpc_error_response_reuses_request_id() {
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let shared_state = shared_state();
    let response = McpHttpResponseService::runtime_response(
        McpHttpDecision::JsonRpcError {
            status: 413,
            code: -32000,
            message: "Request body exceeds limit.".to_string(),
        },
        &registry,
        &mut ui_state,
        &shared_state,
        br#"{"jsonrpc":"2.0","id":"abc"}"#,
    );
    let text = response_text(&response);

    assert!(text.starts_with("HTTP/1.1 413 Payload Too Large"));
    assert!(text.contains("\"id\":\"abc\""));
    assert!(text.contains("\"code\":-32000"));
    assert!(text.contains("Request body exceeds limit."));
}

#[test]
fn forwarded_response_invokes_mcp_runtime() {
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let shared_state = shared_state();
    let response = McpHttpResponseService::runtime_response(
        McpHttpDecision::ForwardToTransport {
            session_id: Some("mcp-session-1".to_string()),
        },
        &registry,
        &mut ui_state,
        &shared_state,
        br#"{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}"#,
    );
    let text = response_text(&response);

    assert!(text.starts_with("HTTP/1.1 200 OK"));
    assert!(text.contains("office.list_sessions"));
    assert!(text.contains("word.get_text"));
}

fn shared_state() -> Arc<RuntimeSharedState> {
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

fn response_text(response: &crate::runtime::http_wire::WireHttpResponse) -> String {
    String::from_utf8(response.to_bytes()).expect("response utf8")
}
