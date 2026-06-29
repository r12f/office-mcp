use super::{RuntimeRequestRouter, test_shared_state};
use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::mcp::{HttpMethod, McpHttpConfig, McpHttpFrontend};
use crate::runtime::http_wire::WireHttpRequest;
use std::collections::BTreeMap;

#[test]
fn healthz_returns_ok_json_without_mcp_session() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let response = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &test_shared_state(),
        Some("127.0.0.1".to_string()),
        request(HttpMethod::Get, "/healthz", Vec::new()),
    );

    assert_eq!(response.status, 200);
    let bytes = response.to_bytes();
    let text = String::from_utf8(bytes).expect("http response is utf8");
    assert!(text.contains("Content-Type: application/json"));
    assert!(text.ends_with("{\"ok\":true}"));
}

#[test]
fn non_mcp_path_returns_not_found() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let response = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &test_shared_state(),
        None,
        request(HttpMethod::Get, "/missing", Vec::new()),
    );

    assert_eq!(response.status, 404);
    let text = String::from_utf8(response.to_bytes()).expect("http response is utf8");
    assert!(text.ends_with("Not found"));
}

#[test]
fn discovery_json_rpc_requests_survive_operation_rate_limit() {
    let mut frontend = McpHttpFrontend::with_config(McpHttpConfig {
        requests_per_minute: 1,
        ..McpHttpConfig::default()
    });
    let mut ui_state = UiStateStore::new();
    let registry = SessionRegistry::new();
    let shared_state = test_shared_state();
    let initialize = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &shared_state,
        Some("127.0.0.1".to_string()),
        request(
            HttpMethod::Post,
            "/mcp",
            br#"{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}"#.to_vec(),
        ),
    );
    let session_id = header_value(&initialize, "MCP-Session-Id").expect("session id");
    let first_call = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &shared_state,
        Some("127.0.0.1".to_string()),
        session_request(
            &session_id,
            br#"{"jsonrpc":"2.0","id":"call-1","method":"tools/call","params":{"name":"word.insert_paragraph","arguments":{"session_id":"missing","text":"x","anchor":{"kind":"end_of_document"}}}}"#.to_vec(),
        ),
    );
    let second_call = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &shared_state,
        Some("127.0.0.1".to_string()),
        session_request(
            &session_id,
            br#"{"jsonrpc":"2.0","id":"call-2","method":"tools/call","params":{"name":"word.insert_paragraph","arguments":{"session_id":"missing","text":"x","anchor":{"kind":"end_of_document"}}}}"#.to_vec(),
        ),
    );
    let resources_list = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &shared_state,
        Some("127.0.0.1".to_string()),
        session_request(
            &session_id,
            br#"{"jsonrpc":"2.0","id":"resources","method":"resources/list","params":{}}"#.to_vec(),
        ),
    );
    let sessions_resource = RuntimeRequestRouter::route(
        &mut frontend,
        &mut ui_state,
        &registry,
        &shared_state,
        Some("127.0.0.1".to_string()),
        session_request(
            &session_id,
            br#"{"jsonrpc":"2.0","id":"sessions","method":"resources/read","params":{"uri":"office://sessions"}}"#.to_vec(),
        ),
    );

    assert_eq!(first_call.status, 200);
    assert_eq!(second_call.status, 429);
    assert_eq!(
        header_value(&second_call, "Retry-After").as_deref(),
        Some("60")
    );
    assert_eq!(resources_list.status, 200);
    assert_eq!(sessions_resource.status, 200);
}

fn request(method: HttpMethod, path: &str, body: Vec<u8>) -> WireHttpRequest {
    WireHttpRequest {
        method,
        path: path.to_string(),
        headers: BTreeMap::new(),
        body,
    }
}

fn session_request(session_id: &str, body: Vec<u8>) -> WireHttpRequest {
    WireHttpRequest {
        method: HttpMethod::Post,
        path: "/mcp".to_string(),
        headers: BTreeMap::from([("mcp-session-id".to_string(), session_id.to_string())]),
        body,
    }
}

fn header_value(
    response: &crate::runtime::http_wire::WireHttpResponse,
    name: &str,
) -> Option<String> {
    response.headers.get(name).cloned()
}
