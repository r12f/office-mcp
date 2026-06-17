use super::{RuntimeRequestRouter, test_shared_state};
use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::mcp::{HttpMethod, McpHttpFrontend};
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

fn request(method: HttpMethod, path: &str, body: Vec<u8>) -> WireHttpRequest {
    WireHttpRequest {
        method,
        path: path.to_string(),
        headers: BTreeMap::new(),
        body,
    }
}
