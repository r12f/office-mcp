use super::AddinHttpService;
use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::mcp::HttpMethod;
use crate::runtime::http_wire::WireHttpRequest;
use crate::runtime::server_config::RuntimeServerConfig;
use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};

#[test]
fn healthz_returns_ok_json() {
    let response = route(request(HttpMethod::Get, "/healthz", BTreeMap::new()));

    assert_eq!(response.status, 200);
    assert!(response_text(&response).contains("{\"ok\":true}"));
}

#[test]
fn websocket_upgrade_accepts_exact_addin_origin() {
    let headers = BTreeMap::from([
        ("origin".to_string(), "https://localhost:8765".to_string()),
        ("upgrade".to_string(), "websocket".to_string()),
        ("connection".to_string(), "Upgrade".to_string()),
        (
            "sec-websocket-key".to_string(),
            "dGhlIHNhbXBsZSBub25jZQ==".to_string(),
        ),
    ]);
    let service = AddinHttpService::from_config(&RuntimeServerConfig::default());
    let request = request(HttpMethod::Get, "/addin", headers);

    assert!(service.is_valid_websocket_upgrade(&request));
    let response = route(request);

    assert_eq!(response.status, 101);
    assert_eq!(
        response.headers.get("Sec-WebSocket-Accept"),
        Some(&"s3pPLMBiTxaQ9kYGzzhZRbK+xOo=".to_string())
    );
}

#[test]
fn websocket_upgrade_rejects_foreign_origin() {
    let response = route(request(
        HttpMethod::Get,
        "/addin",
        BTreeMap::from([
            ("origin".to_string(), "https://evil.example".to_string()),
            ("upgrade".to_string(), "websocket".to_string()),
            ("connection".to_string(), "Upgrade".to_string()),
            ("sec-websocket-key".to_string(), "abc".to_string()),
        ]),
    ));

    assert_eq!(response.status, 403);
    assert!(response_text(&response).contains("Forbidden origin"));
}

#[test]
fn ui_state_uses_origin_guard_and_snapshot_service() {
    let allowed = route(request(
        HttpMethod::Get,
        "/ui/state",
        BTreeMap::from([("origin".to_string(), "https://localhost:8765".to_string())]),
    ));
    assert_eq!(allowed.status, 200);
    assert!(response_text(&allowed).contains("\"status\":\"up\""));

    let forbidden = route(request(
        HttpMethod::Get,
        "/ui/state",
        BTreeMap::from([("origin".to_string(), "https://evil.example".to_string())]),
    ));
    assert_eq!(forbidden.status, 403);
}

#[test]
fn addin_diagnostics_accepts_local_events_and_rejects_foreign_origins() {
    let mut allowed = request(
        HttpMethod::Post,
        "/addin/diagnostics",
        BTreeMap::from([("origin".to_string(), "https://localhost:8765".to_string())]),
    );
    allowed.body = br#"{"host_app":"word","event":"websocket.error"}"#.to_vec();

    let response = route(allowed);

    assert_eq!(response.status, 200);
    assert!(response_text(&response).contains("{\"ok\":true}"));

    let forbidden = route(request(
        HttpMethod::Post,
        "/addin/diagnostics",
        BTreeMap::from([("origin".to_string(), "https://evil.example".to_string())]),
    ));
    assert_eq!(forbidden.status, 403);

    let method_not_allowed = route(request(
        HttpMethod::Get,
        "/addin/diagnostics",
        BTreeMap::new(),
    ));
    assert_eq!(method_not_allowed.status, 405);
}

#[test]
fn non_get_requests_outside_addin_upgrade_are_rejected() {
    let response = route(request(HttpMethod::Post, "/taskpane.html", BTreeMap::new()));

    assert_eq!(response.status, 405);
    assert!(response_text(&response).contains("Method not allowed"));
}

fn route(request: WireHttpRequest) -> crate::runtime::http_wire::WireHttpResponse {
    let service = AddinHttpService::from_config(&RuntimeServerConfig::default());
    let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
    let registry = Arc::new(Mutex::new(SessionRegistry::new()));

    service.route_request(&ui_state, &registry, &request)
}

fn request(method: HttpMethod, path: &str, headers: BTreeMap<String, String>) -> WireHttpRequest {
    WireHttpRequest {
        method,
        path: path.to_string(),
        headers,
        body: Vec::new(),
    }
}

fn response_text(response: &crate::runtime::http_wire::WireHttpResponse) -> String {
    let bytes = response.to_bytes();
    let body = bytes
        .windows(4)
        .position(|window| window == b"\r\n\r\n")
        .map_or(&[][..], |index| &bytes[index + 4..]);
    String::from_utf8(body.to_vec()).expect("response body is UTF-8")
}
