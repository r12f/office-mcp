use super::AddinHttpService;
use crate::addin_mgr::{
    AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher, SessionRegistry,
};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::mcp::{HttpMethod, ToolAccessPolicy};
use crate::runtime::http_wire::WireHttpRequest;
use crate::runtime::mcp_response::RuntimeSharedState;
use crate::runtime::server_config::RuntimeServerConfig;
use std::collections::BTreeMap;
use std::fs;
use std::path::PathBuf;
use std::sync::{Arc, Mutex};

#[test]
fn healthz_returns_ok_json() {
    let response = route(&request(HttpMethod::Get, "/healthz", BTreeMap::new()));

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
    let response = route(&request);

    assert_eq!(response.status, 101);
    assert_eq!(
        response.headers.get("Sec-WebSocket-Accept"),
        Some(&"s3pPLMBiTxaQ9kYGzzhZRbK+xOo=".to_string())
    );
}

#[test]
fn websocket_upgrade_rejects_foreign_origin() {
    let response = route(&request(
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
    let allowed = route(&request(
        HttpMethod::Get,
        "/ui/state",
        BTreeMap::from([("origin".to_string(), "https://localhost:8765".to_string())]),
    ));
    assert_eq!(allowed.status, 200);
    assert!(response_text(&allowed).contains("\"status\":\"up\""));

    let forbidden = route(&request(
        HttpMethod::Get,
        "/ui/state",
        BTreeMap::from([("origin".to_string(), "https://evil.example".to_string())]),
    ));
    assert_eq!(forbidden.status, 403);
}

#[test]
fn tool_access_policy_update_changes_ui_state_and_runtime_policy() {
    let service = AddinHttpService::from_config(&RuntimeServerConfig::default());
    let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
    let shared_state = shared_state();
    let mut update = request(
        HttpMethod::Put,
        "/ui/tool-access-policy",
        BTreeMap::from([("origin".to_string(), "https://localhost:8765".to_string())]),
    );
    update.body = br#"{
        "access_mode":"read",
        "disabled_apps":["powerpoint"],
        "disabled_categories":[{"app":"excel","category":"Range"}],
        "disabled_tools":["word.update_table"]
    }"#
    .to_vec();

    let response = service.route_request(&ui_state, &shared_state, &update);

    assert_eq!(response.status, 200);
    let body = response_text(&response);
    assert!(body.contains("\"access_mode\":\"read\""));
    let snapshot = ui_state
        .lock()
        .expect("ui state")
        .snapshot(&[], std::time::SystemTime::UNIX_EPOCH);
    assert_eq!(
        snapshot.daemon.tool_access_policy.disabled_apps,
        vec!["powerpoint"]
    );
    assert!(
        shared_state
            .tool_access_policy()
            .allows_tool("word.get_text")
    );
    assert!(
        !shared_state
            .tool_access_policy()
            .allows_tool("word.insert_paragraph")
    );
    assert!(
        !shared_state
            .tool_access_policy()
            .allows_tool("excel.read_range")
    );
    assert!(
        !shared_state
            .tool_access_policy()
            .allows_tool("powerpoint.list_slides")
    );
}

#[test]
fn tool_access_policy_update_persists_config_file() {
    let service = AddinHttpService::from_config(&RuntimeServerConfig::default());
    let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
    let config_path = temp_config_path();
    fs::write(
        &config_path,
        r#"
[mcp_http]
port = 8800
"#,
    )
    .expect("write config");
    let shared_state = shared_state_with_config_path(Some(config_path.display().to_string()));
    let mut update = request(
        HttpMethod::Put,
        "/ui/tool-access-policy",
        BTreeMap::from([("origin".to_string(), "https://localhost:8765".to_string())]),
    );
    update.body = br#"{
        "access_mode":"write",
        "disabled_apps":["powerpoint"],
        "disabled_categories":[{"app":"excel","category":"Range"}],
        "disabled_tools":["word.update_table"]
    }"#
    .to_vec();

    let response = service.route_request(&ui_state, &shared_state, &update);

    assert_eq!(response.status, 200);
    let config = fs::read_to_string(&config_path).expect("read config");
    assert!(config.contains("[mcp_http]"));
    assert!(config.contains("[tool_access]"));
    assert!(config.contains("access_mode = \"write\""));
    assert!(config.contains("disabled_apps = \"powerpoint\""));
    assert!(config.contains("disabled_categories = \"excel:Range\""));
    assert!(config.contains("disabled_tools = \"word.update_table\""));
    let _ = fs::remove_file(config_path);
}

#[test]
fn tool_access_policy_update_rejects_foreign_origin() {
    let mut request = request(
        HttpMethod::Put,
        "/ui/tool-access-policy",
        BTreeMap::from([("origin".to_string(), "https://evil.example".to_string())]),
    );
    request.body = br#"{"access_mode":"read"}"#.to_vec();

    let response = route(&request);

    assert_eq!(response.status, 403);
}

#[test]
fn addin_diagnostics_accepts_local_events_and_rejects_foreign_origins() {
    let mut allowed = request(
        HttpMethod::Post,
        "/addin/diagnostics",
        BTreeMap::from([("origin".to_string(), "https://localhost:8765".to_string())]),
    );
    allowed.body = br#"{"host_app":"word","event":"websocket.error"}"#.to_vec();

    let response = route(&allowed);

    assert_eq!(response.status, 200);
    assert!(response_text(&response).contains("{\"ok\":true}"));

    let forbidden = route(&request(
        HttpMethod::Post,
        "/addin/diagnostics",
        BTreeMap::from([("origin".to_string(), "https://evil.example".to_string())]),
    ));
    assert_eq!(forbidden.status, 403);

    let method_not_allowed = route(&request(
        HttpMethod::Get,
        "/addin/diagnostics",
        BTreeMap::new(),
    ));
    assert_eq!(method_not_allowed.status, 405);
}

#[test]
fn non_get_requests_outside_addin_upgrade_are_rejected() {
    let response = route(&request(
        HttpMethod::Post,
        "/taskpane.html",
        BTreeMap::new(),
    ));

    assert_eq!(response.status, 405);
    assert!(response_text(&response).contains("Method not allowed"));
}

fn route(request: &WireHttpRequest) -> crate::runtime::http_wire::WireHttpResponse {
    let service = AddinHttpService::from_config(&RuntimeServerConfig::default());
    let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
    let shared_state = shared_state();

    service.route_request(&ui_state, &shared_state, request)
}

fn shared_state() -> Arc<RuntimeSharedState> {
    shared_state_with_config_path(None)
}

fn shared_state_with_config_path(config_path: Option<String>) -> Arc<RuntimeSharedState> {
    Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: std::time::Duration::from_secs(60),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::new())),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
        tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
        config_path,
    })
}

fn temp_config_path() -> PathBuf {
    std::env::temp_dir().join(format!(
        "office-mcp-ui-tool-access-{}-{}.toml",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("time")
            .as_nanos()
    ))
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
