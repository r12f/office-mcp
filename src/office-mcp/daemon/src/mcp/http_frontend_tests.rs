use super::{HttpMethod, McpHttpConfig, McpHttpDecision, McpHttpFrontend, McpHttpRequest};
use crate::api::UiStateStore;
use std::collections::BTreeMap;
use std::time::{Duration, SystemTime};

#[test]
fn rejects_foreign_browser_origins() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();

    let decision = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Get, [("origin", "https://evil.example")], false),
        SystemTime::UNIX_EPOCH,
    );

    assert_eq!(decision, McpHttpDecision::reject(403, "Forbidden origin"));
}

#[test]
fn allows_configured_loopback_origins() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();

    let decision = frontend.handle_request(
        &mut ui_state,
        &request(
            HttpMethod::Post,
            [("origin", "http://localhost:8800")],
            true,
        ),
        SystemTime::UNIX_EPOCH,
    );

    assert!(matches!(
        decision,
        McpHttpDecision::InitializeTransport { .. }
    ));
}

#[test]
fn rate_limits_per_source_and_resets_after_window() {
    let mut frontend = McpHttpFrontend::with_config(McpHttpConfig {
        requests_per_minute: 1,
        ..McpHttpConfig::default()
    });
    let mut ui_state = UiStateStore::new();
    let first = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Get, [], false),
        SystemTime::UNIX_EPOCH,
    );
    let second = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Get, [], false),
        SystemTime::UNIX_EPOCH,
    );
    let third = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Get, [], false),
        SystemTime::UNIX_EPOCH + Duration::from_mins(1),
    );

    assert_eq!(
        first,
        McpHttpDecision::reject(400, "Invalid or missing MCP session ID")
    );
    assert_eq!(
        second,
        McpHttpDecision::reject_with_header(429, "Rate limit exceeded", "Retry-After", "60")
    );
    assert_eq!(
        third,
        McpHttpDecision::reject(400, "Invalid or missing MCP session ID")
    );
}

#[test]
fn initializes_session_and_tracks_ui_client() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();

    let decision = frontend.handle_request(
        &mut ui_state,
        &request(
            HttpMethod::Post,
            [("x-office-mcp-client", "copilot-cli/1.0")],
            true,
        ),
        SystemTime::UNIX_EPOCH,
    );

    let McpHttpDecision::InitializeTransport { session_id } = decision else {
        panic!("expected initialization");
    };
    assert_eq!(frontend.active_session_count(), 1);
    assert_eq!(
        ui_state.snapshot(&[], SystemTime::UNIX_EPOCH).clients[0].client_id,
        session_id
    );
    assert_eq!(
        ui_state.snapshot(&[], SystemTime::UNIX_EPOCH).clients[0]
            .name
            .as_deref(),
        Some("copilot-cli/1.0")
    );
}

#[test]
fn post_without_session_or_initialize_is_bad_request() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();

    let decision = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Post, [], false),
        SystemTime::UNIX_EPOCH,
    );

    assert_eq!(
        decision,
        McpHttpDecision::json_rpc_error(400, -32000, "Bad Request: missing MCP session ID.")
    );
}

#[test]
fn unknown_post_session_is_not_found() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();

    let decision = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Post, [("mcp-session-id", "missing")], false),
        SystemTime::UNIX_EPOCH,
    );

    assert_eq!(
        decision,
        McpHttpDecision::json_rpc_error(404, -32000, "Unknown MCP session ID.")
    );
}

#[test]
fn get_and_delete_require_known_session_id() {
    let mut frontend = McpHttpFrontend::new();
    let mut ui_state = UiStateStore::new();
    let McpHttpDecision::InitializeTransport { session_id } = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Post, [], true),
        SystemTime::UNIX_EPOCH,
    ) else {
        panic!("expected initialization");
    };

    let decision = frontend.handle_request(
        &mut ui_state,
        &request(
            HttpMethod::Get,
            [("mcp-session-id", session_id.as_str())],
            false,
        ),
        SystemTime::UNIX_EPOCH + Duration::from_secs(1),
    );

    assert_eq!(
        decision,
        McpHttpDecision::ForwardToTransport {
            session_id: Some(session_id)
        }
    );
}

#[test]
fn rejects_unsupported_methods_and_large_bodies() {
    let mut frontend = McpHttpFrontend::with_config(McpHttpConfig {
        max_request_bytes: 4,
        ..McpHttpConfig::default()
    });
    let mut ui_state = UiStateStore::new();

    let large = frontend.handle_request(
        &mut ui_state,
        &McpHttpRequest {
            body_bytes: 5,
            ..request(HttpMethod::Post, [], true)
        },
        SystemTime::UNIX_EPOCH,
    );
    let unsupported = frontend.handle_request(
        &mut ui_state,
        &request(HttpMethod::Patch, [], false),
        SystemTime::UNIX_EPOCH + Duration::from_mins(1),
    );

    assert!(matches!(
        large,
        McpHttpDecision::JsonRpcError { status: 413, .. }
    ));
    assert_eq!(
        unsupported,
        McpHttpDecision::reject(405, "Method not allowed")
    );
}

fn request<const N: usize>(
    method: HttpMethod,
    headers: [(&str, &str); N],
    is_initialize: bool,
) -> McpHttpRequest {
    McpHttpRequest {
        method,
        headers: BTreeMap::from(
            headers.map(|(name, value)| (name.to_ascii_lowercase(), value.to_string())),
        ),
        remote_addr: Some("127.0.0.1".to_string()),
        body_bytes: 0,
        is_initialize,
    }
}
