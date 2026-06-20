use super::{HttpMethod, McpHttpConfig, McpHttpDecision, McpHttpError, McpHttpRequest};
use std::collections::{BTreeMap, BTreeSet};

#[test]
fn config_allows_loopback_origins_for_configured_port() {
    let config = McpHttpConfig {
        host: "127.0.0.1".to_string(),
        port: 9010,
        ..McpHttpConfig::default()
    };

    assert_eq!(
        config.allowed_origins(),
        BTreeSet::from([
            "http://127.0.0.1:9010".to_string(),
            "http://localhost:9010".to_string(),
        ])
    );
}

#[test]
fn default_rate_limit_supports_batched_office_tool_e2e() {
    assert_eq!(McpHttpConfig::default().requests_per_minute, 1000);
}

#[test]
fn request_client_key_prefers_forwarded_for_first_hop() {
    let request = McpHttpRequest {
        method: HttpMethod::Post,
        headers: BTreeMap::from([(
            "x-forwarded-for".to_string(),
            "10.0.0.1, 10.0.0.2".to_string(),
        )]),
        remote_addr: Some("127.0.0.1".to_string()),
        body_bytes: 0,
        is_initialize: true,
    };

    assert_eq!(request.client_key(), "10.0.0.1");
}

#[test]
fn request_client_name_prefers_explicit_office_mcp_header() {
    let request = McpHttpRequest {
        method: HttpMethod::Post,
        headers: BTreeMap::from([
            (
                "x-office-mcp-client".to_string(),
                "Claude Desktop".to_string(),
            ),
            ("user-agent".to_string(), "Generic Agent".to_string()),
        ]),
        remote_addr: None,
        body_bytes: 0,
        is_initialize: true,
    };

    assert_eq!(request.client_name().as_deref(), Some("Claude Desktop"));
}

#[test]
fn http_errors_map_to_mcp_decisions() {
    assert_eq!(
        McpHttpDecision::from(McpHttpError::RateLimited),
        McpHttpDecision::Reject {
            status: 429,
            body: "Rate limit exceeded".to_string(),
            headers: BTreeMap::from([("Retry-After".to_string(), "60".to_string())]),
        }
    );
    assert_eq!(
        McpHttpDecision::from(McpHttpError::RequestTooLarge {
            max_request_bytes: 12
        }),
        McpHttpDecision::JsonRpcError {
            status: 413,
            code: -32000,
            message: "Request body exceeds 12 bytes.".to_string(),
        }
    );
}
