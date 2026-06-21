use super::WebSocketHeartbeatService;
use crate::addin_mgr::{
    AddInInfo, AddinChannelConfig, AddinChannelServer, AddinConnectionHub, CommandRouter, HostInfo,
    ImageFetcher, JsonRpcId, RegisterRequest, SessionRegistry,
};
use crate::common::AuditLog;
use crate::runtime::mcp_response::{HeartbeatLoopDecision, RuntimeSharedState};
use serde_json::Value;
use std::sync::{Arc, Mutex};
use std::time::Duration;

#[test]
fn start_ping_returns_json_payload_for_known_connection() {
    let shared_state = shared_state(Duration::from_secs(30), Duration::from_secs(10));
    register_connection(&shared_state, "connection-1");

    let payload = WebSocketHeartbeatService::start_ping(&shared_state, "connection-1")
        .expect("start ping")
        .expect("ping payload");
    let json: Value = serde_json::from_str(&payload).expect("ping json");

    assert_eq!(json["method"], "ping");
    assert!(json["id"].as_str().expect("id").starts_with("ping-"));
}

#[test]
fn start_ping_ignores_unknown_connection() {
    let shared_state = shared_state(Duration::from_secs(30), Duration::from_secs(10));

    let payload =
        WebSocketHeartbeatService::start_ping(&shared_state, "missing").expect("start ping");

    assert_eq!(payload, None);
}

#[test]
fn handle_response_accepts_matching_pong() {
    let shared_state = shared_state(Duration::from_secs(30), Duration::from_secs(10));
    register_connection(&shared_state, "connection-1");
    let payload = WebSocketHeartbeatService::start_ping(&shared_state, "connection-1")
        .expect("start ping")
        .expect("ping payload");
    let ping: Value = serde_json::from_str(&payload).expect("ping json");
    let ping_id = ping["id"].as_str().expect("ping id");

    let handled = WebSocketHeartbeatService::handle_response(
        &shared_state,
        "connection-1",
        &format!(r#"{{"jsonrpc":"2.0","id":"{ping_id}","result":{{}}}}"#),
    )
    .expect("handle pong");

    assert!(handled);
}

#[test]
fn handle_response_ignores_non_pong_messages() {
    let shared_state = shared_state(Duration::from_secs(30), Duration::from_secs(10));

    assert!(
        !WebSocketHeartbeatService::handle_response(
            &shared_state,
            "connection-1",
            r#"{"jsonrpc":"2.0","method":"session.added","params":{}}"#,
        )
        .expect("handle message")
    );
    assert!(
        !WebSocketHeartbeatService::handle_response(&shared_state, "connection-1", "not json")
            .expect("handle message")
    );
}

#[test]
fn timeout_decision_closes_after_third_missed_ping() {
    let shared_state = shared_state(Duration::from_secs(30), Duration::from_secs(10));
    register_connection(&shared_state, "connection-1");

    assert_eq!(
        WebSocketHeartbeatService::record_timeout(&shared_state, "connection-1")
            .expect("first timeout"),
        HeartbeatLoopDecision::KeepOpen
    );
    assert_eq!(
        WebSocketHeartbeatService::record_timeout(&shared_state, "connection-1")
            .expect("second timeout"),
        HeartbeatLoopDecision::KeepOpen
    );
    assert_eq!(
        WebSocketHeartbeatService::record_timeout(&shared_state, "connection-1")
            .expect("third timeout"),
        HeartbeatLoopDecision::Close
    );
}

fn shared_state(interval: Duration, timeout: Duration) -> RuntimeSharedState {
    RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: Duration::from_secs(60),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::with_config(
            AddinChannelConfig {
                heartbeat_interval: interval,
                heartbeat_timeout: timeout,
                ..AddinChannelConfig::default()
            },
        ))),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
    }
}

fn register_connection(shared_state: &RuntimeSharedState, connection_id: &str) {
    let mut registry = shared_state.registry.lock().expect("registry");
    let mut addin_channel = shared_state.addin_channel.lock().expect("addin channel");
    addin_channel
        .register_runtime(
            &mut registry,
            connection_id,
            RegisterRequest {
                id: JsonRpcId::String("register-1".to_string()),
                instance_id: "instance-1".to_string(),
                host: HostInfo {
                    app: "word".to_string(),
                    version: Some("16.0".to_string()),
                    platform: Some("windows".to_string()),
                    build: None,
                },
                add_in: AddInInfo {
                    version: "0.1.0".to_string(),
                    protocol_version: "1.0".to_string(),
                    supported_features: vec!["doc.read".to_string()],
                },
            },
            std::time::SystemTime::UNIX_EPOCH,
        )
        .expect("register runtime");
}
