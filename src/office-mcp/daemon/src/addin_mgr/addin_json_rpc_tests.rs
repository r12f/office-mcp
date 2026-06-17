use super::{JsonRpcEnvelope, JsonRpcId, RegisterResult};
use std::collections::BTreeMap;

#[test]
fn request_envelope_sets_id_method_and_params() {
    let envelope = JsonRpcEnvelope::request(
        "request-1".to_string(),
        "tool.invoke",
        BTreeMap::from([("session_id".to_string(), "session-1".to_string())]),
    );

    assert_eq!(
        envelope.id,
        Some(JsonRpcId::String("request-1".to_string()))
    );
    assert_eq!(envelope.method.as_deref(), Some("tool.invoke"));
    assert_eq!(envelope.params["session_id"], "session-1");
    assert_eq!(envelope.result, None);
}

#[test]
fn notification_envelope_has_no_id() {
    let envelope = JsonRpcEnvelope::notification(
        "tool.cancel",
        BTreeMap::from([("reason".to_string(), "timeout".to_string())]),
    );

    assert_eq!(envelope.id, None);
    assert_eq!(envelope.method.as_deref(), Some("tool.cancel"));
    assert_eq!(envelope.params["reason"], "timeout");
}

#[test]
fn success_envelope_carries_register_result() {
    let result = RegisterResult {
        server_version: "0.1.0".to_string(),
        protocol_version: "1.0".to_string(),
        session_grace_sec: 60,
        heartbeat_interval_sec: 30,
        max_pending_per_session: 4,
        assigned_instance_id: "instance-1".to_string(),
    };
    let envelope = JsonRpcEnvelope::success(JsonRpcId::Number(7), result.clone());

    assert_eq!(envelope.id, Some(JsonRpcId::Number(7)));
    assert_eq!(envelope.method, None);
    assert_eq!(envelope.result, Some(result));
}
