use super::{HeartbeatDecision, RegisterRequest, SessionRemovedEvent, SessionRemovedReason};
use crate::addin_mgr::{AddInInfo, HostInfo, JsonRpcId};

#[test]
fn register_request_is_cloneable_protocol_data() {
    let request = RegisterRequest {
        id: JsonRpcId::String("register-1".to_string()),
        instance_id: "instance-1".to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["doc.read".to_string()],
        },
    };

    assert_eq!(request.clone(), request);
}

#[test]
fn session_removed_reason_preserves_reported_close_cause() {
    let event = SessionRemovedEvent {
        session_id: "session-1".to_string(),
        reason: SessionRemovedReason::Replaced,
    };

    assert_eq!(event.reason, SessionRemovedReason::Replaced);
    assert_ne!(event.reason, SessionRemovedReason::Closed);
}

#[test]
fn heartbeat_decision_carries_close_code() {
    assert_eq!(HeartbeatDecision::KeepOpen, HeartbeatDecision::KeepOpen);
    assert_eq!(
        HeartbeatDecision::Close { code: 4002 },
        HeartbeatDecision::Close { code: 4002 }
    );
}
