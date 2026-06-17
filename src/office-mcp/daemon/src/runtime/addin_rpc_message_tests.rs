use super::AddinRpcMessage;
use crate::addin_mgr::{JsonRpcEnvelope, JsonRpcId, RegisterResult, SessionRemovedReason};
use serde_json::json;

#[test]
fn parses_register_request_from_json_params() {
    let request = AddinRpcMessage::register_request(
        &json!({
            "instance_id": "instance-1",
            "host": { "app": "word", "version": "16.0", "platform": "windows" },
            "add_in": {
                "version": "0.1.0",
                "protocol_version": "1.0",
                "supported_features": ["doc.read"]
            }
        }),
        JsonRpcId::String("register-1".to_string()),
    );

    assert_eq!(request.instance_id, "instance-1");
    assert_eq!(request.host.app, "word");
    assert_eq!(request.host.platform.as_deref(), Some("windows"));
    assert_eq!(request.add_in.supported_features, vec!["doc.read"]);
}

#[test]
fn parses_session_events_from_json_params() {
    let added = AddinRpcMessage::session_added_event(&json!({
        "session_id": "session-1",
        "instance_id": "instance-1",
        "document": { "filename": "Draft.docx", "is_read_only": false },
        "available_tools": ["word.get_text"],
        "is_active": true
    }));
    assert_eq!(added.session_id, "session-1");
    assert_eq!(added.document.filename.as_deref(), Some("Draft.docx"));
    assert_eq!(added.available_tools, vec!["word.get_text"]);
    assert_eq!(added.is_active, Some(true));

    let updated = AddinRpcMessage::session_updated_event(&json!({
        "session_id": "session-1",
        "patch": {
            "document": { "title": "Final" },
            "available_tools": ["word.add_comment"],
            "is_active": false
        }
    }));
    assert_eq!(
        updated.patch.document.expect("document").title.as_deref(),
        Some("Final")
    );
    assert_eq!(
        updated.patch.available_tools,
        Some(vec!["word.add_comment".to_string()])
    );
    assert_eq!(updated.patch.is_active, Some(Some(false)));

    let removed = AddinRpcMessage::session_removed_event(&json!({
        "session_id": "session-1",
        "reason": "closed"
    }));
    assert_eq!(removed.reason, SessionRemovedReason::Closed);
}

#[test]
fn register_reply_json_contains_only_public_runtime_fields() {
    let json = AddinRpcMessage::register_reply_to_json(JsonRpcEnvelope::success(
        JsonRpcId::String("register-1".to_string()),
        RegisterResult {
            server_version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            session_grace_sec: 60,
            heartbeat_interval_sec: 30,
            max_pending_per_session: 4,
            assigned_instance_id: "instance-1".to_string(),
        },
    ));

    assert!(json.contains("assigned_instance_id"));
    let lowercase = json.to_lowercase();
    for forbidden in ["api_key", "apikey", "secret", "token", "bearer", "pairing"] {
        assert!(!lowercase.contains(forbidden));
    }
}
