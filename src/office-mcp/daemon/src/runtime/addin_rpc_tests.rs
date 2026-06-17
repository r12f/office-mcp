use super::AddinJsonRpcRuntime;
use crate::addin_mgr::{AddinChannelConfig, AddinChannelServer, SessionRegistry};
use serde_json::Value;
use std::collections::BTreeSet;
use std::sync::{Arc, Mutex};

#[test]
fn register_and_session_added_update_registry() {
    let registry = Arc::new(Mutex::new(SessionRegistry::new()));
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::with_config(
        AddinChannelConfig::default(),
    )));

    let register_reply = addin_handle_text(
        r#"{"jsonrpc":"2.0","id":"register-1","method":"register","params":{"instance_id":"instance-1","host":{"app":"word","version":"16.0","platform":"windows"},"add_in":{"version":"0.1.0","protocol_version":"1.0","supported_features":["doc.read"]}}}"#,
        "connection-1",
        &registry,
        &addin_channel,
    )
    .expect("register reply");

    assert!(register_reply.contains("assigned_instance_id"));
    let register_json: Value = serde_json::from_str(&register_reply).expect("register reply json");
    let result = register_json
        .get("result")
        .and_then(Value::as_object)
        .expect("register result object");
    let fields = result.keys().cloned().collect::<BTreeSet<_>>();
    assert_eq!(
        fields,
        BTreeSet::from([
            "assigned_instance_id".to_string(),
            "heartbeat_interval_sec".to_string(),
            "max_pending_per_session".to_string(),
            "protocol_version".to_string(),
            "server_version".to_string(),
            "session_grace_sec".to_string(),
        ])
    );
    let register_reply_lower = register_reply.to_lowercase();
    for forbidden in ["api_key", "apikey", "secret", "token", "bearer", "pairing"] {
        assert!(!register_reply_lower.contains(forbidden));
    }
    assert!(
        registry
            .lock()
            .expect("registry")
            .list_sessions()
            .is_empty()
    );

    let session_reply = addin_handle_text(
        r#"{"jsonrpc":"2.0","method":"session.added","params":{"session_id":"session-1","instance_id":"instance-1","document":{"filename":"Draft.docx","is_read_only":false},"available_tools":["word.get_text"],"is_active":true}}"#,
        "connection-1",
        &registry,
        &addin_channel,
    );

    assert_eq!(session_reply, None);
    let sessions = registry.lock().expect("registry").list_sessions();
    assert_eq!(sessions.len(), 1);
    assert_eq!(sessions[0].document.title.as_deref(), Some("Draft.docx"));
    assert_eq!(sessions[0].available_tool_count, 1);
}

#[test]
fn session_updated_and_removed_update_registry() {
    let registry = Arc::new(Mutex::new(SessionRegistry::new()));
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::with_config(
        AddinChannelConfig::default(),
    )));

    addin_handle_text(
        r#"{"jsonrpc":"2.0","id":"register-1","method":"register","params":{"instance_id":"instance-1","host":{"app":"word","version":"16.0","platform":"windows"},"add_in":{"version":"0.1.0","protocol_version":"1.0","supported_features":["doc.read"]}}}"#,
        "connection-1",
        &registry,
        &addin_channel,
    )
    .expect("register reply");
    addin_handle_text(
        r#"{"jsonrpc":"2.0","method":"session.added","params":{"session_id":"session-1","instance_id":"instance-1","document":{"filename":"Draft.docx"},"available_tools":["word.get_text"],"is_active":true}}"#,
        "connection-1",
        &registry,
        &addin_channel,
    );

    let update_reply = addin_handle_text(
        r#"{"jsonrpc":"2.0","method":"session.updated","params":{"session_id":"session-1","patch":{"document":{"title":"Final","filename":"Final.docx","is_dirty":true},"available_tools":["word.get_text","word.add_comment"],"is_active":false}}}"#,
        "connection-1",
        &registry,
        &addin_channel,
    );

    assert_eq!(update_reply, None);
    let sessions = registry.lock().expect("registry").list_sessions();
    assert_eq!(sessions.len(), 1);
    assert_eq!(sessions[0].document.title.as_deref(), Some("Final"));
    assert_eq!(sessions[0].document.filename.as_deref(), Some("Final.docx"));
    assert_eq!(sessions[0].document.is_dirty, Some(true));
    assert_eq!(sessions[0].available_tool_count, 2);
    assert_eq!(sessions[0].is_active, Some(false));
    drop(sessions);

    let remove_reply = addin_handle_text(
        r#"{"jsonrpc":"2.0","method":"session.removed","params":{"session_id":"session-1","reason":"closed"}}"#,
        "connection-1",
        &registry,
        &addin_channel,
    );

    assert_eq!(remove_reply, None);
    assert!(
        registry
            .lock()
            .expect("registry")
            .list_sessions()
            .is_empty()
    );
}

fn addin_handle_text(
    text: &str,
    connection_id: &str,
    registry: &Arc<Mutex<SessionRegistry>>,
    addin_channel: &Arc<Mutex<AddinChannelServer>>,
) -> Option<String> {
    AddinJsonRpcRuntime::handle_text(
        text,
        connection_id,
        registry,
        addin_channel,
        &crate::addin_mgr::AddinConnectionHub::new(),
    )
}
