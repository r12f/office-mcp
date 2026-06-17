use super::AddinSessionEventHandler;
use crate::addin_mgr::{
    AddinChannelError, AddinConnectionState, DocumentInfo, SessionAddedEvent, SessionPatch,
    SessionRegistry, SessionRemovedEvent, SessionRemovedReason, SessionUpdatedEvent,
};
use std::collections::BTreeMap;
use std::time::{Duration, SystemTime};

#[test]
fn adds_session_and_binds_connection() {
    let mut registry = SessionRegistry::new();
    let mut connections = connections();
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);

    AddinSessionEventHandler::add_session(
        &mut registry,
        &mut connections,
        "connection-1",
        session_added("session-1", "instance-1"),
        now,
    )
    .expect("session added");

    assert_eq!(
        connections
            .get("connection-1")
            .and_then(|connection| connection.session_id.as_deref()),
        Some("session-1")
    );
    assert_eq!(registry.list_sessions()[0].session_id, "session-1");
}

#[test]
fn rejects_wrong_instance_without_registering_session() {
    let mut registry = SessionRegistry::new();
    let mut connections = connections();

    let error = AddinSessionEventHandler::add_session(
        &mut registry,
        &mut connections,
        "connection-1",
        session_added("session-1", "different"),
        SystemTime::UNIX_EPOCH,
    )
    .expect_err("instance mismatch");

    assert_eq!(
        error,
        AddinChannelError::InstanceMismatch {
            expected: "instance-1".to_string(),
            actual: "different".to_string(),
        }
    );
    assert!(registry.list_sessions().is_empty());
}

#[test]
fn updates_existing_session_metadata() {
    let mut registry = SessionRegistry::new();
    let mut connections = connections();
    AddinSessionEventHandler::add_session(
        &mut registry,
        &mut connections,
        "connection-1",
        session_added("session-1", "instance-1"),
        SystemTime::UNIX_EPOCH,
    )
    .expect("session added");

    AddinSessionEventHandler::update_session(
        &mut registry,
        SessionUpdatedEvent {
            session_id: "session-1".to_string(),
            patch: SessionPatch {
                document: Some(DocumentInfo {
                    title: Some("Updated".to_string()),
                    ..DocumentInfo::default()
                }),
                ..SessionPatch::default()
            },
        },
    )
    .expect("session updated");

    assert_eq!(
        registry
            .get_session_info("session-1")
            .expect("session")
            .descriptor
            .document
            .title
            .as_deref(),
        Some("Updated")
    );
}

#[test]
fn removes_session_and_unbinds_connections() {
    let mut registry = SessionRegistry::new();
    let mut connections = connections();
    AddinSessionEventHandler::add_session(
        &mut registry,
        &mut connections,
        "connection-1",
        session_added("session-1", "instance-1"),
        SystemTime::UNIX_EPOCH,
    )
    .expect("session added");

    AddinSessionEventHandler::remove_session(
        &mut registry,
        &mut connections,
        SessionRemovedEvent {
            session_id: "session-1".to_string(),
            reason: SessionRemovedReason::Closed,
        },
    )
    .expect("session removed");

    assert!(registry.list_sessions().is_empty());
    assert_eq!(
        connections
            .get("connection-1")
            .and_then(|connection| connection.session_id.as_deref()),
        None
    );
}

fn connections() -> BTreeMap<String, AddinConnectionState> {
    BTreeMap::from([(
        "connection-1".to_string(),
        AddinConnectionState::new("instance-1".to_string()),
    )])
}

fn session_added(session_id: &str, instance_id: &str) -> SessionAddedEvent {
    SessionAddedEvent {
        session_id: session_id.to_string(),
        instance_id: instance_id.to_string(),
        document: DocumentInfo::default(),
        available_tools: vec!["word.get_text".to_string()],
        is_active: Some(true),
    }
}
