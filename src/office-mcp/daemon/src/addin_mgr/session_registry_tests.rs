use super::SessionRegistry;
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, OfficeMcpCode, RuntimeInfo, SessionPatch,
    SessionStatus,
};
use std::time::{Duration, SystemTime};

#[test]
fn registers_runtime_and_describes_session_metadata() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(runtime("instance-1", "word", now));
    registry.add_session(session("session-1", "instance-1", "Doc.docx"), now);
    registry.set_connection_pending("instance-1", 2);

    let sessions = registry.list_sessions();

    assert_eq!(sessions.len(), 1);
    assert_eq!(sessions[0].app, "word");
    assert_eq!(sessions[0].document.title.as_deref(), Some("Doc.docx"));
    assert_eq!(sessions[0].available_tool_count, 3);
    assert_eq!(sessions[0].queue_depth, 2);
    assert_eq!(
        registry
            .get_session_info("session-1")
            .expect("session info")
            .available_tools,
        vec!["word.get_text", "word.add_comment", "word.accept_change"]
    );
    assert_eq!(
        sessions[0].capability_tiers,
        ["core", "review", "tracked_changes"]
    );
}

#[test]
fn stale_sessions_are_pruned_after_grace_period() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(runtime("instance-1", "word", now));
    registry.add_session(session("session-1", "instance-1", "Doc.docx"), now);

    assert!(registry.remove_runtime("instance-1", now + Duration::from_secs(10)));
    assert_eq!(registry.list_sessions()[0].status, SessionStatus::Stale);
    assert_eq!(
        registry.prune_stale_sessions(now + Duration::from_secs(20), Duration::from_mins(1)),
        0
    );
    assert_eq!(
        registry.prune_stale_sessions(now + Duration::from_secs(80), Duration::from_mins(1)),
        1
    );
    assert!(registry.list_sessions().is_empty());
}

#[test]
fn invocation_preflight_returns_protocol_errors() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
    let mut registry = SessionRegistry::with_limits(1);

    let error = registry
        .prepare_invocation("missing", "word.get_text", true)
        .expect_err("no sessions");
    assert_eq!(error.failure.office_mcp_code, OfficeMcpCode::NoSessions);
    assert_eq!(error.failure.office_mcp_code.as_str(), "NO_SESSIONS");
    assert!(error.failure.retriable);

    registry.register_runtime(runtime("instance-1", "word", now));
    registry.add_session(session("session-1", "instance-1", "Doc.docx"), now);

    let error = registry
        .prepare_invocation("session-1", "word.unsupported", true)
        .expect_err("missing capability");
    assert_eq!(
        error.failure.office_mcp_code,
        OfficeMcpCode::ToolNotEnabledForDocument
    );
    assert_eq!(
        error.failure.office_mcp_code.as_str(),
        "TOOL_NOT_ENABLED_FOR_DOCUMENT"
    );
    assert_eq!(
        error.failure.message,
        "Tool word.unsupported is disabled for this document session. Refresh office.get_session_info or office.list_sessions before retrying."
    );

    let error = registry
        .prepare_invocation("missing-session", "word.get_text", true)
        .expect_err("unknown session");
    assert_eq!(
        error.failure.office_mcp_code,
        OfficeMcpCode::SessionNotFound
    );
    assert_eq!(error.failure.office_mcp_code.as_str(), "SESSION_NOT_FOUND");

    registry.set_connection_pending("instance-1", 1);
    let error = registry
        .prepare_invocation("session-1", "word.get_text", true)
        .expect_err("too many pending");
    assert_eq!(
        error.failure.office_mcp_code,
        OfficeMcpCode::MaxPendingExceeded
    );
}

#[test]
fn session_patch_updates_descriptor_fields() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(runtime("instance-1", "unknown-host", now));
    registry.add_session(session("session-1", "instance-1", "Draft.docx"), now);

    registry.update_session(
        "session-1",
        SessionPatch {
            document: Some(DocumentInfo {
                title: Some("Final".to_string()),
                filename: Some("Final.docx".to_string()),
                ..DocumentInfo::default()
            }),
            is_active: Some(Some(false)),
            ..SessionPatch::default()
        },
    );
    let session = registry
        .get_session_info("session-1")
        .expect("session details");

    assert_eq!(session.descriptor.app, "other");
    assert_eq!(session.descriptor.document.title.as_deref(), Some("Final"));
    assert_eq!(session.descriptor.is_active, Some(false));
}

fn runtime(instance_id: &str, app: &str, registered_at: SystemTime) -> RuntimeInfo {
    RuntimeInfo {
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: app.to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["doc.read".to_string()],
        },
        registered_at,
    }
}

fn session(session_id: &str, instance_id: &str, filename: &str) -> NewSessionInfo {
    NewSessionInfo {
        session_id: session_id.to_string(),
        instance_id: instance_id.to_string(),
        document: DocumentInfo {
            filename: Some(filename.to_string()),
            ..DocumentInfo::default()
        },
        available_tools: vec![
            "word.get_text".to_string(),
            "word.add_comment".to_string(),
            "word.accept_change".to_string(),
        ],
        is_active: Some(true),
    }
}
