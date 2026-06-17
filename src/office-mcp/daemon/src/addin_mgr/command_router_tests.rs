use super::{CommandRouter, CommandRouterError, ToolCallRequest, ToolResponse};
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, OfficeMcpCode, RuntimeInfo, SessionRegistry,
};
use crate::api::{RegisterClientInput, UiClientTransport, UiCommandStatus, UiStateStore};
use std::time::{Duration, SystemTime};

#[test]
fn enqueue_starts_ui_task_and_serializes_per_session() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let registry = registry_with_session(now);
    let mut ui_state = UiStateStore::new();
    let client_id = ui_state.register_client(RegisterClientInput {
        client_id: Some("client-1".to_string()),
        transport: UiClientTransport::Http,
        name: Some("copilot-cli/1.0".to_string()),
    });
    let mut router = CommandRouter::new();

    let first = router
        .enqueue(
            &registry,
            &mut ui_state,
            request("session-1", "word.get_text", Some(client_id.clone())),
            now,
        )
        .expect("enqueue first");
    let second = router
        .enqueue(
            &registry,
            &mut ui_state,
            request("session-1", "word.add_comment", Some(client_id)),
            now,
        )
        .expect("enqueue second");

    assert_eq!(first.sequence, 0);
    assert_eq!(second.sequence, 1);
    assert_eq!(router.queue_depth("session-1"), 2);
    assert_eq!(router.queue_depth("word"), 0);
    let snapshot = ui_state.snapshot(&[], now);
    assert_eq!(snapshot.current_tasks.len(), 2);
    assert_eq!(snapshot.clients[0].in_flight_request_count, 2);
}

#[test]
fn preflight_errors_do_not_start_ui_task() {
    let registry = SessionRegistry::new();
    let mut ui_state = UiStateStore::new();
    let mut router = CommandRouter::new();

    let error = router
        .enqueue(
            &registry,
            &mut ui_state,
            request("missing", "word.get_text", None),
            SystemTime::UNIX_EPOCH,
        )
        .expect_err("preflight error");

    assert!(matches!(
        error,
        CommandRouterError::Preflight(error)
            if error.failure.office_mcp_code == OfficeMcpCode::NoSessions
    ));
    assert!(
        ui_state
            .snapshot(&[], SystemTime::UNIX_EPOCH)
            .current_tasks
            .is_empty()
    );
}

#[test]
fn complete_records_success_and_removes_queue_entry() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let registry = registry_with_session(now);
    let mut ui_state = UiStateStore::new();
    let mut router = CommandRouter::new();
    let queued = router
        .enqueue(
            &registry,
            &mut ui_state,
            request("session-1", "word.get_text", None),
            now,
        )
        .expect("enqueue");

    router
        .complete(
            &mut ui_state,
            "session-1",
            &queued.request_id,
            ToolResponse::Success {
                json: "{\"ok\":true}".to_string(),
            },
            now + Duration::from_secs(1),
        )
        .expect("complete");

    let snapshot = ui_state.snapshot(&[], now + Duration::from_secs(1));
    assert_eq!(router.queue_depth("session-1"), 0);
    assert!(snapshot.current_tasks.is_empty());
    assert_eq!(snapshot.recent_commands[0].status, UiCommandStatus::Success);
}

#[test]
fn oversized_response_returns_max_response_size() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let registry = registry_with_session(now);
    let mut ui_state = UiStateStore::new();
    let mut router = CommandRouter::with_limits(4, Duration::from_secs(30));
    let queued = router
        .enqueue(
            &registry,
            &mut ui_state,
            request("session-1", "word.get_text", None),
            now,
        )
        .expect("enqueue");

    let error = router
        .complete(
            &mut ui_state,
            "session-1",
            &queued.request_id,
            ToolResponse::Success {
                json: "too large".to_string(),
            },
            now,
        )
        .expect_err("too large");
    let failure = error.as_command_failure("word.get_text");

    assert_eq!(failure.office_mcp_code, "MAX_RESPONSE_SIZE");
}

#[test]
fn timeout_expires_command_and_returns_cancel_message() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let registry = registry_with_session(now);
    let mut ui_state = UiStateStore::new();
    let mut router = CommandRouter::new();
    let queued = router
        .enqueue(
            &registry,
            &mut ui_state,
            ToolCallRequest {
                timeout: Some(Duration::from_millis(5)),
                ..request("session-1", "word.get_text", None)
            },
            now,
        )
        .expect("enqueue");

    let cancels = router.expire_timeouts(&mut ui_state, now + Duration::from_millis(6));

    assert_eq!(cancels.len(), 1);
    assert_eq!(cancels[0].request_id, queued.request_id);
    assert_eq!(cancels[0].reason, "deadline_expired");
    assert_eq!(
        ui_state.snapshot(&[], now).recent_commands[0].status,
        UiCommandStatus::Timeout
    );
}

fn registry_with_session(now: SystemTime) -> SessionRegistry {
    let mut registry = SessionRegistry::new();
    registry.register_runtime(RuntimeInfo {
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
        registered_at: now,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: "session-1".to_string(),
            instance_id: "instance-1".to_string(),
            document: DocumentInfo::default(),
            available_tools: vec!["word.get_text".to_string(), "word.add_comment".to_string()],
            is_active: Some(true),
        },
        now,
    );
    registry
}

fn request(session_id: &str, tool: &str, client_id: Option<String>) -> ToolCallRequest {
    ToolCallRequest {
        request_id: None,
        command_id: None,
        client_id,
        client_name: Some("copilot-cli/1.0".to_string()),
        session_id: session_id.to_string(),
        tool: tool.to_string(),
        arguments_json: "{}".to_string(),
        user_intent: Some("read current selection".to_string()),
        timeout: None,
        check_capability: true,
    }
}
