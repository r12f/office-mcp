use super::{
    ADDIN_PROTOCOL_VERSION, AddinChannelError, AddinChannelServer, HeartbeatDecision,
    RegisterRequest, SessionAddedEvent, SessionRemovedEvent, SessionUpdatedEvent,
};
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, JsonRpcId, SessionPatch, SessionRegistry,
    SessionRemovedReason,
};
use crate::addin_mgr::{CancelCommand, QueuedCommand};
use crate::common::{Logger, LoggerLogLevel};
use std::fs::{read_to_string, remove_dir_all};
use std::thread::sleep;
use std::time::{Duration, Instant, SystemTime};

#[test]
fn validates_upgrade_path_and_origin() {
    let server = AddinChannelServer::new();

    assert!(
        server
            .validate_upgrade("/addin", Some("https://localhost:8765"))
            .is_ok()
    );
    assert!(matches!(
        server.validate_upgrade("/wrong", Some("https://localhost:8765")),
        Err(AddinChannelError::InvalidUpgradePath(_))
    ));
    assert!(matches!(
        server.validate_upgrade("/addin", Some("https://example.invalid")),
        Err(AddinChannelError::ForbiddenOrigin(_))
    ));
}

#[test]
fn register_runtime_returns_server_settings_and_updates_registry() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut server = AddinChannelServer::new();
    let mut registry = SessionRegistry::new();

    let reply = server
        .register_runtime(
            &mut registry,
            "connection-1",
            register_request("instance-1", ADDIN_PROTOCOL_VERSION),
            now,
        )
        .expect("register");
    let result = reply.result.expect("register result");

    assert_eq!(result.assigned_instance_id, "instance-1");
    assert_eq!(result.max_pending_per_session, 4);
    assert!(registry.list_sessions().is_empty());
}

#[test]
fn register_rejects_protocol_major_mismatch() {
    let mut server = AddinChannelServer::new();
    let mut registry = SessionRegistry::new();

    let error = server
        .register_runtime(
            &mut registry,
            "",
            register_request("instance-1", "2.0"),
            SystemTime::UNIX_EPOCH,
        )
        .expect_err("protocol mismatch");

    assert!(matches!(
        error,
        AddinChannelError::ProtocolVersionMismatch { .. }
    ));
}

#[test]
fn session_events_update_registry_with_instance_binding() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut server = AddinChannelServer::new();
    let mut registry = SessionRegistry::new();
    server
        .register_runtime(
            &mut registry,
            "connection-1",
            register_request("instance-1", ADDIN_PROTOCOL_VERSION),
            now,
        )
        .expect("register");

    server
        .add_session(
            &mut registry,
            "connection-1",
            SessionAddedEvent {
                session_id: "session-1".to_string(),
                instance_id: "instance-1".to_string(),
                document: DocumentInfo {
                    filename: Some("Draft.docx".to_string()),
                    ..DocumentInfo::default()
                },
                available_tools: vec!["word.get_text".to_string()],
                is_active: Some(true),
            },
            now,
        )
        .expect("session added");
    server
        .update_session(
            &mut registry,
            SessionUpdatedEvent {
                session_id: "session-1".to_string(),
                patch: SessionPatch {
                    document: Some(DocumentInfo {
                        title: Some("Final".to_string()),
                        ..DocumentInfo::default()
                    }),
                    ..SessionPatch::default()
                },
            },
        )
        .expect("session updated");

    let session = registry.get_session_info("session-1").expect("session");
    assert_eq!(session.descriptor.document.title.as_deref(), Some("Final"));

    server
        .remove_session(
            &mut registry,
            SessionRemovedEvent {
                session_id: "session-1".to_string(),
                reason: SessionRemovedReason::Closed,
            },
        )
        .expect("session removed");
    assert!(registry.list_sessions().is_empty());
}

#[test]
fn session_added_rejects_wrong_instance() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut server = AddinChannelServer::new();
    let mut registry = SessionRegistry::new();
    server
        .register_runtime(
            &mut registry,
            "connection-1",
            register_request("instance-1", ADDIN_PROTOCOL_VERSION),
            now,
        )
        .expect("register");

    let error = server
        .add_session(
            &mut registry,
            "connection-1",
            SessionAddedEvent {
                session_id: "session-1".to_string(),
                instance_id: "different".to_string(),
                document: DocumentInfo::default(),
                available_tools: Vec::new(),
                is_active: None,
            },
            now,
        )
        .expect_err("instance mismatch");

    assert!(matches!(error, AddinChannelError::InstanceMismatch { .. }));
}

#[test]
fn heartbeat_marks_session_stale_after_third_miss() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut server = AddinChannelServer::new();
    let mut registry = SessionRegistry::new();
    server
        .register_runtime(
            &mut registry,
            "connection-1",
            register_request("instance-1", ADDIN_PROTOCOL_VERSION),
            now,
        )
        .expect("register");
    server
        .add_session(
            &mut registry,
            "connection-1",
            SessionAddedEvent {
                session_id: "session-1".to_string(),
                instance_id: "instance-1".to_string(),
                document: DocumentInfo::default(),
                available_tools: vec!["word.get_text".to_string()],
                is_active: None,
            },
            now,
        )
        .expect("session");

    let ping = server.start_ping("connection-1", now).expect("ping");
    assert_eq!(ping.method.as_deref(), Some("ping"));
    assert_eq!(
        server.record_heartbeat_timeout(&mut registry, "connection-1", now),
        Ok(HeartbeatDecision::KeepOpen)
    );
    assert_eq!(
        server.record_heartbeat_timeout(&mut registry, "connection-1", now),
        Ok(HeartbeatDecision::KeepOpen)
    );
    assert_eq!(
        server.record_heartbeat_timeout(&mut registry, "connection-1", now),
        Ok(HeartbeatDecision::Close { code: 4002 })
    );

    assert_eq!(
        registry.list_sessions()[0].status,
        crate::addin_mgr::SessionStatus::Stale
    );
}

#[test]
fn builds_tool_invoke_and_cancel_payloads() {
    let server = AddinChannelServer::new();
    let command = QueuedCommand {
        command_id: "command-1".to_string(),
        request_id: "request-1".to_string(),
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        tool: "word.get_text".to_string(),
        arguments_json: "{}".to_string(),
        timeout: Duration::from_secs(30),
        enqueued_at: SystemTime::UNIX_EPOCH,
        deadline_at: SystemTime::UNIX_EPOCH + Duration::from_secs(30),
        sequence: 0,
        dispatched: false,
    };

    let invoke = server.tool_invoke_payload(&command);
    assert_eq!(invoke.method.as_deref(), Some("tool.invoke"));
    assert_eq!(invoke.params["session_id"], "session-1");
    assert_eq!(invoke.params["timeout_ms"], "30000");

    let cancel = server.tool_cancel_payload(&CancelCommand {
        request_id: "request-1".to_string(),
        reason: "timeout".to_string(),
    });
    assert_eq!(cancel.id, None);
    assert_eq!(cancel.method.as_deref(), Some("tool.cancel"));
    assert_eq!(cancel.params["reason"], "timeout");
}

#[test]
fn writes_structured_tracing_events_for_addin_session_lifecycle() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-addin-channel-log-{}-{}",
        std::process::id(),
        SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .expect("system clock after epoch")
            .as_nanos()
    ));
    let path = dir.join("office-mcp.log");
    let (subscriber, guard) =
        Logger::tracing_file_default(LoggerLogLevel::Debug, &path).expect("init tracing");
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut server = AddinChannelServer::new();
    let mut registry = SessionRegistry::new();

    tracing::subscriber::with_default(subscriber, || {
        server
            .validate_upgrade("/addin", Some("https://example.invalid"))
            .expect_err("origin rejected");
        server
            .register_runtime(
                &mut registry,
                "connection-1",
                register_request("instance-1", ADDIN_PROTOCOL_VERSION),
                now,
            )
            .expect("register");
        server
            .add_session(
                &mut registry,
                "connection-1",
                SessionAddedEvent {
                    session_id: "session-1".to_string(),
                    instance_id: "instance-1".to_string(),
                    document: DocumentInfo::default(),
                    available_tools: vec!["word.get_text".to_string()],
                    is_active: Some(true),
                },
                now,
            )
            .expect("session added");
        server
            .update_session(
                &mut registry,
                SessionUpdatedEvent {
                    session_id: "session-1".to_string(),
                    patch: SessionPatch {
                        is_active: Some(Some(false)),
                        ..SessionPatch::default()
                    },
                },
            )
            .expect("session updated");
        server
            .remove_session(
                &mut registry,
                SessionRemovedEvent {
                    session_id: "session-1".to_string(),
                    reason: SessionRemovedReason::Closed,
                },
            )
            .expect("session removed");
    });
    drop(guard);

    let expected = [
        "rejected add-in websocket origin",
        "registered add-in runtime",
        "added add-in document session",
        "updated add-in document session",
        "removed add-in document session",
        "\"component\":\"addin_channel\"",
        "instance-1",
        "session-1",
    ];
    let contents = wait_for_log_contents(&path, &expected);
    for expected_text in expected {
        assert!(
            contents.contains(expected_text),
            "missing {expected_text:?} in tracing log:\n{contents}"
        );
    }
    let _ = remove_dir_all(dir);
}

fn wait_for_log_contents(path: &std::path::Path, expected: &[&str]) -> String {
    let deadline = Instant::now() + Duration::from_secs(2);
    let mut contents = String::new();
    while Instant::now() < deadline {
        contents = read_to_string(path).expect("read tracing log file");
        if expected.iter().all(|value| contents.contains(value)) {
            return contents;
        }
        sleep(Duration::from_millis(10));
    }
    read_to_string(path).unwrap_or(contents)
}

fn register_request(instance_id: &str, protocol_version: &str) -> RegisterRequest {
    RegisterRequest {
        id: JsonRpcId::String("register-1".to_string()),
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: protocol_version.to_string(),
            supported_features: vec!["doc.read".to_string()],
        },
    }
}
