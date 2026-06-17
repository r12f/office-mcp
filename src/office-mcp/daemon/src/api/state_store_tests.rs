use super::{
    CommandFailure, CommandResult, RegisterClientInput, StartCommandInput, UiClientTransport,
    UiCommandStatus, UiStateOptions, UiStateStore,
};
use crate::addin_mgr::{
    DocumentDescriptor, HostDescriptor, PartialEffect, SessionDescriptor, SessionStatus,
};
use std::time::{Duration, SystemTime};

#[test]
fn snapshots_redact_sensitive_text_and_cap_history() {
    let mut store = UiStateStore::with_options(options());
    let session_id = "11111111-1111-4111-8111-111111111111";

    for index in 0..12 {
        let command_id = store.start_command(StartCommandInput {
            command_id: Some(format!("command-{index}")),
            session_id: Some(session_id.to_string()),
            tool: "word.insert_paragraph".to_string(),
            user_intent: Some(format!("token=secret-{index} insert private body text")),
            started_at: Some(SystemTime::UNIX_EPOCH + Duration::from_secs(index)),
            ..start_input()
        });
        store.finish_command(
            &command_id,
            CommandResult::Failure(CommandFailure {
                office_mcp_code: "IRM_DENIED".to_string(),
                message: format!("certificate_passphrase=secret-{index} Word denied the edit."),
                tool: Some("word.insert_paragraph".to_string()),
                retriable: false,
                partial_effect: Some(PartialEffect::None),
            }),
            SystemTime::UNIX_EPOCH + Duration::from_secs(index + 1),
        );
    }

    let snapshot = store.snapshot(&[], SystemTime::UNIX_EPOCH + Duration::from_secs(20));

    assert_eq!(snapshot.recent_commands.len(), 10);
    assert_eq!(snapshot.document_command_history[session_id].len(), 10);
    assert!(snapshot.current_tasks.is_empty());
    let debug = format!("{snapshot:?}");
    assert!(!debug.contains("secret-11"));
    assert!(!debug.contains("certificate_passphrase=secret"));
    assert!(debug.contains("certificate_passphrase=[redacted]"));
    assert!(debug.contains("token=[redacted]"));
}

#[test]
fn snapshots_group_document_metadata_for_console() {
    let store = UiStateStore::with_options(options());
    let session = descriptor("excel", "Budget.xlsx");

    let snapshot = store.snapshot(&[session], SystemTime::UNIX_EPOCH + Duration::from_secs(10));

    assert_eq!(snapshot.documents["excel"].len(), 1);
    assert_eq!(
        snapshot.documents["excel"][0].host.version.as_deref(),
        Some("16.0")
    );
    assert_eq!(
        snapshot.documents["excel"][0].document.is_read_only,
        Some(false)
    );
    assert_eq!(snapshot.documents["excel"][0].queue_depth, 1);
}

#[test]
fn tracks_clients_and_in_flight_request_counts() {
    let mut store = UiStateStore::with_options(options());
    let client_id = store.register_client(RegisterClientInput {
        client_id: None,
        transport: UiClientTransport::Http,
        name: Some("test-client".to_string()),
    });
    let command_id = store.start_command(StartCommandInput {
        client_id: Some(client_id.clone()),
        client_name: Some("test-client".to_string()),
        tool: "word.get_text".to_string(),
        ..start_input()
    });

    let snapshot = store.snapshot(&[], SystemTime::UNIX_EPOCH);
    assert_eq!(snapshot.clients.len(), 1);
    assert_eq!(snapshot.clients[0].name.as_deref(), Some("test-client"));
    assert_eq!(snapshot.clients[0].in_flight_request_count, 1);

    store.finish_command(
        &command_id,
        CommandResult::Success,
        SystemTime::UNIX_EPOCH + Duration::from_secs(1),
    );
    assert_eq!(
        store.snapshot(&[], SystemTime::UNIX_EPOCH).clients[0].in_flight_request_count,
        0
    );
    assert!(store.unregister_client(&client_id));
    assert!(
        store
            .snapshot(&[], SystemTime::UNIX_EPOCH)
            .clients
            .is_empty()
    );
}

#[test]
fn maps_timeout_cancelled_and_thrown_statuses() {
    let mut store = UiStateStore::with_options(options());
    let timeout = store.start_command(StartCommandInput {
        command_id: Some("timeout".to_string()),
        tool: "word.get_text".to_string(),
        ..start_input()
    });
    store.finish_command(
        &timeout,
        CommandResult::Failure(CommandFailure {
            office_mcp_code: "TIMEOUT".to_string(),
            message: "Bearer abc timeout".to_string(),
            tool: Some("word.get_text".to_string()),
            retriable: true,
            partial_effect: None,
        }),
        SystemTime::UNIX_EPOCH + Duration::from_secs(1),
    );

    let cancelled = store.start_command(StartCommandInput {
        command_id: Some("cancelled".to_string()),
        tool: "word.get_text".to_string(),
        ..start_input()
    });
    store.finish_command(
        &cancelled,
        CommandResult::Failure(CommandFailure {
            office_mcp_code: "CANCELLED".to_string(),
            message: "cancelled".to_string(),
            tool: None,
            retriable: false,
            partial_effect: None,
        }),
        SystemTime::UNIX_EPOCH + Duration::from_secs(1),
    );

    let thrown = store.start_command(StartCommandInput {
        command_id: Some("thrown".to_string()),
        tool: "word.get_text".to_string(),
        ..start_input()
    });
    store.finish_command(
        &thrown,
        CommandResult::Thrown("password=hunter2 failed".to_string()),
        SystemTime::UNIX_EPOCH + Duration::from_secs(1),
    );

    let snapshot = store.snapshot(&[], SystemTime::UNIX_EPOCH + Duration::from_secs(2));
    assert_eq!(snapshot.recent_commands[2].status, UiCommandStatus::Timeout);
    assert_eq!(
        snapshot.recent_commands[1].status,
        UiCommandStatus::Cancelled
    );
    assert_eq!(snapshot.recent_commands[0].status, UiCommandStatus::Failure);
    assert!(format!("{snapshot:?}").contains("password=[redacted]"));
}

fn options() -> UiStateOptions {
    UiStateOptions {
        version: "0.1.0".to_string(),
        mcp_endpoint: "http://127.0.0.1:8800/mcp".to_string(),
        addin_endpoint: "https://localhost:8765/addin".to_string(),
        config_path: None,
        log_path: None,
        now: SystemTime::UNIX_EPOCH,
    }
}

fn start_input() -> StartCommandInput {
    StartCommandInput {
        command_id: None,
        mcp_request_id: None,
        client_id: None,
        client_name: None,
        session_id: None,
        host_app: None,
        tool: "word.get_text".to_string(),
        user_intent: None,
        timeout_ms: None,
        started_at: None,
    }
}

fn descriptor(app: &str, filename: &str) -> SessionDescriptor {
    SessionDescriptor {
        session_id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa".to_string(),
        instance_id: "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb".to_string(),
        app: app.to_string(),
        host: HostDescriptor {
            app: app.to_string(),
            version: Some("16.0".to_string()),
            platform: Some("pc".to_string()),
            build: Some("Desktop".to_string()),
        },
        document: DocumentDescriptor {
            title: Some(filename.to_string()),
            url: None,
            filename: Some(filename.to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(false),
            protection_kind: None,
            rights: None,
            rights_source: Some("unavailable".to_string()),
        },
        is_active: Some(true),
        capability_tiers: vec!["core".to_string()],
        available_tool_count: 3,
        queue_depth: 1,
        registered_at: SystemTime::UNIX_EPOCH,
        status: SessionStatus::Active,
    }
}
