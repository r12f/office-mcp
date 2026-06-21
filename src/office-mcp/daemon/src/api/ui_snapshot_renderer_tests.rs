use super::UiSnapshotRenderer;
use crate::addin_mgr::{
    DocumentDescriptor, HostDescriptor, PartialEffect, SessionDescriptor, SessionStatus,
};
use crate::api::{
    UiClientRecord, UiClientTransport, UiCommandError, UiCommandRecord, UiCommandStatus,
    UiDaemonSnapshot, UiHealth, UiSnapshot,
};
use crate::mcp::{AccessMode, ToolAccessPolicy, WORD_V1_TOOLS};
use serde_json::Value;
use std::collections::BTreeMap;
use std::time::{Duration, SystemTime};

#[test]
fn renders_daemon_endpoints_and_grouped_documents() {
    let snapshot = UiSnapshot {
        daemon: daemon_snapshot(),
        clients: vec![client_record()],
        documents: BTreeMap::from([("word".to_string(), vec![descriptor()])]),
        current_tasks: vec![running_command()],
        recent_commands: Vec::new(),
        document_command_history: BTreeMap::new(),
    };

    let rendered = UiSnapshotRenderer::new().render_value(&snapshot);

    assert_eq!(rendered["daemon"]["status"], "degraded");
    assert_eq!(
        rendered["daemon"]["mcp_endpoint"],
        "http://127.0.0.1:8800/mcp"
    );
    assert_eq!(
        rendered["daemon"]["addin_endpoint"],
        "https://localhost:8765/addin"
    );
    assert_eq!(
        rendered["daemon"]["config_path"],
        "C:/office-mcp/config.toml"
    );
    assert_eq!(rendered["daemon"]["log_path"], "C:/office-mcp/daemon.jsonl");
    assert_eq!(
        rendered["daemon"]["tool_access_policy"]["access_mode"],
        "read"
    );
    assert_eq!(
        rendered["daemon"]["tool_access_policy"]["disabled_apps"][0],
        "powerpoint"
    );
    assert_eq!(
        rendered["daemon"]["tool_access_policy"]["disabled_categories"][0]["app"],
        "excel"
    );
    assert_eq!(
        rendered["daemon"]["tool_access_policy"]["disabled_categories"][0]["category"],
        "Range"
    );
    assert_eq!(
        rendered["daemon"]["tool_access_policy"]["disabled_tools"][0],
        "word.update_table"
    );
    assert_eq!(rendered["clients"][0]["transport"], "stdio-bridge");
    assert_eq!(rendered["clients"][0]["connected_at"], 1_000);
    assert_eq!(rendered["documents"]["word"][0]["registered_at"], "unix:3");
    assert_eq!(rendered["documents"]["word"][0]["status"], "active");
    assert_eq!(rendered["documents"]["word"][0]["available_tool_count"], 25);
    assert_eq!(rendered["current_tasks"][0]["status"], "running");
    assert_eq!(rendered["current_tasks"][0]["deadline_at"], 7_000);
}

#[test]
fn renders_command_history_errors_and_partial_effects() {
    let snapshot = UiSnapshot {
        daemon: daemon_snapshot(),
        clients: Vec::new(),
        documents: BTreeMap::new(),
        current_tasks: Vec::new(),
        recent_commands: vec![failed_command()],
        document_command_history: BTreeMap::from([(
            "session-1".to_string(),
            vec![failed_command()],
        )]),
    };

    let text = UiSnapshotRenderer::new().render_text(&snapshot);
    let rendered: Value = serde_json::from_str(&text).expect("valid UI snapshot JSON");

    assert_eq!(rendered["recent_commands"][0]["status"], "timeout");
    assert_eq!(rendered["recent_commands"][0]["completed_at"], 11_000);
    assert_eq!(rendered["recent_commands"][0]["elapsed_ms"], 10_000);
    assert_eq!(
        rendered["recent_commands"][0]["error"]["office_mcp_code"],
        "TIMEOUT"
    );
    assert_eq!(
        rendered["recent_commands"][0]["error"]["partial_effect"],
        "possible"
    );
    assert_eq!(
        rendered["document_command_history"]["session-1"][0]["error"]["message"],
        "Timed out"
    );
}

fn daemon_snapshot() -> UiDaemonSnapshot {
    UiDaemonSnapshot {
        status: UiHealth::Degraded,
        version: "0.1.0-test".to_string(),
        uptime_ms: 42,
        mcp_endpoint: "http://127.0.0.1:8800/mcp".to_string(),
        addin_endpoint: "https://localhost:8765/addin".to_string(),
        config_path: Some("C:/office-mcp/config.toml".to_string()),
        log_path: Some("C:/office-mcp/daemon.jsonl".to_string()),
        last_error: Some("last error".to_string()),
        tool_access_policy: ToolAccessPolicy::default()
            .with_access_mode(AccessMode::Read)
            .with_disabled_app("powerpoint")
            .with_disabled_category("excel", "Range")
            .with_disabled_tool("word.update_table")
            .snapshot(),
    }
}

fn client_record() -> UiClientRecord {
    UiClientRecord {
        client_id: "client-1".to_string(),
        transport: UiClientTransport::StdioBridge,
        name: Some("Codex".to_string()),
        connected_at: SystemTime::UNIX_EPOCH + Duration::from_secs(1),
        last_activity_at: SystemTime::UNIX_EPOCH + Duration::from_secs(2),
        in_flight_request_count: 3,
    }
}

fn running_command() -> UiCommandRecord {
    UiCommandRecord {
        command_id: "command-1".to_string(),
        mcp_request_id: Some("request-1".to_string()),
        client_id: Some("client-1".to_string()),
        client_name: Some("Codex".to_string()),
        session_id: Some("session-1".to_string()),
        host_app: Some("word".to_string()),
        tool: "word.insert_paragraph".to_string(),
        user_intent: Some("insert text".to_string()),
        status: UiCommandStatus::Running,
        started_at: SystemTime::UNIX_EPOCH + Duration::from_secs(4),
        deadline_at: Some(SystemTime::UNIX_EPOCH + Duration::from_secs(7)),
        timeout_ms: Some(3_000),
        completed_at: None,
        elapsed_ms: None,
        error: None,
    }
}

fn failed_command() -> UiCommandRecord {
    UiCommandRecord {
        command_id: "command-2".to_string(),
        status: UiCommandStatus::Timeout,
        completed_at: Some(SystemTime::UNIX_EPOCH + Duration::from_secs(11)),
        elapsed_ms: Some(10_000),
        error: Some(UiCommandError {
            office_mcp_code: "TIMEOUT".to_string(),
            message: "Timed out".to_string(),
            tool: Some("word.insert_paragraph".to_string()),
            retriable: true,
            partial_effect: Some(PartialEffect::Possible),
        }),
        ..running_command()
    }
}

fn descriptor() -> SessionDescriptor {
    SessionDescriptor {
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        app: "word".to_string(),
        host: HostDescriptor {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("pc".to_string()),
            build: Some("Desktop".to_string()),
        },
        document: DocumentDescriptor {
            title: Some("Doc.docx".to_string()),
            url: None,
            filename: Some("Doc.docx".to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(false),
            protection_kind: None,
            rights: None,
            rights_source: Some("unavailable".to_string()),
        },
        is_active: Some(true),
        capability_tiers: vec!["core".to_string()],
        available_tools: vec!["word.get_text".to_string()],
        available_tool_count: WORD_V1_TOOLS.len(),
        queue_depth: 0,
        registered_at: SystemTime::UNIX_EPOCH + Duration::from_secs(3),
        status: SessionStatus::Active,
    }
}
