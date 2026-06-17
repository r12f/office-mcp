use super::McpToolAuditRecorder;
use crate::addin_mgr::{CommandRouterError, PartialEffect, ToolResponse};
use crate::api::CommandFailure;
use crate::common::AuditLog;
use std::time::{Duration, SystemTime};

#[test]
fn records_success_without_document_payload() {
    let (_dir, path) = temp_audit_path("success");
    let audit_log = AuditLog::enabled(&path);
    let started = SystemTime::UNIX_EPOCH + Duration::from_secs(1);
    let completed = started + Duration::from_millis(42);

    McpToolAuditRecorder::record_completed(
        &audit_log,
        "word.get_text",
        "session-1",
        &Ok(ToolResponse::Success {
            json: "{\"text\":\"document body\"}".to_string(),
        }),
        started,
        completed,
    );

    let contents = std::fs::read_to_string(path).expect("audit file");
    assert!(contents.contains("\"tool\":\"word.get_text\""));
    assert!(contents.contains("\"session_id\":\"session-1\""));
    assert!(contents.contains("\"duration_ms\":42"));
    assert!(contents.contains("\"ok\":true"));
    assert!(!contents.contains("document body"));
}

#[test]
fn records_command_failure_with_redacted_message() {
    let (_dir, path) = temp_audit_path("failure");
    let audit_log = AuditLog::enabled(&path);
    let failure = CommandFailure {
        office_mcp_code: "HOST_ERROR".to_string(),
        message: "text=secret body=hidden detail".to_string(),
        tool: Some("word.insert_paragraph".to_string()),
        retriable: false,
        partial_effect: Some(PartialEffect::None),
    };

    McpToolAuditRecorder::record_failure(
        &audit_log,
        "word.insert_paragraph",
        "session-1",
        &failure,
        SystemTime::UNIX_EPOCH,
        SystemTime::UNIX_EPOCH + Duration::from_millis(7),
    );

    let contents = std::fs::read_to_string(path).expect("audit file");
    assert!(contents.contains("HOST_ERROR"));
    assert!(contents.contains("text=[redacted]"));
    assert!(contents.contains("body=[redacted]"));
    assert!(!contents.contains("secret"));
    assert!(!contents.contains("hidden"));
}

#[test]
fn records_router_error_as_command_failure() {
    let (_dir, path) = temp_audit_path("router-error");
    let audit_log = AuditLog::enabled(&path);
    let completed = Err(CommandRouterError::ResponseTooLarge {
        max_response_bytes: 128,
        actual_bytes: 256,
    });

    McpToolAuditRecorder::record_completed(
        &audit_log,
        "word.get_text",
        "session-1",
        &completed,
        SystemTime::UNIX_EPOCH,
        SystemTime::UNIX_EPOCH + Duration::from_millis(9),
    );

    let contents = std::fs::read_to_string(path).expect("audit file");
    assert!(contents.contains("MAX_RESPONSE_SIZE"));
    assert!(contents.contains("\"ok\":false"));
}

fn temp_audit_path(label: &str) -> (std::path::PathBuf, std::path::PathBuf) {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-tool-audit-{label}-{}-{:?}",
        std::process::id(),
        std::thread::current().id()
    ));
    let _ = std::fs::remove_dir_all(&dir);
    let path = dir.join("audit.jsonl");
    (dir, path)
}
