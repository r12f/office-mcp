use super::{AuditLog, AuditRecord};
use std::fs::{read_to_string, remove_dir_all};
use std::time::{Duration, SystemTime};

#[test]
fn disabled_audit_log_does_not_write() {
    let log = AuditLog::new();

    let line = log
        .record(&AuditRecord::success(
            SystemTime::UNIX_EPOCH,
            "word.get_text",
            None,
            1,
        ))
        .expect("record");

    assert_eq!(line, None);
}

#[test]
fn audit_success_line_excludes_document_payload() {
    let line = AuditRecord::success(
        SystemTime::UNIX_EPOCH + Duration::from_secs(1),
        "word.get_text",
        Some("session-1"),
        12,
    )
    .to_json_line();

    assert!(line.contains("\"tool\":\"word.get_text\""));
    assert!(line.contains("\"session_id\":\"session-1\""));
    assert!(line.contains("\"ok\":true"));
    assert!(!line.contains("document body"));
}

#[test]
fn audit_failure_redacts_body_like_error_text() {
    let line = AuditRecord::failure(
        SystemTime::UNIX_EPOCH,
        "word.insert_paragraph",
        Some("session-1"),
        9,
        "HOST_CAPABILITY_UNAVAILABLE",
        "text=secret-body failed",
    )
    .to_json_line();

    assert!(line.contains("HOST_CAPABILITY_UNAVAILABLE"));
    assert!(line.contains("text=[redacted]"));
    assert!(!line.contains("secret-body"));
}

#[test]
fn enabled_audit_log_appends_jsonl() {
    let dir = std::env::temp_dir().join(format!("office-mcp-audit-rust-{}", std::process::id()));
    let path = dir.join("audit.jsonl");
    let log = AuditLog::enabled(&path);

    log.record(&AuditRecord::success(
        SystemTime::UNIX_EPOCH,
        "office.list_sessions",
        None,
        3,
    ))
    .expect("write audit");

    let contents = read_to_string(&path).expect("read audit file");
    assert!(contents.contains("office.list_sessions"));
    let _ = remove_dir_all(dir);
}
