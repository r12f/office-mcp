use super::{LogLevel, LogRecord, append_line};
use std::fs::{read_to_string, remove_dir_all};
use std::time::SystemTime;

#[test]
fn log_record_redacts_sensitive_fields() {
    let line = LogRecord::new(
        LogLevel::Warn,
        "Bearer abc token=secret base64,QUJDREVGRw==",
    )
    .at(SystemTime::UNIX_EPOCH)
    .with_field("event", "warn_event")
    .with_field("certificate_passphrase", "secret-value")
    .to_json_line();

    assert!(line.contains("\"level\":\"warn\""));
    assert!(line.contains("warn_event"));
    assert!(line.contains("Bearer [redacted]"));
    assert!(line.contains("token=[redacted]"));
    assert!(line.contains("base64,[redacted]"));
    assert!(!line.contains("secret-value"));
    assert!(!line.contains("QUJDREVGRw"));
}

#[test]
fn append_line_writes_jsonl_file() {
    let dir =
        std::env::temp_dir().join(format!("office-mcp-log-record-rust-{}", std::process::id()));
    let path = dir.join("office-mcp.log");
    let line = LogRecord::new(LogLevel::Info, "service started")
        .at(SystemTime::UNIX_EPOCH)
        .with_field("component", "daemon")
        .to_json_line();

    append_line(&path, &line).expect("write log");

    let contents = read_to_string(&path).expect("read log file");
    assert!(contents.contains("service started"));
    assert!(contents.contains("daemon"));
    let _ = remove_dir_all(dir);
}
