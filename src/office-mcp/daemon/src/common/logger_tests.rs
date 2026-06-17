use super::{LogLevel, LogRecord, Logger};
use std::fs::{read_to_string, remove_dir_all};
use std::time::SystemTime;

#[test]
fn filters_records_below_configured_level() {
    let logger = Logger::with_file(
        LogLevel::Warn,
        std::env::temp_dir().join("unused-office-mcp.log"),
    );

    assert_eq!(logger.info("hidden info").expect("info"), None);
    let warning = logger.warn("visible warning").expect("warn");

    assert!(warning.expect("line").contains("visible warning"));
}

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
fn logger_appends_jsonl_file() {
    let dir = std::env::temp_dir().join(format!("office-mcp-log-rust-{}", std::process::id()));
    let path = dir.join("office-mcp.log");
    let logger = Logger::with_file(LogLevel::Info, &path);

    logger
        .write(
            &LogRecord::new(LogLevel::Info, "service started")
                .at(SystemTime::UNIX_EPOCH)
                .with_field("component", "daemon"),
        )
        .expect("write log");

    let contents = read_to_string(&path).expect("read log file");
    assert!(contents.contains("service started"));
    assert!(contents.contains("daemon"));
    let _ = remove_dir_all(dir);
}

#[test]
fn tracing_file_subscriber_writes_json_events_with_level_filter() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-tracing-log-rust-{}",
        std::process::id()
    ));
    let path = dir.join("office-mcp-tracing.log");

    let guard = Logger::init_tracing_file(LogLevel::Warn, &path).expect("init tracing");
    tracing::info!(component = "daemon", "hidden info");
    tracing::warn!(component = "daemon", "visible warning");
    drop(guard);

    let contents = read_to_string(&path).expect("read tracing log file");
    assert!(contents.contains("visible warning"));
    assert!(contents.contains("\"level\":\"WARN\""));
    assert!(!contents.contains("hidden info"));
    let _ = remove_dir_all(dir);
}
