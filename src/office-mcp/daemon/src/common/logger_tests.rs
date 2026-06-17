use super::{LogLevel, Logger};
use std::fs::{read_to_string, remove_dir_all};

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
fn tracing_file_subscriber_writes_json_events_with_level_filter() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-tracing-log-rust-{}",
        std::process::id()
    ));
    let path = dir.join("office-mcp-tracing.log");

    let (subscriber, guard) =
        Logger::tracing_file_default(LogLevel::Warn, &path).expect("init tracing");
    tracing::subscriber::with_default(subscriber, || {
        tracing::info!(component = "daemon", "hidden info");
        tracing::warn!(component = "daemon", "visible warning");
    });
    drop(guard);

    let contents = read_to_string(&path).expect("read tracing log file");
    assert!(contents.contains("visible warning"));
    assert!(contents.contains("\"level\":\"WARN\""));
    assert!(!contents.contains("hidden info"));
    let _ = remove_dir_all(dir);
}
