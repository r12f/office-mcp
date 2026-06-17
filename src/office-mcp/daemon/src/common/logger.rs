use std::collections::BTreeMap;
use std::fmt::{Display, Formatter};
use std::fs::{OpenOptions, create_dir_all};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Logger {
    min_level: LogLevel,
    file: Option<PathBuf>,
}

impl Logger {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            min_level: LogLevel::Info,
            file: None,
        }
    }

    #[must_use]
    pub fn with_file(min_level: LogLevel, file: impl Into<PathBuf>) -> Self {
        Self {
            min_level,
            file: Some(file.into()),
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns operational logs with boundary redaction"
    }

    /// Writes a structured log record when it passes the configured level filter.
    ///
    /// # Errors
    ///
    /// Returns an error when the log file directory cannot be created or the
    /// log record cannot be appended.
    pub fn write(&self, record: &LogRecord) -> Result<Option<String>, LoggerError> {
        if record.level < self.min_level {
            return Ok(None);
        }
        let line = record.to_json_line();
        if let Some(path) = &self.file {
            append_line(path, &line)?;
        }
        Ok(Some(line))
    }

    /// Writes a trace-level message.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured log file cannot be written.
    pub fn trace(&self, message: &str) -> Result<Option<String>, LoggerError> {
        self.write(&LogRecord::new(LogLevel::Trace, message))
    }

    /// Writes a debug-level message.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured log file cannot be written.
    pub fn debug(&self, message: &str) -> Result<Option<String>, LoggerError> {
        self.write(&LogRecord::new(LogLevel::Debug, message))
    }

    /// Writes an info-level message.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured log file cannot be written.
    pub fn info(&self, message: &str) -> Result<Option<String>, LoggerError> {
        self.write(&LogRecord::new(LogLevel::Info, message))
    }

    /// Writes a warning-level message.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured log file cannot be written.
    pub fn warn(&self, message: &str) -> Result<Option<String>, LoggerError> {
        self.write(&LogRecord::new(LogLevel::Warn, message))
    }

    /// Writes an error-level message.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured log file cannot be written.
    pub fn error(&self, message: &str) -> Result<Option<String>, LoggerError> {
        self.write(&LogRecord::new(LogLevel::Error, message))
    }
}

impl Default for Logger {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum LogLevel {
    Trace,
    Debug,
    Info,
    Warn,
    Error,
}

impl LogLevel {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Trace => "trace",
            Self::Debug => "debug",
            Self::Info => "info",
            Self::Warn => "warn",
            Self::Error => "error",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LogRecord {
    pub ts: SystemTime,
    pub level: LogLevel,
    pub message: String,
    pub fields: BTreeMap<String, String>,
}

impl LogRecord {
    #[must_use]
    pub fn new(level: LogLevel, message: &str) -> Self {
        Self {
            ts: SystemTime::now(),
            level,
            message: redact_log_text(message),
            fields: BTreeMap::new(),
        }
    }

    #[must_use]
    pub fn at(mut self, ts: SystemTime) -> Self {
        self.ts = ts;
        self
    }

    #[must_use]
    pub fn with_field(mut self, name: &str, value: &str) -> Self {
        let value = if is_sensitive_key(name) {
            "[redacted]".to_string()
        } else {
            redact_log_text(value)
        };
        self.fields.insert(name.to_string(), value);
        self
    }

    #[must_use]
    pub fn to_json_line(&self) -> String {
        let mut fields = vec![
            json_field("ts", &format_system_time(self.ts)),
            json_field("level", self.level.as_str()),
            json_field("message", &self.message),
        ];
        for (name, value) in &self.fields {
            fields.push(json_field(name, value));
        }
        format!("{{{}}}", fields.join(","))
    }
}

#[derive(Debug)]
pub enum LoggerError {
    Io(std::io::Error),
}

impl Display for LoggerError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
        }
    }
}

impl std::error::Error for LoggerError {}

impl From<std::io::Error> for LoggerError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

fn append_line(path: &Path, line: &str) -> Result<(), LoggerError> {
    if let Some(parent) = path.parent() {
        create_dir_all(parent)?;
    }
    let mut file = OpenOptions::new().create(true).append(true).open(path)?;
    writeln!(file, "{line}")?;
    Ok(())
}

fn redact_log_text(value: &str) -> String {
    let mut output = redact_bearer_tokens(value);
    output = redact_key_values(&output);
    output = redact_base64(&output);
    output.truncate(500);
    output
}

fn redact_bearer_tokens(value: &str) -> String {
    let parts = value.split_whitespace().collect::<Vec<_>>();
    let mut output = Vec::with_capacity(parts.len());
    let mut redact_next = false;
    for part in parts {
        if redact_next {
            output.push("[redacted]".to_string());
            redact_next = false;
            continue;
        }
        if part.eq_ignore_ascii_case("bearer") {
            output.push("Bearer".to_string());
            redact_next = true;
        } else {
            output.push(part.to_string());
        }
    }
    output.join(" ")
}

fn redact_key_values(value: &str) -> String {
    value
        .split_whitespace()
        .map(|part| {
            let Some((key, _value)) = part.split_once('=') else {
                return part.to_string();
            };
            if is_sensitive_key(key) {
                format!("{key}=[redacted]")
            } else {
                part.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn is_sensitive_key(key: &str) -> bool {
    let normalized = key.replace(['_', '-'], "").to_ascii_lowercase();
    matches!(normalized.as_str(), "password" | "passphrase" | "token")
        || normalized.ends_with("passphrase")
}

fn redact_base64(value: &str) -> String {
    let mut output = String::new();
    let mut rest = value;
    while let Some(index) = rest.find("base64,") {
        output.push_str(&rest[..index]);
        output.push_str("base64,[redacted]");
        let after = &rest[index + "base64,".len()..];
        let end = after
            .find(|character: char| character.is_whitespace())
            .unwrap_or(after.len());
        rest = &after[end..];
    }
    output.push_str(rest);
    output
}

fn json_field(name: &str, value: &str) -> String {
    format!("\"{}\":\"{}\"", escape_json(name), escape_json(value))
}

fn escape_json(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

fn format_system_time(value: SystemTime) -> String {
    let millis = value
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis();
    format!("unix_ms:{millis}")
}

#[cfg(test)]
mod tests {
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
}
