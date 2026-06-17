use crate::common::logger_redaction::{redact_log_field, redact_log_text};
use std::collections::BTreeMap;
use std::fmt::{Display, Formatter};
use std::fs::{OpenOptions, create_dir_all};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::SystemTime;
use tracing_appender::non_blocking::WorkerGuard;
use tracing_subscriber::EnvFilter;
use tracing_subscriber::fmt::MakeWriter;

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

    /// Installs the process-wide tracing subscriber and writes tracing events to a file.
    ///
    /// The returned guard must be kept alive for the lifetime of the daemon so
    /// buffered log records are flushed on shutdown.
    ///
    /// # Errors
    ///
    /// Returns an error when the log directory cannot be created or when a
    /// process-wide tracing subscriber has already been installed.
    pub fn init_tracing_file(
        min_level: LogLevel,
        file: impl AsRef<Path>,
    ) -> Result<TracingLogGuard, LoggerError> {
        let file = file.as_ref();
        if let Some(parent) = file.parent() {
            create_dir_all(parent)?;
        }
        let appender = OpenOptions::new().create(true).append(true).open(file)?;
        let (writer, guard) = tracing_appender::non_blocking(appender);
        let subscriber = tracing_file_subscriber(min_level, writer);
        tracing::subscriber::set_global_default(subscriber)
            .map_err(|error| LoggerError::TracingInit(error.to_string()))?;
        Ok(TracingLogGuard { _guard: guard })
    }

    /// Creates a tracing file subscriber and guard without installing it globally.
    ///
    /// This is intended for tests that need isolated tracing capture with
    /// `tracing::subscriber::with_default`.
    ///
    /// # Errors
    ///
    /// Returns an error when the log directory or file cannot be created.
    #[cfg(test)]
    pub(crate) fn tracing_file_default(
        min_level: LogLevel,
        file: impl AsRef<Path>,
    ) -> Result<(impl tracing::Subscriber + Send + Sync, TracingLogGuard), LoggerError> {
        let file = file.as_ref();
        if let Some(parent) = file.parent() {
            create_dir_all(parent)?;
        }
        let appender = OpenOptions::new().create(true).append(true).open(file)?;
        let (writer, guard) = tracing_appender::non_blocking(appender);
        Ok((
            tracing_file_subscriber(min_level, writer),
            TracingLogGuard { _guard: guard },
        ))
    }
}

fn tracing_file_subscriber<W>(
    min_level: LogLevel,
    writer: W,
) -> impl tracing::Subscriber + Send + Sync
where
    W: for<'writer> MakeWriter<'writer> + Send + Sync + 'static,
{
    let filter = EnvFilter::new(min_level.as_str());
    tracing_subscriber::fmt()
        .json()
        .with_env_filter(filter)
        .with_writer(writer)
        .with_current_span(false)
        .with_span_list(false)
        .finish()
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

#[derive(Debug)]
pub struct TracingLogGuard {
    _guard: WorkerGuard,
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
        self.fields
            .insert(name.to_string(), redact_log_field(name, value));
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
    TracingInit(String),
}

impl Display for LoggerError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::TracingInit(message) => formatter.write_str(message),
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
#[path = "logger_tests.rs"]
mod logger_tests;
