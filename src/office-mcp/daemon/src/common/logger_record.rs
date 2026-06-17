use crate::common::logger::LoggerError;
use crate::common::logger_redaction::{redact_log_field, redact_log_text};
use std::collections::BTreeMap;
use std::fs::{OpenOptions, create_dir_all};
use std::io::Write;
use std::path::Path;
use std::time::SystemTime;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum LogLevel {
    Trace,
    Debug,
    Info,
    Warn,
    Error,
}

impl LogLevel {
    #[must_use]
    pub(crate) const fn as_str(self) -> &'static str {
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

pub(crate) fn append_line(path: &Path, line: &str) -> Result<(), LoggerError> {
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
#[path = "logger_record_tests.rs"]
mod logger_record_tests;
