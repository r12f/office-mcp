use std::fmt::{Display, Formatter};
use std::fs::{OpenOptions, create_dir_all};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuditLog {
    enabled: bool,
    path: Option<PathBuf>,
}

impl AuditLog {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            enabled: false,
            path: None,
        }
    }

    #[must_use]
    pub fn enabled(path: impl Into<PathBuf>) -> Self {
        Self {
            enabled: true,
            path: Some(path.into()),
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns optional audit records and document payload redaction"
    }

    /// Writes one audit record when audit logging is enabled.
    ///
    /// # Errors
    ///
    /// Returns an error when the audit file directory cannot be created or the
    /// audit record cannot be appended.
    pub fn record(&self, record: &AuditRecord) -> Result<Option<String>, AuditLogError> {
        if !self.enabled {
            return Ok(None);
        }
        let line = record.to_json_line();
        if let Some(path) = &self.path {
            append_line(path, &line)?;
        }
        Ok(Some(line))
    }
}

impl Default for AuditLog {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuditRecord {
    pub ts: SystemTime,
    pub tool: String,
    pub session_id: Option<String>,
    pub duration_ms: u64,
    pub ok: bool,
    pub error_code: Option<String>,
    pub error_message: Option<String>,
}

impl AuditRecord {
    #[must_use]
    pub fn success(ts: SystemTime, tool: &str, session_id: Option<&str>, duration_ms: u64) -> Self {
        Self {
            ts,
            tool: tool.to_string(),
            session_id: session_id.map(str::to_string),
            duration_ms,
            ok: true,
            error_code: None,
            error_message: None,
        }
    }

    #[must_use]
    pub fn failure(
        ts: SystemTime,
        tool: &str,
        session_id: Option<&str>,
        duration_ms: u64,
        error_code: &str,
        error_message: &str,
    ) -> Self {
        Self {
            ts,
            tool: tool.to_string(),
            session_id: session_id.map(str::to_string),
            duration_ms,
            ok: false,
            error_code: Some(error_code.to_string()),
            error_message: Some(redact_audit_text(error_message)),
        }
    }

    #[must_use]
    pub fn to_json_line(&self) -> String {
        let mut fields = vec![
            json_field("ts", &format_system_time(self.ts)),
            json_field("tool", &self.tool),
        ];
        if let Some(session_id) = &self.session_id {
            fields.push(json_field("session_id", session_id));
        }
        fields.push(format!("\"duration_ms\":{}", self.duration_ms));
        fields.push(format!("\"ok\":{}", self.ok));
        if let Some(error_code) = &self.error_code {
            fields.push(json_field("error_code", error_code));
        }
        if let Some(error_message) = &self.error_message {
            fields.push(json_field("error_message", error_message));
        }
        format!("{{{}}}", fields.join(","))
    }
}

#[derive(Debug)]
pub enum AuditLogError {
    Io(std::io::Error),
}

impl Display for AuditLogError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
        }
    }
}

impl std::error::Error for AuditLogError {}

impl From<std::io::Error> for AuditLogError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

fn append_line(path: &Path, line: &str) -> Result<(), AuditLogError> {
    if let Some(parent) = path.parent() {
        create_dir_all(parent)?;
    }
    let mut file = OpenOptions::new().create(true).append(true).open(path)?;
    writeln!(file, "{line}")?;
    Ok(())
}

fn redact_audit_text(value: &str) -> String {
    let mut result = value
        .split_whitespace()
        .map(|part| {
            let Some((key, _payload)) = part.split_once('=') else {
                return part.to_string();
            };
            if matches!(
                key.to_ascii_lowercase().as_str(),
                "text" | "body" | "content"
            ) {
                format!("{key}=[redacted]")
            } else {
                part.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ");
    result.truncate(500);
    result
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
        let dir =
            std::env::temp_dir().join(format!("office-mcp-audit-rust-{}", std::process::id()));
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
}
