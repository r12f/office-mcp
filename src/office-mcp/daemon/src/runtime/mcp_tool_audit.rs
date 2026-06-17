use crate::addin_mgr::{CommandRouterError, ToolResponse};
use crate::api::CommandFailure;
use crate::common::{AuditLog, AuditRecord};
use std::time::SystemTime;

pub(crate) struct McpToolAuditRecorder;

impl McpToolAuditRecorder {
    pub(crate) fn record_completed(
        audit_log: &AuditLog,
        tool: &str,
        session_id: &str,
        completed: &Result<ToolResponse, CommandRouterError>,
        started_at: SystemTime,
        completed_at: SystemTime,
    ) {
        let duration_ms = duration_millis(started_at, completed_at);
        let record = match completed {
            Ok(ToolResponse::Success { .. }) => {
                AuditRecord::success(SystemTime::now(), tool, Some(session_id), duration_ms)
            }
            Ok(ToolResponse::Failure(failure)) => AuditRecord::failure(
                SystemTime::now(),
                tool,
                Some(session_id),
                duration_ms,
                &failure.office_mcp_code,
                &failure.message,
            ),
            Err(error) => {
                let failure = error.as_command_failure(tool);
                AuditRecord::failure(
                    SystemTime::now(),
                    tool,
                    Some(session_id),
                    duration_ms,
                    &failure.office_mcp_code,
                    &failure.message,
                )
            }
        };
        write_record(audit_log, &record);
    }

    pub(crate) fn record_failure(
        audit_log: &AuditLog,
        tool: &str,
        session_id: &str,
        failure: &CommandFailure,
        started_at: SystemTime,
        completed_at: SystemTime,
    ) {
        let record = AuditRecord::failure(
            SystemTime::now(),
            tool,
            Some(session_id),
            duration_millis(started_at, completed_at),
            &failure.office_mcp_code,
            &failure.message,
        );
        write_record(audit_log, &record);
    }
}

fn write_record(audit_log: &AuditLog, record: &AuditRecord) {
    if let Err(error) = audit_log.record(record) {
        tracing::error!(%error, "failed to write audit record");
        eprintln!("office-mcp-daemon failed to write audit record: {error}");
    }
}

fn duration_millis(started_at: SystemTime, completed_at: SystemTime) -> u64 {
    completed_at
        .duration_since(started_at)
        .unwrap_or_default()
        .as_millis()
        .try_into()
        .unwrap_or(u64::MAX)
}

#[cfg(test)]
#[path = "mcp_tool_audit_tests.rs"]
mod mcp_tool_audit_tests;
