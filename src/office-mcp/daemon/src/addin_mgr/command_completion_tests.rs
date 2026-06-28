use super::{cancelled_failure, duration_millis, result_from_tool_response, timeout_failure};
use crate::addin_mgr::{PartialEffect, QueuedCommand, ToolResponse};
use crate::api::{CommandFailure, CommandResult};
use std::time::{Duration, SystemTime};

#[test]
fn maps_tool_response_to_ui_command_result() {
    assert_eq!(
        result_from_tool_response(&ToolResponse::Success {
            json: "{}".to_string()
        }),
        CommandResult::Success
    );
    let failure = CommandFailure {
        office_mcp_code: "SESSION_LOST".to_string(),
        message: "Session session-1 lost its add-in connection.".to_string(),
        tool: Some("word.get_text".to_string()),
        retriable: false,
        partial_effect: Some(PartialEffect::Unknown),
        debug: None,
    };
    assert_eq!(
        result_from_tool_response(&ToolResponse::Failure(failure.clone())),
        CommandResult::Failure(failure)
    );
}

#[test]
fn builds_cancelled_failure_from_command() {
    let failure = cancelled_failure(&command(Duration::from_secs(30)));

    assert_eq!(failure.office_mcp_code, "CANCELLED");
    assert_eq!(failure.tool.as_deref(), Some("word.get_text"));
    assert!(failure.retriable);
    assert_eq!(failure.partial_effect, Some(PartialEffect::Unknown));
}

#[test]
fn builds_timeout_failure_with_duration_message() {
    let failure = timeout_failure(&command(Duration::from_secs(30)));

    assert_eq!(failure.office_mcp_code, "TIMEOUT");
    assert_eq!(
        failure.message,
        "Tool word.get_text timed out after 30000ms."
    );
    assert_eq!(failure.tool.as_deref(), Some("word.get_text"));
}

#[test]
fn duration_millis_saturates_large_values() {
    assert_eq!(duration_millis(Duration::from_millis(12)), 12);
    assert_eq!(duration_millis(Duration::MAX), u64::MAX);
}

fn command(timeout: Duration) -> QueuedCommand {
    QueuedCommand {
        command_id: "command-1".to_string(),
        request_id: "request-1".to_string(),
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        tool: "word.get_text".to_string(),
        arguments_json: "{}".to_string(),
        timeout,
        enqueued_at: SystemTime::UNIX_EPOCH,
        deadline_at: SystemTime::UNIX_EPOCH + Duration::from_secs(30),
        dispatched: false,
        sequence: 0,
    }
}
