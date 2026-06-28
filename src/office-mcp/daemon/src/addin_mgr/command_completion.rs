use crate::addin_mgr::{PartialEffect, QueuedCommand, ToolResponse};
use crate::api::{CommandFailure, CommandResult};
use std::time::Duration;

pub(crate) const CLIENT_CANCELLED_REASON: &str = "client_cancelled";
pub(crate) const DEADLINE_EXPIRED_REASON: &str = "deadline_expired";

#[must_use]
pub(crate) fn result_from_tool_response(response: &ToolResponse) -> CommandResult {
    match response {
        ToolResponse::Success { .. } => CommandResult::Success,
        ToolResponse::Failure(failure) => CommandResult::Failure(failure.clone()),
    }
}

#[must_use]
pub(crate) fn cancelled_failure(command: &QueuedCommand) -> CommandFailure {
    CommandFailure {
        office_mcp_code: "CANCELLED".to_string(),
        message: "The client cancelled the command.".to_string(),
        tool: Some(command.tool.clone()),
        retriable: true,
        partial_effect: Some(PartialEffect::Unknown),
        debug: None,
    }
}

#[must_use]
pub(crate) fn timeout_failure(command: &QueuedCommand) -> CommandFailure {
    CommandFailure {
        office_mcp_code: "TIMEOUT".to_string(),
        message: format!(
            "Tool {} timed out after {}ms.",
            command.tool,
            duration_millis(command.timeout)
        ),
        tool: Some(command.tool.clone()),
        retriable: true,
        partial_effect: Some(PartialEffect::Unknown),
        debug: None,
    }
}

#[must_use]
pub(crate) fn duration_millis(duration: Duration) -> u64 {
    duration.as_millis().try_into().unwrap_or(u64::MAX)
}

#[cfg(test)]
#[path = "command_completion_tests.rs"]
mod command_completion_tests;
