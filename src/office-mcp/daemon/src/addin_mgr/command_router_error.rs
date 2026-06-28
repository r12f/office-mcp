use crate::addin_mgr::{PartialEffect, ToolInvocationError};
use crate::api::CommandFailure;
use std::fmt::{Display, Formatter};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CommandRouterError {
    Preflight(ToolInvocationError),
    UnknownRequest(String),
    ResponseTooLarge {
        max_response_bytes: usize,
        actual_bytes: usize,
    },
}

impl CommandRouterError {
    #[must_use]
    pub fn as_command_failure(&self, tool: &str) -> CommandFailure {
        match self {
            Self::Preflight(error) => CommandFailure {
                office_mcp_code: error.failure.office_mcp_code.as_str().to_string(),
                message: error.failure.message.clone(),
                tool: Some(tool.to_string()),
                retriable: error.failure.retriable,
                partial_effect: error.failure.partial_effect,
                debug: None,
            },
            Self::UnknownRequest(request_id) => CommandFailure {
                office_mcp_code: "INTERNAL_BUG".to_string(),
                message: format!("Unknown command request {request_id}."),
                tool: Some(tool.to_string()),
                retriable: false,
                partial_effect: None,
                debug: None,
            },
            Self::ResponseTooLarge {
                max_response_bytes, ..
            } => CommandFailure {
                office_mcp_code: "MAX_RESPONSE_SIZE".to_string(),
                message: format!("Tool response exceeds {max_response_bytes} bytes."),
                tool: Some(tool.to_string()),
                retriable: false,
                partial_effect: Some(PartialEffect::None),
                debug: None,
            },
        }
    }
}

impl Display for CommandRouterError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Preflight(error) => write!(formatter, "{error}"),
            Self::UnknownRequest(request_id) => {
                write!(formatter, "Unknown command request {request_id}.")
            }
            Self::ResponseTooLarge {
                max_response_bytes,
                actual_bytes,
            } => write!(
                formatter,
                "Response exceeded {max_response_bytes} bytes: {actual_bytes} bytes."
            ),
        }
    }
}

impl std::error::Error for CommandRouterError {}

#[cfg(test)]
#[path = "command_router_error_tests.rs"]
mod command_router_error_tests;
