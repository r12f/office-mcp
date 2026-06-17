use crate::api::CommandFailure;
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToolCallRequest {
    pub request_id: Option<String>,
    pub command_id: Option<String>,
    pub client_id: Option<String>,
    pub client_name: Option<String>,
    pub session_id: String,
    pub tool: String,
    pub arguments_json: String,
    pub user_intent: Option<String>,
    pub timeout: Option<Duration>,
    pub check_capability: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueuedCommand {
    pub command_id: String,
    pub request_id: String,
    pub session_id: String,
    pub instance_id: String,
    pub tool: String,
    pub arguments_json: String,
    pub timeout: Duration,
    pub enqueued_at: SystemTime,
    pub deadline_at: SystemTime,
    pub sequence: u64,
    pub dispatched: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ToolResponse {
    Success { json: String },
    Failure(CommandFailure),
}

impl ToolResponse {
    pub(crate) fn estimated_json_bytes(&self) -> usize {
        match self {
            Self::Success { json } => json.len(),
            Self::Failure(failure) => failure.message.len() + failure.office_mcp_code.len() + 128,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CancelCommand {
    pub request_id: String,
    pub reason: String,
}

#[cfg(test)]
#[path = "command_model_tests.rs"]
mod command_model_tests;
