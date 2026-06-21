use crate::addin_mgr::{PartialEffect, SessionDescriptor};
use crate::api::ui_redaction::redact_text;
use crate::mcp::{ToolAccessPolicy, UiToolAccessPolicySnapshot};
use std::collections::BTreeMap;
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiStateOptions {
    pub version: String,
    pub mcp_endpoint: String,
    pub addin_endpoint: String,
    pub config_path: Option<String>,
    pub log_path: Option<String>,
    pub tool_access_policy: ToolAccessPolicy,
    pub now: SystemTime,
}

impl Default for UiStateOptions {
    fn default() -> Self {
        Self {
            version: "0.1.0".to_string(),
            mcp_endpoint: "http://127.0.0.1:8800/mcp".to_string(),
            addin_endpoint: "https://localhost:8765/addin".to_string(),
            config_path: None,
            log_path: None,
            tool_access_policy: ToolAccessPolicy::default(),
            now: SystemTime::UNIX_EPOCH,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UiHealth {
    Up,
    Degraded,
    Down,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UiClientTransport {
    Http,
    StdioBridge,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiClientRecord {
    pub client_id: String,
    pub transport: UiClientTransport,
    pub name: Option<String>,
    pub connected_at: SystemTime,
    pub last_activity_at: SystemTime,
    pub in_flight_request_count: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RegisterClientInput {
    pub client_id: Option<String>,
    pub transport: UiClientTransport,
    pub name: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StartCommandInput {
    pub command_id: Option<String>,
    pub mcp_request_id: Option<String>,
    pub client_id: Option<String>,
    pub client_name: Option<String>,
    pub session_id: Option<String>,
    pub host_app: Option<String>,
    pub tool: String,
    pub user_intent: Option<String>,
    pub timeout_ms: Option<u64>,
    pub started_at: Option<SystemTime>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiCommandRecord {
    pub command_id: String,
    pub mcp_request_id: Option<String>,
    pub client_id: Option<String>,
    pub client_name: Option<String>,
    pub session_id: Option<String>,
    pub host_app: Option<String>,
    pub tool: String,
    pub user_intent: Option<String>,
    pub status: UiCommandStatus,
    pub started_at: SystemTime,
    pub deadline_at: Option<SystemTime>,
    pub timeout_ms: Option<u64>,
    pub completed_at: Option<SystemTime>,
    pub elapsed_ms: Option<u64>,
    pub error: Option<UiCommandError>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UiCommandStatus {
    Running,
    Success,
    Failure,
    Cancelled,
    Timeout,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiCommandError {
    pub office_mcp_code: String,
    pub message: String,
    pub tool: Option<String>,
    pub retriable: bool,
    pub partial_effect: Option<PartialEffect>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CommandResult {
    Success,
    Failure(CommandFailure),
    Thrown(String),
}

impl CommandResult {
    pub(crate) fn into_status(self) -> CommandStatusUpdate {
        match self {
            Self::Success => CommandStatusUpdate {
                status: UiCommandStatus::Success,
                error: None,
            },
            Self::Thrown(message) => CommandStatusUpdate {
                status: UiCommandStatus::Failure,
                error: Some(UiCommandError {
                    office_mcp_code: "THROWN".to_string(),
                    message: redact_text(&message),
                    tool: None,
                    retriable: false,
                    partial_effect: None,
                }),
            },
            Self::Failure(failure) => failure.into_status(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommandFailure {
    pub office_mcp_code: String,
    pub message: String,
    pub tool: Option<String>,
    pub retriable: bool,
    pub partial_effect: Option<PartialEffect>,
}

impl CommandFailure {
    fn into_status(self) -> CommandStatusUpdate {
        let status = match self.office_mcp_code.as_str() {
            "TIMEOUT" => UiCommandStatus::Timeout,
            "CANCELLED" => UiCommandStatus::Cancelled,
            _ => UiCommandStatus::Failure,
        };
        CommandStatusUpdate {
            status,
            error: Some(UiCommandError {
                office_mcp_code: self.office_mcp_code,
                message: redact_text(&self.message),
                tool: self.tool,
                retriable: self.retriable,
                partial_effect: self.partial_effect,
            }),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct CommandStatusUpdate {
    pub(crate) status: UiCommandStatus,
    pub(crate) error: Option<UiCommandError>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiSnapshot {
    pub daemon: UiDaemonSnapshot,
    pub clients: Vec<UiClientRecord>,
    pub documents: BTreeMap<String, Vec<SessionDescriptor>>,
    pub current_tasks: Vec<UiCommandRecord>,
    pub recent_commands: Vec<UiCommandRecord>,
    pub document_command_history: BTreeMap<String, Vec<UiCommandRecord>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiDaemonSnapshot {
    pub status: UiHealth,
    pub version: String,
    pub uptime_ms: u64,
    pub mcp_endpoint: String,
    pub addin_endpoint: String,
    pub config_path: Option<String>,
    pub log_path: Option<String>,
    pub last_error: Option<String>,
    pub tool_access_policy: UiToolAccessPolicySnapshot,
}

#[cfg(test)]
#[path = "state_model_tests.rs"]
mod state_model_tests;
