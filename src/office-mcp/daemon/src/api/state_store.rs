use crate::addin_mgr::{PartialEffect, SessionDescriptor};
use std::collections::BTreeMap;
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiStateStore {
    options: UiStateOptions,
    clients: BTreeMap<String, UiClientRecord>,
    current_tasks: BTreeMap<String, UiCommandRecord>,
    recent_commands: Vec<UiCommandRecord>,
    commands_by_session: BTreeMap<String, Vec<UiCommandRecord>>,
    started_at: SystemTime,
    health: UiHealth,
    last_error: Option<String>,
    next_command_id: u64,
    next_client_id: u64,
}

impl UiStateStore {
    #[must_use]
    pub fn new() -> Self {
        Self::with_options(UiStateOptions::default())
    }

    #[must_use]
    pub fn with_options(options: UiStateOptions) -> Self {
        let started_at = options.now;
        Self {
            options,
            clients: BTreeMap::new(),
            current_tasks: BTreeMap::new(),
            recent_commands: Vec::new(),
            commands_by_session: BTreeMap::new(),
            started_at,
            health: UiHealth::Up,
            last_error: None,
            next_command_id: 1,
            next_client_id: 1,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns redacted UI snapshots current tasks bounded history and subscriptions"
    }

    pub fn set_health(&mut self, status: UiHealth, last_error: Option<&str>) {
        self.health = status;
        self.last_error = last_error.map(redact_text);
    }

    pub fn register_client(&mut self, input: RegisterClientInput) -> String {
        let now = self.options.now;
        let client_id = input.client_id.unwrap_or_else(|| self.next_client_id());
        self.clients.insert(
            client_id.clone(),
            UiClientRecord {
                client_id: client_id.clone(),
                transport: input.transport,
                name: input.name.as_deref().map(redact_text),
                connected_at: now,
                last_activity_at: now,
                in_flight_request_count: 0,
            },
        );
        client_id
    }

    pub fn unregister_client(&mut self, client_id: &str) -> bool {
        self.clients.remove(client_id).is_some()
    }

    pub fn touch_client(&mut self, client_id: &str, now: SystemTime) -> bool {
        let Some(client) = self.clients.get_mut(client_id) else {
            return false;
        };
        client.last_activity_at = now;
        true
    }

    pub fn start_command(&mut self, input: StartCommandInput) -> String {
        let command_id = input.command_id.unwrap_or_else(|| self.next_command_id());
        let started_at = input.started_at.unwrap_or(self.options.now);
        let deadline_at = input
            .timeout_ms
            .map(|timeout_ms| started_at + Duration::from_millis(timeout_ms));
        let command = UiCommandRecord {
            command_id: command_id.clone(),
            mcp_request_id: input.mcp_request_id.as_deref().map(redact_text),
            client_id: input.client_id.as_deref().map(redact_text),
            client_name: input.client_name.as_deref().map(redact_text),
            session_id: input.session_id.as_deref().map(redact_text),
            host_app: input.host_app.as_deref().map(redact_text),
            tool: input.tool,
            user_intent: input.user_intent.as_deref().map(redact_text),
            status: UiCommandStatus::Running,
            started_at,
            deadline_at,
            timeout_ms: input.timeout_ms,
            completed_at: None,
            elapsed_ms: None,
            error: None,
        };
        if let Some(client_id) = input.client_id {
            self.increment_client(client_id.as_str(), ClientCountDelta::Increment, started_at);
        }
        self.current_tasks.insert(command_id.clone(), command);
        command_id
    }

    pub fn finish_command(
        &mut self,
        command_id: &str,
        result: CommandResult,
        completed_at: SystemTime,
    ) -> bool {
        let Some(command) = self.current_tasks.remove(command_id) else {
            return false;
        };
        if let Some(client_id) = command.client_id.as_deref() {
            self.increment_client(client_id, ClientCountDelta::Decrement, completed_at);
        }
        let elapsed_ms = completed_at
            .duration_since(command.started_at)
            .unwrap_or_default()
            .as_millis()
            .try_into()
            .unwrap_or(u64::MAX);
        let mut finished = command;
        finished.completed_at = Some(completed_at);
        finished.elapsed_ms = Some(elapsed_ms);
        let status = result.into_status();
        finished.status = status.status;
        finished.error = status.error;
        self.push_recent(finished.clone());
        if let Some(session_id) = finished.session_id.clone() {
            let session_commands = self.commands_by_session.entry(session_id).or_default();
            session_commands.insert(0, finished);
            session_commands.truncate(10);
        }
        true
    }

    #[must_use]
    pub fn snapshot(&self, sessions: &[SessionDescriptor], now: SystemTime) -> UiSnapshot {
        UiSnapshot {
            daemon: UiDaemonSnapshot {
                status: self.health,
                version: self.options.version.clone(),
                uptime_ms: now
                    .duration_since(self.started_at)
                    .unwrap_or_default()
                    .as_millis()
                    .try_into()
                    .unwrap_or(u64::MAX),
                mcp_endpoint: self.options.mcp_endpoint.clone(),
                addin_endpoint: self.options.addin_endpoint.clone(),
                config_path: self.options.config_path.clone(),
                log_path: self.options.log_path.clone(),
                last_error: self.last_error.clone(),
            },
            clients: self.clients.values().cloned().collect(),
            documents: group_sessions_by_app(sessions),
            current_tasks: self.current_tasks.values().cloned().collect(),
            recent_commands: self.recent_commands.clone(),
            document_command_history: self.commands_by_session.clone(),
        }
    }

    fn push_recent(&mut self, command: UiCommandRecord) {
        self.recent_commands.insert(0, command);
        self.recent_commands.truncate(10);
    }

    fn increment_client(&mut self, client_id: &str, delta: ClientCountDelta, now: SystemTime) {
        let Some(client) = self.clients.get_mut(client_id) else {
            return;
        };
        client.in_flight_request_count = match delta {
            ClientCountDelta::Increment => client.in_flight_request_count.saturating_add(1),
            ClientCountDelta::Decrement => client.in_flight_request_count.saturating_sub(1),
        };
        client.last_activity_at = now;
    }

    fn next_client_id(&mut self) -> String {
        let value = format!("client-{}", self.next_client_id);
        self.next_client_id += 1;
        value
    }

    fn next_command_id(&mut self) -> String {
        let value = format!("command-{}", self.next_command_id);
        self.next_command_id += 1;
        value
    }
}

impl Default for UiStateStore {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiStateOptions {
    pub version: String,
    pub mcp_endpoint: String,
    pub addin_endpoint: String,
    pub config_path: Option<String>,
    pub log_path: Option<String>,
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
    fn into_status(self) -> CommandStatusUpdate {
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
struct CommandStatusUpdate {
    status: UiCommandStatus,
    error: Option<UiCommandError>,
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
}

fn group_sessions_by_app(
    sessions: &[SessionDescriptor],
) -> BTreeMap<String, Vec<SessionDescriptor>> {
    let mut grouped = BTreeMap::from([
        ("word".to_string(), Vec::new()),
        ("excel".to_string(), Vec::new()),
        ("powerpoint".to_string(), Vec::new()),
        ("outlook".to_string(), Vec::new()),
        ("other".to_string(), Vec::new()),
    ]);
    for session in sessions {
        let key = match session.app.as_str() {
            "word" | "excel" | "powerpoint" | "outlook" => session.app.clone(),
            _ => "other".to_string(),
        };
        grouped.entry(key).or_default().push(session.clone());
    }
    grouped
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ClientCountDelta {
    Increment,
    Decrement,
}

fn redact_text(value: &str) -> String {
    let mut result = redact_bearer_tokens(value);
    result = redact_key_value_secret(&result);
    result = redact_base64_data(&result);
    result.truncate(500);
    result
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

fn redact_key_value_secret(value: &str) -> String {
    value
        .split_whitespace()
        .map(|part| {
            let Some((key, _secret)) = part.split_once('=') else {
                return part.to_string();
            };
            let normalized = key.replace(['_', '-'], "").to_ascii_lowercase();
            if matches!(normalized.as_str(), "password" | "passphrase" | "token")
                || normalized.ends_with("passphrase")
            {
                format!("{key}=[redacted]")
            } else {
                part.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn redact_base64_data(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    let mut rest = value;
    while let Some(index) = rest.find("base64,") {
        output.push_str(&rest[..index]);
        output.push_str("base64,[redacted]");
        let after_marker = &rest[index + "base64,".len()..];
        let end = after_marker
            .find(|character: char| character.is_whitespace())
            .unwrap_or(after_marker.len());
        rest = &after_marker[end..];
    }
    output.push_str(rest);
    output
}

#[cfg(test)]
#[path = "state_store_tests.rs"]
mod state_store_tests;
