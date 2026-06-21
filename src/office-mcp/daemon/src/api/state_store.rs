use crate::addin_mgr::SessionDescriptor;
use crate::api::ui_redaction::redact_text;
use crate::api::{
    CommandResult, RegisterClientInput, StartCommandInput, UiClientRecord, UiCommandRecord,
    UiCommandStatus, UiDaemonSnapshot, UiHealth, UiSnapshot, UiStateOptions,
};
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
                tool_access_policy: self.options.tool_access_policy.snapshot(),
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

#[cfg(test)]
#[path = "state_store_tests.rs"]
mod state_store_tests;
