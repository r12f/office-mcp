use crate::session_registry::{PartialEffect, SessionRegistry, ToolInvocationError};
use crate::ui::{CommandFailure, CommandResult, StartCommandInput, UiStateStore};
use std::collections::{BTreeMap, VecDeque};
use std::fmt::{Display, Formatter};
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommandRouter {
    max_response_bytes: usize,
    default_tool_timeout: Duration,
    queues_by_session: BTreeMap<String, SessionCommandQueue>,
    next_request_id: u64,
}

impl CommandRouter {
    #[must_use]
    pub fn new() -> Self {
        Self::with_limits(1024 * 1024, Duration::from_secs(30))
    }

    #[must_use]
    pub fn with_limits(max_response_bytes: usize, default_tool_timeout: Duration) -> Self {
        Self {
            max_response_bytes,
            default_tool_timeout,
            queues_by_session: BTreeMap::new(),
            next_request_id: 1,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns tool dispatch per session serialization timeouts cancellation and errors"
    }

    /// Preflights a tool call and enqueues it for the owning document session.
    ///
    /// # Errors
    ///
    /// Returns an error when the session registry rejects the call before
    /// dispatch, for example because the session is missing, stale,
    /// disconnected, saturated, or lacks the requested capability.
    pub fn enqueue(
        &mut self,
        registry: &SessionRegistry,
        ui_state: &mut UiStateStore,
        request: ToolCallRequest,
        now: SystemTime,
    ) -> Result<QueuedCommand, CommandRouterError> {
        let permit = registry
            .prepare_invocation(&request.session_id, &request.tool, request.check_capability)
            .map_err(CommandRouterError::Preflight)?;
        let request_id = request
            .request_id
            .unwrap_or_else(|| self.next_protocol_request_id());
        let timeout = request.timeout.unwrap_or(self.default_tool_timeout);
        let command_id = ui_state.start_command(StartCommandInput {
            command_id: request.command_id,
            mcp_request_id: Some(request_id.clone()),
            client_id: request.client_id,
            client_name: request.client_name,
            session_id: Some(request.session_id.clone()),
            host_app: Some(permit.host_app.clone()),
            tool: request.tool.clone(),
            user_intent: request.user_intent,
            timeout_ms: Some(duration_millis(timeout)),
            started_at: Some(now),
        });
        let sequence = self
            .queues_by_session
            .entry(request.session_id.clone())
            .or_default()
            .next_sequence();
        let queued = QueuedCommand {
            command_id,
            request_id,
            session_id: request.session_id,
            instance_id: permit.instance_id,
            tool: request.tool,
            arguments_json: request.arguments_json,
            timeout,
            enqueued_at: now,
            deadline_at: now + timeout,
            dispatched: false,
            sequence,
        };
        self.queues_by_session
            .entry(queued.session_id.clone())
            .or_default()
            .push(queued.clone());
        Ok(queued)
    }

    pub fn mark_dispatched(&mut self, session_id: &str, request_id: &str) -> bool {
        let Some(queue) = self.queues_by_session.get_mut(session_id) else {
            return false;
        };
        queue.mark_dispatched(request_id)
    }

    /// Completes a queued command and records the result in UI state.
    ///
    /// # Errors
    ///
    /// Returns an error when no queued command matches the request ID or when
    /// the response exceeds the configured maximum response size.
    pub fn complete(
        &mut self,
        ui_state: &mut UiStateStore,
        session_id: &str,
        request_id: &str,
        response: ToolResponse,
        completed_at: SystemTime,
    ) -> Result<ToolResponse, CommandRouterError> {
        let command = self
            .remove_command(session_id, request_id)
            .ok_or_else(|| CommandRouterError::UnknownRequest(request_id.to_string()))?;
        let response = self.validate_response_size(response)?;
        let result = match &response {
            ToolResponse::Success { .. } => CommandResult::Success,
            ToolResponse::Failure(failure) => CommandResult::Failure(failure.clone()),
        };
        ui_state.finish_command(&command.command_id, result, completed_at);
        Ok(response)
    }

    pub fn fail(
        &mut self,
        ui_state: &mut UiStateStore,
        session_id: &str,
        request_id: &str,
        failure: CommandFailure,
        completed_at: SystemTime,
    ) -> bool {
        let Some(command) = self.remove_command(session_id, request_id) else {
            return false;
        };
        ui_state.finish_command(
            &command.command_id,
            CommandResult::Failure(failure),
            completed_at,
        );
        true
    }

    pub fn cancel(
        &mut self,
        ui_state: &mut UiStateStore,
        session_id: &str,
        request_id: &str,
        completed_at: SystemTime,
    ) -> Option<CancelCommand> {
        let command = self.remove_command(session_id, request_id)?;
        ui_state.finish_command(
            &command.command_id,
            CommandResult::Failure(CommandFailure {
                office_mcp_code: "CANCELLED".to_string(),
                message: "The client cancelled the command.".to_string(),
                tool: Some(command.tool.clone()),
                retriable: true,
                partial_effect: Some(PartialEffect::Unknown),
            }),
            completed_at,
        );
        Some(CancelCommand {
            request_id: command.request_id,
            reason: "client_cancelled".to_string(),
        })
    }

    pub fn expire_timeouts(
        &mut self,
        ui_state: &mut UiStateStore,
        now: SystemTime,
    ) -> Vec<CancelCommand> {
        let mut expired = Vec::new();
        let sessions = self.queues_by_session.keys().cloned().collect::<Vec<_>>();
        for session_id in sessions {
            let request_ids = self
                .queues_by_session
                .get(&session_id)
                .map(|queue| queue.expired_request_ids(now))
                .unwrap_or_default();
            for request_id in request_ids {
                if let Some(command) = self.remove_command(&session_id, &request_id) {
                    ui_state.finish_command(
                        &command.command_id,
                        CommandResult::Failure(CommandFailure {
                            office_mcp_code: "TIMEOUT".to_string(),
                            message: format!(
                                "Tool {} timed out after {}ms.",
                                command.tool,
                                duration_millis(command.timeout)
                            ),
                            tool: Some(command.tool),
                            retriable: true,
                            partial_effect: Some(PartialEffect::Unknown),
                        }),
                        now,
                    );
                    expired.push(CancelCommand {
                        request_id,
                        reason: "deadline_expired".to_string(),
                    });
                }
            }
        }
        expired
    }

    #[must_use]
    pub fn queue_depth(&self, session_id: &str) -> usize {
        self.queues_by_session
            .get(session_id)
            .map_or(0, SessionCommandQueue::len)
    }

    fn validate_response_size(
        &self,
        response: ToolResponse,
    ) -> Result<ToolResponse, CommandRouterError> {
        let bytes = response.estimated_json_bytes();
        if bytes <= self.max_response_bytes {
            return Ok(response);
        }
        Err(CommandRouterError::ResponseTooLarge {
            max_response_bytes: self.max_response_bytes,
            actual_bytes: bytes,
        })
    }

    fn remove_command(&mut self, session_id: &str, request_id: &str) -> Option<QueuedCommand> {
        let queue = self.queues_by_session.get_mut(session_id)?;
        let command = queue.remove(request_id);
        if queue.is_empty() {
            self.queues_by_session.remove(session_id);
        }
        command
    }

    fn next_protocol_request_id(&mut self) -> String {
        let value = format!("request-{}", self.next_request_id);
        self.next_request_id += 1;
        value
    }
}

impl Default for CommandRouter {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
struct SessionCommandQueue {
    commands: VecDeque<QueuedCommand>,
    next_sequence: u64,
}

impl SessionCommandQueue {
    fn push(&mut self, command: QueuedCommand) {
        self.commands.push_back(command);
    }

    fn remove(&mut self, request_id: &str) -> Option<QueuedCommand> {
        let index = self
            .commands
            .iter()
            .position(|command| command.request_id == request_id)?;
        self.commands.remove(index)
    }

    fn mark_dispatched(&mut self, request_id: &str) -> bool {
        let Some(command) = self
            .commands
            .iter_mut()
            .find(|command| command.request_id == request_id)
        else {
            return false;
        };
        command.dispatched = true;
        true
    }

    fn expired_request_ids(&self, now: SystemTime) -> Vec<String> {
        self.commands
            .iter()
            .filter(|command| command.deadline_at <= now)
            .map(|command| command.request_id.clone())
            .collect()
    }

    fn next_sequence(&mut self) -> u64 {
        let value = self.next_sequence;
        self.next_sequence += 1;
        value
    }

    fn len(&self) -> usize {
        self.commands.len()
    }

    fn is_empty(&self) -> bool {
        self.commands.is_empty()
    }
}

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
    fn estimated_json_bytes(&self) -> usize {
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
            },
            Self::UnknownRequest(request_id) => CommandFailure {
                office_mcp_code: "INTERNAL_BUG".to_string(),
                message: format!("Unknown command request {request_id}."),
                tool: Some(tool.to_string()),
                retriable: false,
                partial_effect: None,
            },
            Self::ResponseTooLarge {
                max_response_bytes, ..
            } => CommandFailure {
                office_mcp_code: "MAX_RESPONSE_SIZE".to_string(),
                message: format!("Tool response exceeds {max_response_bytes} bytes."),
                tool: Some(tool.to_string()),
                retriable: false,
                partial_effect: Some(PartialEffect::None),
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

fn duration_millis(duration: Duration) -> u64 {
    duration.as_millis().try_into().unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::{CommandRouter, CommandRouterError, ToolCallRequest, ToolResponse};
    use crate::session_registry::{
        AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, OfficeMcpCode, RuntimeInfo,
        SessionRegistry,
    };
    use crate::ui::{RegisterClientInput, UiClientTransport, UiCommandStatus, UiStateStore};
    use std::time::{Duration, SystemTime};

    #[test]
    fn enqueue_starts_ui_task_and_serializes_per_session() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let registry = registry_with_session(now);
        let mut ui_state = UiStateStore::new();
        let client_id = ui_state.register_client(RegisterClientInput {
            client_id: Some("client-1".to_string()),
            transport: UiClientTransport::Http,
            name: Some("copilot-cli/1.0".to_string()),
        });
        let mut router = CommandRouter::new();

        let first = router
            .enqueue(
                &registry,
                &mut ui_state,
                request("session-1", "word.get_text", Some(client_id.clone())),
                now,
            )
            .expect("enqueue first");
        let second = router
            .enqueue(
                &registry,
                &mut ui_state,
                request("session-1", "word.add_comment", Some(client_id)),
                now,
            )
            .expect("enqueue second");

        assert_eq!(first.sequence, 0);
        assert_eq!(second.sequence, 1);
        assert_eq!(router.queue_depth("session-1"), 2);
        assert_eq!(router.queue_depth("word"), 0);
        let snapshot = ui_state.snapshot(&[], now);
        assert_eq!(snapshot.current_tasks.len(), 2);
        assert_eq!(snapshot.clients[0].in_flight_request_count, 2);
    }

    #[test]
    fn preflight_errors_do_not_start_ui_task() {
        let registry = SessionRegistry::new();
        let mut ui_state = UiStateStore::new();
        let mut router = CommandRouter::new();

        let error = router
            .enqueue(
                &registry,
                &mut ui_state,
                request("missing", "word.get_text", None),
                SystemTime::UNIX_EPOCH,
            )
            .expect_err("preflight error");

        assert!(matches!(
            error,
            CommandRouterError::Preflight(error)
                if error.failure.office_mcp_code == OfficeMcpCode::NoSessions
        ));
        assert!(
            ui_state
                .snapshot(&[], SystemTime::UNIX_EPOCH)
                .current_tasks
                .is_empty()
        );
    }

    #[test]
    fn complete_records_success_and_removes_queue_entry() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let registry = registry_with_session(now);
        let mut ui_state = UiStateStore::new();
        let mut router = CommandRouter::new();
        let queued = router
            .enqueue(
                &registry,
                &mut ui_state,
                request("session-1", "word.get_text", None),
                now,
            )
            .expect("enqueue");

        router
            .complete(
                &mut ui_state,
                "session-1",
                &queued.request_id,
                ToolResponse::Success {
                    json: "{\"ok\":true}".to_string(),
                },
                now + Duration::from_secs(1),
            )
            .expect("complete");

        let snapshot = ui_state.snapshot(&[], now + Duration::from_secs(1));
        assert_eq!(router.queue_depth("session-1"), 0);
        assert!(snapshot.current_tasks.is_empty());
        assert_eq!(snapshot.recent_commands[0].status, UiCommandStatus::Success);
    }

    #[test]
    fn oversized_response_returns_max_response_size() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let registry = registry_with_session(now);
        let mut ui_state = UiStateStore::new();
        let mut router = CommandRouter::with_limits(4, Duration::from_secs(30));
        let queued = router
            .enqueue(
                &registry,
                &mut ui_state,
                request("session-1", "word.get_text", None),
                now,
            )
            .expect("enqueue");

        let error = router
            .complete(
                &mut ui_state,
                "session-1",
                &queued.request_id,
                ToolResponse::Success {
                    json: "too large".to_string(),
                },
                now,
            )
            .expect_err("too large");
        let failure = error.as_command_failure("word.get_text");

        assert_eq!(failure.office_mcp_code, "MAX_RESPONSE_SIZE");
    }

    #[test]
    fn timeout_expires_command_and_returns_cancel_message() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let registry = registry_with_session(now);
        let mut ui_state = UiStateStore::new();
        let mut router = CommandRouter::new();
        let queued = router
            .enqueue(
                &registry,
                &mut ui_state,
                ToolCallRequest {
                    timeout: Some(Duration::from_millis(5)),
                    ..request("session-1", "word.get_text", None)
                },
                now,
            )
            .expect("enqueue");

        let cancels = router.expire_timeouts(&mut ui_state, now + Duration::from_millis(6));

        assert_eq!(cancels.len(), 1);
        assert_eq!(cancels[0].request_id, queued.request_id);
        assert_eq!(cancels[0].reason, "deadline_expired");
        assert_eq!(
            ui_state.snapshot(&[], now).recent_commands[0].status,
            UiCommandStatus::Timeout
        );
    }

    fn registry_with_session(now: SystemTime) -> SessionRegistry {
        let mut registry = SessionRegistry::new();
        registry.register_runtime(RuntimeInfo {
            instance_id: "instance-1".to_string(),
            host: HostInfo {
                app: "word".to_string(),
                version: Some("16.0".to_string()),
                platform: Some("windows".to_string()),
                build: Some("Desktop".to_string()),
            },
            add_in: AddInInfo {
                version: "0.1.0".to_string(),
                protocol_version: "1.0".to_string(),
                supported_features: vec!["doc.read".to_string()],
            },
            registered_at: now,
        });
        registry.add_session(
            NewSessionInfo {
                session_id: "session-1".to_string(),
                instance_id: "instance-1".to_string(),
                document: DocumentInfo::default(),
                available_tools: vec!["word.get_text".to_string(), "word.add_comment".to_string()],
                is_active: Some(true),
            },
            now,
        );
        registry
    }

    fn request(session_id: &str, tool: &str, client_id: Option<String>) -> ToolCallRequest {
        ToolCallRequest {
            request_id: None,
            command_id: None,
            client_id,
            client_name: Some("copilot-cli/1.0".to_string()),
            session_id: session_id.to_string(),
            tool: tool.to_string(),
            arguments_json: "{}".to_string(),
            user_intent: Some("read current selection".to_string()),
            timeout: None,
            check_capability: true,
        }
    }
}
