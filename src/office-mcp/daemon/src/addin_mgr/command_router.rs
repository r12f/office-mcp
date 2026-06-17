use crate::addin_mgr::{CommandRouterError, PartialEffect, SessionCommandQueue, SessionRegistry};
use crate::api::{CommandFailure, CommandResult, StartCommandInput, UiStateStore};
use std::collections::BTreeMap;
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
        let permit = match registry.prepare_invocation(
            &request.session_id,
            &request.tool,
            request.check_capability,
        ) {
            Ok(permit) => permit,
            Err(error) => {
                tracing::warn!(
                    component = "command_router",
                    session_id = %request.session_id,
                    tool = %request.tool,
                    client_id = ?request.client_id,
                    code = %error.failure.office_mcp_code.as_str(),
                    "rejected tool command during preflight"
                );
                return Err(CommandRouterError::Preflight(error));
            }
        };
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
        tracing::info!(
            component = "command_router",
            command_id = %queued.command_id,
            request_id = %queued.request_id,
            session_id = %queued.session_id,
            instance_id = %queued.instance_id,
            tool = %queued.tool,
            sequence = queued.sequence,
            timeout_ms = duration_millis(queued.timeout),
            "queued tool command"
        );
        Ok(queued)
    }

    pub fn mark_dispatched(&mut self, session_id: &str, request_id: &str) -> bool {
        let Some(queue) = self.queues_by_session.get_mut(session_id) else {
            tracing::warn!(
                component = "command_router",
                session_id = %session_id,
                request_id = %request_id,
                "failed to mark missing command queue dispatched"
            );
            return false;
        };
        let dispatched = queue.mark_dispatched(request_id);
        if dispatched {
            tracing::debug!(
                component = "command_router",
                session_id = %session_id,
                request_id = %request_id,
                "marked tool command dispatched"
            );
        } else {
            tracing::warn!(
                component = "command_router",
                session_id = %session_id,
                request_id = %request_id,
                "failed to mark unknown tool command dispatched"
            );
        }
        dispatched
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
        let command = self.remove_command(session_id, request_id).ok_or_else(|| {
            tracing::warn!(
                component = "command_router",
                session_id = %session_id,
                request_id = %request_id,
                "received completion for unknown tool command"
            );
            CommandRouterError::UnknownRequest(request_id.to_string())
        })?;
        let response = self.validate_response_size(&command, response)?;
        let result = match &response {
            ToolResponse::Success { .. } => CommandResult::Success,
            ToolResponse::Failure(failure) => CommandResult::Failure(failure.clone()),
        };
        ui_state.finish_command(&command.command_id, result, completed_at);
        match &response {
            ToolResponse::Success { .. } => tracing::info!(
                component = "command_router",
                command_id = %command.command_id,
                request_id = %command.request_id,
                session_id = %command.session_id,
                tool = %command.tool,
                "completed tool command successfully"
            ),
            ToolResponse::Failure(failure) => tracing::warn!(
                component = "command_router",
                command_id = %command.command_id,
                request_id = %command.request_id,
                session_id = %command.session_id,
                tool = %command.tool,
                code = %failure.office_mcp_code,
                retriable = failure.retriable,
                "completed tool command with failure"
            ),
        }
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
            tracing::warn!(
                component = "command_router",
                session_id = %session_id,
                request_id = %request_id,
                "failed unknown tool command"
            );
            return false;
        };
        let code = failure.office_mcp_code.clone();
        ui_state.finish_command(
            &command.command_id,
            CommandResult::Failure(failure),
            completed_at,
        );
        tracing::warn!(
            component = "command_router",
            command_id = %command.command_id,
            request_id = %command.request_id,
            session_id = %command.session_id,
            tool = %command.tool,
            code = %code,
            "failed tool command"
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
        tracing::info!(
            component = "command_router",
            command_id = %command.command_id,
            request_id = %command.request_id,
            session_id = %command.session_id,
            tool = %command.tool,
            "cancelled tool command"
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
                    let tool = command.tool.clone();
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
                    tracing::warn!(
                        component = "command_router",
                        command_id = %command.command_id,
                        request_id = %request_id,
                        session_id = %session_id,
                        tool = %tool,
                        timeout_ms = duration_millis(command.timeout),
                        "expired tool command timeout"
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
        command: &QueuedCommand,
        response: ToolResponse,
    ) -> Result<ToolResponse, CommandRouterError> {
        let bytes = response.estimated_json_bytes();
        if bytes <= self.max_response_bytes {
            return Ok(response);
        }
        tracing::warn!(
            component = "command_router",
            command_id = %command.command_id,
            request_id = %command.request_id,
            session_id = %command.session_id,
            tool = %command.tool,
            max_response_bytes = self.max_response_bytes,
            actual_bytes = bytes,
            "rejected oversized tool response"
        );
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

fn duration_millis(duration: Duration) -> u64 {
    duration.as_millis().try_into().unwrap_or(u64::MAX)
}

#[cfg(test)]
#[path = "command_router_tests.rs"]
mod command_router_tests;
