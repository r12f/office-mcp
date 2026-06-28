use crate::addin_mgr::{AddinConnectionHubError, QueuedCommand, ToolCallRequest, ToolResponse};
use crate::api::CommandFailure;
use crate::mcp::{tool_failure, tool_failure_from_command, tool_success};
use crate::runtime::addin_tool_response::AddinToolResponseMapper;
use crate::runtime::json_rpc;
use crate::runtime::mcp_rpc::McpDispatchContext;
use crate::runtime::mcp_tool_arguments::McpToolArgumentPreprocessor;
use crate::runtime::mcp_tool_audit::McpToolAuditRecorder;
use serde_json::{Value, json};
use std::time::SystemTime;

pub(crate) struct McpForwardedToolInvoker;

impl McpForwardedToolInvoker {
    pub(crate) fn call(
        context: &mut McpDispatchContext<'_>,
        request_value: &Value,
        tool: &str,
        arguments: &Value,
        check_capability: bool,
    ) -> Value {
        let Some(session_id) = arguments.get("session_id").and_then(Value::as_str) else {
            return tool_failure(
                "INVALID_ARGUMENTS",
                "Forwarded Office tools require session_id.",
                Some(tool),
                false,
            );
        };
        let audit_started_at = SystemTime::now();
        let request_id = request_value
            .get("id")
            .and_then(Value::as_str)
            .map(str::to_string);
        let arguments =
            match McpToolArgumentPreprocessor::preprocess(context.image_fetcher, tool, arguments) {
                Ok(arguments) => arguments,
                Err(failure) => return tool_failure_from_command(&failure),
            };
        let arguments_json = arguments.to_string();
        let queued = {
            let mut router = context.command_router.lock().expect("command router lock");
            match router.enqueue(
                context.registry,
                context.ui_state,
                ToolCallRequest {
                    request_id,
                    command_id: None,
                    client_id: None,
                    client_name: None,
                    session_id: session_id.to_string(),
                    tool: tool.to_string(),
                    arguments_json,
                    user_intent: None,
                    timeout: None,
                    check_capability,
                },
                SystemTime::now(),
            ) {
                Ok(queued) => queued,
                Err(error) => {
                    let failure = error.as_command_failure(tool);
                    McpToolAuditRecorder::record_failure(
                        context.audit_log,
                        tool,
                        session_id,
                        &failure,
                        audit_started_at,
                        SystemTime::now(),
                    );
                    return tool_failure_from_command(&failure);
                }
            }
        };
        let tool_response = Self::invoke_queued_tool(context, &queued, tool);
        Self::complete_queued_tool(context, &queued, tool_response, tool, audit_started_at)
    }

    fn invoke_queued_tool(
        context: &McpDispatchContext<'_>,
        queued: &QueuedCommand,
        tool: &str,
    ) -> ToolResponse {
        let payload = {
            let addin_channel = context.addin_channel.lock().expect("addin channel lock");
            json_rpc::envelope_to_text(&addin_channel.tool_invoke_payload(queued))
        };
        match context.connection_hub.invoke(
            &queued.instance_id,
            &queued.request_id,
            payload,
            queued.timeout,
        ) {
            Ok(response) => AddinToolResponseMapper::map(&response),
            Err(AddinConnectionHubError::NoConnection) => ToolResponse::Failure(CommandFailure {
                office_mcp_code: "SESSION_LOST".to_string(),
                message: format!("Session {} lost its add-in connection.", queued.session_id),
                tool: Some(tool.to_string()),
                retriable: true,
                partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
                debug: None,
            }),
            Err(AddinConnectionHubError::Timeout) => {
                Self::send_timeout_cancel(context, queued);
                ToolResponse::Failure(CommandFailure {
                    office_mcp_code: "TIMEOUT".to_string(),
                    message: format!(
                        "Tool {tool} timed out after {}ms.",
                        queued.timeout.as_millis()
                    ),
                    tool: Some(tool.to_string()),
                    retriable: true,
                    partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
                    debug: None,
                })
            }
        }
    }

    fn send_timeout_cancel(context: &McpDispatchContext<'_>, queued: &QueuedCommand) {
        let cancel_payload = {
            let addin_channel = context.addin_channel.lock().expect("addin channel lock");
            json_rpc::envelope_to_text(&addin_channel.tool_cancel_payload(
                &crate::addin_mgr::CancelCommand {
                    request_id: queued.request_id.clone(),
                    reason: "deadline_expired".to_string(),
                },
            ))
        };
        context
            .connection_hub
            .send_to_instance(&queued.instance_id, cancel_payload);
    }

    fn complete_queued_tool(
        context: &mut McpDispatchContext<'_>,
        queued: &QueuedCommand,
        tool_response: ToolResponse,
        tool: &str,
        audit_started_at: SystemTime,
    ) -> Value {
        let completed = {
            let mut router = context.command_router.lock().expect("command router lock");
            router.complete(
                context.ui_state,
                &queued.session_id,
                &queued.request_id,
                tool_response,
                SystemTime::now(),
            )
        };
        McpToolAuditRecorder::record_completed(
            context.audit_log,
            tool,
            &queued.session_id,
            &completed,
            audit_started_at,
            SystemTime::now(),
        );
        match completed {
            Ok(ToolResponse::Success { json }) => serde_json::from_str::<Value>(&json)
                .map_or_else(|_| tool_success(&json!(json)), |value| tool_success(&value)),
            Ok(ToolResponse::Failure(failure)) => tool_failure_from_command(&failure),
            Err(error) => tool_failure_from_command(&error.as_command_failure(tool)),
        }
    }
}

#[cfg(test)]
#[path = "mcp_forwarded_tool_tests.rs"]
mod mcp_forwarded_tool_tests;
