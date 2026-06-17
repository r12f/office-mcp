use crate::addin_mgr::{
    AddinChannelServer, AddinConnectionHub, AddinConnectionHubError, CommandRouter, ImageFetcher,
    QueuedCommand, SessionDescriptorView, SessionRegistry, ToolCallRequest, ToolResponse,
};
use crate::api::{CommandFailure, UiStateStore};
use crate::common::AuditLog;
use crate::mcp::{
    ExcelToolCatalog, ResourceReadRequest, WORD_V1_TOOLS, resource_request_from_uri, tool_failure,
    tool_failure_from_command, tool_success,
};
use crate::runtime::addin_tool_response::AddinToolResponseMapper;
use crate::runtime::json_rpc;
use crate::runtime::mcp_catalog_response::McpCatalogResponder;
use crate::runtime::mcp_tool_arguments::McpToolArgumentPreprocessor;
use crate::runtime::mcp_tool_audit::McpToolAuditRecorder;
use serde_json::{Value, json};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

pub(crate) struct McpDispatchContext<'a> {
    pub(crate) registry: &'a SessionRegistry,
    pub(crate) ui_state: &'a mut UiStateStore,
    pub(crate) addin_channel: &'a Arc<Mutex<AddinChannelServer>>,
    pub(crate) connection_hub: &'a Arc<AddinConnectionHub>,
    pub(crate) command_router: &'a Arc<Mutex<CommandRouter>>,
    pub(crate) audit_log: &'a AuditLog,
    pub(crate) image_fetcher: &'a ImageFetcher,
}

pub(crate) struct McpJsonRpcRuntime;

impl McpJsonRpcRuntime {
    pub(crate) fn handle_body(context: &mut McpDispatchContext<'_>, body: &[u8]) -> String {
        let Ok(value) = serde_json::from_slice::<Value>(body) else {
            return json_rpc::error(&Value::Null, -32700, "Parse error");
        };
        let id = value.get("id").cloned().unwrap_or(Value::Null);
        let Some(method) = value.get("method").and_then(Value::as_str) else {
            return json_rpc::error(&id, -32600, "Invalid Request");
        };
        match method {
            "tools/list" => McpCatalogResponder::tools_list(&id),
            "tools/call" => Self::handle_tools_call(context, &id, &value),
            "resources/list" => McpCatalogResponder::resources_list(context.registry, &id),
            "resources/templates/list" => McpCatalogResponder::resource_templates_list(&id),
            "resources/read" => Self::handle_resources_read(context, &id, &value),
            "prompts/list" => McpCatalogResponder::prompts_list(&id),
            "prompts/get" => McpCatalogResponder::prompts_get(&id, &value),
            _ => json_rpc::error(&id, -32601, &format!("Unknown method {method}")),
        }
    }

    fn handle_resources_read(
        context: &mut McpDispatchContext<'_>,
        id: &Value,
        value: &Value,
    ) -> String {
        let Some(uri) = value
            .get("params")
            .and_then(|params| params.get("uri"))
            .and_then(Value::as_str)
        else {
            return json_rpc::error(id, -32602, "resources/read requires params.uri");
        };
        match resource_request_from_uri(context.registry, uri) {
            Ok(ResourceReadRequest::Sessions) => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": {
                    "contents": [{
                        "uri": "office://sessions",
                        "mimeType": "application/json",
                        "text": json!({
                            "sessions": context
                                .registry
                                .list_sessions()
                                .iter()
                                .map(|session| SessionDescriptorView::new(session).to_json())
                                .collect::<Vec<_>>()
                        }).to_string()
                    }]
                }
            })
            .to_string(),
            Ok(ResourceReadRequest::Forwarded {
                uri,
                tool,
                arguments,
                check_capability,
            }) => {
                let result = Self::call_forwarded_tool_with_capability(
                    context,
                    value,
                    tool,
                    &arguments,
                    check_capability,
                );
                json!({
                    "jsonrpc": "2.0",
                    "id": id,
                    "result": {
                        "contents": [{
                            "uri": uri,
                            "mimeType": "application/json",
                            "text": result.get("structuredContent").cloned().unwrap_or(result).to_string()
                        }]
                    }
                })
                .to_string()
            }
            Err(message) => json_rpc::error(id, -32602, &message),
        }
    }

    fn handle_tools_call(
        context: &mut McpDispatchContext<'_>,
        id: &Value,
        value: &Value,
    ) -> String {
        let Some(params) = value.get("params") else {
            return json_rpc::error(id, -32602, "Missing tool call params");
        };
        let Some(name) = params.get("name").and_then(Value::as_str) else {
            return json_rpc::error(id, -32602, "Missing tool name");
        };
        let arguments = params.get("arguments").unwrap_or(&Value::Null);
        let result = match name {
            "office.list_sessions" => tool_success(&json!({
                "sessions": context
                    .registry
                    .list_sessions()
                    .iter()
                    .map(|session| SessionDescriptorView::new(session).to_json())
                    .collect::<Vec<_>>()
            })),
            "office.get_session_info" => Self::get_session_info(context.registry, arguments),
            tool if WORD_V1_TOOLS.contains(&tool) => {
                Self::call_forwarded_tool(context, value, tool, arguments)
            }
            tool if ExcelToolCatalog::contains(tool) => {
                Self::call_forwarded_tool(context, value, tool, arguments)
            }
            _ => tool_failure(
                "UNKNOWN_TOOL",
                &format!("Unknown tool {name}."),
                Some(name),
                false,
            ),
        };
        json!({ "jsonrpc": "2.0", "id": id, "result": result }).to_string()
    }

    fn call_forwarded_tool(
        context: &mut McpDispatchContext<'_>,
        request_value: &Value,
        tool: &str,
        arguments: &Value,
    ) -> Value {
        Self::call_forwarded_tool_with_capability(context, request_value, tool, arguments, true)
    }

    fn call_forwarded_tool_with_capability(
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

    fn get_session_info(registry: &SessionRegistry, arguments: &Value) -> Value {
        let Some(session_id) = arguments.get("session_id").and_then(Value::as_str) else {
            return tool_failure(
                "INVALID_ARGUMENTS",
                "office.get_session_info requires session_id.",
                Some("office.get_session_info"),
                false,
            );
        };
        let Some(info) = registry.get_session_info(session_id) else {
            return tool_failure(
                "SESSION_NOT_FOUND",
                &format!("Session {session_id} is not registered."),
                Some("office.get_session_info"),
                false,
            );
        };
        tool_success(&json!({
            "descriptor": SessionDescriptorView::new(&info.descriptor).to_json(),
            "available_tools": info.available_tools
        }))
    }
}

#[cfg(test)]
#[path = "mcp_rpc_tests.rs"]
mod mcp_rpc_tests;
