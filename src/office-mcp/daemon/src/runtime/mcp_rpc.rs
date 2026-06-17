use crate::addin_mgr::{
    AddinChannelServer, AddinConnectionHub, AddinConnectionHubError, CommandRouter,
    CommandRouterError, ImageFetcher, QueuedCommand, SessionDescriptorView, SessionRegistry,
    ToolCallRequest, ToolResponse,
};
use crate::api::{CommandFailure, UiStateStore};
use crate::common::{AuditLog, AuditRecord};
use crate::mcp::{
    ExcelToolCatalog, ResourceReadRequest, WORD_V1_TOOLS, prompt_catalog_json, prompt_description,
    prompt_messages, resource_request_from_uri, tool_catalog_json, tool_failure,
    tool_failure_from_command, tool_success, word_resource_catalog_for_session,
    word_resource_templates,
};
use crate::runtime::json_rpc;
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
            "tools/list" => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": { "tools": tool_catalog_json() }
            })
            .to_string(),
            "tools/call" => Self::handle_tools_call(context, &id, &value),
            "resources/list" => Self::handle_resources_list(context.registry, &id),
            "resources/templates/list" => Self::handle_resource_templates_list(&id),
            "resources/read" => Self::handle_resources_read(context, &id, &value),
            "prompts/list" => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": { "prompts": prompt_catalog_json() }
            })
            .to_string(),
            "prompts/get" => Self::handle_prompts_get(&id, &value),
            _ => json_rpc::error(&id, -32601, &format!("Unknown method {method}")),
        }
    }

    fn handle_resources_list(registry: &SessionRegistry, id: &Value) -> String {
        let mut resources = vec![json!({
            "uri": "office://sessions",
            "name": "office.sessions",
            "title": "Office Sessions",
            "mimeType": "application/json"
        })];
        for session in registry.list_sessions() {
            if session.app == "word" {
                resources.extend(word_resource_catalog_for_session(&session.session_id));
            }
        }
        json!({ "jsonrpc": "2.0", "id": id, "result": { "resources": resources } }).to_string()
    }

    fn handle_resource_templates_list(id: &Value) -> String {
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": { "resourceTemplates": word_resource_templates() }
        })
        .to_string()
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

    fn handle_prompts_get(id: &Value, value: &Value) -> String {
        let Some(name) = value
            .get("params")
            .and_then(|params| params.get("name"))
            .and_then(Value::as_str)
        else {
            return json_rpc::error(id, -32602, "prompts/get requires params.name");
        };
        let arguments = value
            .get("params")
            .and_then(|params| params.get("arguments"));
        match prompt_messages(name, arguments) {
            Some(messages) => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": { "description": prompt_description(name), "messages": messages }
            })
            .to_string(),
            None => json_rpc::error(id, -32602, &format!("Unknown prompt {name}")),
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
        let arguments = match preprocess_tool_arguments(context.image_fetcher, tool, arguments) {
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
                    record_failure_audit(
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
            Ok(response) => addin_response_to_tool_response(&response),
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
        record_tool_audit(
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

fn preprocess_tool_arguments(
    image_fetcher: &ImageFetcher,
    tool: &str,
    arguments: &Value,
) -> Result<Value, CommandFailure> {
    if tool != "word.insert_image" {
        return Ok(arguments.clone());
    }
    let Some(image) = arguments.get("image") else {
        return Ok(arguments.clone());
    };
    let processed = if let Some(base64) = image.get("base64").and_then(Value::as_str) {
        image_fetcher.validate_base64(base64)
    } else if let Some(url) = image.get("url").and_then(Value::as_str) {
        image_fetcher.fetch_url(url)
    } else {
        return Ok(arguments.clone());
    };
    let fetched = processed.map_err(|error| CommandFailure {
        office_mcp_code: "IMAGE_FETCH_FAILED".to_string(),
        message: error.to_string(),
        tool: Some(tool.to_string()),
        retriable: false,
        partial_effect: Some(crate::addin_mgr::PartialEffect::None),
    })?;
    let mut updated = arguments.clone();
    if let Some(object) = updated.as_object_mut() {
        object.insert(
            "image".to_string(),
            json!({
                "base64": fetched.base64,
                "mime_type": fetched.mime_type.as_str(),
                "byte_length": fetched.byte_length
            }),
        );
    }
    Ok(updated)
}

fn record_tool_audit(
    audit_log: &AuditLog,
    tool: &str,
    session_id: &str,
    completed: &Result<ToolResponse, CommandRouterError>,
    started_at: SystemTime,
    completed_at: SystemTime,
) {
    let duration_ms = duration_millis(started_at, completed_at);
    let record = match completed {
        Ok(ToolResponse::Success { .. }) => {
            AuditRecord::success(SystemTime::now(), tool, Some(session_id), duration_ms)
        }
        Ok(ToolResponse::Failure(failure)) => AuditRecord::failure(
            SystemTime::now(),
            tool,
            Some(session_id),
            duration_ms,
            &failure.office_mcp_code,
            &failure.message,
        ),
        Err(error) => {
            let failure = error.as_command_failure(tool);
            AuditRecord::failure(
                SystemTime::now(),
                tool,
                Some(session_id),
                duration_ms,
                &failure.office_mcp_code,
                &failure.message,
            )
        }
    };
    if let Err(error) = audit_log.record(&record) {
        tracing::error!(%error, "failed to write audit record");
        eprintln!("office-mcp-daemon failed to write audit record: {error}");
    }
}

fn record_failure_audit(
    audit_log: &AuditLog,
    tool: &str,
    session_id: &str,
    failure: &CommandFailure,
    started_at: SystemTime,
    completed_at: SystemTime,
) {
    let record = AuditRecord::failure(
        SystemTime::now(),
        tool,
        Some(session_id),
        duration_millis(started_at, completed_at),
        &failure.office_mcp_code,
        &failure.message,
    );
    if let Err(error) = audit_log.record(&record) {
        tracing::error!(%error, "failed to write audit record");
        eprintln!("office-mcp-daemon failed to write audit record: {error}");
    }
}

fn duration_millis(started_at: SystemTime, completed_at: SystemTime) -> u64 {
    completed_at
        .duration_since(started_at)
        .unwrap_or_default()
        .as_millis()
        .try_into()
        .unwrap_or(u64::MAX)
}

fn addin_response_to_tool_response(response: &Value) -> ToolResponse {
    if let Some(error) = response.get("error") {
        return ToolResponse::Failure(CommandFailure {
            office_mcp_code: error
                .get("office_mcp_code")
                .or_else(|| error.get("code"))
                .and_then(Value::as_str)
                .unwrap_or("ADDIN_ERROR")
                .to_string(),
            message: error
                .get("message")
                .and_then(Value::as_str)
                .unwrap_or("Add-in tool call failed.")
                .to_string(),
            tool: error
                .get("tool")
                .and_then(Value::as_str)
                .map(str::to_string),
            retriable: error
                .get("retriable")
                .and_then(Value::as_bool)
                .unwrap_or(false),
            partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
        });
    }
    let result = response.get("result").cloned().unwrap_or(Value::Null);
    if result.get("ok").and_then(Value::as_bool) == Some(false) {
        let error = result.get("error").cloned().unwrap_or(Value::Null);
        return ToolResponse::Failure(CommandFailure {
            office_mcp_code: error
                .get("office_mcp_code")
                .and_then(Value::as_str)
                .unwrap_or("ADDIN_ERROR")
                .to_string(),
            message: error
                .get("message")
                .and_then(Value::as_str)
                .unwrap_or("Add-in tool call failed.")
                .to_string(),
            tool: error
                .get("tool")
                .and_then(Value::as_str)
                .map(str::to_string),
            retriable: error
                .get("retriable")
                .and_then(Value::as_bool)
                .unwrap_or(false),
            partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
        });
    }
    let data = result.get("data").cloned().unwrap_or(result);
    ToolResponse::Success {
        json: data.to_string(),
    }
}

#[cfg(test)]
#[path = "mcp_rpc_tests.rs"]
mod mcp_rpc_tests;
