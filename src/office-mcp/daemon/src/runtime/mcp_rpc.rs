use crate::addin_mgr::{
    AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher, SessionDescriptorView,
    SessionRegistry,
};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::mcp::{
    ResourceReadRequest, ToolAccessPolicy, canonical_tool_name, describe_tool_contract,
    is_office_tool, resource_request_from_uri, tool_failure, tool_failure_without_effect,
    tool_not_available_by_policy, tool_success, unknown_tool_contract, validate_tool_arguments,
};
use crate::runtime::json_rpc;
use crate::runtime::mcp_catalog_response::McpCatalogResponder;
use crate::runtime::mcp_forwarded_tool::McpForwardedToolInvoker;
use serde_json::{Value, json};
use std::sync::{Arc, Mutex};

pub(crate) struct McpDispatchContext<'a> {
    pub(crate) registry: &'a SessionRegistry,
    pub(crate) ui_state: &'a mut UiStateStore,
    pub(crate) addin_channel: &'a Arc<Mutex<AddinChannelServer>>,
    pub(crate) connection_hub: &'a Arc<AddinConnectionHub>,
    pub(crate) command_router: &'a Arc<Mutex<CommandRouter>>,
    pub(crate) audit_log: &'a AuditLog,
    pub(crate) image_fetcher: &'a ImageFetcher,
    pub(crate) tool_access_policy: &'a ToolAccessPolicy,
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
            "tools/list" => McpCatalogResponder::tools_list(&id, context.tool_access_policy),
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
                if Self::is_forwarded_tool(tool) && !context.tool_access_policy.allows_tool(tool) {
                    let result = tool_not_available_by_policy(tool);
                    return json!({
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
                    .to_string();
                }
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
        let canonical_name = canonical_tool_name(name);
        let arguments = params.get("arguments").unwrap_or(&Value::Null);
        if Self::is_forwarded_tool(canonical_name)
            && !context.tool_access_policy.allows_tool(canonical_name)
        {
            return json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": tool_not_available_by_policy(canonical_name)
            })
            .to_string();
        }
        if Self::is_known_tool(canonical_name)
            && let Err(message) = validate_tool_arguments(canonical_name, arguments)
        {
            return json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": tool_failure_without_effect("INVALID_ARGUMENTS", &message, Some(canonical_name), false)
            })
            .to_string();
        }
        let result = match canonical_name {
            "office.list_sessions" => tool_success(&json!({
                "sessions": context
                    .registry
                    .list_sessions()
                    .iter()
                    .map(|session| SessionDescriptorView::new(session).to_json())
                    .collect::<Vec<_>>()
            })),
            "office.get_session_info" => Self::get_session_info(context.registry, arguments),
            "office.describe_tools" => Self::describe_tools(arguments),
            tool if is_office_tool(tool) => {
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
        McpForwardedToolInvoker::call(context, request_value, tool, arguments, check_capability)
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

    fn describe_tools(arguments: &Value) -> Value {
        let Some(tools) = arguments.get("tools").and_then(Value::as_array) else {
            return tool_failure(
                "INVALID_ARGUMENTS",
                "office.describe_tools requires tools.",
                Some("office.describe_tools"),
                false,
            );
        };
        let mut contracts = Vec::new();
        for tool in tools {
            let Some(tool) = tool.as_str() else {
                return tool_failure(
                    "INVALID_ARGUMENTS",
                    "office.describe_tools requires tools to contain only strings.",
                    Some("office.describe_tools"),
                    false,
                );
            };
            contracts
                .push(describe_tool_contract(tool).unwrap_or_else(|| unknown_tool_contract(tool)));
        }
        tool_success(&json!({ "tools": contracts }))
    }

    fn is_forwarded_tool(name: &str) -> bool {
        is_office_tool(name)
    }

    fn is_known_tool(name: &str) -> bool {
        matches!(
            name,
            "office.list_sessions" | "office.get_session_info" | "office.describe_tools"
        ) || Self::is_forwarded_tool(name)
    }
}

#[cfg(test)]
#[path = "mcp_rpc_tests.rs"]
mod mcp_rpc_tests;
