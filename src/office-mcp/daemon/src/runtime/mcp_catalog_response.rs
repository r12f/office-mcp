use crate::addin_mgr::SessionRegistry;
use crate::mcp::{
    prompt_catalog_json, prompt_description, prompt_messages, tool_catalog_json,
    word_resource_catalog_for_session, word_resource_templates,
};
use crate::runtime::json_rpc;
use serde_json::{Value, json};

pub(crate) struct McpCatalogResponder;

impl McpCatalogResponder {
    #[must_use]
    pub(crate) fn tools_list(id: &Value) -> String {
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": { "tools": tool_catalog_json() }
        })
        .to_string()
    }

    #[must_use]
    pub(crate) fn resources_list(registry: &SessionRegistry, id: &Value) -> String {
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

    #[must_use]
    pub(crate) fn resource_templates_list(id: &Value) -> String {
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": { "resourceTemplates": word_resource_templates() }
        })
        .to_string()
    }

    #[must_use]
    pub(crate) fn prompts_list(id: &Value) -> String {
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": { "prompts": prompt_catalog_json() }
        })
        .to_string()
    }

    #[must_use]
    pub(crate) fn prompts_get(id: &Value, value: &Value) -> String {
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
}

#[cfg(test)]
#[path = "mcp_catalog_response_tests.rs"]
mod mcp_catalog_response_tests;
