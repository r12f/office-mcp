use crate::addin_mgr::{PartialEffect, ToolResponse};
use crate::api::CommandFailure;
use serde_json::Value;

pub(crate) struct AddinToolResponseMapper;

impl AddinToolResponseMapper {
    #[must_use]
    pub(crate) fn map(response: &Value) -> ToolResponse {
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
                partial_effect: Some(PartialEffect::Unknown),
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
                partial_effect: Some(PartialEffect::Unknown),
            });
        }
        let data = result.get("data").cloned().unwrap_or(result);
        ToolResponse::Success {
            json: data.to_string(),
        }
    }
}

#[cfg(test)]
#[path = "addin_tool_response_tests.rs"]
mod addin_tool_response_tests;
