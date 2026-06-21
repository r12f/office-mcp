use crate::addin_mgr::PartialEffect;
use crate::api::CommandFailure;
use serde_json::{Value, json};

#[must_use]
pub fn tool_success(data: &Value) -> Value {
    json!({
        "content": [{ "type": "text", "text": data.to_string() }],
        "structuredContent": data
    })
}

#[must_use]
pub fn tool_failure(code: &str, message: &str, tool: Option<&str>, retriable: bool) -> Value {
    let error = json!({
        "office_mcp_code": code,
        "message": message,
        "tool": tool,
        "retriable": retriable,
        "partial_effect": null
    });
    tool_error_result(&error)
}

#[must_use]
pub fn tool_failure_from_command(failure: &CommandFailure) -> Value {
    let partial_effect = failure.partial_effect.map(partial_effect_json);
    let mut error = json!({
        "office_mcp_code": failure.office_mcp_code,
        "message": failure.message,
        "tool": failure.tool,
        "retriable": failure.retriable,
        "partial_effect": partial_effect
    });
    if failure.office_mcp_code == "TOOL_NOT_ENABLED_FOR_DOCUMENT" {
        error["refresh_session_info"] = json!(true);
    }
    tool_error_result(&error)
}

fn tool_error_result(error: &Value) -> Value {
    json!({
        "isError": true,
        "content": [{ "type": "text", "text": error.to_string() }],
        "structuredContent": { "error": error }
    })
}

const fn partial_effect_json(effect: PartialEffect) -> &'static str {
    match effect {
        PartialEffect::None => "none",
        PartialEffect::Possible => "possible",
        PartialEffect::Unknown => "unknown",
    }
}

#[cfg(test)]
#[path = "tool_result_tests.rs"]
mod tool_result_tests;
