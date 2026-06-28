use crate::addin_mgr::{PartialEffect, ToolResponse};
use crate::api::CommandFailure;
use serde_json::{Map, Value};

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
                partial_effect: Some(partial_effect_from_error(error)),
                debug: safe_debug_context(error.get("debug")),
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
                partial_effect: Some(partial_effect_from_error(&error)),
                debug: safe_debug_context(error.get("debug")),
            });
        }
        let data = result.get("data").cloned().unwrap_or(result);
        ToolResponse::Success {
            json: data.to_string(),
        }
    }
}

fn partial_effect_from_error(error: &Value) -> PartialEffect {
    match error.get("partial_effect").and_then(Value::as_str) {
        Some("none") => PartialEffect::None,
        Some("possible") => PartialEffect::Possible,
        _ => PartialEffect::Unknown,
    }
}

fn safe_debug_context(debug: Option<&Value>) -> Option<Value> {
    let object = debug?.as_object()?;
    let mut filtered = Map::new();
    for key in [
        "office_error_code",
        "office_error_message",
        "office_error_location",
        "error_location",
        "statement",
        "tool",
        "anchor_kind",
        "target_object_type",
        "placement",
        "extent",
        "action",
        "image_mime_type",
        "image_byte_length",
        "width_pt",
        "height_pt",
        "index",
        "table_index",
        "row",
        "col",
        "content_control_id",
        "hint",
    ] {
        if let Some(value) = object.get(key).and_then(safe_debug_value) {
            filtered.insert(key.to_string(), value);
        }
    }
    (!filtered.is_empty()).then_some(Value::Object(filtered))
}

fn safe_debug_value(value: &Value) -> Option<Value> {
    match value {
        Value::String(text) if looks_sensitive(text) => None,
        Value::String(text) => Some(Value::String(text.chars().take(240).collect())),
        Value::Number(_) | Value::Bool(_) => Some(value.clone()),
        _ => None,
    }
}

fn looks_sensitive(value: &str) -> bool {
    value.contains("base64") || value.contains("data:image") || has_long_base64_like_run(value)
}

fn has_long_base64_like_run(value: &str) -> bool {
    let mut run = 0;
    for character in value.chars() {
        if character.is_ascii_alphanumeric() || matches!(character, '+' | '/' | '=') {
            run += 1;
            if run >= 80 {
                return true;
            }
        } else {
            run = 0;
        }
    }
    false
}

#[cfg(test)]
#[path = "addin_tool_response_tests.rs"]
mod addin_tool_response_tests;
