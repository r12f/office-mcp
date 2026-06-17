use super::AddinToolResponseMapper;
use crate::addin_mgr::{PartialEffect, ToolResponse};
use serde_json::json;

#[test]
fn maps_success_data_to_tool_response_json() {
    let response = AddinToolResponseMapper::map(&json!({
        "result": {
            "ok": true,
            "data": { "text": "Hello" }
        }
    }));

    assert_eq!(
        response,
        ToolResponse::Success {
            json: json!({ "text": "Hello" }).to_string()
        }
    );
}

#[test]
fn maps_json_rpc_error_to_command_failure() {
    let response = AddinToolResponseMapper::map(&json!({
        "error": {
            "code": "HOST_ERROR",
            "message": "Host rejected the command.",
            "tool": "word.get_text",
            "retriable": true
        }
    }));

    let ToolResponse::Failure(failure) = response else {
        panic!("expected failure response");
    };
    assert_eq!(failure.office_mcp_code, "HOST_ERROR");
    assert_eq!(failure.message, "Host rejected the command.");
    assert_eq!(failure.tool.as_deref(), Some("word.get_text"));
    assert!(failure.retriable);
    assert_eq!(failure.partial_effect, Some(PartialEffect::Unknown));
}

#[test]
fn maps_ok_false_result_to_command_failure() {
    let response = AddinToolResponseMapper::map(&json!({
        "result": {
            "ok": false,
            "error": {
                "office_mcp_code": "INVALID_ARGUMENTS",
                "message": "Range is invalid.",
                "tool": "word.replace_text"
            }
        }
    }));

    let ToolResponse::Failure(failure) = response else {
        panic!("expected failure response");
    };
    assert_eq!(failure.office_mcp_code, "INVALID_ARGUMENTS");
    assert_eq!(failure.message, "Range is invalid.");
    assert_eq!(failure.tool.as_deref(), Some("word.replace_text"));
    assert!(!failure.retriable);
    assert_eq!(failure.partial_effect, Some(PartialEffect::Unknown));
}
