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

#[test]
fn maps_ok_false_result_debug_to_command_failure() {
    let response = AddinToolResponseMapper::map(&json!({
        "result": {
            "ok": false,
            "error": {
                "office_mcp_code": "INVALID_ARGUMENT",
                "message": "Word.js InvalidArgument while running word.insert_image.",
                "tool": "word.insert_image",
                "debug": {
                    "office_error_code": "InvalidArgument",
                    "error_location": "Range.insertInlinePictureFromBase64",
                    "anchor_kind": "after_paragraph_index"
                }
            }
        }
    }));

    let ToolResponse::Failure(failure) = response else {
        panic!("expected failure response");
    };
    let debug = failure.debug.expect("debug context");
    assert_eq!(debug["office_error_code"], "InvalidArgument");
    assert_eq!(debug["anchor_kind"], "after_paragraph_index");
}

#[test]
fn maps_structured_debug_context_from_addin_error() {
    let response = AddinToolResponseMapper::map(&json!({
        "result": {
            "ok": false,
            "error": {
                "office_mcp_code": "INVALID_ARGUMENT",
                "message": "Word.js InvalidArgument while running word.insert_image.",
                "tool": "word.insert_image",
                "debug": {
                    "office_error_code": "InvalidArgument",
                    "error_location": "Range.insertInlinePictureFromBase64",
                    "anchor_kind": "after_paragraph_index",
                    "image_byte_length": 1024
                }
            }
        }
    }));

    let ToolResponse::Failure(failure) = response else {
        panic!("expected failure response");
    };
    let debug = failure.debug.expect("debug context");
    assert_eq!(debug["office_error_code"], "InvalidArgument");
    assert_eq!(debug["anchor_kind"], "after_paragraph_index");
    assert_eq!(debug["image_byte_length"], 1024);
}
