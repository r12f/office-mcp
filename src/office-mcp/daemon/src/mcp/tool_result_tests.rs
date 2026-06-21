use super::{tool_failure, tool_failure_from_command, tool_success};
use crate::addin_mgr::PartialEffect;
use crate::api::CommandFailure;
use serde_json::json;

#[test]
fn success_result_contains_text_and_structured_content() {
    let data = json!({ "session_id": "session-1", "count": 2 });

    let result = tool_success(&data);

    assert_eq!(result["structuredContent"], data);
    assert_eq!(result["content"][0]["type"], "text");
    assert_eq!(result["content"][0]["text"], data.to_string());
}

#[test]
fn failure_result_contains_structured_error() {
    let result = tool_failure("UNKNOWN_TOOL", "No such tool.", Some("word.nope"), false);

    assert_eq!(result["isError"], true);
    assert_eq!(
        result["structuredContent"]["error"]["office_mcp_code"],
        "UNKNOWN_TOOL"
    );
    assert_eq!(result["structuredContent"]["error"]["tool"], "word.nope");
    assert_eq!(
        result["structuredContent"]["error"]["partial_effect"],
        json!(null)
    );
    assert_eq!(
        result["content"][0]["text"],
        result["structuredContent"]["error"].to_string()
    );
}

#[test]
fn command_failure_result_maps_partial_effect() {
    let result = tool_failure_from_command(&CommandFailure {
        office_mcp_code: "TIMEOUT".to_string(),
        message: "Timed out".to_string(),
        tool: Some("word.insert_paragraph".to_string()),
        retriable: true,
        partial_effect: Some(PartialEffect::Possible),
    });

    assert_eq!(result["isError"], true);
    assert_eq!(
        result["structuredContent"]["error"]["office_mcp_code"],
        "TIMEOUT"
    );
    assert_eq!(result["structuredContent"]["error"]["retriable"], true);
    assert_eq!(
        result["structuredContent"]["error"]["partial_effect"],
        "possible"
    );
}

#[test]
fn disabled_document_tool_failure_includes_session_refresh_hint() {
    let result = tool_failure_from_command(&CommandFailure {
        office_mcp_code: "TOOL_NOT_ENABLED_FOR_DOCUMENT".to_string(),
        message: "Tool word.get_text is disabled for this document session. Refresh office.get_session_info or office.list_sessions before retrying.".to_string(),
        tool: Some("word.get_text".to_string()),
        retriable: false,
        partial_effect: None,
    });

    assert_eq!(
        result["structuredContent"]["error"]["office_mcp_code"],
        "TOOL_NOT_ENABLED_FOR_DOCUMENT"
    );
    assert_eq!(
        result["structuredContent"]["error"]["refresh_session_info"],
        true
    );
    assert!(
        result["structuredContent"]["error"]["message"]
            .as_str()
            .expect("message")
            .starts_with("Tool word.get_text is disabled")
    );
}
