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
