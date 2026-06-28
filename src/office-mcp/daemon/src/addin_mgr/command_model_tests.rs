use super::ToolResponse;
use crate::addin_mgr::PartialEffect;
use crate::api::CommandFailure;

#[test]
fn success_response_estimates_json_payload_bytes() {
    let response = ToolResponse::Success {
        json: "{\"ok\":true}".to_string(),
    };

    assert_eq!(response.estimated_json_bytes(), 11);
}

#[test]
fn failure_response_estimate_includes_error_code_message_and_overhead() {
    let response = ToolResponse::Failure(CommandFailure {
        office_mcp_code: "MAX_RESPONSE_SIZE".to_string(),
        message: "Tool response exceeds 4096 bytes.".to_string(),
        tool: Some("word.get_text".to_string()),
        retriable: false,
        partial_effect: Some(PartialEffect::None),
        debug: None,
    });

    assert_eq!(
        response.estimated_json_bytes(),
        "MAX_RESPONSE_SIZE".len() + "Tool response exceeds 4096 bytes.".len() + 128
    );
}
