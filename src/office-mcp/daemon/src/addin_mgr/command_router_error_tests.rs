use super::CommandRouterError;
use crate::addin_mgr::{OfficeMcpCode, PartialEffect, ToolFailure, ToolInvocationError};

#[test]
fn preflight_error_maps_registry_failure_fields() {
    let error = CommandRouterError::Preflight(ToolInvocationError {
        failure: ToolFailure {
            office_mcp_code: OfficeMcpCode::SessionLost,
            message: "Session session-1 is disconnected.".to_string(),
            session_id: Some("session-1".to_string()),
            tool: Some("word.get_text".to_string()),
            retriable: true,
            partial_effect: Some(PartialEffect::Unknown),
        },
    });

    let failure = error.as_command_failure("word.get_text");

    assert_eq!(failure.office_mcp_code, "SESSION_LOST");
    assert_eq!(failure.message, "Session session-1 is disconnected.");
    assert_eq!(failure.tool.as_deref(), Some("word.get_text"));
    assert!(failure.retriable);
    assert_eq!(failure.partial_effect, Some(PartialEffect::Unknown));
    assert_eq!(error.to_string(), "Session session-1 is disconnected.");
}

#[test]
fn unknown_request_maps_to_internal_bug() {
    let error = CommandRouterError::UnknownRequest("request-404".to_string());

    let failure = error.as_command_failure("word.add_comment");

    assert_eq!(failure.office_mcp_code, "INTERNAL_BUG");
    assert_eq!(failure.message, "Unknown command request request-404.");
    assert_eq!(failure.tool.as_deref(), Some("word.add_comment"));
    assert!(!failure.retriable);
    assert_eq!(failure.partial_effect, None);
    assert_eq!(error.to_string(), "Unknown command request request-404.");
}

#[test]
fn oversized_response_maps_to_max_response_size() {
    let error = CommandRouterError::ResponseTooLarge {
        max_response_bytes: 4096,
        actual_bytes: 8192,
    };

    let failure = error.as_command_failure("word.insert_image");

    assert_eq!(failure.office_mcp_code, "MAX_RESPONSE_SIZE");
    assert_eq!(failure.message, "Tool response exceeds 4096 bytes.");
    assert_eq!(failure.tool.as_deref(), Some("word.insert_image"));
    assert!(!failure.retriable);
    assert_eq!(failure.partial_effect, Some(PartialEffect::None));
    assert_eq!(
        error.to_string(),
        "Response exceeded 4096 bytes: 8192 bytes."
    );
}
