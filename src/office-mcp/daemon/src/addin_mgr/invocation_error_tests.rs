use super::{OfficeMcpCode, PartialEffect, ToolInvocationError};

#[test]
fn office_mcp_codes_keep_protocol_strings() {
    assert_eq!(OfficeMcpCode::NoSessions.as_str(), "NO_SESSIONS");
    assert_eq!(OfficeMcpCode::SessionNotFound.as_str(), "SESSION_NOT_FOUND");
    assert_eq!(OfficeMcpCode::SessionStale.as_str(), "SESSION_STALE");
    assert_eq!(OfficeMcpCode::SessionLost.as_str(), "SESSION_LOST");
    assert_eq!(
        OfficeMcpCode::MaxPendingExceeded.as_str(),
        "MAX_PENDING_EXCEEDED"
    );
    assert_eq!(
        OfficeMcpCode::HostCapabilityUnavailable.as_str(),
        "HOST_CAPABILITY_UNAVAILABLE"
    );
    assert_eq!(
        OfficeMcpCode::ToolNotEnabledForDocument.as_str(),
        "TOOL_NOT_ENABLED_FOR_DOCUMENT"
    );
}

#[test]
fn invocation_error_formats_user_actionable_message() {
    let error = ToolInvocationError::new(OfficeMcpCode::NoSessions, "session-1", "word.get_text");

    assert_eq!(error.failure.office_mcp_code, OfficeMcpCode::NoSessions);
    assert!(error.failure.retriable);
    assert_eq!(error.failure.partial_effect, None);
    assert_eq!(
        error.to_string(),
        "No Office document sessions are connected. Open Office MCP Control in Word and try again."
    );
    assert!(!error.to_string().contains("office-mcp add-in"));
}

#[test]
fn lost_session_records_unknown_partial_effect() {
    let error = ToolInvocationError::new(OfficeMcpCode::SessionLost, "session-1", "word.get_text");

    assert_eq!(error.failure.session_id.as_deref(), Some("session-1"));
    assert_eq!(error.failure.tool.as_deref(), Some("word.get_text"));
    assert!(!error.failure.retriable);
    assert_eq!(error.failure.partial_effect, Some(PartialEffect::Unknown));
    assert_eq!(
        error.failure.message,
        "Session session-1 lost its add-in connection."
    );
}

#[test]
fn missing_capability_message_names_tool() {
    let error = ToolInvocationError::new(
        OfficeMcpCode::HostCapabilityUnavailable,
        "session-1",
        "word.insert_image",
    );

    assert_eq!(
        error.failure.message,
        "The selected Office session does not support word.insert_image."
    );
}

#[test]
fn disabled_document_tool_message_uses_unified_wording() {
    let error = ToolInvocationError::new(
        OfficeMcpCode::ToolNotEnabledForDocument,
        "session-1",
        "word.get_text",
    );

    assert_eq!(
        error.failure.message,
        "Tool word.get_text is disabled for this document session. Refresh office.get_session_info or office.list_sessions before retrying."
    );
    assert!(!error.failure.message.contains("not enabled"));
}
