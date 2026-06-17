use super::{tool_cancel_payload, tool_invoke_payload};
use crate::addin_mgr::{CancelCommand, QueuedCommand};
use std::time::{Duration, SystemTime};

#[test]
fn builds_tool_invoke_payload() {
    let command = QueuedCommand {
        command_id: "command-1".to_string(),
        request_id: "request-1".to_string(),
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        tool: "word.get_text".to_string(),
        arguments_json: "{}".to_string(),
        timeout: Duration::from_secs(30),
        enqueued_at: SystemTime::UNIX_EPOCH,
        deadline_at: SystemTime::UNIX_EPOCH + Duration::from_secs(30),
        dispatched: false,
        sequence: 0,
    };

    let payload = tool_invoke_payload(&command);

    assert_eq!(payload.method.as_deref(), Some("tool.invoke"));
    assert_eq!(payload.params["session_id"], "session-1");
    assert_eq!(payload.params["tool"], "word.get_text");
    assert_eq!(payload.params["timeout_ms"], "30000");
}

#[test]
fn builds_tool_cancel_payload() {
    let payload = tool_cancel_payload(&CancelCommand {
        request_id: "request-1".to_string(),
        reason: "timeout".to_string(),
    });

    assert_eq!(payload.id, None);
    assert_eq!(payload.method.as_deref(), Some("tool.cancel"));
    assert_eq!(payload.params["request_id"], "request-1");
    assert_eq!(payload.params["reason"], "timeout");
}
