use crate::addin_mgr::{CancelCommand, JsonRpcEnvelope, QueuedCommand};
use std::collections::BTreeMap;

#[must_use]
pub(crate) fn tool_invoke_payload(command: &QueuedCommand) -> JsonRpcEnvelope {
    JsonRpcEnvelope::request(
        command.request_id.clone(),
        "tool.invoke",
        BTreeMap::from([
            ("session_id".to_string(), command.session_id.clone()),
            ("tool".to_string(), command.tool.clone()),
            ("args".to_string(), command.arguments_json.clone()),
            (
                "timeout_ms".to_string(),
                command.timeout.as_millis().to_string(),
            ),
        ]),
    )
}

#[must_use]
pub(crate) fn tool_cancel_payload(cancel: &CancelCommand) -> JsonRpcEnvelope {
    JsonRpcEnvelope::notification(
        "tool.cancel",
        BTreeMap::from([
            ("request_id".to_string(), cancel.request_id.clone()),
            ("reason".to_string(), cancel.reason.clone()),
        ]),
    )
}

#[cfg(test)]
#[path = "addin_tool_payload_tests.rs"]
mod addin_tool_payload_tests;
