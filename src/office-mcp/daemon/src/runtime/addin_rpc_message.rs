use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, JsonRpcEnvelope, RegisterRequest, SessionAddedEvent,
    SessionPatch, SessionRemovedEvent, SessionRemovedReason, SessionUpdatedEvent,
};
use crate::runtime::json_rpc;
use serde_json::{Value, json};

pub(crate) struct AddinRpcMessage;

impl AddinRpcMessage {
    pub(crate) fn register_request(
        value: &Value,
        id: crate::addin_mgr::JsonRpcId,
    ) -> RegisterRequest {
        RegisterRequest {
            id,
            instance_id: string_field(value, "instance_id"),
            host: HostInfo {
                app: nested_string(value, "host", "app"),
                version: nested_optional_string(value, "host", "version"),
                platform: nested_optional_string(value, "host", "platform"),
                build: nested_optional_string(value, "host", "build"),
            },
            add_in: AddInInfo {
                version: nested_string(value, "add_in", "version"),
                protocol_version: nested_string(value, "add_in", "protocol_version"),
                supported_features: nested_string_array(value, "add_in", "supported_features"),
            },
        }
    }

    pub(crate) fn session_added_event(value: &Value) -> SessionAddedEvent {
        SessionAddedEvent {
            session_id: string_field(value, "session_id"),
            instance_id: string_field(value, "instance_id"),
            document: document_info(value.get("document")),
            available_tools: string_array_field(value, "available_tools"),
            is_active: optional_bool_field(value, "is_active"),
        }
    }

    pub(crate) fn session_updated_event(value: &Value) -> SessionUpdatedEvent {
        let patch_value = value.get("patch").unwrap_or(value);
        SessionUpdatedEvent {
            session_id: string_field(value, "session_id"),
            patch: SessionPatch {
                document: parse_optional_document(patch_value),
                available_tools: optional_string_array_field(patch_value, "available_tools"),
                is_active: patch_value.get("is_active").map(serde_json::Value::as_bool),
            },
        }
    }

    pub(crate) fn session_removed_event(value: &Value) -> SessionRemovedEvent {
        SessionRemovedEvent {
            session_id: string_field(value, "session_id"),
            reason: parse_session_removed_reason(value.get("reason").and_then(Value::as_str)),
        }
    }

    pub(crate) fn register_reply_to_json(reply: JsonRpcEnvelope) -> String {
        let id = reply.id.map_or(Value::Null, json_rpc::id_value);
        let Some(result) = reply.result else {
            return json!({ "jsonrpc": "2.0", "id": id, "result": null }).to_string();
        };
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": {
                "server_version": result.server_version,
                "protocol_version": result.protocol_version,
                "session_grace_sec": result.session_grace_sec,
                "heartbeat_interval_sec": result.heartbeat_interval_sec,
                "max_pending_per_session": result.max_pending_per_session,
                "assigned_instance_id": result.assigned_instance_id
            }
        })
        .to_string()
    }
}

fn document_info(value: Option<&Value>) -> DocumentInfo {
    let Some(document) = value else {
        return DocumentInfo::default();
    };
    DocumentInfo {
        title: optional_string_field(document, "title"),
        url: optional_string_field(document, "url"),
        filename: optional_string_field(document, "filename"),
        is_dirty: optional_bool_field(document, "is_dirty"),
        is_read_only: optional_bool_field(document, "is_read_only"),
        is_protected: optional_bool_field(document, "is_protected"),
        protection: None,
    }
}

fn string_field(value: &Value, name: &str) -> String {
    value
        .get(name)
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_string()
}

fn optional_bool_field(value: &Value, name: &str) -> Option<bool> {
    value.get(name).and_then(Value::as_bool)
}

fn string_array_field(value: &Value, name: &str) -> Vec<String> {
    value
        .get(name)
        .and_then(Value::as_array)
        .map(|items| {
            items
                .iter()
                .filter_map(Value::as_str)
                .map(str::to_string)
                .collect()
        })
        .unwrap_or_default()
}

fn optional_string_array_field(value: &Value, name: &str) -> Option<Vec<String>> {
    value.get(name).and_then(Value::as_array).map(|items| {
        items
            .iter()
            .filter_map(Value::as_str)
            .map(str::to_string)
            .collect()
    })
}

fn parse_optional_document(value: &Value) -> Option<DocumentInfo> {
    value
        .get("document")
        .map(|document| document_info(Some(document)))
}

fn optional_string_field(value: &Value, name: &str) -> Option<String> {
    value.get(name).and_then(Value::as_str).map(str::to_string)
}

fn parse_session_removed_reason(value: Option<&str>) -> SessionRemovedReason {
    match value {
        Some("closed") => SessionRemovedReason::Closed,
        Some("crashed") => SessionRemovedReason::Crashed,
        Some("replaced") => SessionRemovedReason::Replaced,
        _ => SessionRemovedReason::Unknown,
    }
}

fn nested_string(value: &Value, object: &str, name: &str) -> String {
    nested_optional_string(value, object, name).unwrap_or_default()
}

fn nested_optional_string(value: &Value, object: &str, name: &str) -> Option<String> {
    value
        .get(object)
        .and_then(|nested| nested.get(name))
        .and_then(Value::as_str)
        .map(str::to_string)
}

fn nested_string_array(value: &Value, object: &str, name: &str) -> Vec<String> {
    value
        .get(object)
        .map_or_else(Vec::new, |nested| string_array_field(nested, name))
}

#[cfg(test)]
#[path = "addin_rpc_message_tests.rs"]
mod addin_rpc_message_tests;
