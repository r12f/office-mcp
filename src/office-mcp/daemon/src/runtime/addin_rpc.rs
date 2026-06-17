use crate::addin_mgr::{
    AddInInfo, AddinChannelServer, AddinConnectionHub, DocumentInfo, HostInfo, RegisterRequest,
    SessionAddedEvent, SessionPatch, SessionRegistry, SessionRemovedEvent, SessionRemovedReason,
    SessionUpdatedEvent,
};
use crate::runtime::json_rpc;
use serde_json::{Value, json};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

pub(crate) struct AddinJsonRpcRuntime;

impl AddinJsonRpcRuntime {
    pub(crate) fn handle_text(
        text: &str,
        connection_id: &str,
        registry: &Arc<Mutex<SessionRegistry>>,
        addin_channel: &Arc<Mutex<AddinChannelServer>>,
        connection_hub: &AddinConnectionHub,
    ) -> Option<String> {
        let value = serde_json::from_str::<Value>(text).ok()?;
        let method = value.get("method")?.as_str()?;
        match method {
            "register" => Some(Self::handle_register(
                &value,
                connection_id,
                registry,
                addin_channel,
                connection_hub,
            )),
            "session.added" => {
                Self::handle_session_added(&value, connection_id, registry, addin_channel);
                None
            }
            "session.updated" => {
                Self::handle_session_updated(&value, registry, addin_channel);
                None
            }
            "session.removed" => {
                Self::handle_session_removed(&value, registry, addin_channel);
                None
            }
            _ => value.get("id").map(|id| {
                json!({
                    "jsonrpc": "2.0",
                    "id": id,
                    "error": { "code": -32601, "message": format!("Unknown method {method}") }
                })
                .to_string()
            }),
        }
    }

    fn handle_register(
        value: &Value,
        connection_id: &str,
        registry: &Arc<Mutex<SessionRegistry>>,
        addin_channel: &Arc<Mutex<AddinChannelServer>>,
        connection_hub: &AddinConnectionHub,
    ) -> String {
        let id_value = value.get("id").cloned().unwrap_or(Value::Null);
        let id = json_rpc::id(&id_value);
        let Some(params) = value.get("params") else {
            return json_rpc::error(&id_value, -32602, "Malformed register request");
        };
        let request = RegisterRequest {
            id,
            instance_id: string_field(params, "instance_id"),
            host: HostInfo {
                app: nested_string(params, "host", "app"),
                version: nested_optional_string(params, "host", "version"),
                platform: nested_optional_string(params, "host", "platform"),
                build: nested_optional_string(params, "host", "build"),
            },
            add_in: AddInInfo {
                version: nested_string(params, "add_in", "version"),
                protocol_version: nested_string(params, "add_in", "protocol_version"),
                supported_features: nested_string_array(params, "add_in", "supported_features"),
            },
        };
        let mut registry = registry.lock().expect("session registry lock");
        let mut addin_channel = addin_channel.lock().expect("addin channel lock");
        match addin_channel.register_runtime(
            &mut registry,
            connection_id.to_string(),
            request,
            SystemTime::now(),
        ) {
            Ok(reply) => {
                if let Some(result) = reply.result.as_ref() {
                    connection_hub.bind_instance(connection_id, &result.assigned_instance_id);
                }
                register_reply_to_json(reply)
            }
            Err(error) => json_rpc::error(&id_value, -32602, &error.to_string()),
        }
    }

    fn handle_session_added(
        value: &Value,
        connection_id: &str,
        registry: &Arc<Mutex<SessionRegistry>>,
        addin_channel: &Arc<Mutex<AddinChannelServer>>,
    ) {
        let Some(params) = value.get("params") else {
            return;
        };
        let event = SessionAddedEvent {
            session_id: string_field(params, "session_id"),
            instance_id: string_field(params, "instance_id"),
            document: DocumentInfo {
                title: nested_optional_string(params, "document", "title"),
                url: nested_optional_string(params, "document", "url"),
                filename: nested_optional_string(params, "document", "filename"),
                is_dirty: nested_optional_bool(params, "document", "is_dirty"),
                is_read_only: nested_optional_bool(params, "document", "is_read_only"),
                is_protected: nested_optional_bool(params, "document", "is_protected"),
                protection: None,
            },
            available_tools: string_array_field(params, "available_tools"),
            is_active: optional_bool_field(params, "is_active"),
        };
        let mut registry = registry.lock().expect("session registry lock");
        let mut addin_channel = addin_channel.lock().expect("addin channel lock");
        let _ = addin_channel.add_session(&mut registry, connection_id, event, SystemTime::now());
    }

    fn handle_session_updated(
        value: &Value,
        registry: &Arc<Mutex<SessionRegistry>>,
        addin_channel: &Arc<Mutex<AddinChannelServer>>,
    ) {
        let Some(params) = value.get("params") else {
            return;
        };
        let patch_value = params.get("patch").unwrap_or(params);
        let event = SessionUpdatedEvent {
            session_id: string_field(params, "session_id"),
            patch: SessionPatch {
                document: parse_optional_document(patch_value),
                available_tools: optional_string_array_field(patch_value, "available_tools"),
                is_active: patch_value.get("is_active").map(serde_json::Value::as_bool),
            },
        };
        let mut registry = registry.lock().expect("session registry lock");
        let addin_channel = addin_channel.lock().expect("addin channel lock");
        let _ = addin_channel.update_session(&mut registry, event);
    }

    fn handle_session_removed(
        value: &Value,
        registry: &Arc<Mutex<SessionRegistry>>,
        addin_channel: &Arc<Mutex<AddinChannelServer>>,
    ) {
        let Some(params) = value.get("params") else {
            return;
        };
        let event = SessionRemovedEvent {
            session_id: string_field(params, "session_id"),
            reason: parse_session_removed_reason(params.get("reason").and_then(Value::as_str)),
        };
        let mut registry = registry.lock().expect("session registry lock");
        let mut addin_channel = addin_channel.lock().expect("addin channel lock");
        let _ = addin_channel.remove_session(&mut registry, event);
    }
}

fn register_reply_to_json(reply: crate::addin_mgr::JsonRpcEnvelope) -> String {
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
    let document = value.get("document")?;
    Some(DocumentInfo {
        title: optional_string_field(document, "title"),
        url: optional_string_field(document, "url"),
        filename: optional_string_field(document, "filename"),
        is_dirty: optional_bool_field(document, "is_dirty"),
        is_read_only: optional_bool_field(document, "is_read_only"),
        is_protected: optional_bool_field(document, "is_protected"),
        protection: None,
    })
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

fn nested_optional_bool(value: &Value, object: &str, name: &str) -> Option<bool> {
    value
        .get(object)
        .and_then(|nested| nested.get(name))
        .and_then(Value::as_bool)
}

fn nested_string_array(value: &Value, object: &str, name: &str) -> Vec<String> {
    value
        .get(object)
        .map_or_else(Vec::new, |nested| string_array_field(nested, name))
}

#[cfg(test)]
#[path = "addin_rpc_tests.rs"]
mod addin_rpc_tests;
