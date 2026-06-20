use crate::addin_mgr::{AddinChannelServer, AddinConnectionHub, SessionRegistry};
use crate::runtime::addin_rpc_message::AddinRpcMessage;
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
        let request = AddinRpcMessage::register_request(params, id);
        let mut registry = registry.lock().expect("session registry lock");
        let mut addin_channel = addin_channel.lock().expect("addin channel lock");
        match addin_channel.register_runtime(
            &mut registry,
            connection_id,
            request,
            SystemTime::now(),
        ) {
            Ok(reply) => {
                if let Some(result) = reply.result.as_ref() {
                    connection_hub.bind_instance(connection_id, &result.assigned_instance_id);
                }
                AddinRpcMessage::register_reply_to_json(reply)
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
        let event = AddinRpcMessage::session_added_event(params);
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
        let event = AddinRpcMessage::session_updated_event(params);
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
        let event = AddinRpcMessage::session_removed_event(params);
        let mut registry = registry.lock().expect("session registry lock");
        let mut addin_channel = addin_channel.lock().expect("addin channel lock");
        let _ = addin_channel.remove_session(&mut registry, event);
    }
}

#[cfg(test)]
#[path = "addin_rpc_tests.rs"]
mod addin_rpc_tests;
