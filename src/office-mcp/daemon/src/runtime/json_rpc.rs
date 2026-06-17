use crate::addin_mgr::JsonRpcId;
use serde_json::{Value, json};

#[must_use]
pub(crate) fn envelope_to_text(envelope: &crate::addin_mgr::JsonRpcEnvelope) -> String {
    let mut value = json!({ "jsonrpc": "2.0" });
    if let Some(id) = envelope.id.clone() {
        value["id"] = id_value(id);
    }
    if let Some(method) = envelope.method.as_ref() {
        value["method"] = Value::String(method.clone());
    }
    if !envelope.params.is_empty() {
        let params = envelope
            .params
            .iter()
            .map(|(key, value)| {
                let parsed = serde_json::from_str::<Value>(value)
                    .unwrap_or_else(|_| Value::String(value.clone()));
                (key.clone(), parsed)
            })
            .collect::<serde_json::Map<_, _>>();
        value["params"] = Value::Object(params);
    }
    value.to_string()
}

#[must_use]
pub(crate) fn id(value: &Value) -> JsonRpcId {
    if let Some(text) = value.as_str() {
        JsonRpcId::String(text.to_string())
    } else if let Some(number) = value.as_i64() {
        JsonRpcId::Number(number)
    } else {
        JsonRpcId::Null
    }
}

#[must_use]
pub(crate) fn id_value(id: JsonRpcId) -> Value {
    match id {
        JsonRpcId::String(value) => Value::String(value),
        JsonRpcId::Number(value) => Value::Number(value.into()),
        JsonRpcId::Null => Value::Null,
    }
}

#[must_use]
pub(crate) fn error(id: &Value, code: i64, message: &str) -> String {
    json!({
        "jsonrpc": "2.0",
        "id": id,
        "error": { "code": code, "message": message }
    })
    .to_string()
}
