use std::collections::BTreeMap;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum JsonRpcId {
    String(String),
    Number(i64),
    Null,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RegisterResult {
    pub server_version: String,
    pub protocol_version: String,
    pub session_grace_sec: u64,
    pub heartbeat_interval_sec: u64,
    pub max_pending_per_session: usize,
    pub assigned_instance_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct JsonRpcEnvelope {
    pub id: Option<JsonRpcId>,
    pub method: Option<String>,
    pub params: BTreeMap<String, String>,
    pub result: Option<RegisterResult>,
}

impl JsonRpcEnvelope {
    #[must_use]
    pub(crate) fn request(id: String, method: &str, params: BTreeMap<String, String>) -> Self {
        Self {
            id: Some(JsonRpcId::String(id)),
            method: Some(method.to_string()),
            params,
            result: None,
        }
    }

    #[must_use]
    pub(crate) fn notification(method: &str, params: BTreeMap<String, String>) -> Self {
        Self {
            id: None,
            method: Some(method.to_string()),
            params,
            result: None,
        }
    }

    #[must_use]
    pub(crate) fn success(id: JsonRpcId, result: RegisterResult) -> Self {
        Self {
            id: Some(id),
            method: None,
            params: BTreeMap::new(),
            result: Some(result),
        }
    }
}

#[cfg(test)]
#[path = "addin_json_rpc_tests.rs"]
mod addin_json_rpc_tests;
