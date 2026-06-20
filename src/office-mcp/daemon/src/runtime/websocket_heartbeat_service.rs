use crate::addin_mgr::{AddinChannelError, HeartbeatDecision};
use crate::runtime::RuntimeServerError;
use crate::runtime::json_rpc;
use crate::runtime::mcp_response::{HeartbeatLoopDecision, RuntimeSharedState};
use serde_json::Value;
use std::time::SystemTime;

pub(crate) struct WebSocketHeartbeatService;

impl WebSocketHeartbeatService {
    pub(crate) fn start_ping(
        shared_state: &RuntimeSharedState,
        connection_id: &str,
    ) -> Result<Option<String>, RuntimeServerError> {
        let mut addin_channel = shared_state.addin_channel.lock().map_err(|_| {
            RuntimeServerError::Internal("Add-in channel lock poisoned.".to_string())
        })?;
        match addin_channel.start_ping(connection_id, SystemTime::now()) {
            Ok(ping) => Ok(Some(json_rpc::envelope_to_text(&ping))),
            Err(AddinChannelError::UnknownConnection(_)) => Ok(None),
            Err(error) => Err(RuntimeServerError::Internal(error.to_string())),
        }
    }

    pub(crate) fn handle_response(
        shared_state: &RuntimeSharedState,
        connection_id: &str,
        text: &str,
    ) -> Result<bool, RuntimeServerError> {
        let Ok(value) = serde_json::from_str::<Value>(text) else {
            return Ok(false);
        };
        if value.get("method").is_some() || value.get("result").is_none() {
            return Ok(false);
        }
        let Some(response_id) = value.get("id").and_then(Value::as_str) else {
            return Ok(false);
        };
        let mut addin_channel = shared_state.addin_channel.lock().map_err(|_| {
            RuntimeServerError::Internal("Add-in channel lock poisoned.".to_string())
        })?;
        addin_channel
            .handle_pong(connection_id, response_id)
            .map_err(|error| RuntimeServerError::Internal(error.to_string()))
    }

    pub(crate) fn record_timeout(
        shared_state: &RuntimeSharedState,
        connection_id: &str,
    ) -> Result<HeartbeatLoopDecision, RuntimeServerError> {
        let mut registry = shared_state.registry.lock().map_err(|_| {
            RuntimeServerError::Internal("Session registry lock poisoned.".to_string())
        })?;
        let mut addin_channel = shared_state.addin_channel.lock().map_err(|_| {
            RuntimeServerError::Internal("Add-in channel lock poisoned.".to_string())
        })?;
        match addin_channel.record_heartbeat_timeout(
            &mut registry,
            connection_id,
            SystemTime::now(),
        ) {
            Ok(HeartbeatDecision::Close { .. }) => Ok(HeartbeatLoopDecision::Close),
            Ok(HeartbeatDecision::KeepOpen) | Err(AddinChannelError::UnknownConnection(_)) => {
                Ok(HeartbeatLoopDecision::KeepOpen)
            }
            Err(error) => Err(RuntimeServerError::Internal(error.to_string())),
        }
    }
}

#[cfg(test)]
#[path = "websocket_heartbeat_service_tests.rs"]
mod websocket_heartbeat_service_tests;
