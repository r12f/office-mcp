use serde_json::Value;
use std::collections::{BTreeMap, VecDeque};
use std::sync::{Condvar, Mutex};
use std::time::{Duration, SystemTime};

#[derive(Debug, Default)]
pub struct AddinConnectionHub {
    state: Mutex<AddinConnectionHubState>,
    response_available: Condvar,
}

impl AddinConnectionHub {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    pub fn register_connection(&self, connection_id: &str) {
        let mut state = self.state.lock().expect("addin connection hub lock");
        state
            .connections
            .entry(connection_id.to_string())
            .or_default();
    }

    pub fn bind_instance(&self, connection_id: &str, instance_id: &str) {
        let mut state = self.state.lock().expect("addin connection hub lock");
        state
            .connections
            .entry(connection_id.to_string())
            .or_default()
            .instance_id = Some(instance_id.to_string());
        state
            .connection_by_instance
            .insert(instance_id.to_string(), connection_id.to_string());
    }

    pub fn remove_connection(&self, connection_id: &str) {
        let mut state = self.state.lock().expect("addin connection hub lock");
        if let Some(connection) = state.connections.remove(connection_id)
            && let Some(instance_id) = connection.instance_id
        {
            state.connection_by_instance.remove(&instance_id);
        }
        self.response_available.notify_all();
    }

    pub fn invoke(
        &self,
        instance_id: &str,
        request_id: &str,
        payload: String,
        timeout: Duration,
    ) -> Result<Value, AddinConnectionHubError> {
        let deadline = SystemTime::now() + timeout;
        let mut state = self.state.lock().expect("addin connection hub lock");
        let connection_id = state
            .connection_by_instance
            .get(instance_id)
            .cloned()
            .ok_or(AddinConnectionHubError::NoConnection)?;
        let connection = state
            .connections
            .get_mut(&connection_id)
            .ok_or(AddinConnectionHubError::NoConnection)?;
        connection.outbound.push_back(payload);
        state.pending.insert(request_id.to_string(), None);
        self.response_available.notify_all();

        loop {
            if let Some(response) = state.pending.get_mut(request_id).and_then(Option::take) {
                state.pending.remove(request_id);
                return Ok(response);
            }
            let now = SystemTime::now();
            let remaining = deadline
                .duration_since(now)
                .map_err(|_| AddinConnectionHubError::Timeout)?;
            let (next_state, wait_result) = self
                .response_available
                .wait_timeout(state, remaining)
                .expect("addin connection hub condvar");
            state = next_state;
            if wait_result.timed_out() {
                state.pending.remove(request_id);
                return Err(AddinConnectionHubError::Timeout);
            }
        }
    }

    #[must_use]
    pub fn take_outbound(&self, connection_id: &str) -> Vec<String> {
        let mut state = self.state.lock().expect("addin connection hub lock");
        state
            .connections
            .get_mut(connection_id)
            .map(|connection| connection.outbound.drain(..).collect())
            .unwrap_or_default()
    }

    pub fn send_to_instance(&self, instance_id: &str, payload: String) -> bool {
        let mut state = self.state.lock().expect("addin connection hub lock");
        let Some(connection_id) = state.connection_by_instance.get(instance_id).cloned() else {
            return false;
        };
        let Some(connection) = state.connections.get_mut(&connection_id) else {
            return false;
        };
        connection.outbound.push_back(payload);
        self.response_available.notify_all();
        true
    }

    pub fn complete_from_text(&self, text: &str) -> bool {
        let Ok(value) = serde_json::from_str::<Value>(text) else {
            return false;
        };
        let Some(request_id) = value.get("id").and_then(Value::as_str) else {
            return false;
        };
        let mut state = self.state.lock().expect("addin connection hub lock");
        if let Some(slot) = state.pending.get_mut(request_id) {
            *slot = Some(value);
            self.response_available.notify_all();
            return true;
        }
        false
    }
}

#[derive(Debug, Default)]
struct AddinConnectionHubState {
    connections: BTreeMap<String, AddinConnectionHubConnection>,
    connection_by_instance: BTreeMap<String, String>,
    pending: BTreeMap<String, Option<Value>>,
}

#[derive(Debug, Default)]
struct AddinConnectionHubConnection {
    instance_id: Option<String>,
    outbound: VecDeque<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AddinConnectionHubError {
    NoConnection,
    Timeout,
}

#[cfg(test)]
#[path = "connection_hub_tests.rs"]
mod connection_hub_tests;
