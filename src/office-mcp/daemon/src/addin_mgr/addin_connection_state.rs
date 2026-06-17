use crate::addin_mgr::addin_heartbeat::AddinHeartbeatState;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct AddinConnectionState {
    pub(crate) instance_id: String,
    pub(crate) session_id: Option<String>,
    pub(crate) heartbeat: AddinHeartbeatState,
}

impl AddinConnectionState {
    pub(crate) fn new(instance_id: String) -> Self {
        Self {
            instance_id,
            session_id: None,
            heartbeat: AddinHeartbeatState::default(),
        }
    }
}

#[cfg(test)]
#[path = "addin_connection_state_tests.rs"]
mod addin_connection_state_tests;
