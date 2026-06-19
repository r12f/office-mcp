use std::time::Duration;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddinChannelConfig {
    pub origin: String,
    pub session_grace: Duration,
    pub heartbeat_interval: Duration,
    pub heartbeat_timeout: Duration,
    pub max_pending_per_session: usize,
}

impl Default for AddinChannelConfig {
    fn default() -> Self {
        Self {
            origin: "https://localhost:8765".to_string(),
            session_grace: Duration::from_mins(1),
            heartbeat_interval: Duration::from_secs(30),
            heartbeat_timeout: Duration::from_secs(20),
            max_pending_per_session: 4,
        }
    }
}

#[cfg(test)]
#[path = "addin_channel_config_tests.rs"]
mod addin_channel_config_tests;
