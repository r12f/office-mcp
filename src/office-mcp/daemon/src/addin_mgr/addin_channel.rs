use crate::addin_mgr::JsonRpcEnvelope;
use crate::addin_mgr::addin_channel_clock::format_system_time;
use crate::addin_mgr::addin_heartbeat::AddinHeartbeatTimeout;
use crate::addin_mgr::addin_registration::AddinRegistrationPolicy;
use crate::addin_mgr::addin_session_events::AddinSessionEventHandler;
use crate::addin_mgr::addin_tool_payload;
use crate::addin_mgr::{
    AddinChannelConfig, AddinChannelError, AddinConnectionState, AddinUpgradeGuard,
    HeartbeatDecision, RegisterRequest, SessionAddedEvent, SessionRegistry, SessionRemovedEvent,
    SessionUpdatedEvent,
};
use crate::addin_mgr::{CancelCommand, QueuedCommand};
use std::collections::BTreeMap;
use std::time::SystemTime;

pub const SERVER_VERSION: &str = "0.1.0";
pub const ADDIN_PROTOCOL_VERSION: &str = "1.0";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddinChannelServer {
    config: AddinChannelConfig,
    upgrade_guard: AddinUpgradeGuard,
    connections: BTreeMap<String, AddinConnectionState>,
    next_ping_id: u64,
}

impl AddinChannelServer {
    #[must_use]
    pub fn new() -> Self {
        Self::with_config(AddinChannelConfig::default())
    }

    #[must_use]
    pub fn with_config(config: AddinChannelConfig) -> Self {
        let upgrade_guard = AddinUpgradeGuard::new(&config);
        Self {
            config,
            upgrade_guard,
            connections: BTreeMap::new(),
            next_ping_id: 1,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns HTTPS WSS origin checks JSON RPC heartbeat and registration"
    }

    /// Validates that a WebSocket upgrade targets the add-in endpoint and origin.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is not `/addin` or the browser `Origin`
    /// does not exactly match the configured add-in origin.
    pub fn validate_upgrade(
        &self,
        path: &str,
        origin: Option<&str>,
    ) -> Result<(), AddinChannelError> {
        self.upgrade_guard.validate(path, origin)
    }

    /// Registers a document-scoped add-in runtime.
    ///
    /// # Errors
    ///
    /// Returns an error when the register request is malformed or when the
    /// add-in protocol major version does not match the daemon version.
    pub fn register_runtime(
        &mut self,
        registry: &mut SessionRegistry,
        connection_id: &str,
        request: RegisterRequest,
        now: SystemTime,
    ) -> Result<JsonRpcEnvelope, AddinChannelError> {
        let request_id = request.id.clone();
        let registration = AddinRegistrationPolicy::new(SERVER_VERSION, ADDIN_PROTOCOL_VERSION);
        registration.validate(connection_id, &request)?;
        let runtime = AddinRegistrationPolicy::runtime_from(request, now);
        registry.register_runtime(runtime.clone());
        self.connections.insert(
            connection_id.to_string(),
            AddinConnectionState::new(runtime.instance_id.clone()),
        );
        tracing::info!(
            component = "addin_channel",
            instance_id = %runtime.instance_id,
            host_app = %runtime.host.app,
            host_platform = ?runtime.host.platform,
            addin_version = %runtime.add_in.version,
            protocol_version = %runtime.add_in.protocol_version,
            "registered add-in runtime"
        );
        Ok(JsonRpcEnvelope::success(
            request_id,
            registration.register_result(&runtime, &self.config),
        ))
    }

    /// Adds the current document session for a registered add-in connection.
    ///
    /// # Errors
    ///
    /// Returns an error when the connection is unknown, the session event is
    /// malformed, or the event instance ID does not match the runtime.
    pub fn add_session(
        &mut self,
        registry: &mut SessionRegistry,
        connection_id: &str,
        event: SessionAddedEvent,
        now: SystemTime,
    ) -> Result<(), AddinChannelError> {
        AddinSessionEventHandler::add_session(
            registry,
            &mut self.connections,
            connection_id,
            event,
            now,
        )
    }

    /// Applies add-in session metadata updates.
    ///
    /// # Errors
    ///
    /// Returns an error when the event is malformed or references an unknown session.
    pub fn update_session(
        &self,
        registry: &mut SessionRegistry,
        event: SessionUpdatedEvent,
    ) -> Result<(), AddinChannelError> {
        AddinSessionEventHandler::update_session(registry, event)
    }

    /// Removes a document session after the add-in reports it closed or was replaced.
    ///
    /// # Errors
    ///
    /// Returns an error when the event is malformed or references an unknown session.
    pub fn remove_session(
        &mut self,
        registry: &mut SessionRegistry,
        event: SessionRemovedEvent,
    ) -> Result<(), AddinChannelError> {
        AddinSessionEventHandler::remove_session(registry, &mut self.connections, event)
    }

    pub fn remove_connection(
        &mut self,
        registry: &mut SessionRegistry,
        connection_id: &str,
        stale_since: SystemTime,
    ) -> bool {
        let Some(connection) = self.connections.remove(connection_id) else {
            return false;
        };
        let removed = registry.remove_runtime(&connection.instance_id, stale_since);
        tracing::info!(
            component = "addin_channel",
            connection_id = %connection_id,
            instance_id = %connection.instance_id,
            had_session = connection.session_id.is_some(),
            removed_runtime = removed,
            "removed add-in connection"
        );
        removed
    }

    /// Builds the next heartbeat ping request for an add-in connection.
    ///
    /// # Errors
    ///
    /// Returns an error when the connection ID is unknown.
    pub fn start_ping(
        &mut self,
        connection_id: &str,
        now: SystemTime,
    ) -> Result<JsonRpcEnvelope, AddinChannelError> {
        let ping_id = self.next_ping_id();
        let connection = self
            .connections
            .get_mut(connection_id)
            .ok_or_else(|| AddinChannelError::UnknownConnection(connection_id.to_string()))?;
        connection.heartbeat.start_ping(ping_id.clone());
        Ok(JsonRpcEnvelope::request(
            ping_id,
            "ping",
            BTreeMap::from([("ts".to_string(), format_system_time(now))]),
        ))
    }

    /// Handles a heartbeat pong response.
    ///
    /// # Errors
    ///
    /// Returns an error when the connection ID is unknown.
    pub fn handle_pong(
        &mut self,
        connection_id: &str,
        response_id: &str,
    ) -> Result<bool, AddinChannelError> {
        let connection = self
            .connections
            .get_mut(connection_id)
            .ok_or_else(|| AddinChannelError::UnknownConnection(connection_id.to_string()))?;
        Ok(connection.heartbeat.handle_pong(response_id))
    }

    /// Records a missed heartbeat and decides whether to close the socket.
    ///
    /// # Errors
    ///
    /// Returns an error when the connection ID is unknown.
    pub fn record_heartbeat_timeout(
        &mut self,
        registry: &mut SessionRegistry,
        connection_id: &str,
        stale_since: SystemTime,
    ) -> Result<HeartbeatDecision, AddinChannelError> {
        let connection = self
            .connections
            .get_mut(connection_id)
            .ok_or_else(|| AddinChannelError::UnknownConnection(connection_id.to_string()))?;
        match connection.heartbeat.record_timeout() {
            AddinHeartbeatTimeout::KeepOpen { missed_pongs } => {
                tracing::warn!(
                    component = "addin_channel",
                    connection_id = %connection_id,
                    instance_id = %connection.instance_id,
                    missed_pongs = missed_pongs,
                    "add-in heartbeat missed"
                );
                Ok(HeartbeatDecision::KeepOpen)
            }
            AddinHeartbeatTimeout::Close {
                missed_pongs,
                close_code,
            } => {
                registry.mark_instance_stale(&connection.instance_id, stale_since);
                tracing::warn!(
                    component = "addin_channel",
                    connection_id = %connection_id,
                    instance_id = %connection.instance_id,
                    missed_pongs = missed_pongs,
                    close_code = close_code,
                    "closing stale add-in connection after heartbeat misses"
                );
                Ok(HeartbeatDecision::Close { code: close_code })
            }
        }
    }

    #[must_use]
    pub fn tool_invoke_payload(&self, command: &QueuedCommand) -> JsonRpcEnvelope {
        addin_tool_payload::tool_invoke_payload(command)
    }

    #[must_use]
    pub fn tool_cancel_payload(&self, cancel: &CancelCommand) -> JsonRpcEnvelope {
        addin_tool_payload::tool_cancel_payload(cancel)
    }

    fn next_ping_id(&mut self) -> String {
        let value = format!("ping-{}", self.next_ping_id);
        self.next_ping_id += 1;
        value
    }
}

impl Default for AddinChannelServer {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
#[path = "addin_channel_tests.rs"]
mod addin_channel_tests;
