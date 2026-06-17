use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionPatch, SessionRegistry,
};
use crate::addin_mgr::{CancelCommand, QueuedCommand};
use crate::addin_mgr::{JsonRpcEnvelope, JsonRpcId, RegisterResult};
use std::collections::BTreeMap;
use std::fmt::{Display, Formatter};
use std::time::{Duration, SystemTime};

pub const SERVER_VERSION: &str = "0.1.0";
pub const ADDIN_PROTOCOL_VERSION: &str = "1.0";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddinChannelServer {
    config: AddinChannelConfig,
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
        Self {
            config,
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
        if path != "/addin" {
            tracing::warn!(
                component = "addin_channel",
                path = %path,
                origin = ?origin,
                "rejected add-in websocket upgrade path"
            );
            return Err(AddinChannelError::InvalidUpgradePath(path.to_string()));
        }
        if origin != Some(self.config.origin.as_str()) {
            tracing::warn!(
                component = "addin_channel",
                path = %path,
                origin = ?origin,
                expected_origin = %self.config.origin,
                "rejected add-in websocket origin"
            );
            return Err(AddinChannelError::ForbiddenOrigin(
                origin.unwrap_or_default().to_string(),
            ));
        }
        tracing::debug!(
            component = "addin_channel",
            path = %path,
            origin = ?origin,
            "accepted add-in websocket upgrade"
        );
        Ok(())
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
        connection_id: String,
        request: RegisterRequest,
        now: SystemTime,
    ) -> Result<JsonRpcEnvelope, AddinChannelError> {
        if request.instance_id.is_empty()
            || request.host.app.is_empty()
            || request.add_in.protocol_version.is_empty()
        {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                instance_id = %request.instance_id,
                host_app = %request.host.app,
                "rejected malformed add-in register request"
            );
            return Err(AddinChannelError::MalformedRegister);
        }
        if !same_major(&request.add_in.protocol_version, ADDIN_PROTOCOL_VERSION) {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                instance_id = %request.instance_id,
                offered = %request.add_in.protocol_version,
                supported = %ADDIN_PROTOCOL_VERSION,
                "rejected add-in protocol version"
            );
            return Err(AddinChannelError::ProtocolVersionMismatch {
                offered: request.add_in.protocol_version,
                supported: ADDIN_PROTOCOL_VERSION.to_string(),
            });
        }
        let runtime = RuntimeInfo {
            instance_id: request.instance_id.clone(),
            host: request.host,
            add_in: request.add_in,
            registered_at: now,
        };
        registry.register_runtime(runtime.clone());
        self.connections.insert(
            connection_id,
            AddinConnectionState {
                instance_id: runtime.instance_id.clone(),
                session_id: None,
                pending_ping_id: None,
                missed_pongs: 0,
            },
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
            request.id,
            RegisterResult {
                server_version: SERVER_VERSION.to_string(),
                protocol_version: ADDIN_PROTOCOL_VERSION.to_string(),
                session_grace_sec: self.config.session_grace.as_secs(),
                heartbeat_interval_sec: self.config.heartbeat_interval.as_secs(),
                max_pending_per_session: self.config.max_pending_per_session,
                assigned_instance_id: runtime.instance_id,
            },
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
        let connection = self
            .connections
            .get_mut(connection_id)
            .ok_or_else(|| AddinChannelError::UnknownConnection(connection_id.to_string()))?;
        if event.instance_id != connection.instance_id {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                expected_instance_id = %connection.instance_id,
                actual_instance_id = %event.instance_id,
                session_id = %event.session_id,
                "rejected add-in session for wrong instance"
            );
            return Err(AddinChannelError::InstanceMismatch {
                expected: connection.instance_id.clone(),
                actual: event.instance_id,
            });
        }
        if event.session_id.is_empty() {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                instance_id = %connection.instance_id,
                "rejected malformed add-in session added event"
            );
            return Err(AddinChannelError::MalformedSessionEvent);
        }
        let tool_count = event.available_tools.len();
        let session = registry.add_session(
            NewSessionInfo {
                session_id: event.session_id.clone(),
                instance_id: connection.instance_id.clone(),
                document: event.document,
                available_tools: event.available_tools,
                is_active: event.is_active,
            },
            now,
        );
        connection.session_id = Some(session.session_id);
        tracing::info!(
            component = "addin_channel",
            connection_id = %connection_id,
            instance_id = %connection.instance_id,
            session_id = %event.session_id,
            tool_count = tool_count,
            is_active = ?event.is_active,
            "added add-in document session"
        );
        Ok(())
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
        if event.session_id.is_empty() {
            tracing::warn!(
                component = "addin_channel",
                "rejected malformed add-in session updated event"
            );
            return Err(AddinChannelError::MalformedSessionEvent);
        }
        if registry.update_session(&event.session_id, event.patch) {
            tracing::debug!(
                component = "addin_channel",
                session_id = %event.session_id,
                "updated add-in document session"
            );
            Ok(())
        } else {
            tracing::warn!(
                component = "addin_channel",
                session_id = %event.session_id,
                "rejected update for unknown add-in session"
            );
            Err(AddinChannelError::UnknownSession(event.session_id))
        }
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
        if event.session_id.is_empty() {
            tracing::warn!(
                component = "addin_channel",
                reason = ?event.reason,
                "rejected malformed add-in session removed event"
            );
            return Err(AddinChannelError::MalformedSessionEvent);
        }
        for connection in self.connections.values_mut() {
            if connection.session_id.as_deref() == Some(event.session_id.as_str()) {
                connection.session_id = None;
            }
        }
        if registry.remove_session(&event.session_id) {
            tracing::info!(
                component = "addin_channel",
                session_id = %event.session_id,
                reason = ?event.reason,
                "removed add-in document session"
            );
            Ok(())
        } else {
            tracing::warn!(
                component = "addin_channel",
                session_id = %event.session_id,
                reason = ?event.reason,
                "rejected removal for unknown add-in session"
            );
            Err(AddinChannelError::UnknownSession(event.session_id))
        }
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
        connection.pending_ping_id = Some(ping_id.clone());
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
        if connection.pending_ping_id.as_deref() != Some(response_id) {
            return Ok(false);
        }
        connection.pending_ping_id = None;
        connection.missed_pongs = 0;
        Ok(true)
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
        connection.pending_ping_id = None;
        connection.missed_pongs += 1;
        if connection.missed_pongs < 2 {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                instance_id = %connection.instance_id,
                missed_pongs = connection.missed_pongs,
                "add-in heartbeat missed"
            );
            return Ok(HeartbeatDecision::KeepOpen);
        }
        registry.mark_instance_stale(&connection.instance_id, stale_since);
        tracing::warn!(
            component = "addin_channel",
            connection_id = %connection_id,
            instance_id = %connection.instance_id,
            missed_pongs = connection.missed_pongs,
            close_code = 4002,
            "closing stale add-in connection after heartbeat misses"
        );
        Ok(HeartbeatDecision::Close { code: 4002 })
    }

    #[must_use]
    pub fn tool_invoke_payload(&self, command: &QueuedCommand) -> JsonRpcEnvelope {
        JsonRpcEnvelope::request(
            command.request_id.clone(),
            "tool.invoke",
            BTreeMap::from([
                ("session_id".to_string(), command.session_id.clone()),
                ("tool".to_string(), command.tool.clone()),
                ("args".to_string(), command.arguments_json.clone()),
                (
                    "timeout_ms".to_string(),
                    command.timeout.as_millis().to_string(),
                ),
            ]),
        )
    }

    #[must_use]
    pub fn tool_cancel_payload(&self, cancel: &CancelCommand) -> JsonRpcEnvelope {
        JsonRpcEnvelope::notification(
            "tool.cancel",
            BTreeMap::from([
                ("request_id".to_string(), cancel.request_id.clone()),
                ("reason".to_string(), cancel.reason.clone()),
            ]),
        )
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
            heartbeat_timeout: Duration::from_secs(10),
            max_pending_per_session: 4,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AddinConnectionState {
    instance_id: String,
    session_id: Option<String>,
    pending_ping_id: Option<String>,
    missed_pongs: u8,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RegisterRequest {
    pub id: JsonRpcId,
    pub instance_id: String,
    pub host: HostInfo,
    pub add_in: AddInInfo,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionAddedEvent {
    pub session_id: String,
    pub instance_id: String,
    pub document: DocumentInfo,
    pub available_tools: Vec<String>,
    pub is_active: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionUpdatedEvent {
    pub session_id: String,
    pub patch: SessionPatch,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionRemovedEvent {
    pub session_id: String,
    pub reason: SessionRemovedReason,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SessionRemovedReason {
    Closed,
    Crashed,
    Replaced,
    Unknown,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeartbeatDecision {
    KeepOpen,
    Close { code: u16 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AddinChannelError {
    InvalidUpgradePath(String),
    ForbiddenOrigin(String),
    MalformedRegister,
    ProtocolVersionMismatch { offered: String, supported: String },
    UnknownConnection(String),
    InstanceMismatch { expected: String, actual: String },
    MalformedSessionEvent,
    UnknownSession(String),
}

impl Display for AddinChannelError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::InvalidUpgradePath(path) => write!(formatter, "Invalid add-in path {path}."),
            Self::ForbiddenOrigin(origin) => write!(formatter, "Forbidden add-in origin {origin}."),
            Self::MalformedRegister => formatter.write_str("Malformed register request."),
            Self::ProtocolVersionMismatch { offered, supported } => write!(
                formatter,
                "Protocol version mismatch: server supports {supported}, add-in offered {offered}."
            ),
            Self::UnknownConnection(connection_id) => {
                write!(formatter, "Unknown add-in connection {connection_id}.")
            }
            Self::InstanceMismatch { expected, actual } => write!(
                formatter,
                "Add-in instance mismatch: expected {expected}, got {actual}."
            ),
            Self::MalformedSessionEvent => formatter.write_str("Malformed session event."),
            Self::UnknownSession(session_id) => write!(formatter, "Unknown session {session_id}."),
        }
    }
}

impl std::error::Error for AddinChannelError {}

fn same_major(left: &str, right: &str) -> bool {
    left.split('.').next() == right.split('.').next()
}

fn format_system_time(value: SystemTime) -> String {
    let seconds = value
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("unix:{seconds}")
}

#[cfg(test)]
#[path = "addin_channel_tests.rs"]
mod addin_channel_tests;
