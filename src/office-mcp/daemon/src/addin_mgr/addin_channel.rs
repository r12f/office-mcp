use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionPatch, SessionRegistry,
};
use crate::addin_mgr::{CancelCommand, QueuedCommand};
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
            return Err(AddinChannelError::InvalidUpgradePath(path.to_string()));
        }
        if origin != Some(self.config.origin.as_str()) {
            return Err(AddinChannelError::ForbiddenOrigin(
                origin.unwrap_or_default().to_string(),
            ));
        }
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
            return Err(AddinChannelError::MalformedRegister);
        }
        if !same_major(&request.add_in.protocol_version, ADDIN_PROTOCOL_VERSION) {
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
            return Err(AddinChannelError::InstanceMismatch {
                expected: connection.instance_id.clone(),
                actual: event.instance_id,
            });
        }
        if event.session_id.is_empty() {
            return Err(AddinChannelError::MalformedSessionEvent);
        }
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
            return Err(AddinChannelError::MalformedSessionEvent);
        }
        if registry.update_session(&event.session_id, event.patch) {
            Ok(())
        } else {
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
            return Err(AddinChannelError::MalformedSessionEvent);
        }
        for connection in self.connections.values_mut() {
            if connection.session_id.as_deref() == Some(event.session_id.as_str()) {
                connection.session_id = None;
            }
        }
        if registry.remove_session(&event.session_id) {
            Ok(())
        } else {
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
        registry.remove_runtime(&connection.instance_id, stale_since)
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
            return Ok(HeartbeatDecision::KeepOpen);
        }
        registry.mark_instance_stale(&connection.instance_id, stale_since);
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
pub struct RegisterResult {
    pub server_version: String,
    pub protocol_version: String,
    pub session_grace_sec: u64,
    pub heartbeat_interval_sec: u64,
    pub max_pending_per_session: usize,
    pub assigned_instance_id: String,
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
pub enum JsonRpcId {
    String(String),
    Number(i64),
    Null,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct JsonRpcEnvelope {
    pub id: Option<JsonRpcId>,
    pub method: Option<String>,
    pub params: BTreeMap<String, String>,
    pub result: Option<RegisterResult>,
}

impl JsonRpcEnvelope {
    fn request(id: String, method: &str, params: BTreeMap<String, String>) -> Self {
        Self {
            id: Some(JsonRpcId::String(id)),
            method: Some(method.to_string()),
            params,
            result: None,
        }
    }

    fn notification(method: &str, params: BTreeMap<String, String>) -> Self {
        Self {
            id: None,
            method: Some(method.to_string()),
            params,
            result: None,
        }
    }

    fn success(id: JsonRpcId, result: RegisterResult) -> Self {
        Self {
            id: Some(id),
            method: None,
            params: BTreeMap::new(),
            result: Some(result),
        }
    }
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
mod tests {
    use super::{
        ADDIN_PROTOCOL_VERSION, AddinChannelError, AddinChannelServer, HeartbeatDecision,
        JsonRpcId, RegisterRequest, SessionAddedEvent, SessionRemovedEvent, SessionRemovedReason,
        SessionUpdatedEvent,
    };
    use crate::addin_mgr::{AddInInfo, DocumentInfo, HostInfo, SessionPatch, SessionRegistry};
    use crate::addin_mgr::{CancelCommand, QueuedCommand};
    use std::time::{Duration, SystemTime};

    #[test]
    fn validates_upgrade_path_and_origin() {
        let server = AddinChannelServer::new();

        assert!(
            server
                .validate_upgrade("/addin", Some("https://localhost:8765"))
                .is_ok()
        );
        assert!(matches!(
            server.validate_upgrade("/wrong", Some("https://localhost:8765")),
            Err(AddinChannelError::InvalidUpgradePath(_))
        ));
        assert!(matches!(
            server.validate_upgrade("/addin", Some("https://example.invalid")),
            Err(AddinChannelError::ForbiddenOrigin(_))
        ));
    }

    #[test]
    fn register_runtime_returns_server_settings_and_updates_registry() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let mut server = AddinChannelServer::new();
        let mut registry = SessionRegistry::new();

        let reply = server
            .register_runtime(
                &mut registry,
                "connection-1".to_string(),
                register_request("instance-1", ADDIN_PROTOCOL_VERSION),
                now,
            )
            .expect("register");
        let result = reply.result.expect("register result");

        assert_eq!(result.assigned_instance_id, "instance-1");
        assert_eq!(result.max_pending_per_session, 4);
        assert!(registry.list_sessions().is_empty());
    }

    #[test]
    fn register_rejects_protocol_major_mismatch() {
        let mut server = AddinChannelServer::new();
        let mut registry = SessionRegistry::new();

        let error = server
            .register_runtime(
                &mut registry,
                "connection-1".to_string(),
                register_request("instance-1", "2.0"),
                SystemTime::UNIX_EPOCH,
            )
            .expect_err("protocol mismatch");

        assert!(matches!(
            error,
            AddinChannelError::ProtocolVersionMismatch { .. }
        ));
    }

    #[test]
    fn session_events_update_registry_with_instance_binding() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let mut server = AddinChannelServer::new();
        let mut registry = SessionRegistry::new();
        server
            .register_runtime(
                &mut registry,
                "connection-1".to_string(),
                register_request("instance-1", ADDIN_PROTOCOL_VERSION),
                now,
            )
            .expect("register");

        server
            .add_session(
                &mut registry,
                "connection-1",
                SessionAddedEvent {
                    session_id: "session-1".to_string(),
                    instance_id: "instance-1".to_string(),
                    document: DocumentInfo {
                        filename: Some("Draft.docx".to_string()),
                        ..DocumentInfo::default()
                    },
                    available_tools: vec!["word.get_text".to_string()],
                    is_active: Some(true),
                },
                now,
            )
            .expect("session added");
        server
            .update_session(
                &mut registry,
                SessionUpdatedEvent {
                    session_id: "session-1".to_string(),
                    patch: SessionPatch {
                        document: Some(DocumentInfo {
                            title: Some("Final".to_string()),
                            ..DocumentInfo::default()
                        }),
                        ..SessionPatch::default()
                    },
                },
            )
            .expect("session updated");

        let session = registry.get_session_info("session-1").expect("session");
        assert_eq!(session.descriptor.document.title.as_deref(), Some("Final"));

        server
            .remove_session(
                &mut registry,
                SessionRemovedEvent {
                    session_id: "session-1".to_string(),
                    reason: SessionRemovedReason::Closed,
                },
            )
            .expect("session removed");
        assert!(registry.list_sessions().is_empty());
    }

    #[test]
    fn session_added_rejects_wrong_instance() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let mut server = AddinChannelServer::new();
        let mut registry = SessionRegistry::new();
        server
            .register_runtime(
                &mut registry,
                "connection-1".to_string(),
                register_request("instance-1", ADDIN_PROTOCOL_VERSION),
                now,
            )
            .expect("register");

        let error = server
            .add_session(
                &mut registry,
                "connection-1",
                SessionAddedEvent {
                    session_id: "session-1".to_string(),
                    instance_id: "different".to_string(),
                    document: DocumentInfo::default(),
                    available_tools: Vec::new(),
                    is_active: None,
                },
                now,
            )
            .expect_err("instance mismatch");

        assert!(matches!(error, AddinChannelError::InstanceMismatch { .. }));
    }

    #[test]
    fn heartbeat_marks_session_stale_after_second_miss() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
        let mut server = AddinChannelServer::new();
        let mut registry = SessionRegistry::new();
        server
            .register_runtime(
                &mut registry,
                "connection-1".to_string(),
                register_request("instance-1", ADDIN_PROTOCOL_VERSION),
                now,
            )
            .expect("register");
        server
            .add_session(
                &mut registry,
                "connection-1",
                SessionAddedEvent {
                    session_id: "session-1".to_string(),
                    instance_id: "instance-1".to_string(),
                    document: DocumentInfo::default(),
                    available_tools: vec!["word.get_text".to_string()],
                    is_active: None,
                },
                now,
            )
            .expect("session");

        let ping = server.start_ping("connection-1", now).expect("ping");
        assert_eq!(ping.method.as_deref(), Some("ping"));
        assert_eq!(
            server.record_heartbeat_timeout(&mut registry, "connection-1", now),
            Ok(HeartbeatDecision::KeepOpen)
        );
        assert_eq!(
            server.record_heartbeat_timeout(&mut registry, "connection-1", now),
            Ok(HeartbeatDecision::Close { code: 4002 })
        );

        assert_eq!(
            registry.list_sessions()[0].status,
            crate::addin_mgr::SessionStatus::Stale
        );
    }

    #[test]
    fn builds_tool_invoke_and_cancel_payloads() {
        let server = AddinChannelServer::new();
        let command = QueuedCommand {
            command_id: "command-1".to_string(),
            request_id: "request-1".to_string(),
            session_id: "session-1".to_string(),
            instance_id: "instance-1".to_string(),
            tool: "word.get_text".to_string(),
            arguments_json: "{}".to_string(),
            timeout: Duration::from_secs(30),
            enqueued_at: SystemTime::UNIX_EPOCH,
            deadline_at: SystemTime::UNIX_EPOCH + Duration::from_secs(30),
            sequence: 0,
            dispatched: false,
        };

        let invoke = server.tool_invoke_payload(&command);
        assert_eq!(invoke.method.as_deref(), Some("tool.invoke"));
        assert_eq!(invoke.params["session_id"], "session-1");
        assert_eq!(invoke.params["timeout_ms"], "30000");

        let cancel = server.tool_cancel_payload(&CancelCommand {
            request_id: "request-1".to_string(),
            reason: "timeout".to_string(),
        });
        assert_eq!(cancel.id, None);
        assert_eq!(cancel.method.as_deref(), Some("tool.cancel"));
        assert_eq!(cancel.params["reason"], "timeout");
    }

    fn register_request(instance_id: &str, protocol_version: &str) -> RegisterRequest {
        RegisterRequest {
            id: JsonRpcId::String("register-1".to_string()),
            instance_id: instance_id.to_string(),
            host: HostInfo {
                app: "word".to_string(),
                version: Some("16.0".to_string()),
                platform: Some("windows".to_string()),
                build: Some("Desktop".to_string()),
            },
            add_in: AddInInfo {
                version: "0.1.0".to_string(),
                protocol_version: protocol_version.to_string(),
                supported_features: vec!["doc.read".to_string()],
            },
        }
    }
}
