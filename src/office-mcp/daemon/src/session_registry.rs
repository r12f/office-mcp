use std::collections::BTreeMap;
use std::fmt::{Display, Formatter};
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionRegistry {
    connections_by_instance: BTreeMap<String, AddinConnectionState>,
    sessions_by_id: BTreeMap<String, SessionInfo>,
    instance_by_session: BTreeMap<String, String>,
    max_pending_per_session: usize,
}

impl SessionRegistry {
    #[must_use]
    pub fn new() -> Self {
        Self::with_limits(4)
    }

    #[must_use]
    pub fn with_limits(max_pending_per_session: usize) -> Self {
        Self {
            connections_by_instance: BTreeMap::new(),
            sessions_by_id: BTreeMap::new(),
            instance_by_session: BTreeMap::new(),
            max_pending_per_session,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns Office runtime identity stale session grace capabilities and queue depth"
    }

    pub fn register_runtime(&mut self, runtime: RuntimeInfo) -> RegistrationOutcome {
        let instance_id = runtime.instance_id.clone();
        let replaced = self.connections_by_instance.contains_key(&instance_id);
        if replaced {
            self.mark_instance_stale(&instance_id, SystemTime::now());
        }
        self.connections_by_instance
            .insert(instance_id, AddinConnectionState::new(runtime));
        RegistrationOutcome { replaced }
    }

    pub fn add_session(
        &mut self,
        session: NewSessionInfo,
        registered_at: SystemTime,
    ) -> SessionInfo {
        let full = SessionInfo::new(session, registered_at);
        if let Some(connection) = self.connections_by_instance.get_mut(&full.instance_id) {
            connection.session_id = Some(full.session_id.clone());
        }
        self.instance_by_session
            .insert(full.session_id.clone(), full.instance_id.clone());
        self.sessions_by_id
            .insert(full.session_id.clone(), full.clone());
        full
    }

    pub fn update_session(&mut self, session_id: &str, patch: SessionPatch) -> bool {
        let Some(session) = self.sessions_by_id.get_mut(session_id) else {
            return false;
        };
        session.apply_patch(patch);
        true
    }

    pub fn remove_session(&mut self, session_id: &str) -> bool {
        let Some(instance_id) = self.instance_by_session.remove(session_id) else {
            return false;
        };
        if let Some(connection) = self.connections_by_instance.get_mut(&instance_id)
            && connection.session_id.as_deref() == Some(session_id)
        {
            connection.session_id = None;
        }
        self.sessions_by_id.remove(session_id).is_some()
    }

    pub fn remove_runtime(&mut self, instance_id: &str, stale_since: SystemTime) -> bool {
        let Some(connection) = self.connections_by_instance.remove(instance_id) else {
            return false;
        };
        if let Some(session_id) = connection.session_id {
            self.mark_session_stale(&session_id, stale_since);
        }
        true
    }

    pub fn mark_instance_stale(&mut self, instance_id: &str, stale_since: SystemTime) -> bool {
        let Some(connection) = self.connections_by_instance.get(instance_id) else {
            return false;
        };
        let Some(session_id) = connection.session_id.clone() else {
            return false;
        };
        self.mark_session_stale(&session_id, stale_since)
    }

    pub fn mark_session_stale(&mut self, session_id: &str, stale_since: SystemTime) -> bool {
        let Some(session) = self.sessions_by_id.get_mut(session_id) else {
            return false;
        };
        if session.status == SessionStatus::Stale {
            return false;
        }
        session.status = SessionStatus::Stale;
        session.stale_since = Some(stale_since);
        true
    }

    pub fn prune_stale_sessions(&mut self, now: SystemTime, grace: Duration) -> usize {
        let expired = self
            .sessions_by_id
            .values()
            .filter(|session| session.is_stale_past(now, grace))
            .map(|session| session.session_id.clone())
            .collect::<Vec<_>>();
        let count = expired.len();
        for session_id in expired {
            self.remove_session(&session_id);
        }
        count
    }

    #[must_use]
    pub fn list_sessions(&self) -> Vec<SessionDescriptor> {
        self.sessions_by_id
            .values()
            .map(|session| self.describe(session))
            .collect()
    }

    #[must_use]
    pub fn get_session_info(&self, session_id: &str) -> Option<SessionDetails> {
        let session = self.sessions_by_id.get(session_id)?;
        Some(SessionDetails {
            descriptor: self.describe(session),
            available_tools: session.available_tools.clone(),
        })
    }

    /// Checks whether a tool call can be sent to the owning add-in runtime.
    ///
    /// # Errors
    ///
    /// Returns a protocol-shaped error when the session is missing, stale,
    /// disconnected, over its pending-call limit, or lacks the requested tool.
    pub fn prepare_invocation(
        &self,
        session_id: &str,
        tool: &str,
        check_capability: bool,
    ) -> Result<InvocationPermit, ToolInvocationError> {
        let Some(session) = self.sessions_by_id.get(session_id) else {
            let code = if self.sessions_by_id.is_empty() {
                OfficeMcpCode::NoSessions
            } else {
                OfficeMcpCode::SessionNotFound
            };
            return Err(ToolInvocationError::new(code, session_id, tool));
        };
        if session.status == SessionStatus::Stale {
            return Err(ToolInvocationError::new(
                OfficeMcpCode::SessionStale,
                session_id,
                tool,
            ));
        }
        if check_capability && !session.available_tools.iter().any(|value| value == tool) {
            return Err(ToolInvocationError::new(
                OfficeMcpCode::HostCapabilityUnavailable,
                session_id,
                tool,
            ));
        }
        let Some(connection) = self.connection_for_session(session_id) else {
            return Err(ToolInvocationError::new(
                OfficeMcpCode::SessionLost,
                session_id,
                tool,
            ));
        };
        if !connection.connected {
            return Err(ToolInvocationError::new(
                OfficeMcpCode::SessionLost,
                session_id,
                tool,
            ));
        }
        if connection.pending_count >= self.max_pending_per_session {
            return Err(ToolInvocationError::new(
                OfficeMcpCode::MaxPendingExceeded,
                session_id,
                tool,
            ));
        }
        Ok(InvocationPermit {
            instance_id: connection.runtime.instance_id.clone(),
            host_app: normalize_host_app(&connection.runtime.host.app),
            queue_depth: connection.pending_count,
        })
    }

    pub fn set_connection_pending(&mut self, instance_id: &str, pending_count: usize) -> bool {
        let Some(connection) = self.connections_by_instance.get_mut(instance_id) else {
            return false;
        };
        connection.pending_count = pending_count;
        true
    }

    pub fn set_connection_connected(&mut self, instance_id: &str, connected: bool) -> bool {
        let Some(connection) = self.connections_by_instance.get_mut(instance_id) else {
            return false;
        };
        connection.connected = connected;
        true
    }

    fn connection_for_session(&self, session_id: &str) -> Option<&AddinConnectionState> {
        let instance_id = self.instance_by_session.get(session_id)?;
        self.connections_by_instance.get(instance_id)
    }

    fn describe(&self, session: &SessionInfo) -> SessionDescriptor {
        let connection = self.connection_for_session(&session.session_id);
        let host = connection.map(|connection| &connection.runtime.host);
        SessionDescriptor {
            session_id: session.session_id.clone(),
            instance_id: session.instance_id.clone(),
            app: normalize_host_app(host.map_or("other", |host| &host.app)),
            host: HostDescriptor {
                app: normalize_host_app(host.map_or("other", |host| &host.app)),
                version: host.and_then(|host| host.version.clone()),
                platform: host.and_then(|host| host.platform.clone()),
                build: host.and_then(|host| host.build.clone()),
            },
            document: DocumentDescriptor::from(&session.document),
            is_active: session.is_active,
            capability_tiers: infer_capability_tiers(&session.available_tools),
            available_tool_count: session.available_tools.len(),
            queue_depth: connection.map_or(0, |connection| connection.pending_count),
            registered_at: session.registered_at,
            status: session.status,
        }
    }
}

impl Default for SessionRegistry {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AddinConnectionState {
    runtime: RuntimeInfo,
    session_id: Option<String>,
    connected: bool,
    pending_count: usize,
}

impl AddinConnectionState {
    fn new(runtime: RuntimeInfo) -> Self {
        Self {
            runtime,
            session_id: None,
            connected: true,
            pending_count: 0,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RegistrationOutcome {
    pub replaced: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeInfo {
    pub instance_id: String,
    pub host: HostInfo,
    pub add_in: AddInInfo,
    pub registered_at: SystemTime,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HostInfo {
    pub app: String,
    pub version: Option<String>,
    pub platform: Option<String>,
    pub build: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddInInfo {
    pub version: String,
    pub protocol_version: String,
    pub supported_features: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewSessionInfo {
    pub session_id: String,
    pub instance_id: String,
    pub document: DocumentInfo,
    pub available_tools: Vec<String>,
    pub is_active: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionInfo {
    pub session_id: String,
    pub instance_id: String,
    pub document: DocumentInfo,
    pub available_tools: Vec<String>,
    pub is_active: Option<bool>,
    pub status: SessionStatus,
    pub registered_at: SystemTime,
    pub stale_since: Option<SystemTime>,
}

impl SessionInfo {
    fn new(session: NewSessionInfo, registered_at: SystemTime) -> Self {
        Self {
            session_id: session.session_id,
            instance_id: session.instance_id,
            document: session.document,
            available_tools: session.available_tools,
            is_active: session.is_active,
            status: SessionStatus::Active,
            registered_at,
            stale_since: None,
        }
    }

    fn apply_patch(&mut self, patch: SessionPatch) {
        if let Some(document) = patch.document {
            self.document = document;
        }
        if let Some(available_tools) = patch.available_tools {
            self.available_tools = available_tools;
        }
        if let Some(is_active) = patch.is_active {
            self.is_active = is_active;
        }
    }

    fn is_stale_past(&self, now: SystemTime, grace: Duration) -> bool {
        self.status == SessionStatus::Stale
            && self
                .stale_since
                .and_then(|stale_since| now.duration_since(stale_since).ok())
                .is_some_and(|elapsed| elapsed > grace)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SessionPatch {
    pub document: Option<DocumentInfo>,
    pub available_tools: Option<Vec<String>>,
    pub is_active: Option<Option<bool>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SessionStatus {
    Active,
    Stale,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DocumentInfo {
    pub title: Option<String>,
    pub url: Option<String>,
    pub filename: Option<String>,
    pub is_dirty: Option<bool>,
    pub is_read_only: Option<bool>,
    pub is_protected: Option<bool>,
    pub protection: Option<ProtectionInfo>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionInfo {
    pub kind: Option<String>,
    pub rights: Option<Vec<String>>,
    pub rights_source: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionDescriptor {
    pub session_id: String,
    pub instance_id: String,
    pub app: String,
    pub host: HostDescriptor,
    pub document: DocumentDescriptor,
    pub is_active: Option<bool>,
    pub capability_tiers: Vec<String>,
    pub available_tool_count: usize,
    pub queue_depth: usize,
    pub registered_at: SystemTime,
    pub status: SessionStatus,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HostDescriptor {
    pub app: String,
    pub version: Option<String>,
    pub platform: Option<String>,
    pub build: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentDescriptor {
    pub title: Option<String>,
    pub url: Option<String>,
    pub filename: Option<String>,
    pub is_dirty: Option<bool>,
    pub is_read_only: Option<bool>,
    pub is_protected: Option<bool>,
    pub protection_kind: Option<String>,
    pub rights: Option<Vec<String>>,
    pub rights_source: Option<String>,
}

impl From<&DocumentInfo> for DocumentDescriptor {
    fn from(document: &DocumentInfo) -> Self {
        Self {
            title: document.title.clone().or_else(|| document.filename.clone()),
            url: document.url.clone(),
            filename: document.filename.clone(),
            is_dirty: document.is_dirty,
            is_read_only: document.is_read_only,
            is_protected: document.is_protected,
            protection_kind: document
                .protection
                .as_ref()
                .and_then(|protection| protection.kind.clone()),
            rights: document
                .protection
                .as_ref()
                .and_then(|protection| protection.rights.clone()),
            rights_source: document
                .protection
                .as_ref()
                .and_then(|protection| protection.rights_source.clone()),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SessionDetails {
    pub descriptor: SessionDescriptor,
    pub available_tools: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InvocationPermit {
    pub instance_id: String,
    pub host_app: String,
    pub queue_depth: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToolInvocationError {
    pub failure: ToolFailure,
}

impl ToolInvocationError {
    fn new(code: OfficeMcpCode, session_id: &str, tool: &str) -> Self {
        Self {
            failure: ToolFailure {
                office_mcp_code: code,
                message: code.message(session_id, tool),
                session_id: Some(session_id.to_string()),
                tool: Some(tool.to_string()),
                retriable: code.retriable(),
                partial_effect: code.partial_effect(),
            },
        }
    }
}

impl Display for ToolInvocationError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.failure.message)
    }
}

impl std::error::Error for ToolInvocationError {}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToolFailure {
    pub office_mcp_code: OfficeMcpCode,
    pub message: String,
    pub session_id: Option<String>,
    pub tool: Option<String>,
    pub retriable: bool,
    pub partial_effect: Option<PartialEffect>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PartialEffect {
    None,
    Possible,
    Unknown,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfficeMcpCode {
    NoSessions,
    SessionNotFound,
    SessionStale,
    SessionLost,
    MaxPendingExceeded,
    HostCapabilityUnavailable,
}

impl OfficeMcpCode {
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::NoSessions => "NO_SESSIONS",
            Self::SessionNotFound => "SESSION_NOT_FOUND",
            Self::SessionStale => "SESSION_STALE",
            Self::SessionLost => "SESSION_LOST",
            Self::MaxPendingExceeded => "MAX_PENDING_EXCEEDED",
            Self::HostCapabilityUnavailable => "HOST_CAPABILITY_UNAVAILABLE",
        }
    }

    fn message(self, session_id: &str, tool: &str) -> String {
        match self {
            Self::NoSessions => "No Office document sessions are connected. Activate the office-mcp add-in in Word and try again.".to_string(),
            Self::SessionNotFound => format!("Session {session_id} is not registered."),
            Self::SessionStale => format!("Session {session_id} is stale while the add-in reconnects."),
            Self::SessionLost => format!("Session {session_id} lost its add-in connection."),
            Self::MaxPendingExceeded => format!("Session {session_id} has too many pending tool calls."),
            Self::HostCapabilityUnavailable => format!("The selected Office session does not support {tool}."),
        }
    }

    const fn retriable(self) -> bool {
        matches!(
            self,
            Self::NoSessions | Self::SessionStale | Self::MaxPendingExceeded
        )
    }

    const fn partial_effect(self) -> Option<PartialEffect> {
        match self {
            Self::SessionLost => Some(PartialEffect::Unknown),
            _ => None,
        }
    }
}

fn normalize_host_app(app: &str) -> String {
    let value = app.to_ascii_lowercase();
    match value.as_str() {
        "word" | "excel" | "powerpoint" | "outlook" => value,
        _ => "other".to_string(),
    }
}

fn infer_capability_tiers(tools: &[String]) -> Vec<String> {
    let mut tiers = vec!["core".to_string()];
    if tools.iter().any(|tool| tool == "word.add_comment") {
        tiers.push("review".to_string());
    }
    if tools.iter().any(|tool| tool == "word.accept_change") {
        tiers.push("tracked_changes".to_string());
    }
    tiers
}

#[cfg(test)]
mod tests {
    use super::{
        AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, OfficeMcpCode, RuntimeInfo,
        SessionPatch, SessionRegistry, SessionStatus,
    };
    use std::time::{Duration, SystemTime};

    #[test]
    fn registers_runtime_and_describes_session_metadata() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
        let mut registry = SessionRegistry::new();
        registry.register_runtime(runtime("instance-1", "word", now));
        registry.add_session(session("session-1", "instance-1", "Doc.docx"), now);
        registry.set_connection_pending("instance-1", 2);

        let sessions = registry.list_sessions();

        assert_eq!(sessions.len(), 1);
        assert_eq!(sessions[0].app, "word");
        assert_eq!(sessions[0].document.title.as_deref(), Some("Doc.docx"));
        assert_eq!(sessions[0].available_tool_count, 3);
        assert_eq!(sessions[0].queue_depth, 2);
        assert_eq!(
            registry
                .get_session_info("session-1")
                .expect("session info")
                .available_tools,
            vec!["word.get_text", "word.add_comment", "word.accept_change"]
        );
        assert_eq!(
            sessions[0].capability_tiers,
            ["core", "review", "tracked_changes"]
        );
    }

    #[test]
    fn stale_sessions_are_pruned_after_grace_period() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
        let mut registry = SessionRegistry::new();
        registry.register_runtime(runtime("instance-1", "word", now));
        registry.add_session(session("session-1", "instance-1", "Doc.docx"), now);

        assert!(registry.remove_runtime("instance-1", now + Duration::from_secs(10)));
        assert_eq!(registry.list_sessions()[0].status, SessionStatus::Stale);
        assert_eq!(
            registry.prune_stale_sessions(now + Duration::from_secs(20), Duration::from_mins(1)),
            0
        );
        assert_eq!(
            registry.prune_stale_sessions(now + Duration::from_secs(80), Duration::from_mins(1)),
            1
        );
        assert!(registry.list_sessions().is_empty());
    }

    #[test]
    fn invocation_preflight_returns_protocol_errors() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
        let mut registry = SessionRegistry::with_limits(1);

        let error = registry
            .prepare_invocation("missing", "word.get_text", true)
            .expect_err("no sessions");
        assert_eq!(error.failure.office_mcp_code, OfficeMcpCode::NoSessions);
        assert_eq!(error.failure.office_mcp_code.as_str(), "NO_SESSIONS");
        assert!(error.failure.retriable);

        registry.register_runtime(runtime("instance-1", "word", now));
        registry.add_session(session("session-1", "instance-1", "Doc.docx"), now);

        let error = registry
            .prepare_invocation("session-1", "word.unsupported", true)
            .expect_err("missing capability");
        assert_eq!(
            error.failure.office_mcp_code,
            OfficeMcpCode::HostCapabilityUnavailable
        );
        assert_eq!(
            error.failure.office_mcp_code.as_str(),
            "HOST_CAPABILITY_UNAVAILABLE"
        );

        let error = registry
            .prepare_invocation("missing-session", "word.get_text", true)
            .expect_err("unknown session");
        assert_eq!(
            error.failure.office_mcp_code,
            OfficeMcpCode::SessionNotFound
        );
        assert_eq!(error.failure.office_mcp_code.as_str(), "SESSION_NOT_FOUND");

        registry.set_connection_pending("instance-1", 1);
        let error = registry
            .prepare_invocation("session-1", "word.get_text", true)
            .expect_err("too many pending");
        assert_eq!(
            error.failure.office_mcp_code,
            OfficeMcpCode::MaxPendingExceeded
        );
    }

    #[test]
    fn session_patch_updates_descriptor_fields() {
        let now = SystemTime::UNIX_EPOCH + Duration::from_secs(100);
        let mut registry = SessionRegistry::new();
        registry.register_runtime(runtime("instance-1", "unknown-host", now));
        registry.add_session(session("session-1", "instance-1", "Draft.docx"), now);

        registry.update_session(
            "session-1",
            SessionPatch {
                document: Some(DocumentInfo {
                    title: Some("Final".to_string()),
                    filename: Some("Final.docx".to_string()),
                    ..DocumentInfo::default()
                }),
                is_active: Some(Some(false)),
                ..SessionPatch::default()
            },
        );
        let session = registry
            .get_session_info("session-1")
            .expect("session details");

        assert_eq!(session.descriptor.app, "other");
        assert_eq!(session.descriptor.document.title.as_deref(), Some("Final"));
        assert_eq!(session.descriptor.is_active, Some(false));
    }

    fn runtime(instance_id: &str, app: &str, registered_at: SystemTime) -> RuntimeInfo {
        RuntimeInfo {
            instance_id: instance_id.to_string(),
            host: HostInfo {
                app: app.to_string(),
                version: Some("16.0".to_string()),
                platform: Some("windows".to_string()),
                build: Some("Desktop".to_string()),
            },
            add_in: AddInInfo {
                version: "0.1.0".to_string(),
                protocol_version: "1.0".to_string(),
                supported_features: vec!["doc.read".to_string()],
            },
            registered_at,
        }
    }

    fn session(session_id: &str, instance_id: &str, filename: &str) -> NewSessionInfo {
        NewSessionInfo {
            session_id: session_id.to_string(),
            instance_id: instance_id.to_string(),
            document: DocumentInfo {
                filename: Some(filename.to_string()),
                ..DocumentInfo::default()
            },
            available_tools: vec![
                "word.get_text".to_string(),
                "word.add_comment".to_string(),
                "word.accept_change".to_string(),
            ],
            is_active: Some(true),
        }
    }
}
