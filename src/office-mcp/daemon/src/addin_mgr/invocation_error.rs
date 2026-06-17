use std::fmt::{Display, Formatter};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ToolInvocationError {
    pub failure: ToolFailure,
}

impl ToolInvocationError {
    pub(crate) fn new(code: OfficeMcpCode, session_id: &str, tool: &str) -> Self {
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

#[cfg(test)]
#[path = "invocation_error_tests.rs"]
mod invocation_error_tests;
