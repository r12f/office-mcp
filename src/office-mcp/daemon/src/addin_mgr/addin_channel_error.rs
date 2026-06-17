use std::fmt::{Display, Formatter};

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

#[cfg(test)]
#[path = "addin_channel_error_tests.rs"]
mod addin_channel_error_tests;
