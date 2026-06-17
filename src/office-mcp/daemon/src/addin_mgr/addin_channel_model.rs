use crate::addin_mgr::{AddInInfo, DocumentInfo, HostInfo, JsonRpcId, SessionPatch};

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

#[cfg(test)]
#[path = "addin_channel_model_tests.rs"]
mod addin_channel_model_tests;
