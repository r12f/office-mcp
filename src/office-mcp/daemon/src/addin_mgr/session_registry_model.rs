use crate::addin_mgr::{DocumentDescriptor, DocumentInfo};
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct AddinConnectionState {
    pub(crate) runtime: RuntimeInfo,
    pub(crate) session_id: Option<String>,
    pub(crate) connected: bool,
    pub(crate) pending_count: usize,
}

impl AddinConnectionState {
    #[must_use]
    pub(crate) fn new(runtime: RuntimeInfo) -> Self {
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

#[must_use]
pub(crate) fn normalize_host_app(app: &str) -> String {
    let value = app.to_ascii_lowercase();
    match value.as_str() {
        "word" | "excel" | "powerpoint" | "outlook" => value,
        _ => "other".to_string(),
    }
}

#[must_use]
pub(crate) fn infer_capability_tiers(tools: &[String]) -> Vec<String> {
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
#[path = "session_registry_model_tests.rs"]
mod session_registry_model_tests;
