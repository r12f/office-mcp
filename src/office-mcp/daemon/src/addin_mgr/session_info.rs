use crate::addin_mgr::session_registry::{
    DocumentInfo, NewSessionInfo, SessionPatch, SessionStatus,
};
use std::time::{Duration, SystemTime};

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
    pub(crate) fn new(session: NewSessionInfo, registered_at: SystemTime) -> Self {
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

    pub(crate) fn apply_patch(&mut self, patch: SessionPatch) {
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

    pub(crate) fn is_stale_past(&self, now: SystemTime, grace: Duration) -> bool {
        self.status == SessionStatus::Stale
            && self
                .stale_since
                .and_then(|stale_since| now.duration_since(stale_since).ok())
                .is_some_and(|elapsed| elapsed > grace)
    }
}

#[cfg(test)]
#[path = "session_info_tests.rs"]
mod session_info_tests;
