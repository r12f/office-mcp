use crate::addin_mgr::{SessionDescriptor, SessionStatus};
use serde_json::{Value, json};
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Debug, Clone, Copy)]
pub struct SessionDescriptorView<'a> {
    descriptor: &'a SessionDescriptor,
}

impl<'a> SessionDescriptorView<'a> {
    #[must_use]
    pub const fn new(descriptor: &'a SessionDescriptor) -> Self {
        Self { descriptor }
    }

    #[must_use]
    pub fn to_json(&self) -> Value {
        let session = self.descriptor;
        json!({
            "session_id": session.session_id,
            "instance_id": session.instance_id,
            "app": session.app,
            "host": {
                "app": session.host.app,
                "version": session.host.version,
                "platform": session.host.platform,
                "build": session.host.build
            },
            "document": {
                "title": session.document.title,
                "url": session.document.url,
                "filename": session.document.filename,
                "is_dirty": session.document.is_dirty,
                "is_read_only": session.document.is_read_only,
                "is_protected": session.document.is_protected,
                "protection_kind": session.document.protection_kind,
                "rights": session.document.rights,
                "rights_source": session.document.rights_source
            },
            "is_active": session.is_active,
            "capability_tiers": session.capability_tiers,
            "available_tool_count": session.available_tool_count,
            "queue_depth": session.queue_depth,
            "registered_at": format_unix_time(session.registered_at),
            "status": match session.status {
                SessionStatus::Active => "active",
                SessionStatus::Stale => "stale",
            }
        })
    }
}

fn format_unix_time(value: SystemTime) -> String {
    let seconds = value
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("unix:{seconds}")
}

#[cfg(test)]
#[path = "session_descriptor_view_tests.rs"]
mod session_descriptor_view_tests;
