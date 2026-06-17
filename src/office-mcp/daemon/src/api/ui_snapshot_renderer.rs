use crate::addin_mgr::{PartialEffect, SessionDescriptor, SessionStatus};
use crate::api::{
    UiClientRecord, UiClientTransport, UiCommandError, UiCommandRecord, UiCommandStatus, UiHealth,
    UiSnapshot,
};
use serde_json::{Value, json};
use std::collections::BTreeMap;
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct UiSnapshotRenderer;

impl UiSnapshotRenderer {
    #[must_use]
    pub const fn new() -> Self {
        Self
    }

    #[must_use]
    pub fn render_text(&self, snapshot: &UiSnapshot) -> String {
        self.render_value(snapshot).to_string()
    }

    #[must_use]
    pub fn render_value(&self, snapshot: &UiSnapshot) -> Value {
        json!({
            "daemon": {
                "status": ui_health_json(snapshot.daemon.status),
                "version": snapshot.daemon.version,
                "uptime_ms": snapshot.daemon.uptime_ms,
                "mcp_endpoint": snapshot.daemon.mcp_endpoint,
                "addin_endpoint": snapshot.daemon.addin_endpoint,
                "config_path": snapshot.daemon.config_path,
                "log_path": snapshot.daemon.log_path,
                "last_error": snapshot.daemon.last_error,
            },
            "clients": snapshot.clients.iter().map(ui_client_json).collect::<Vec<_>>(),
            "documents": snapshot.documents.iter().map(|(app, sessions)| {
                (app.clone(), sessions.iter().map(session_descriptor_json).collect::<Vec<_>>())
            }).collect::<BTreeMap<_, _>>(),
            "current_tasks": snapshot.current_tasks.iter().map(ui_command_json).collect::<Vec<_>>(),
            "recent_commands": snapshot.recent_commands.iter().map(ui_command_json).collect::<Vec<_>>(),
            "document_command_history": snapshot.document_command_history.iter().map(|(session_id, commands)| {
                (session_id.clone(), commands.iter().map(ui_command_json).collect::<Vec<_>>())
            }).collect::<BTreeMap<_, _>>(),
        })
    }
}

fn ui_health_json(value: UiHealth) -> &'static str {
    match value {
        UiHealth::Up => "up",
        UiHealth::Degraded => "degraded",
        UiHealth::Down => "down",
    }
}

fn ui_command_status_json(value: UiCommandStatus) -> &'static str {
    match value {
        UiCommandStatus::Running => "running",
        UiCommandStatus::Success => "success",
        UiCommandStatus::Failure => "failure",
        UiCommandStatus::Cancelled => "cancelled",
        UiCommandStatus::Timeout => "timeout",
    }
}

fn ui_client_transport_json(value: UiClientTransport) -> &'static str {
    match value {
        UiClientTransport::Http => "http",
        UiClientTransport::StdioBridge => "stdio-bridge",
    }
}

fn ui_client_json(client: &UiClientRecord) -> Value {
    json!({
        "client_id": client.client_id,
        "transport": ui_client_transport_json(client.transport),
        "name": client.name,
        "connected_at": system_time_millis(client.connected_at),
        "last_activity_at": system_time_millis(client.last_activity_at),
        "in_flight_request_count": client.in_flight_request_count,
    })
}

fn ui_command_json(command: &UiCommandRecord) -> Value {
    json!({
        "command_id": command.command_id,
        "mcp_request_id": command.mcp_request_id,
        "client_id": command.client_id,
        "client_name": command.client_name,
        "session_id": command.session_id,
        "host_app": command.host_app,
        "tool": command.tool,
        "user_intent": command.user_intent,
        "status": ui_command_status_json(command.status),
        "started_at": system_time_millis(command.started_at),
        "deadline_at": command.deadline_at.map(system_time_millis),
        "timeout_ms": command.timeout_ms,
        "completed_at": command.completed_at.map(system_time_millis),
        "elapsed_ms": command.elapsed_ms,
        "error": command.error.as_ref().map(ui_command_error_json),
    })
}

fn ui_command_error_json(error: &UiCommandError) -> Value {
    json!({
        "office_mcp_code": error.office_mcp_code,
        "message": error.message,
        "tool": error.tool,
        "retriable": error.retriable,
        "partial_effect": error.partial_effect.map(partial_effect_json),
    })
}

fn session_descriptor_json(session: &SessionDescriptor) -> Value {
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

fn partial_effect_json(value: PartialEffect) -> &'static str {
    match value {
        PartialEffect::None => "none",
        PartialEffect::Possible => "possible",
        PartialEffect::Unknown => "unknown",
    }
}

fn format_unix_time(value: SystemTime) -> String {
    let seconds = value
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("unix:{seconds}")
}

fn system_time_millis(value: SystemTime) -> u128 {
    value
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis()
}

#[cfg(test)]
#[path = "ui_snapshot_renderer_tests.rs"]
mod ui_snapshot_renderer_tests;
