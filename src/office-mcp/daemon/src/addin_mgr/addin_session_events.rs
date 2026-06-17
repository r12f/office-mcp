use crate::addin_mgr::{
    AddinChannelError, AddinConnectionState, NewSessionInfo, SessionAddedEvent, SessionRegistry,
    SessionRemovedEvent, SessionUpdatedEvent,
};
use std::collections::BTreeMap;
use std::time::SystemTime;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct AddinSessionEventHandler;

impl AddinSessionEventHandler {
    pub(crate) fn add_session(
        registry: &mut SessionRegistry,
        connections: &mut BTreeMap<String, AddinConnectionState>,
        connection_id: &str,
        event: SessionAddedEvent,
        now: SystemTime,
    ) -> Result<(), AddinChannelError> {
        let connection = connections
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

    pub(crate) fn update_session(
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

    pub(crate) fn remove_session(
        registry: &mut SessionRegistry,
        connections: &mut BTreeMap<String, AddinConnectionState>,
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
        for connection in connections.values_mut() {
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
}

#[cfg(test)]
#[path = "addin_session_events_tests.rs"]
mod addin_session_events_tests;
