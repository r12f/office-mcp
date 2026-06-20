use crate::addin_mgr::addin_protocol_version::same_major_version;
use crate::addin_mgr::{
    AddinChannelConfig, AddinChannelError, RegisterRequest, RegisterResult, RuntimeInfo,
};
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct AddinRegistrationPolicy {
    server_version: &'static str,
    protocol_version: &'static str,
}

impl AddinRegistrationPolicy {
    #[must_use]
    pub(crate) const fn new(server_version: &'static str, protocol_version: &'static str) -> Self {
        Self {
            server_version,
            protocol_version,
        }
    }

    pub(crate) fn validate(
        &self,
        connection_id: &str,
        request: &RegisterRequest,
    ) -> Result<(), AddinChannelError> {
        if request.instance_id.is_empty()
            || request.host.app.is_empty()
            || request.add_in.protocol_version.is_empty()
        {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                instance_id = %request.instance_id,
                host_app = %request.host.app,
                "rejected malformed add-in register request"
            );
            return Err(AddinChannelError::MalformedRegister);
        }
        if !same_major_version(&request.add_in.protocol_version, self.protocol_version) {
            tracing::warn!(
                component = "addin_channel",
                connection_id = %connection_id,
                instance_id = %request.instance_id,
                offered = %request.add_in.protocol_version,
                supported = %self.protocol_version,
                "rejected add-in protocol version"
            );
            return Err(AddinChannelError::ProtocolVersionMismatch {
                offered: request.add_in.protocol_version.clone(),
                supported: self.protocol_version.to_string(),
            });
        }
        Ok(())
    }

    #[must_use]
    pub(crate) fn runtime_from(request: RegisterRequest, now: SystemTime) -> RuntimeInfo {
        RuntimeInfo {
            instance_id: request.instance_id,
            host: request.host,
            add_in: request.add_in,
            registered_at: now,
        }
    }

    #[must_use]
    pub(crate) fn register_result(
        &self,
        runtime: &RuntimeInfo,
        config: &AddinChannelConfig,
    ) -> RegisterResult {
        RegisterResult {
            server_version: self.server_version.to_string(),
            protocol_version: self.protocol_version.to_string(),
            session_grace_sec: config.session_grace.as_secs(),
            heartbeat_interval_sec: config.heartbeat_interval.as_secs(),
            max_pending_per_session: config.max_pending_per_session,
            assigned_instance_id: runtime.instance_id.clone(),
        }
    }
}

#[cfg(test)]
#[path = "addin_registration_tests.rs"]
mod addin_registration_tests;
