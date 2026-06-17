use crate::addin_mgr::{AddinChannelConfig, AddinChannelError};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AddinUpgradeGuard {
    origin: String,
}

impl AddinUpgradeGuard {
    #[must_use]
    pub fn new(config: &AddinChannelConfig) -> Self {
        Self {
            origin: config.origin.clone(),
        }
    }

    /// Validates that a WebSocket upgrade targets the add-in endpoint and origin.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is not `/addin` or the browser `Origin`
    /// does not exactly match the configured add-in origin.
    pub fn validate(&self, path: &str, origin: Option<&str>) -> Result<(), AddinChannelError> {
        if path != "/addin" {
            tracing::warn!(
                component = "addin_channel",
                path = %path,
                origin = ?origin,
                "rejected add-in websocket upgrade path"
            );
            return Err(AddinChannelError::InvalidUpgradePath(path.to_string()));
        }
        if origin != Some(self.origin.as_str()) {
            tracing::warn!(
                component = "addin_channel",
                path = %path,
                origin = ?origin,
                expected_origin = %self.origin,
                "rejected add-in websocket origin"
            );
            return Err(AddinChannelError::ForbiddenOrigin(
                origin.unwrap_or_default().to_string(),
            ));
        }
        tracing::debug!(
            component = "addin_channel",
            path = %path,
            origin = ?origin,
            "accepted add-in websocket upgrade"
        );
        Ok(())
    }
}

#[cfg(test)]
#[path = "addin_upgrade_guard_tests.rs"]
mod addin_upgrade_guard_tests;
