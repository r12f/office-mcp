use crate::tray::model::{TrayPlatform, TraySnapshot, TrayStatusInput};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayController {
    description: &'static str,
    platform: TrayPlatform,
}

impl TrayController {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            description: "owns native tray menu bar status menu show UI and graceful quit",
            platform: TrayPlatform::detect(),
        }
    }

    #[must_use]
    pub const fn with_platform(platform: TrayPlatform) -> Self {
        Self {
            description: "owns native tray menu bar status menu show UI and graceful quit",
            platform,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        self.description
    }

    #[must_use]
    pub const fn platform(&self) -> TrayPlatform {
        self.platform
    }

    #[must_use]
    pub fn snapshot(&self, input: TrayStatusInput) -> TraySnapshot {
        TraySnapshot::from_input(self.platform, input)
    }

    #[must_use]
    pub fn snapshot_from_ui_state(&self, ui_state: Option<&serde_json::Value>) -> TraySnapshot {
        self.snapshot(TrayStatusInput::from_ui_state(ui_state))
    }
}

impl Default for TrayController {
    fn default() -> Self {
        Self::new()
    }
}

pub trait TrayPlatformAdapter {
    /// Applies a tray snapshot to the platform-specific tray surface.
    ///
    /// # Errors
    ///
    /// Returns an error when the platform tray surface cannot be created or updated.
    fn apply_snapshot(&mut self, snapshot: &TraySnapshot) -> Result<(), TrayPlatformError>;

    /// Opens or focuses the daemon UI.
    ///
    /// # Errors
    ///
    /// Returns an error when the UI URL cannot be opened.
    fn show_ui(&mut self) -> Result<(), TrayPlatformError>;

    /// Requests graceful daemon shutdown.
    ///
    /// # Errors
    ///
    /// Returns an error when the daemon cannot be asked to stop.
    fn quit(&mut self) -> Result<(), TrayPlatformError>;
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayPlatformError {
    message: String,
}

impl TrayPlatformError {
    #[must_use]
    pub fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

impl std::fmt::Display for TrayPlatformError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for TrayPlatformError {}

#[cfg(test)]
#[path = "controller_tests.rs"]
mod controller_tests;
