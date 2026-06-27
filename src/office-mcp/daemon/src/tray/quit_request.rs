use crate::api::{DaemonControlError, DaemonController};
use crate::tray::TrayPlatformError;

const QUIT_ACTION: &str = "quit";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayQuitRequest {
    action: &'static str,
    source: &'static str,
    process_id: u32,
}

impl TrayQuitRequest {
    #[must_use]
    pub fn new(source: &'static str) -> Self {
        Self {
            action: QUIT_ACTION,
            source,
            process_id: std::process::id(),
        }
    }

    /// Requests daemon shutdown through the provided controller.
    ///
    /// # Errors
    ///
    /// Returns an error when the daemon shutdown controller cannot stop the
    /// daemon process.
    pub fn shutdown_with(
        &self,
        controller: &impl ShutdownController,
    ) -> Result<(), TrayPlatformError> {
        tracing::info!(
            component = "tray_host",
            action = self.action,
            source = self.source,
            pid = self.process_id,
            "stopping daemon from tray"
        );
        controller.stop_daemon().map_err(|error| {
            let message = format!(
                "action={} source={} pid={} failed to stop daemon: {error}",
                self.action, self.source, self.process_id
            );
            tracing::error!(
                component = "tray_host",
                action = self.action,
                source = self.source,
                pid = self.process_id,
                %error,
                "failed to stop daemon from tray"
            );
            TrayPlatformError::new(message)
        })
    }

    #[must_use]
    pub const fn action(&self) -> &'static str {
        self.action
    }

    #[must_use]
    pub const fn source(&self) -> &'static str {
        self.source
    }

    #[must_use]
    pub const fn process_id(&self) -> u32 {
        self.process_id
    }
}

pub trait ShutdownController {
    /// Stops the daemon process.
    ///
    /// # Errors
    ///
    /// Returns an error when the daemon cannot be stopped.
    fn stop_daemon(&self) -> Result<(), DaemonControlError>;
}

#[derive(Debug, Clone, Copy, Default)]
pub struct EnvShutdownController;

impl ShutdownController for EnvShutdownController {
    fn stop_daemon(&self) -> Result<(), DaemonControlError> {
        DaemonController::from_env().stop()
    }
}

#[cfg(test)]
#[derive(Debug, Default)]
pub struct RecordingShutdownController {
    stop_count: std::cell::Cell<usize>,
    error: Option<String>,
}

#[cfg(test)]
impl RecordingShutdownController {
    #[must_use]
    pub fn failing(error: impl Into<String>) -> Self {
        Self {
            stop_count: std::cell::Cell::new(0),
            error: Some(error.into()),
        }
    }

    #[must_use]
    pub fn stop_count(&self) -> usize {
        self.stop_count.get()
    }
}

#[cfg(test)]
impl ShutdownController for RecordingShutdownController {
    fn stop_daemon(&self) -> Result<(), DaemonControlError> {
        self.stop_count.set(self.stop_count.get() + 1);
        if let Some(error) = self.error.as_ref() {
            return Err(DaemonControlError::CommandFailed(error.clone()));
        }
        Ok(())
    }
}

#[cfg(test)]
#[path = "quit_request_tests.rs"]
mod quit_request_tests;
