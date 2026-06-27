use crate::tray::{TrayHostOptions, TrayPlatformError};
use crate::ui::UiRuntimeFile;
use std::path::PathBuf;

const DEFAULT_UI_URL: &str = "https://localhost:8765/ui/";
const SHOW_UI_ACTION: &str = "show_ui";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayUiOpenRequest {
    action: &'static str,
    url: String,
    source: &'static str,
    process_id: u32,
}

impl TrayUiOpenRequest {
    /// Resolves the daemon UI URL from the tray runtime options.
    ///
    /// # Errors
    ///
    /// Returns an error only when a present runtime file is malformed.
    pub fn from_runtime(options: &TrayHostOptions) -> Result<Self, TrayPlatformError> {
        let file = options
            .runtime_path
            .as_ref()
            .map_or_else(UiRuntimeFile::default_path, PathBuf::clone);
        match UiRuntimeFile::read_path(&file) {
            Ok(runtime) => Ok(Self::new(runtime.ui_url, "runtime_file")),
            Err(error) if !file.exists() => {
                tracing::warn!(
                    component = "tray_host",
                    action = SHOW_UI_ACTION,
                    ui_url = DEFAULT_UI_URL,
                    source = "fallback",
                    pid = std::process::id(),
                    runtime_path = %file.display(),
                    %error,
                    "opening daemon UI from tray without runtime file"
                );
                Ok(Self::new(DEFAULT_UI_URL.to_string(), "fallback"))
            }
            Err(error) => Err(TrayPlatformError::new(format!(
                "action={SHOW_UI_ACTION} source=runtime_file url=unknown pid={} failed to read runtime file {}: {error}",
                std::process::id(),
                file.display()
            ))),
        }
    }

    #[must_use]
    pub fn new(url: String, source: &'static str) -> Self {
        Self {
            action: SHOW_UI_ACTION,
            url,
            source,
            process_id: std::process::id(),
        }
    }

    /// Opens the resolved UI URL through the provided launcher.
    ///
    /// # Errors
    ///
    /// Returns an error when the platform launcher fails.
    pub fn open_with(&self, launcher: &impl UiLauncher) -> Result<(), TrayPlatformError> {
        tracing::info!(
            component = "tray_host",
            action = self.action,
            ui_url = %self.url,
            source = self.source,
            pid = self.process_id,
            "opening daemon UI from tray"
        );
        launcher.open_url(self.url.as_str()).map_err(|error| {
            let message = format!(
                "action={} source={} url={} pid={} failed to open UI: {error}",
                self.action, self.source, self.url, self.process_id
            );
            tracing::error!(
                component = "tray_host",
                action = self.action,
                ui_url = %self.url,
                source = self.source,
                pid = self.process_id,
                %error,
                "failed to open daemon UI from tray"
            );
            TrayPlatformError::new(message)
        })
    }

    #[must_use]
    pub const fn action(&self) -> &'static str {
        self.action
    }

    #[must_use]
    pub fn url(&self) -> &str {
        self.url.as_str()
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

pub trait UiLauncher {
    /// Opens a URL in the user's desktop environment.
    ///
    /// # Errors
    ///
    /// Returns an error when the platform launcher command cannot be started or
    /// reports failure.
    fn open_url(&self, url: &str) -> Result<(), TrayPlatformError>;
}

#[derive(Debug, Clone, Copy, Default)]
pub struct PlatformUiLauncher;

impl UiLauncher for PlatformUiLauncher {
    fn open_url(&self, url: &str) -> Result<(), TrayPlatformError> {
        open_platform_url(url)
    }
}

fn open_platform_url(url: &str) -> Result<(), TrayPlatformError> {
    #[cfg(windows)]
    {
        std::process::Command::new("cmd")
            .args(["/C", "start", "", url])
            .spawn()
            .and_then(|mut child| child.wait())
            .map_err(|error| TrayPlatformError::new(error.to_string()))?;
        return Ok(());
    }
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("open")
            .arg(url)
            .spawn()
            .and_then(|mut child| child.wait())
            .map_err(|error| TrayPlatformError::new(error.to_string()))?;
        return Ok(());
    }
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        std::process::Command::new("xdg-open")
            .arg(url)
            .spawn()
            .and_then(|mut child| child.wait())
            .map_err(|error| TrayPlatformError::new(error.to_string()))?;
        return Ok(());
    }
    #[allow(unreachable_code)]
    Err(TrayPlatformError::new(
        "opening URLs is unsupported on this platform",
    ))
}

#[cfg(test)]
#[derive(Debug, Default)]
pub struct RecordingUiLauncher {
    opened_urls: std::cell::RefCell<Vec<String>>,
    error: Option<String>,
}

#[cfg(test)]
impl RecordingUiLauncher {
    #[must_use]
    pub fn failing(error: impl Into<String>) -> Self {
        Self {
            opened_urls: std::cell::RefCell::new(Vec::new()),
            error: Some(error.into()),
        }
    }

    #[must_use]
    pub fn opened_urls(&self) -> Vec<String> {
        self.opened_urls.borrow().clone()
    }
}

#[cfg(test)]
impl UiLauncher for RecordingUiLauncher {
    fn open_url(&self, url: &str) -> Result<(), TrayPlatformError> {
        if let Some(error) = self.error.as_ref() {
            return Err(TrayPlatformError::new(error.clone()));
        }
        self.opened_urls.borrow_mut().push(url.to_string());
        Ok(())
    }
}

#[cfg(test)]
#[path = "ui_launch_tests.rs"]
mod ui_launch_tests;
