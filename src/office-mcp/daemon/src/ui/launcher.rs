use crate::ui::{UiRuntimeError, UiRuntimeFile};
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiLauncher {
    runtime_path: PathBuf,
}

impl UiLauncher {
    #[must_use]
    pub fn default_runtime_path() -> Self {
        Self::with_runtime_path(UiRuntimeFile::default_path())
    }

    #[must_use]
    pub const fn with_runtime_path(runtime_path: PathBuf) -> Self {
        Self { runtime_path }
    }

    #[must_use]
    pub fn runtime_path(&self) -> &Path {
        &self.runtime_path
    }

    /// Reads the daemon-published UI URL from this launcher's runtime file.
    ///
    /// # Errors
    ///
    /// Returns an error when the runtime file is missing or malformed.
    pub fn ui_url(&self) -> Result<String, UiRuntimeError> {
        Ok(UiRuntimeFile::read_path(&self.runtime_path)?.ui_url)
    }

    /// Opens the daemon UI in the user's default browser.
    ///
    /// # Errors
    ///
    /// Returns an error when no daemon runtime file is available or the
    /// platform browser launcher fails.
    pub fn open(&self) -> Result<String, UiLaunchError> {
        let url = self.ui_url().map_err(UiLaunchError::Runtime)?;
        open_system_url(&url).map_err(UiLaunchError::Open)?;
        Ok(url)
    }
}

impl Default for UiLauncher {
    fn default() -> Self {
        Self::default_runtime_path()
    }
}

#[derive(Debug)]
pub enum UiLaunchError {
    Runtime(UiRuntimeError),
    Open(std::io::Error),
}

impl Display for UiLaunchError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Runtime(error) => write!(
                formatter,
                "No running office-mcp daemon UI was found. Start the daemon with `office-mcp-daemon daemon run` and try again. {error}"
            ),
            Self::Open(error) => write!(formatter, "Failed to open office-mcp daemon UI: {error}"),
        }
    }
}

impl Error for UiLaunchError {}

fn open_system_url(url: &str) -> Result<(), std::io::Error> {
    #[cfg(windows)]
    {
        std::process::Command::new("cmd")
            .args(["/C", "start", "", url])
            .spawn()?
            .wait()?;
        return Ok(());
    }
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("open")
            .arg(url)
            .spawn()?
            .wait()?;
        return Ok(());
    }
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        std::process::Command::new("xdg-open")
            .arg(url)
            .spawn()?
            .wait()?;
        return Ok(());
    }
    #[allow(unreachable_code)]
    Ok(())
}

#[cfg(test)]
#[path = "launcher_tests.rs"]
mod launcher_tests;
