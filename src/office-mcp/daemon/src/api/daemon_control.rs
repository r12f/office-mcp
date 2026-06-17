use crate::api::daemon_status::DaemonStatusReporter;
use crate::ui::UiRuntimeFile;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::path::PathBuf;
use std::process::Command;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DaemonController {
    task_name: String,
    install_root: Option<PathBuf>,
    runtime_path: PathBuf,
}

pub trait PowerShellExecutor {
    /// Runs a PowerShell command.
    ///
    /// # Errors
    ///
    /// Returns an error when PowerShell is unavailable or the command exits unsuccessfully.
    fn run(&self, command: &str) -> Result<(), DaemonControlError>;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SystemPowerShellExecutor;

impl PowerShellExecutor for SystemPowerShellExecutor {
    fn run(&self, command: &str) -> Result<(), DaemonControlError> {
        if !cfg!(windows) {
            return Err(DaemonControlError::Unavailable(
                "daemon start/stop is currently implemented for Windows only.".to_string(),
            ));
        }
        let status = Command::new("powershell.exe")
            .args(["-NoProfile", "-Command", command])
            .status()
            .map_err(|error| DaemonControlError::CommandFailed(error.to_string()))?;
        if status.success() {
            Ok(())
        } else {
            Err(DaemonControlError::CommandFailed(format!(
                "PowerShell command failed with status {status}."
            )))
        }
    }
}

impl DaemonController {
    #[must_use]
    pub fn from_env() -> Self {
        Self {
            task_name: std::env::var("OFFICE_MCP_TASK_NAME")
                .unwrap_or_else(|_| "office-mcp".to_string()),
            install_root: std::env::var_os("OFFICE_MCP_INSTALL_ROOT").map(PathBuf::from),
            runtime_path: UiRuntimeFile::default_path(),
        }
    }

    #[must_use]
    pub fn with_runtime_path(runtime_path: PathBuf) -> Self {
        Self {
            task_name: "office-mcp".to_string(),
            install_root: None,
            runtime_path,
        }
    }

    /// Starts the installed daemon integration.
    ///
    /// # Errors
    ///
    /// Returns an error when neither the Windows Scheduled Task nor installed
    /// launcher can be started.
    pub fn start(&self) -> Result<(), DaemonControlError> {
        self.start_with_executor(&SystemPowerShellExecutor)
    }

    /// Starts the daemon through an injectable PowerShell executor.
    ///
    /// # Errors
    ///
    /// Returns an error when neither the Windows Scheduled Task nor installed
    /// launcher can be started.
    pub fn start_with_executor(
        &self,
        executor: &impl PowerShellExecutor,
    ) -> Result<(), DaemonControlError> {
        if run_windows_task("Start-ScheduledTask", &self.task_name, executor).is_ok() {
            return Ok(());
        }
        let launcher = self.installed_launcher()?;
        executor.run(&format!(
            "Start-Process -WindowStyle Hidden powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File','{}')",
            escape_power_shell(&launcher.display().to_string())
        ))
    }

    /// Stops the installed daemon integration.
    ///
    /// # Errors
    ///
    /// Returns an error when neither the Windows Scheduled Task nor installed
    /// launcher process can be stopped.
    pub fn stop(&self) -> Result<(), DaemonControlError> {
        self.stop_with_executor(&SystemPowerShellExecutor)
    }

    /// Stops the daemon through an injectable PowerShell executor.
    ///
    /// # Errors
    ///
    /// Returns an error when neither the Windows Scheduled Task nor installed
    /// launcher process can be stopped.
    pub fn stop_with_executor(
        &self,
        executor: &impl PowerShellExecutor,
    ) -> Result<(), DaemonControlError> {
        if run_windows_task("Stop-ScheduledTask", &self.task_name, executor).is_ok() {
            return Ok(());
        }
        let launcher = self.installed_launcher()?;
        executor.run(&format!(
            "$launcher='{}'; Get-CimInstance Win32_Process | Where-Object {{ $_.CommandLine -like \"*$launcher*\" }} | ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}",
            escape_power_shell(&launcher.display().to_string())
        ))
    }

    #[must_use]
    pub fn status_json(&self) -> String {
        DaemonStatusReporter::new(self.runtime_path.clone()).status_json()
    }

    fn installed_launcher(&self) -> Result<PathBuf, DaemonControlError> {
        let Some(root) = &self.install_root else {
            return Err(DaemonControlError::Unavailable(
                "OFFICE_MCP_INSTALL_ROOT is not set and the Scheduled Task is unavailable."
                    .to_string(),
            ));
        };
        let launcher = root.join("office-mcp-daemon.ps1");
        if !launcher.exists() {
            return Err(DaemonControlError::Unavailable(format!(
                "Cannot find installed daemon launcher: {}",
                launcher.display()
            )));
        }
        Ok(launcher)
    }
}

#[derive(Debug)]
pub enum DaemonControlError {
    CommandFailed(String),
    Unavailable(String),
}

impl Display for DaemonControlError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::CommandFailed(message) | Self::Unavailable(message) => {
                formatter.write_str(message)
            }
        }
    }
}

impl Error for DaemonControlError {}

fn run_windows_task(
    command_name: &str,
    task_name: &str,
    executor: &impl PowerShellExecutor,
) -> Result<(), DaemonControlError> {
    executor.run(&format!(
        "{command_name} -TaskName '{}'",
        escape_power_shell(task_name)
    ))
}

fn escape_power_shell(value: &str) -> String {
    value.replace('\'', "''")
}

#[cfg(test)]
#[path = "daemon_control_tests.rs"]
mod daemon_control_tests;
