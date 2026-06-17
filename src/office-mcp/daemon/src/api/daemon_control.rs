use crate::ui::UiRuntimeFile;
use serde_json::Value;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::path::{Path, PathBuf};
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
        let runtime = RuntimeStatus::read(&self.runtime_path);
        let running = runtime
            .as_ref()
            .and_then(|status| status.pid)
            .is_some_and(process_exists);
        format!(
            concat!(
                "{{\n",
                "  \"running\": {},\n",
                "  \"runtimePath\": \"{}\",\n",
                "  \"pid\": {},\n",
                "  \"uiUrl\": {},\n",
                "  \"stateUrl\": {},\n",
                "  \"logPath\": {},\n",
                "  \"uiCommand\": \"office-mcp-daemon ui\"\n",
                "}}"
            ),
            running,
            json_escape(&self.runtime_path.display().to_string()),
            runtime
                .as_ref()
                .and_then(|status| status.pid)
                .map_or_else(|| "null".to_string(), |pid| pid.to_string()),
            runtime
                .as_ref()
                .and_then(|status| status.ui_url.as_ref())
                .map_or_else(
                    || "null".to_string(),
                    |url| format!("\"{}\"", json_escape(url))
                ),
            runtime
                .as_ref()
                .and_then(|status| status.state_url.as_ref())
                .map_or_else(
                    || "null".to_string(),
                    |url| format!("\"{}\"", json_escape(url))
                ),
            runtime
                .as_ref()
                .and_then(|status| status.log_path.as_ref())
                .map_or_else(
                    || "null".to_string(),
                    |path| format!("\"{}\"", json_escape(path))
                )
        )
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

#[derive(Debug, Clone, PartialEq, Eq)]
struct RuntimeStatus {
    pid: Option<u32>,
    ui_url: Option<String>,
    state_url: Option<String>,
    log_path: Option<String>,
}

impl RuntimeStatus {
    fn read(path: &Path) -> Option<Self> {
        let value = serde_json::from_str::<Value>(&std::fs::read_to_string(path).ok()?).ok()?;
        Some(Self {
            pid: value
                .get("pid")
                .and_then(serde_json::Value::as_u64)
                .and_then(|pid| u32::try_from(pid).ok()),
            ui_url: value
                .get("uiUrl")
                .and_then(serde_json::Value::as_str)
                .map(ToString::to_string),
            state_url: value
                .get("stateUrl")
                .and_then(serde_json::Value::as_str)
                .map(ToString::to_string),
            log_path: value
                .get("logPath")
                .and_then(serde_json::Value::as_str)
                .map(ToString::to_string),
        })
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

fn process_exists(pid: u32) -> bool {
    if pid == 0 {
        return false;
    }
    #[cfg(windows)]
    {
        Command::new("powershell.exe")
            .args([
                "-NoProfile",
                "-Command",
                &format!("if (Get-Process -Id {pid} -ErrorAction SilentlyContinue) {{ exit 0 }} else {{ exit 1 }}"),
            ])
            .status()
            .is_ok_and(|status| status.success())
    }
    #[cfg(not(windows))]
    {
        Command::new("kill")
            .args(["-0", &pid.to_string()])
            .status()
            .is_ok_and(|status| status.success())
    }
}

fn escape_power_shell(value: &str) -> String {
    value.replace('\'', "''")
}

fn json_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

#[cfg(test)]
#[path = "daemon_control_tests.rs"]
mod daemon_control_tests;
