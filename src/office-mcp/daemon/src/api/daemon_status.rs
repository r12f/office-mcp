use serde_json::Value;
use std::path::{Path, PathBuf};
use std::process::Command;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DaemonStatusReporter {
    runtime_path: PathBuf,
}

impl DaemonStatusReporter {
    #[must_use]
    pub const fn new(runtime_path: PathBuf) -> Self {
        Self { runtime_path }
    }

    #[must_use]
    pub fn status_json(&self) -> String {
        let runtime = RuntimeStatus::read(&self.runtime_path);
        let running = runtime
            .as_ref()
            .and_then(|status| status.pid)
            .is_some_and(process_exists);
        let runtime_stale = runtime.is_some() && !running;
        format!(
            concat!(
                "{{\n",
                "  \"running\": {},\n",
                "  \"runtimeStale\": {},\n",
                "  \"runtimePath\": \"{}\",\n",
                "  \"pid\": {},\n",
                "  \"uiUrl\": {},\n",
                "  \"stateUrl\": {},\n",
                "  \"logPath\": {},\n",
                "  \"uiCommand\": \"office-mcp-daemon ui\"\n",
                "}}"
            ),
            running,
            runtime_stale,
            json_escape(&self.runtime_path.display().to_string()),
            active_value(running, runtime.as_ref().and_then(|status| status.pid)),
            active_string(running, runtime.as_ref().and_then(|status| status.ui_url.as_deref())),
            active_string(
                running,
                runtime.as_ref().and_then(|status| status.state_url.as_deref())
            ),
            active_string(
                running,
                runtime.as_ref().and_then(|status| status.log_path.as_deref())
            )
        )
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

fn json_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

fn active_value(running: bool, value: Option<u32>) -> String {
    if running {
        value.map_or_else(|| "null".to_string(), |pid| pid.to_string())
    } else {
        "null".to_string()
    }
}

fn active_string(running: bool, value: Option<&str>) -> String {
    if running {
        value.map_or_else(
            || "null".to_string(),
            |text| format!("\"{}\"", json_escape(text)),
        )
    } else {
        "null".to_string()
    }
}

#[cfg(test)]
#[path = "daemon_status_tests.rs"]
mod daemon_status_tests;
