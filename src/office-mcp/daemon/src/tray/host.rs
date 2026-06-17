use crate::api::DaemonController;
use crate::tray::ui_state_client::TrayUiStateClient;
use crate::tray::{TrayController, TrayPlatformError, TraySnapshot};
use crate::ui::UiRuntimeFile;
use serde_json::Value;
use std::path::PathBuf;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TrayHostOptions {
    pub runtime_path: Option<PathBuf>,
    pub probe_state_path: Option<PathBuf>,
    pub probe: bool,
}

impl TrayHostOptions {
    #[must_use]
    pub fn from_args(args: &[String]) -> Self {
        Self {
            runtime_path: read_option(args, "--runtime-path").map(PathBuf::from),
            probe_state_path: read_option(args, "--probe-state-path").map(PathBuf::from),
            probe: args.iter().any(|arg| arg == "--probe"),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayHost {
    options: TrayHostOptions,
    controller: TrayController,
}

impl TrayHost {
    #[must_use]
    pub fn new(options: TrayHostOptions) -> Self {
        Self {
            options,
            controller: TrayController::new(),
        }
    }

    /// Runs the platform tray host.
    ///
    /// # Errors
    ///
    /// Returns an error when the native tray surface cannot be created, the UI
    /// URL cannot be opened, or the daemon cannot be stopped.
    pub fn run(&self) -> Result<(), TrayPlatformError> {
        if self.options.probe {
            let probe = self.probe_json();
            tracing::info!(
                component = "tray_host",
                probe = true,
                native_host = true,
                can_read_runtime = probe["can_read_runtime"].as_bool().unwrap_or(false),
                state_fetch_ok = probe["state_fetch_ok"].as_bool().unwrap_or(false),
                "ran tray probe"
            );
            println!("{}", probe);
            return Ok(());
        }
        tracing::info!(
            component = "tray_host",
            probe = false,
            "starting native tray host"
        );
        run_platform_tray(&self.options)
    }

    #[must_use]
    pub fn snapshot(&self) -> TraySnapshot {
        self.controller
            .snapshot_from_ui_state(read_ui_state(&self.options).as_ref())
    }

    #[must_use]
    pub fn probe_json(&self) -> Value {
        let snapshot = self.snapshot();
        serde_json::json!({
            "ok": true,
            "native_host": true,
            "snapshot": snapshot.probe_json(),
            "can_read_runtime": runtime_info(&self.options).is_some(),
            "state_fetch_ok": read_ui_state(&self.options).is_some()
        })
    }
}

/// Opens the daemon UI URL from the runtime file, or the default local URL.
///
/// # Errors
///
/// Returns an error when the platform URL launcher fails.
pub fn open_ui_from_runtime(options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
    if let Some(runtime) = runtime_info(options) {
        tracing::info!(
            component = "tray_host",
            ui_url = %runtime.ui_url,
            source = "runtime_file",
            "opening daemon UI from tray"
        );
        open_url(runtime.ui_url.as_str())?;
        return Ok(());
    }
    tracing::warn!(
        component = "tray_host",
        ui_url = "https://localhost:8765/ui/",
        source = "fallback",
        "opening daemon UI from tray without runtime file"
    );
    open_url("https://localhost:8765/ui/")
}

/// Requests daemon shutdown through the daemon controller.
///
/// # Errors
///
/// Returns an error when the controller cannot stop the daemon.
pub fn stop_daemon() -> Result<(), TrayPlatformError> {
    tracing::info!(component = "tray_host", "stopping daemon from tray");
    DaemonController::from_env()
        .stop()
        .map_err(|error| TrayPlatformError::new(error.to_string()))
}

pub fn start_tray_background() {
    tracing::info!(
        component = "tray_host",
        "starting tray host background thread"
    );
    let _ = std::thread::Builder::new()
        .name("office-mcp-tray".to_string())
        .spawn(|| {
            let result =
                std::panic::catch_unwind(|| TrayHost::new(TrayHostOptions::default()).run());
            match result {
                Ok(Ok(())) => {}
                Ok(Err(error)) => {
                    tracing::error!(%error, "office-mcp tray host stopped with error");
                    eprintln!("office-mcp-daemon tray host stopped with error: {error}");
                }
                Err(payload) => {
                    let message = panic_message(payload.as_ref());
                    tracing::error!(panic = %message, "office-mcp tray host panicked");
                    eprintln!("office-mcp-daemon tray host panicked: {message}");
                }
            }
        })
        .map_err(|error| {
            tracing::error!(%error, "office-mcp tray host thread failed to start");
            eprintln!("office-mcp-daemon failed to start tray host thread: {error}");
        });
}

fn panic_message(payload: &(dyn std::any::Any + Send)) -> String {
    payload.downcast_ref::<&str>().map_or_else(
        || {
            payload
                .downcast_ref::<String>()
                .map_or("unknown panic".to_string(), Clone::clone)
        },
        |message| (*message).to_string(),
    )
}

fn read_ui_state(options: &TrayHostOptions) -> Option<Value> {
    if let Some(path) = options.probe_state_path.as_ref() {
        let body = std::fs::read_to_string(path).ok()?;
        return serde_json::from_str(&body).ok();
    }
    let runtime = runtime_info(options)?;
    match TrayUiStateClient::new(runtime.state_url.as_str()).fetch_json() {
        Ok(value) => Some(value),
        Err(error) => {
            tracing::warn!(
                component = "tray_host",
                state_url = %runtime.state_url,
                %error,
                "failed to read daemon UI state for tray snapshot"
            );
            None
        }
    }
}

fn runtime_info(options: &TrayHostOptions) -> Option<crate::ui::UiRuntimeInfo> {
    let file = options
        .runtime_path
        .as_ref()
        .map_or_else(UiRuntimeFile::default_path, PathBuf::clone);
    UiRuntimeFile::read_path(&file).ok()
}

fn read_option(args: &[String], name: &str) -> Option<String> {
    args.iter()
        .position(|arg| arg == name)
        .and_then(|index| args.get(index + 1))
        .filter(|value| !value.starts_with("--"))
        .cloned()
}

fn open_url(url: &str) -> Result<(), TrayPlatformError> {
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

#[cfg(any(windows, target_os = "macos", target_os = "linux"))]
#[path = "native_tray.rs"]
mod native_tray;

#[cfg(any(windows, target_os = "macos", target_os = "linux"))]
fn run_platform_tray(options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
    native_tray::run(options)
}

#[cfg(not(any(windows, target_os = "macos", target_os = "linux")))]
fn run_platform_tray(_options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
    Err(TrayPlatformError::new(
        "native tray host is not implemented for this platform yet",
    ))
}

#[cfg(test)]
#[path = "host_tests.rs"]
mod host_tests;
