use crate::api::DaemonController;
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
            println!("{}", self.probe_json());
            return Ok(());
        }
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
        open_url(runtime.ui_url.as_str())?;
        return Ok(());
    }
    open_url("https://localhost:8765/ui/")
}

/// Requests daemon shutdown through the daemon controller.
///
/// # Errors
///
/// Returns an error when the controller cannot stop the daemon.
pub fn stop_daemon() -> Result<(), TrayPlatformError> {
    DaemonController::from_env()
        .stop()
        .map_err(|error| TrayPlatformError::new(error.to_string()))
}

fn read_ui_state(options: &TrayHostOptions) -> Option<Value> {
    if let Some(path) = options.probe_state_path.as_ref() {
        let body = std::fs::read_to_string(path).ok()?;
        return serde_json::from_str(&body).ok();
    }
    let runtime = runtime_info(options)?;
    let body = ureq::get(runtime.state_url.as_str())
        .call()
        .ok()?
        .into_string()
        .ok()?;
    serde_json::from_str(&body).ok()
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
fn run_platform_tray(options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
    native_tray::run(options)
}

#[cfg(not(any(windows, target_os = "macos", target_os = "linux")))]
fn run_platform_tray(_options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
    Err(TrayPlatformError::new(
        "native tray host is not implemented for this platform yet",
    ))
}

#[cfg(any(windows, target_os = "macos", target_os = "linux"))]
mod native_tray {
    use super::{TrayHostOptions, open_ui_from_runtime, read_ui_state, stop_daemon};
    use crate::tray::{
        QuitConfirmation, TrayController, TrayMenuItem, TrayPlatformAdapter, TrayPlatformError,
        TraySnapshot,
    };
    use std::time::{Duration, Instant};
    use tao::event::{Event, StartCause};
    use tao::event_loop::{ControlFlow, EventLoop};
    use tray_icon::menu::{Menu, MenuEvent, MenuId, MenuItem, PredefinedMenuItem};
    use tray_icon::{Icon, TrayIcon, TrayIconBuilder};

    const SHOW_ID: &str = "office-mcp-show";
    const QUIT_ID: &str = "office-mcp-quit";

    pub fn run(options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
        let event_loop = EventLoop::new();
        let mut surface = NativeTraySurface::new(options)?;
        let mut next_refresh = Instant::now();
        event_loop.run(move |event, _, control_flow| {
            *control_flow = ControlFlow::WaitUntil(next_refresh);
            match event {
                Event::NewEvents(StartCause::Init | StartCause::ResumeTimeReached { .. }) => {
                    if let Err(error) = surface.refresh() {
                        eprintln!("office-mcp tray refresh failed: {error}");
                    }
                    next_refresh = Instant::now() + Duration::from_secs(2);
                    *control_flow = ControlFlow::WaitUntil(next_refresh);
                }
                Event::MainEventsCleared => {
                    while let Ok(event) = MenuEvent::receiver().try_recv() {
                        let id = event.id.0.as_str();
                        if id == SHOW_ID {
                            if let Err(error) = surface.show_ui() {
                                eprintln!("office-mcp tray show UI failed: {error}");
                            }
                        } else if id == QUIT_ID && confirm_quit(&surface.snapshot.quit_confirmation)
                        {
                            if let Err(error) = surface.quit() {
                                eprintln!("office-mcp tray quit failed: {error}");
                            }
                            *control_flow = ControlFlow::Exit;
                        }
                    }
                }
                _ => {}
            }
        });
    }

    struct NativeTraySurface {
        options: TrayHostOptions,
        controller: TrayController,
        icon: TrayIcon,
        snapshot: TraySnapshot,
    }

    impl NativeTraySurface {
        fn new(options: &TrayHostOptions) -> Result<Self, TrayPlatformError> {
            let controller = TrayController::new();
            let snapshot = controller.snapshot_from_ui_state(read_ui_state(options).as_ref());
            let menu = build_menu(&snapshot);
            let icon = TrayIconBuilder::new()
                .with_menu(Box::new(menu))
                .with_tooltip(snapshot.tooltip.as_str())
                .with_icon(app_icon())
                .build()
                .map_err(|error| TrayPlatformError::new(error.to_string()))?;
            Ok(Self {
                options: options.clone(),
                controller,
                icon,
                snapshot,
            })
        }

        fn refresh(&mut self) -> Result<(), TrayPlatformError> {
            let snapshot = self
                .controller
                .snapshot_from_ui_state(read_ui_state(&self.options).as_ref());
            self.apply_snapshot(&snapshot)
        }
    }

    impl TrayPlatformAdapter for NativeTraySurface {
        fn apply_snapshot(&mut self, snapshot: &TraySnapshot) -> Result<(), TrayPlatformError> {
            self.icon
                .set_tooltip(Some(snapshot.tooltip.as_str()))
                .map_err(|error| TrayPlatformError::new(error.to_string()))?;
            self.icon.set_menu(Some(Box::new(build_menu(snapshot))));
            self.snapshot = snapshot.clone();
            Ok(())
        }

        fn show_ui(&mut self) -> Result<(), TrayPlatformError> {
            open_ui_from_runtime(&self.options)
        }

        fn quit(&mut self) -> Result<(), TrayPlatformError> {
            stop_daemon()
        }
    }

    fn build_menu(snapshot: &TraySnapshot) -> Menu {
        let menu = Menu::new();
        for item in &snapshot.menu {
            match item {
                TrayMenuItem::ReadOnly { label } => {
                    let menu_item = MenuItem::new(label, false, None);
                    let _ = menu.append(&menu_item);
                }
                TrayMenuItem::Action { label, .. } if label == "Show Office MCP" => {
                    let menu_item = MenuItem::with_id(MenuId::new(SHOW_ID), label, true, None);
                    let _ = menu.append(&menu_item);
                }
                TrayMenuItem::Action { label, .. } => {
                    let menu_item = MenuItem::with_id(MenuId::new(QUIT_ID), label, true, None);
                    let _ = menu.append(&menu_item);
                }
                TrayMenuItem::Separator => {
                    let _ = menu.append(&PredefinedMenuItem::separator());
                }
            }
        }
        menu
    }

    fn app_icon() -> Icon {
        let width: u32 = 32;
        let height: u32 = 32;
        let mut rgba = Vec::with_capacity(32 * 32 * 4);
        for y in 0..height {
            for x in 0..width {
                let inside = (4..28).contains(&x) && (4..28).contains(&y);
                if inside {
                    rgba.extend_from_slice(&[43, 87, 154, 255]);
                } else {
                    rgba.extend_from_slice(&[0, 0, 0, 0]);
                }
            }
        }
        Icon::from_rgba(rgba, width, height).expect("generated tray icon is valid")
    }

    fn confirm_quit(confirmation: &QuitConfirmation) -> bool {
        #[cfg(windows)]
        {
            let script = format!(
                concat!(
                    "Add-Type -AssemblyName PresentationFramework; ",
                    "$result = [System.Windows.MessageBox]::Show({body}, {title}, 'YesNo', 'Warning'); ",
                    "if ($result -eq 'Yes') {{ exit 0 }} else {{ exit 1 }}"
                ),
                body = powershell_string(&confirmation.body),
                title = powershell_string(&confirmation.title)
            );
            command_succeeds(
                "powershell.exe",
                &["-NoProfile", "-Command", script.as_str()],
            )
        }
        #[cfg(target_os = "macos")]
        {
            let script = format!(
                concat!(
                    "display dialog {body} with title {title} buttons {{{cancel}, {confirm}}} ",
                    "default button {cancel} cancel button {cancel} with icon caution"
                ),
                body = applescript_string(&confirmation.body),
                title = applescript_string(&confirmation.title),
                cancel = applescript_string(&confirmation.secondary_action),
                confirm = applescript_string(&confirmation.primary_action)
            );
            return command_succeeds("osascript", &["-e", script.as_str()]);
        }
        #[cfg(all(unix, not(target_os = "macos")))]
        {
            command_succeeds(
                "zenity",
                &[
                    "--question",
                    "--title",
                    confirmation.title.as_str(),
                    "--text",
                    confirmation.body.as_str(),
                    "--ok-label",
                    confirmation.primary_action.as_str(),
                    "--cancel-label",
                    confirmation.secondary_action.as_str(),
                ],
            ) || command_succeeds(
                "kdialog",
                &[
                    "--warningyesno",
                    confirmation.body.as_str(),
                    "--title",
                    confirmation.title.as_str(),
                    "--yes-label",
                    confirmation.primary_action.as_str(),
                    "--no-label",
                    confirmation.secondary_action.as_str(),
                ],
            )
        }
        #[cfg(not(any(windows, target_os = "macos", all(unix, not(target_os = "macos")))))]
        {
            false
        }
    }

    fn command_succeeds(command: &str, args: &[&str]) -> bool {
        std::process::Command::new(command)
            .args(args)
            .status()
            .is_ok_and(|status| status.success())
    }

    #[cfg(windows)]
    fn powershell_string(value: &str) -> String {
        format!("'{}'", value.replace('\'', "''"))
    }

    #[cfg(target_os = "macos")]
    fn applescript_string(value: &str) -> String {
        format!("\"{}\"", value.replace('\\', "\\\\").replace('"', "\\\""))
    }
}

#[cfg(test)]
mod tests {
    use super::{TrayHost, TrayHostOptions};
    use std::fs;

    #[test]
    fn probe_snapshot_reads_state_file() {
        let dir =
            std::env::temp_dir().join(format!("office-mcp-tray-host-test-{}", std::process::id()));
        fs::create_dir_all(&dir).expect("create temp dir");
        let state_path = dir.join("state.json");
        fs::write(
            &state_path,
            r#"{"daemon":{"status":"up"},"clients":[{}],"documents":{"word":[{}]},"current_tasks":[]}"#,
        )
        .expect("write state");

        let host = TrayHost::new(TrayHostOptions {
            runtime_path: None,
            probe_state_path: Some(state_path),
            probe: true,
        });
        let probe = host.probe_json();

        assert_eq!(probe["native_host"], true);
        assert_eq!(probe["snapshot"]["menu_items"][0], "Status: Up");
        assert_eq!(probe["snapshot"]["menu_items"][1], "Clients: 1");
        assert_eq!(probe["snapshot"]["menu_items"][2], "Documents: 1");
        let _ = fs::remove_dir_all(dir);
    }

    #[test]
    fn linux_tray_manifest_enables_native_status_notifier_features() {
        let manifest = fs::read_to_string(env!("CARGO_MANIFEST_PATH")).expect("read manifest");

        assert!(manifest.contains("target_os = \"linux\""));
        assert!(manifest.contains("features = [\"dbus\", \"rwh_06\", \"x11\"]"));
        assert!(manifest.contains("features = [\"gtk\", \"libxdo\"]"));
    }

    #[test]
    fn native_tray_quit_uses_platform_confirmation_dialogs() {
        let source_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("src")
            .join("tray")
            .join("host.rs");
        let source = fs::read_to_string(source_path).expect("read source");

        assert!(source.contains("confirm_quit(&surface.snapshot.quit_confirmation)"));
        assert!(source.contains("impl TrayPlatformAdapter for NativeTraySurface"));
        assert!(source.contains("surface.show_ui()"));
        assert!(source.contains("surface.quit()"));
        assert!(source.contains("System.Windows.MessageBox"));
        assert!(source.contains("display dialog"));
        assert!(source.contains("zenity"));
        assert!(source.contains("kdialog"));
    }
}
