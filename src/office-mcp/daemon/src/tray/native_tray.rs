use super::{TrayHostOptions, open_ui_from_runtime, read_ui_state, stop_daemon};
use crate::tray::product_icon::{ICON_HEIGHT, ICON_WIDTH, product_icon_rgba};
use crate::tray::{
    QuitConfirmation, TrayController, TrayMenuItem, TrayPlatformAdapter, TrayPlatformError,
    TraySnapshot,
};
use std::time::{Duration, Instant};
use tao::event::{Event, StartCause};
use tao::event_loop::{ControlFlow, EventLoop, EventLoopBuilder};
use tray_icon::menu::{Menu, MenuEvent, MenuId, MenuItem, PredefinedMenuItem};
use tray_icon::{Icon, TrayIcon, TrayIconBuilder};

const SHOW_ID: &str = "office-mcp-show";
const QUIT_ID: &str = "office-mcp-quit";

pub fn run(options: &TrayHostOptions) -> Result<(), TrayPlatformError> {
    let event_loop = tray_event_loop();
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
                    } else if id == QUIT_ID && confirm_quit(&surface.snapshot.quit_confirmation) {
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

fn tray_event_loop() -> EventLoop<()> {
    let mut builder = EventLoopBuilder::new();
    allow_background_thread_event_loop(&mut builder);
    builder.build()
}

#[cfg(windows)]
fn allow_background_thread_event_loop(builder: &mut EventLoopBuilder<()>) {
    use tao::platform::windows::EventLoopBuilderExtWindows;

    builder.with_any_thread(true);
}

#[cfg(all(unix, not(target_os = "macos")))]
fn allow_background_thread_event_loop(builder: &mut EventLoopBuilder<()>) {
    use tao::platform::unix::EventLoopBuilderExtUnix;

    builder.with_any_thread(true);
}

#[cfg(target_os = "macos")]
fn allow_background_thread_event_loop(_builder: &mut EventLoopBuilder<()>) {}

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
        tracing::info!(
            component = "tray_host",
            tooltip = %snapshot.tooltip,
            platform = %snapshot.platform.label(),
            "created native tray icon"
        );
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
        tracing::debug!(
            component = "tray_host",
            tooltip = %snapshot.tooltip,
            "refreshing tray snapshot"
        );
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
        tracing::debug!(
            component = "tray_host",
            tooltip = %self.snapshot.tooltip,
            menu_items = self.snapshot.menu.len(),
            "applied tray snapshot"
        );
        Ok(())
    }

    fn show_ui(&mut self) -> Result<(), TrayPlatformError> {
        tracing::info!(component = "tray_host", "tray action show UI selected");
        open_ui_from_runtime(&self.options)
    }

    fn quit(&mut self) -> Result<(), TrayPlatformError> {
        tracing::info!(component = "tray_host", "tray action quit selected");
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
    Icon::from_rgba(product_icon_rgba(), ICON_WIDTH, ICON_HEIGHT)
        .expect("generated tray icon is valid")
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
