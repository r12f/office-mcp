pub mod controller;
pub mod host;

pub use controller::{
    QuitConfirmation, TrayAction, TrayController, TrayHealth, TrayMenuItem, TrayPlatform,
    TrayPlatformAdapter, TrayPlatformError, TraySnapshot, TrayStatusInput,
};
pub use host::{
    TrayHost, TrayHostOptions, open_ui_from_runtime, start_tray_background, stop_daemon,
};
