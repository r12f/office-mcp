pub mod controller;
pub mod host;
pub mod model;
pub(crate) mod product_icon;
pub mod ui_launch;
pub(crate) mod ui_state_client;

pub use controller::{TrayController, TrayPlatformAdapter, TrayPlatformError};
pub use host::{
    TrayHost, TrayHostOptions, open_ui_from_runtime, start_tray_background, stop_daemon,
};
pub use model::{
    QuitConfirmation, TrayAction, TrayHealth, TrayMenuItem, TrayPlatform, TraySnapshot,
    TrayStatusInput,
};
