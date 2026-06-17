pub mod daemon_control;
pub mod daemon_status;
pub mod state_model;
pub mod state_store;
pub(crate) mod ui_redaction;
pub mod ui_snapshot_renderer;
pub mod ui_snapshot_service;

pub use daemon_control::{DaemonControlError, DaemonController, PowerShellExecutor};
pub use daemon_status::DaemonStatusReporter;

pub use state_model::{
    CommandFailure, CommandResult, RegisterClientInput, StartCommandInput, UiClientRecord,
    UiClientTransport, UiCommandError, UiCommandRecord, UiCommandStatus, UiDaemonSnapshot,
    UiHealth, UiSnapshot, UiStateOptions,
};
pub use state_store::UiStateStore;
pub use ui_snapshot_renderer::UiSnapshotRenderer;
pub use ui_snapshot_service::{UiSnapshotEndpoints, UiSnapshotService};
