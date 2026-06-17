pub mod daemon_control;
pub mod state_store;
pub(crate) mod ui_redaction;
pub mod ui_snapshot_renderer;
pub mod ui_snapshot_service;

pub use daemon_control::{DaemonControlError, DaemonController, PowerShellExecutor};

pub use state_store::{
    CommandFailure, CommandResult, RegisterClientInput, StartCommandInput, UiClientRecord,
    UiClientTransport, UiCommandError, UiCommandRecord, UiCommandStatus, UiDaemonSnapshot,
    UiHealth, UiSnapshot, UiStateOptions, UiStateStore,
};
pub use ui_snapshot_renderer::UiSnapshotRenderer;
pub use ui_snapshot_service::{UiSnapshotEndpoints, UiSnapshotService};
