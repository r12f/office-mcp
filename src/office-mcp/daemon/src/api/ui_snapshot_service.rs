use crate::addin_mgr::SessionRegistry;
use crate::api::{UiSnapshotRenderer, UiStateStore};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiSnapshotEndpoints {
    pub mcp_endpoint: String,
    pub addin_endpoint: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct UiSnapshotService {
    renderer: UiSnapshotRenderer,
}

impl UiSnapshotService {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            renderer: UiSnapshotRenderer::new(),
        }
    }

    #[must_use]
    pub fn render_runtime_snapshot(
        &self,
        ui_state: &Arc<Mutex<UiStateStore>>,
        registry: &Arc<Mutex<SessionRegistry>>,
        endpoints: &UiSnapshotEndpoints,
    ) -> String {
        let sessions = match registry.lock() {
            Ok(registry) => registry.list_sessions(),
            Err(error) => {
                tracing::warn!(
                    component = "ui_snapshot_service",
                    %error,
                    "rendering UI snapshot without registry after lock failure"
                );
                Vec::new()
            }
        };
        let ui_state_lock_failed;
        let mut snapshot = match ui_state.lock() {
            Ok(ui_state) => {
                ui_state_lock_failed = false;
                ui_state.snapshot(&sessions, SystemTime::now())
            }
            Err(error) => {
                ui_state_lock_failed = true;
                tracing::warn!(
                    component = "ui_snapshot_service",
                    %error,
                    "rendering fallback UI snapshot after state lock failure"
                );
                UiStateStore::new().snapshot(&sessions, SystemTime::now())
            }
        };
        snapshot.daemon.mcp_endpoint = endpoints.mcp_endpoint.clone();
        snapshot.daemon.addin_endpoint = endpoints.addin_endpoint.clone();
        let rendered = self.renderer.render_text(&snapshot);
        tracing::debug!(
            component = "ui_snapshot_service",
            status = ?snapshot.daemon.status,
            clients = snapshot.clients.len(),
            documents = snapshot.documents.values().map(Vec::len).sum::<usize>(),
            current_tasks = snapshot.current_tasks.len(),
            recent_commands = snapshot.recent_commands.len(),
            ui_state_lock_failed,
            bytes = rendered.len(),
            "rendered UI API snapshot"
        );
        rendered
    }
}

#[cfg(test)]
#[path = "ui_snapshot_service_tests.rs"]
mod ui_snapshot_service_tests;
