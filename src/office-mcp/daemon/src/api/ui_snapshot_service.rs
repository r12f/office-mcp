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
        let sessions = registry
            .lock()
            .map(|registry| registry.list_sessions())
            .unwrap_or_default();
        let mut snapshot = ui_state.lock().map_or_else(
            |_| UiStateStore::new().snapshot(&sessions, SystemTime::now()),
            |ui_state| ui_state.snapshot(&sessions, SystemTime::now()),
        );
        snapshot.daemon.mcp_endpoint = endpoints.mcp_endpoint.clone();
        snapshot.daemon.addin_endpoint = endpoints.addin_endpoint.clone();
        self.renderer.render_text(&snapshot)
    }
}

#[cfg(test)]
#[path = "ui_snapshot_service_tests.rs"]
mod ui_snapshot_service_tests;
