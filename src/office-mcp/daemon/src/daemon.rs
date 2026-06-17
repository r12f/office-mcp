use crate::addin_mgr::AddinChannelServer;
use crate::addin_mgr::CommandRouter;
use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::common::DaemonConfigService;
use crate::common::Logger;
use crate::mcp::McpHttpFrontend;
use crate::parity::ParityPlan;
use crate::runtime_server::RuntimeServer;
use crate::tray::TrayController;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ComponentDescription {
    name: &'static str,
    description: &'static str,
}

impl ComponentDescription {
    #[must_use]
    pub const fn new(name: &'static str, description: &'static str) -> Self {
        Self { name, description }
    }

    #[must_use]
    pub const fn name(&self) -> &'static str {
        self.name
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        self.description
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeMcpDaemon {
    config_service: DaemonConfigService,
    mcp_http_frontend: McpHttpFrontend,
    addin_channel: AddinChannelServer,
    session_registry: SessionRegistry,
    command_router: CommandRouter,
    ui_state_store: UiStateStore,
    tray_controller: TrayController,
    audit_log: AuditLog,
    logger: Logger,
    runtime_server: RuntimeServer,
    parity_plan: ParityPlan,
}

impl OfficeMcpDaemon {
    #[must_use]
    pub fn new() -> Self {
        Self {
            config_service: DaemonConfigService::new(),
            mcp_http_frontend: McpHttpFrontend::new(),
            addin_channel: AddinChannelServer::new(),
            session_registry: SessionRegistry::new(),
            command_router: CommandRouter::new(),
            ui_state_store: UiStateStore::new(),
            tray_controller: TrayController::new(),
            audit_log: AuditLog::new(),
            logger: Logger::new(),
            runtime_server: RuntimeServer::new(),
            parity_plan: ParityPlan::rust_runtime_readiness(),
        }
    }

    #[must_use]
    pub fn component_descriptions(&self) -> Vec<ComponentDescription> {
        vec![
            ComponentDescription::new("DaemonConfigService", self.config_service.description()),
            ComponentDescription::new("McpHttpFrontend", self.mcp_http_frontend.description()),
            ComponentDescription::new("AddinChannelServer", self.addin_channel.description()),
            ComponentDescription::new("SessionRegistry", self.session_registry.description()),
            ComponentDescription::new("CommandRouter", self.command_router.description()),
            ComponentDescription::new("UiStateStore", self.ui_state_store.description()),
            ComponentDescription::new("TrayController", self.tray_controller.description()),
            ComponentDescription::new("AuditLog", self.audit_log.description()),
            ComponentDescription::new("Logger", self.logger.description()),
            ComponentDescription::new("RuntimeServer", self.runtime_server.description()),
        ]
    }

    #[must_use]
    pub const fn parity_plan(&self) -> &ParityPlan {
        &self.parity_plan
    }
}

impl Default for OfficeMcpDaemon {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::OfficeMcpDaemon;
    use std::collections::BTreeSet;

    #[test]
    fn daemon_owns_all_required_domain_objects() {
        let daemon = OfficeMcpDaemon::new();
        let components = daemon.component_descriptions();
        let names = components
            .iter()
            .map(super::ComponentDescription::name)
            .collect::<Vec<_>>();

        assert_eq!(names.len(), 10);
        assert!(names.contains(&"DaemonConfigService"));
        assert!(names.contains(&"McpHttpFrontend"));
        assert!(names.contains(&"AddinChannelServer"));
        assert!(names.contains(&"SessionRegistry"));
        assert!(names.contains(&"CommandRouter"));
        assert!(names.contains(&"UiStateStore"));
        assert!(names.contains(&"TrayController"));
        assert!(names.contains(&"AuditLog"));
        assert!(names.contains(&"Logger"));
        assert!(names.contains(&"RuntimeServer"));
    }

    #[test]
    fn daemon_tracks_rust_runtime_readiness_gates() {
        let daemon = OfficeMcpDaemon::new();

        assert_eq!(daemon.parity_plan().gates().len(), 8);
    }

    #[test]
    fn daemon_src_root_only_contains_composition_and_transitional_files() {
        let src_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("src");
        let allowed_root_files = BTreeSet::from([
            "daemon.rs",
            "evidence_fixture.rs",
            "lib.rs",
            "main.rs",
            "parity.rs",
            "runtime_server.rs",
        ]);
        let forbidden_service_files = BTreeSet::from([
            "addin_channel.rs",
            "audit_log.rs",
            "client_config.rs",
            "command_router.rs",
            "config_service.rs",
            "daemon_control.rs",
            "image_fetcher.rs",
            "logger.rs",
            "mcp_http_frontend.rs",
            "mcp_management_client.rs",
            "session_registry.rs",
            "state_store.rs",
            "stdio_bridge.rs",
        ]);

        let root_files = std::fs::read_dir(&src_dir)
            .expect("read daemon src dir")
            .filter_map(Result::ok)
            .filter(|entry| {
                entry
                    .file_type()
                    .map(|file_type| file_type.is_file())
                    .unwrap_or(false)
            })
            .map(|entry| entry.file_name().to_string_lossy().into_owned())
            .collect::<BTreeSet<_>>();

        for file in &forbidden_service_files {
            assert!(
                !root_files.contains(*file),
                "service module {file} must stay under its functional module directory"
            );
        }

        for file in &root_files {
            assert!(
                allowed_root_files.contains(file.as_str()),
                "unexpected daemon src root file {file}; place service code under common, ui, api, mcp, addin_mgr, or tray"
            );
        }
    }
}
