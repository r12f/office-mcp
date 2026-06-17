use crate::addin_mgr::AddinChannelServer;
use crate::addin_mgr::CommandRouter;
use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::common::DaemonConfigService;
use crate::common::Logger;
use crate::mcp::McpHttpFrontend;
use crate::parity::ParityPlan;
use crate::runtime::RuntimeServer;
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
#[path = "daemon_tests.rs"]
mod daemon_tests;
