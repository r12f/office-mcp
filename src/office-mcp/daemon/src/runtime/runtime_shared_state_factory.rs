use crate::addin_mgr::{AddinChannelServer, AddinConnectionHub, CommandRouter, SessionRegistry};
use crate::runtime::RuntimeServerConfig;
use crate::runtime::mcp_response::RuntimeSharedState;
use std::sync::{Arc, Mutex};

pub(crate) struct RuntimeSharedStateFactory;

impl RuntimeSharedStateFactory {
    pub(crate) fn with_registry(
        config: &RuntimeServerConfig,
        registry: SessionRegistry,
    ) -> Arc<RuntimeSharedState> {
        Arc::new(RuntimeSharedState {
            registry: Arc::new(Mutex::new(registry)),
            session_grace: config.session_grace,
            addin_channel: Arc::new(Mutex::new(AddinChannelServer::with_config(
                config.addin_channel_config(),
            ))),
            connection_hub: Arc::new(AddinConnectionHub::new()),
            command_router: Arc::new(Mutex::new(CommandRouter::new())),
            audit_log: config.audit_log.clone(),
            image_fetcher: config.image_fetcher.clone(),
            tool_access_policy: config.tool_access_policy.clone(),
        })
    }
}

#[cfg(test)]
#[path = "runtime_shared_state_factory_tests.rs"]
mod runtime_shared_state_factory_tests;
