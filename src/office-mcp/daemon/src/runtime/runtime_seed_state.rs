use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeSeedState {
    pub ui_state: UiStateStore,
    pub registry: SessionRegistry,
}

#[cfg(test)]
#[path = "runtime_seed_state_tests.rs"]
mod runtime_seed_state_tests;
