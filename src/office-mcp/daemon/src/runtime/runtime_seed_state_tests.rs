use super::RuntimeSeedState;
use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;

#[test]
fn seed_state_keeps_ui_state_and_registry_together() {
    let state = RuntimeSeedState {
        ui_state: UiStateStore::new(),
        registry: SessionRegistry::new(),
    };

    assert_eq!(state.clone(), state);
    assert!(state.registry.list_sessions().is_empty());
}
