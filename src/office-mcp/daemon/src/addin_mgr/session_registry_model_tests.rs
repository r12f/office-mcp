use super::{
    AddinConnectionState, HostInfo, RuntimeInfo, infer_capability_tiers, normalize_host_app,
};
use crate::addin_mgr::AddInInfo;
use std::time::SystemTime;

#[test]
fn normalizes_known_and_unknown_host_apps() {
    assert_eq!(normalize_host_app("Word"), "word");
    assert_eq!(normalize_host_app("Excel"), "excel");
    assert_eq!(normalize_host_app("Visio"), "other");
}

#[test]
fn infers_capability_tiers_from_available_tools() {
    assert_eq!(infer_capability_tiers(&[]), vec!["core".to_string()]);
    assert_eq!(
        infer_capability_tiers(&[
            "word.add_comment".to_string(),
            "word.accept_change".to_string(),
        ]),
        vec![
            "core".to_string(),
            "review".to_string(),
            "tracked_changes".to_string(),
        ]
    );
}

#[test]
fn connection_state_starts_connected_without_session_or_pending_calls() {
    let state = AddinConnectionState::new(RuntimeInfo {
        instance_id: "instance-1".to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: None,
            platform: None,
            build: None,
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: Vec::new(),
        },
        registered_at: SystemTime::UNIX_EPOCH,
    });

    assert_eq!(state.runtime.instance_id, "instance-1");
    assert!(state.connected);
    assert_eq!(state.pending_count, 0);
    assert_eq!(state.session_id, None);
}
