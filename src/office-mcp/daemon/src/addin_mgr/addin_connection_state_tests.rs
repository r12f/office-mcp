use super::AddinConnectionState;
use crate::addin_mgr::addin_heartbeat::AddinHeartbeatState;

#[test]
fn new_connection_state_starts_without_session_and_default_heartbeat() {
    let state = AddinConnectionState::new("instance-1".to_string());

    assert_eq!(state.instance_id, "instance-1");
    assert_eq!(state.session_id, None);
    assert_eq!(state.heartbeat, AddinHeartbeatState::default());
}
