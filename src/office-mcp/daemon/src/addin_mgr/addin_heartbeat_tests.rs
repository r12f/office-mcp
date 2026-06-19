use super::{AddinHeartbeatState, AddinHeartbeatTimeout};

#[test]
fn matching_pong_clears_pending_ping() {
    let mut heartbeat = AddinHeartbeatState::default();
    heartbeat.start_ping("ping-1".to_string());

    assert!(heartbeat.handle_pong("ping-1"));
    assert!(!heartbeat.handle_pong("ping-1"));
}

#[test]
fn mismatched_pong_is_ignored() {
    let mut heartbeat = AddinHeartbeatState::default();
    heartbeat.start_ping("ping-1".to_string());

    assert!(!heartbeat.handle_pong("ping-other"));
    assert!(heartbeat.handle_pong("ping-1"));
}

#[test]
fn first_timeout_keeps_connection_open() {
    let mut heartbeat = AddinHeartbeatState::default();
    heartbeat.start_ping("ping-1".to_string());

    assert_eq!(
        heartbeat.record_timeout(),
        AddinHeartbeatTimeout::KeepOpen { missed_pongs: 1 }
    );
    assert!(!heartbeat.handle_pong("ping-1"));
}

#[test]
fn third_timeout_closes_connection() {
    let mut heartbeat = AddinHeartbeatState::default();

    assert_eq!(
        heartbeat.record_timeout(),
        AddinHeartbeatTimeout::KeepOpen { missed_pongs: 1 }
    );
    assert_eq!(
        heartbeat.record_timeout(),
        AddinHeartbeatTimeout::KeepOpen { missed_pongs: 2 }
    );
    assert_eq!(
        heartbeat.record_timeout(),
        AddinHeartbeatTimeout::Close {
            missed_pongs: 3,
            close_code: 4002
        }
    );
}

#[test]
fn successful_pong_resets_missed_count() {
    let mut heartbeat = AddinHeartbeatState::default();
    assert_eq!(
        heartbeat.record_timeout(),
        AddinHeartbeatTimeout::KeepOpen { missed_pongs: 1 }
    );
    heartbeat.start_ping("ping-2".to_string());
    assert!(heartbeat.handle_pong("ping-2"));

    assert_eq!(
        heartbeat.record_timeout(),
        AddinHeartbeatTimeout::KeepOpen { missed_pongs: 1 }
    );
}
