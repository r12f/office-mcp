use super::WebSocketHeartbeatState;
use std::time::{Duration, Instant};

#[test]
fn waits_until_interval_before_starting_ping() {
    let now = Instant::now();
    let state = WebSocketHeartbeatState::new(now, Duration::from_secs(30), Duration::from_secs(10));

    assert!(!state.should_start_ping(now + Duration::from_secs(29)));
    assert!(state.should_start_ping(now + Duration::from_secs(30)));
}

#[test]
fn ping_sent_starts_timeout_deadline() {
    let now = Instant::now();
    let mut state =
        WebSocketHeartbeatState::new(now, Duration::from_secs(30), Duration::from_secs(10));

    state.mark_ping_sent(now + Duration::from_secs(30));

    assert!(!state.deadline_elapsed(now + Duration::from_secs(39)));
    assert!(state.deadline_elapsed(now + Duration::from_secs(40)));
    assert!(!state.should_start_ping(now + Duration::from_secs(41)));
}

#[test]
fn pong_received_resets_next_ping_interval() {
    let now = Instant::now();
    let mut state =
        WebSocketHeartbeatState::new(now, Duration::from_secs(30), Duration::from_secs(10));

    state.mark_ping_sent(now + Duration::from_secs(30));
    state.mark_pong_received(now + Duration::from_secs(35));

    assert!(!state.deadline_elapsed(now + Duration::from_secs(45)));
    assert!(!state.should_start_ping(now + Duration::from_secs(64)));
    assert!(state.should_start_ping(now + Duration::from_secs(65)));
}

#[test]
fn skipped_ping_defers_until_next_interval() {
    let now = Instant::now();
    let mut state =
        WebSocketHeartbeatState::new(now, Duration::from_secs(30), Duration::from_secs(10));

    state.mark_ping_skipped(now + Duration::from_secs(30));

    assert!(!state.deadline_elapsed(now + Duration::from_mins(1)));
    assert!(!state.should_start_ping(now + Duration::from_secs(59)));
    assert!(state.should_start_ping(now + Duration::from_mins(1)));
}
