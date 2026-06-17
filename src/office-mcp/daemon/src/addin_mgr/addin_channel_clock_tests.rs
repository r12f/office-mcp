use super::format_system_time;
use std::time::{Duration, SystemTime};

#[test]
fn formats_epoch_seconds_for_heartbeat_payloads() {
    let value = SystemTime::UNIX_EPOCH + Duration::from_secs(42);

    assert_eq!(format_system_time(value), "unix:42");
}

#[test]
fn clamps_pre_epoch_times_to_zero() {
    let value = SystemTime::UNIX_EPOCH - Duration::from_secs(1);

    assert_eq!(format_system_time(value), "unix:0");
}
