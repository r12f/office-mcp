use super::AddinChannelConfig;
use std::time::Duration;

#[test]
fn default_config_matches_local_addin_channel_contract() {
    let config = AddinChannelConfig::default();

    assert_eq!(config.origin, "https://localhost:8765");
    assert_eq!(config.session_grace, Duration::from_secs(60));
    assert_eq!(config.heartbeat_interval, Duration::from_secs(30));
    assert_eq!(config.heartbeat_timeout, Duration::from_secs(20));
    assert_eq!(config.max_pending_per_session, 4);
}
