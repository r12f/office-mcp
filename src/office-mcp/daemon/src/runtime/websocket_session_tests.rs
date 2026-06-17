use super::{RuntimeWebSocketSession, RuntimeWebSocketSessionConfig};
use crate::runtime::RuntimeServerConfig;
use std::time::Duration;

#[test]
fn session_config_keeps_websocket_runtime_limits() {
    let server_config = RuntimeServerConfig {
        heartbeat_interval: Duration::from_secs(7),
        heartbeat_timeout: Duration::from_secs(3),
        max_ws_frame_bytes: 1234,
        ..RuntimeServerConfig::default()
    };
    let config = RuntimeWebSocketSessionConfig::from_config(&server_config);

    assert_eq!(config.heartbeat_interval, Duration::from_secs(7));
    assert_eq!(config.heartbeat_timeout, Duration::from_secs(3));
    assert_eq!(config.max_ws_frame_bytes, 1234);
}

#[test]
fn websocket_session_is_built_from_server_config() {
    let session = RuntimeWebSocketSession::from_config(&RuntimeServerConfig::default());

    assert_eq!(
        session,
        RuntimeWebSocketSession {
            config: RuntimeWebSocketSessionConfig::from_config(&RuntimeServerConfig::default())
        }
    );
}
