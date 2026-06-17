use crate::addin_mgr::{WebSocketCodec, WebSocketCodecError, WebSocketFrame};
use crate::runtime::addin_rpc::AddinJsonRpcRuntime;
use crate::runtime::mcp_response::{HeartbeatLoopDecision, RuntimeSharedState};
use crate::runtime::websocket_heartbeat::WebSocketHeartbeatState;
use crate::runtime::websocket_heartbeat_service::WebSocketHeartbeatService;
use crate::runtime::{RuntimeServerConfig, RuntimeServerError};
use native_tls::TlsStream;
use std::net::TcpStream;
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RuntimeWebSocketSession {
    config: RuntimeWebSocketSessionConfig,
}

impl RuntimeWebSocketSession {
    #[must_use]
    pub(crate) fn from_config(config: &RuntimeServerConfig) -> Self {
        Self {
            config: RuntimeWebSocketSessionConfig::from_config(config),
        }
    }

    pub(crate) fn handle(
        &self,
        stream: &mut TlsStream<TcpStream>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        let connection_id = format!("addin-{:?}", thread::current().id());
        shared_state
            .connection_hub
            .register_connection(&connection_id);
        let result = self.run_loop(stream, shared_state, &connection_id);
        self.remove_connection(shared_state, &connection_id);
        result
    }

    fn run_loop(
        &self,
        stream: &mut TlsStream<TcpStream>,
        shared_state: &Arc<RuntimeSharedState>,
        connection_id: &str,
    ) -> Result<(), RuntimeServerError> {
        let _ = stream
            .get_ref()
            .set_read_timeout(Some(Duration::from_millis(100)));
        let mut heartbeat = WebSocketHeartbeatState::new(
            Instant::now(),
            self.config.heartbeat_interval,
            self.config.heartbeat_timeout,
        );
        loop {
            for outbound in shared_state.connection_hub.take_outbound(connection_id) {
                WebSocketCodec::write_text(stream, &outbound)?;
            }
            let now = Instant::now();
            if heartbeat.deadline_elapsed(now) {
                match WebSocketHeartbeatService::record_timeout(shared_state, connection_id)? {
                    HeartbeatLoopDecision::KeepOpen => heartbeat.mark_pong_received(now),
                    HeartbeatLoopDecision::Close => {
                        WebSocketCodec::write_close(stream, 4002, "Heartbeat timeout")?;
                        break;
                    }
                }
            }
            let now = Instant::now();
            if heartbeat.should_start_ping(now) {
                if let Some(ping) =
                    WebSocketHeartbeatService::start_ping(shared_state, connection_id)?
                {
                    WebSocketCodec::write_text(stream, &ping)?;
                    heartbeat.mark_ping_sent(Instant::now());
                } else {
                    heartbeat.mark_ping_skipped(Instant::now());
                }
            }
            match self.read_next_frame(stream)? {
                Some(WebSocketFrame::Text(text)) => {
                    if WebSocketHeartbeatService::handle_response(
                        shared_state,
                        connection_id,
                        &text,
                    )? {
                        heartbeat.mark_pong_received(Instant::now());
                        continue;
                    }
                    if shared_state.connection_hub.complete_from_text(&text) {
                        continue;
                    }
                    let response = AddinJsonRpcRuntime::handle_text(
                        &text,
                        connection_id,
                        &shared_state.registry,
                        &shared_state.addin_channel,
                        &shared_state.connection_hub,
                    );
                    if let Some(response) = response {
                        WebSocketCodec::write_text(stream, &response)?;
                    }
                }
                Some(WebSocketFrame::Close) => {
                    WebSocketCodec::write_close(stream, 1000, "Normal closure")?;
                    break;
                }
                Some(WebSocketFrame::Ping(payload)) => {
                    WebSocketCodec::write_pong(stream, &payload)?
                }
                Some(WebSocketFrame::Pong) => {}
                None => break,
            }
        }
        Ok(())
    }

    fn read_next_frame(
        &self,
        stream: &mut TlsStream<TcpStream>,
    ) -> Result<Option<WebSocketFrame>, RuntimeServerError> {
        match WebSocketCodec::read_frame(stream, self.config.max_ws_frame_bytes) {
            Ok(frame) => Ok(frame),
            Err(WebSocketCodecError::Io(error))
                if matches!(
                    error.kind(),
                    std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut
                ) =>
            {
                Ok(Some(WebSocketFrame::Pong))
            }
            Err(WebSocketCodecError::Protocol(error)) => {
                WebSocketCodec::write_close(stream, error.close_code, &error.reason)?;
                Ok(None)
            }
            Err(error) => Err(error.into()),
        }
    }

    fn remove_connection(&self, shared_state: &Arc<RuntimeSharedState>, connection_id: &str) {
        let stale_since = SystemTime::now();
        let mut registry = shared_state.registry.lock().expect("session registry lock");
        let mut addin_channel = shared_state
            .addin_channel
            .lock()
            .expect("addin channel lock");
        addin_channel.remove_connection(&mut registry, connection_id, stale_since);
        shared_state.connection_hub.remove_connection(connection_id);
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RuntimeWebSocketSessionConfig {
    heartbeat_interval: Duration,
    heartbeat_timeout: Duration,
    max_ws_frame_bytes: usize,
}

impl RuntimeWebSocketSessionConfig {
    #[must_use]
    pub(crate) const fn from_config(config: &RuntimeServerConfig) -> Self {
        Self {
            heartbeat_interval: config.heartbeat_interval,
            heartbeat_timeout: config.heartbeat_timeout,
            max_ws_frame_bytes: config.max_ws_frame_bytes,
        }
    }
}

#[cfg(test)]
#[path = "websocket_session_tests.rs"]
mod websocket_session_tests;
