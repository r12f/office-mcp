use crate::addin_mgr::SessionRegistry;
use crate::addin_mgr::websocket_accept_key;
use crate::addin_mgr::{AddinChannelServer, HeartbeatDecision};
use crate::addin_mgr::{AddinConnectionHub, CommandRouter};
use crate::addin_mgr::{WebSocketCodec, WebSocketCodecError, WebSocketFrame};
use crate::api::{UiStateOptions, UiStateStore};
use crate::common::DaemonConfig;
use crate::mcp::{HttpMethod, McpHttpFrontend, McpHttpRequest};
use crate::runtime::addin_rpc::AddinJsonRpcRuntime;
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::json_rpc;
use crate::runtime::mcp_response::{
    HeartbeatLoopDecision, McpHttpResponseService, RuntimeSharedState,
};
use crate::runtime::static_response::StaticResponseService;
use crate::runtime::ui_http::UiHttpService;
pub use crate::runtime::{RuntimeServerConfig, RuntimeServerError, default_pfx_path};
use crate::ui::UiRuntimeFile;
use native_tls::{TlsAcceptor, TlsStream};
use serde_json::Value;
use std::collections::BTreeMap;
use std::io::Write;
use std::net::{TcpListener, TcpStream};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeSeedState {
    pub ui_state: UiStateStore,
    pub registry: SessionRegistry,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RuntimeServer {
    config: RuntimeServerConfig,
}

impl RuntimeServer {
    #[must_use]
    pub fn new() -> Self {
        Self {
            config: RuntimeServerConfig::default(),
        }
    }

    #[must_use]
    pub const fn with_config(config: RuntimeServerConfig) -> Self {
        Self { config }
    }

    /// Builds a runtime server from validated daemon configuration.
    ///
    /// # Errors
    ///
    /// Returns an error when a configured port or byte limit does not fit the
    /// native runtime type used by the socket server.
    pub fn from_daemon_config(config: &DaemonConfig) -> Result<Self, RuntimeServerError> {
        Ok(Self::with_config(RuntimeServerConfig::from_daemon_config(
            config,
        )?))
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns the long-running local TCP listener for MCP HTTP"
    }

    /// Runs the MCP HTTP listener until the process exits.
    ///
    /// # Errors
    ///
    /// Returns an error when the listener cannot bind or a client connection
    /// cannot be accepted or written.
    pub fn serve_forever(&self) -> Result<(), RuntimeServerError> {
        let mcp_listener = TcpListener::bind(self.config.mcp_bind_addr())?;
        let addin_listener = TcpListener::bind(self.config.addin_bind_addr())?;
        self.serve_bound_forever(&mcp_listener, &addin_listener)
    }

    /// Runs the listeners and publishes a UI runtime file after both sockets bind.
    ///
    /// # Errors
    ///
    /// Returns an error when listeners cannot bind, the runtime file cannot be
    /// written, or a fatal server error occurs.
    pub fn serve_forever_with_runtime_file(
        &self,
        runtime_file: &UiRuntimeFile,
    ) -> Result<(), RuntimeServerError> {
        let mcp_listener = TcpListener::bind(self.config.mcp_bind_addr())?;
        let addin_listener = TcpListener::bind(self.config.addin_bind_addr())?;
        runtime_file.write()?;
        let result = self.serve_bound_forever(&mcp_listener, &addin_listener);
        if let Err(error) = runtime_file.remove() {
            tracing::warn!(%error, "failed to remove daemon UI runtime file");
            eprintln!("{error}");
        }
        result
    }

    fn serve_bound_forever(
        &self,
        mcp_listener: &TcpListener,
        addin_listener: &TcpListener,
    ) -> Result<(), RuntimeServerError> {
        self.serve_bound_with_state_forever(
            mcp_listener,
            addin_listener,
            self.ui_state_store(),
            SessionRegistry::with_limits(self.config.max_pending_per_session),
        )
    }

    fn ui_state_store(&self) -> UiStateStore {
        UiStateStore::with_options(UiStateOptions {
            mcp_endpoint: format!(
                "http://{}:{}/mcp",
                self.config.mcp_host, self.config.mcp_port
            ),
            addin_endpoint: format!("{}/addin", self.config.addin_origin),
            config_path: self.config.config_path.clone(),
            log_path: self.config.log_path.clone(),
            now: SystemTime::now(),
            ..UiStateOptions::default()
        })
    }

    /// Runs already-bound listeners with a caller-provided initial UI and
    /// session state. This is used by evidence fixtures and keeps production
    /// startup on the normal empty-state path.
    ///
    /// # Errors
    ///
    /// Returns an error when a listener cannot be cloned, TLS cannot be loaded,
    /// or a fatal server error occurs.
    pub fn serve_bound_with_state_forever(
        &self,
        mcp_listener: &TcpListener,
        addin_listener: &TcpListener,
        seed_ui_state: UiStateStore,
        seed_registry: SessionRegistry,
    ) -> Result<(), RuntimeServerError> {
        let addin_listener = addin_listener.try_clone()?;
        let tls_acceptor = Arc::new(self.config.tls_acceptor()?);
        let frontend = Arc::new(Mutex::new(McpHttpFrontend::with_config(
            self.config.mcp_http_config(),
        )));
        let ui_state = Arc::new(Mutex::new(seed_ui_state));
        let registry = Arc::new(Mutex::new(seed_registry));
        let addin_channel = Arc::new(Mutex::new(AddinChannelServer::with_config(
            self.config.addin_channel_config(),
        )));
        let connection_hub = Arc::new(AddinConnectionHub::new());
        let command_router = Arc::new(Mutex::new(CommandRouter::new()));
        let shared_state = Arc::new(RuntimeSharedState {
            registry: Arc::clone(&registry),
            addin_channel: Arc::clone(&addin_channel),
            connection_hub: Arc::clone(&connection_hub),
            command_router: Arc::clone(&command_router),
            audit_log: self.config.audit_log.clone(),
            image_fetcher: self.config.image_fetcher.clone(),
        });
        let addin_server = self.clone();
        let addin_ui_state = Arc::clone(&ui_state);
        let addin_tls_acceptor = Arc::clone(&tls_acceptor);
        let addin_shared_state = Arc::clone(&shared_state);
        thread::spawn(move || {
            loop {
                if let Err(error) = addin_server.serve_next_addin(
                    &addin_listener,
                    &addin_ui_state,
                    &addin_tls_acceptor,
                    &addin_shared_state,
                ) {
                    tracing::warn!(%error, "ignored malformed add-in client connection");
                    eprintln!(
                        "office-mcp-daemon ignored malformed add-in client connection: {error}"
                    );
                }
            }
        });
        loop {
            if let Err(error) =
                self.serve_next_mcp(mcp_listener, &frontend, &ui_state, &shared_state)
            {
                tracing::warn!(%error, "ignored malformed MCP client connection");
                eprintln!("office-mcp-daemon ignored malformed MCP client connection: {error}");
            }
        }
    }

    /// Accepts and handles one HTTP connection from an already-bound listener.
    ///
    /// # Errors
    ///
    /// Returns an error when accepting, reading, parsing, or writing the HTTP
    /// connection fails.
    pub fn serve_next(
        &self,
        listener: &TcpListener,
        frontend: &mut McpHttpFrontend,
        ui_state: &mut UiStateStore,
    ) -> Result<(), RuntimeServerError> {
        let (mut stream, peer) = listener.accept()?;
        let registry = SessionRegistry::new();
        let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
        let connection_hub = Arc::new(AddinConnectionHub::new());
        let command_router = Arc::new(Mutex::new(CommandRouter::new()));
        let shared_state = Arc::new(RuntimeSharedState {
            registry: Arc::new(Mutex::new(registry.clone())),
            addin_channel,
            connection_hub,
            command_router,
            audit_log: self.config.audit_log.clone(),
            image_fetcher: self.config.image_fetcher.clone(),
        });
        self.handle_mcp_stream(
            &mut stream,
            frontend,
            ui_state,
            &registry,
            &shared_state,
            Some(peer.ip().to_string()),
        )
    }

    fn serve_next_mcp(
        &self,
        listener: &TcpListener,
        frontend: &Arc<Mutex<McpHttpFrontend>>,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        let (mut stream, peer) = listener.accept()?;
        let mut frontend = frontend
            .lock()
            .map_err(|_| RuntimeServerError::Internal("MCP frontend lock poisoned.".to_string()))?;
        let mut ui_state = ui_state
            .lock()
            .map_err(|_| RuntimeServerError::Internal("UI state lock poisoned.".to_string()))?;
        let registry = shared_state
            .registry
            .lock()
            .map_err(|_| {
                RuntimeServerError::Internal("Session registry lock poisoned.".to_string())
            })?
            .clone();
        self.handle_mcp_stream(
            &mut stream,
            &mut frontend,
            &mut ui_state,
            &registry,
            shared_state,
            Some(peer.ip().to_string()),
        )
    }

    fn serve_next_addin(
        &self,
        listener: &TcpListener,
        ui_state: &Arc<Mutex<UiStateStore>>,
        tls_acceptor: &Arc<TlsAcceptor>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        let (stream, _) = listener.accept()?;
        let server = self.clone();
        let tls_acceptor = Arc::clone(tls_acceptor);
        let ui_state = Arc::clone(ui_state);
        let shared_state = Arc::clone(shared_state);
        thread::spawn(move || {
            let result = (|| -> Result<(), RuntimeServerError> {
                let mut stream = tls_acceptor
                    .accept(stream)
                    .map_err(|error| RuntimeServerError::Tls(error.to_string()))?;
                server.handle_addin_tls_stream(&mut stream, &ui_state, &shared_state)
            })();
            if let Err(error) = result {
                tracing::warn!(%error, "ignored malformed add-in TLS client connection");
                eprintln!("office-mcp-daemon ignored malformed add-in client connection: {error}");
            }
        });
        Ok(())
    }

    fn handle_addin_tls_stream(
        &self,
        stream: &mut TlsStream<TcpStream>,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        let request = WireHttpRequest::read_from(stream, self.config.max_request_bytes)?;
        let websocket_upgrade = self.is_valid_websocket_upgrade(&request);
        let response = self.route_addin_request(ui_state, &shared_state.registry, &request);
        stream.write_all(&response.to_bytes())?;
        stream.flush()?;
        if websocket_upgrade && response.status == 101 {
            self.handle_websocket_messages(stream, shared_state)?;
        }
        Ok(())
    }

    /// Handles a single HTTP request/response exchange on a stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the request cannot be read or the response cannot
    /// be written.
    fn handle_mcp_stream(
        &self,
        stream: &mut TcpStream,
        frontend: &mut McpHttpFrontend,
        ui_state: &mut UiStateStore,
        registry: &SessionRegistry,
        shared_state: &Arc<RuntimeSharedState>,
        remote_addr: Option<String>,
    ) -> Result<(), RuntimeServerError> {
        let request = WireHttpRequest::read_from(stream, self.config.max_request_bytes)?;
        let response = Self::route_request(
            frontend,
            ui_state,
            registry,
            shared_state,
            remote_addr,
            request,
        );
        stream.write_all(&response.to_bytes())?;
        stream.flush()?;
        Ok(())
    }

    /// Handles a single add-in/static/UI HTTP request on a stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the request cannot be read or the response cannot
    /// be written.
    pub fn handle_addin_stream(
        &self,
        stream: &mut TcpStream,
        ui_state: &UiStateStore,
    ) -> Result<(), RuntimeServerError> {
        let request = WireHttpRequest::read_from(stream, self.config.max_request_bytes)?;
        let registry = Arc::new(Mutex::new(SessionRegistry::new()));
        let ui_state = Arc::new(Mutex::new(ui_state.clone()));
        let response = self.route_addin_request(&ui_state, &registry, &request);
        stream.write_all(&response.to_bytes())?;
        stream.flush()?;
        Ok(())
    }

    fn handle_websocket_messages(
        &self,
        stream: &mut TlsStream<TcpStream>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        let connection_id = format!("addin-{:?}", thread::current().id());
        shared_state
            .connection_hub
            .register_connection(&connection_id);
        let _ = stream
            .get_ref()
            .set_read_timeout(Some(Duration::from_millis(100)));
        let mut next_ping_at = Instant::now() + self.config.heartbeat_interval;
        let mut heartbeat_deadline: Option<Instant> = None;
        loop {
            for outbound in shared_state.connection_hub.take_outbound(&connection_id) {
                WebSocketCodec::write_text(stream, &outbound)?;
            }
            if let Some(deadline) = heartbeat_deadline
                && Instant::now() >= deadline
            {
                match Self::record_heartbeat_timeout(shared_state, &connection_id)? {
                    HeartbeatLoopDecision::KeepOpen => {
                        heartbeat_deadline = None;
                        next_ping_at = Instant::now() + self.config.heartbeat_interval;
                    }
                    HeartbeatLoopDecision::Close => {
                        WebSocketCodec::write_close(stream, 4002, "Heartbeat timeout")?;
                        break;
                    }
                }
            }
            if heartbeat_deadline.is_none() && Instant::now() >= next_ping_at {
                if let Some(ping) = Self::start_heartbeat_ping(shared_state, &connection_id)? {
                    WebSocketCodec::write_text(stream, &ping)?;
                    heartbeat_deadline = Some(Instant::now() + self.config.heartbeat_timeout);
                } else {
                    next_ping_at = Instant::now() + self.config.heartbeat_interval;
                }
            }
            let frame = match WebSocketCodec::read_frame(stream, self.config.max_ws_frame_bytes) {
                Ok(Some(frame)) => frame,
                Ok(None) => break,
                Err(WebSocketCodecError::Io(error))
                    if matches!(
                        error.kind(),
                        std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut
                    ) =>
                {
                    continue;
                }
                Err(WebSocketCodecError::Protocol(error)) => {
                    WebSocketCodec::write_close(stream, error.close_code, &error.reason)?;
                    break;
                }
                Err(error) => return Err(error.into()),
            };
            match frame {
                WebSocketFrame::Text(text) => {
                    if Self::handle_heartbeat_response(shared_state, &connection_id, &text)? {
                        heartbeat_deadline = None;
                        next_ping_at = Instant::now() + self.config.heartbeat_interval;
                        continue;
                    }
                    if shared_state.connection_hub.complete_from_text(&text) {
                        continue;
                    }
                    let response = AddinJsonRpcRuntime::handle_text(
                        &text,
                        &connection_id,
                        &shared_state.registry,
                        &shared_state.addin_channel,
                        &shared_state.connection_hub,
                    );
                    if let Some(response) = response {
                        WebSocketCodec::write_text(stream, &response)?;
                    }
                }
                WebSocketFrame::Close => {
                    WebSocketCodec::write_close(stream, 1000, "Normal closure")?;
                    break;
                }
                WebSocketFrame::Ping(payload) => WebSocketCodec::write_pong(stream, &payload)?,
                WebSocketFrame::Pong => {}
            }
        }
        let stale_since = SystemTime::now();
        let mut registry = shared_state.registry.lock().expect("session registry lock");
        let mut addin_channel = shared_state
            .addin_channel
            .lock()
            .expect("addin channel lock");
        addin_channel.remove_connection(&mut registry, &connection_id, stale_since);
        shared_state
            .connection_hub
            .remove_connection(&connection_id);
        Ok(())
    }

    fn start_heartbeat_ping(
        shared_state: &RuntimeSharedState,
        connection_id: &str,
    ) -> Result<Option<String>, RuntimeServerError> {
        let mut addin_channel = shared_state.addin_channel.lock().map_err(|_| {
            RuntimeServerError::Internal("Add-in channel lock poisoned.".to_string())
        })?;
        match addin_channel.start_ping(connection_id, SystemTime::now()) {
            Ok(ping) => Ok(Some(json_rpc::envelope_to_text(&ping))),
            Err(crate::addin_mgr::AddinChannelError::UnknownConnection(_)) => Ok(None),
            Err(error) => Err(RuntimeServerError::Internal(error.to_string())),
        }
    }

    fn handle_heartbeat_response(
        shared_state: &RuntimeSharedState,
        connection_id: &str,
        text: &str,
    ) -> Result<bool, RuntimeServerError> {
        let Ok(value) = serde_json::from_str::<Value>(text) else {
            return Ok(false);
        };
        if value.get("method").is_some() || value.get("result").is_none() {
            return Ok(false);
        }
        let Some(response_id) = value.get("id").and_then(Value::as_str) else {
            return Ok(false);
        };
        let mut addin_channel = shared_state.addin_channel.lock().map_err(|_| {
            RuntimeServerError::Internal("Add-in channel lock poisoned.".to_string())
        })?;
        addin_channel
            .handle_pong(connection_id, response_id)
            .map_err(|error| RuntimeServerError::Internal(error.to_string()))
    }

    fn record_heartbeat_timeout(
        shared_state: &RuntimeSharedState,
        connection_id: &str,
    ) -> Result<HeartbeatLoopDecision, RuntimeServerError> {
        let mut registry = shared_state.registry.lock().map_err(|_| {
            RuntimeServerError::Internal("Session registry lock poisoned.".to_string())
        })?;
        let mut addin_channel = shared_state.addin_channel.lock().map_err(|_| {
            RuntimeServerError::Internal("Add-in channel lock poisoned.".to_string())
        })?;
        match addin_channel.record_heartbeat_timeout(
            &mut registry,
            connection_id,
            SystemTime::now(),
        ) {
            Ok(HeartbeatDecision::KeepOpen) => Ok(HeartbeatLoopDecision::KeepOpen),
            Ok(HeartbeatDecision::Close { .. }) => Ok(HeartbeatLoopDecision::Close),
            Err(crate::addin_mgr::AddinChannelError::UnknownConnection(_)) => {
                Ok(HeartbeatLoopDecision::KeepOpen)
            }
            Err(error) => Err(RuntimeServerError::Internal(error.to_string())),
        }
    }

    fn route_addin_request(
        &self,
        ui_state: &Arc<Mutex<UiStateStore>>,
        registry: &Arc<Mutex<SessionRegistry>>,
        request: &WireHttpRequest,
    ) -> WireHttpResponse {
        if request.path == "/healthz" && request.method == HttpMethod::Get {
            return WireHttpResponse::json(200, BTreeMap::new(), "{\"ok\":true}".to_string());
        }
        if request.path == "/addin" {
            return self.route_websocket_upgrade(request);
        }
        if request.method == HttpMethod::Get
            && let Some(response) = self
                .ui_http_service()
                .try_handle(request, ui_state, registry)
        {
            return response;
        }
        if request.method != HttpMethod::Get {
            return WireHttpResponse::text(405, "Method not allowed".to_string());
        }
        self.static_response_service()
            .serve_addin_asset(&request.path)
    }

    fn route_websocket_upgrade(&self, request: &WireHttpRequest) -> WireHttpResponse {
        if request.method != HttpMethod::Get {
            return WireHttpResponse::text(405, "Method not allowed".to_string());
        }
        if request.headers.get("origin") != Some(&self.config.addin_origin) {
            return WireHttpResponse::text(403, "Forbidden origin".to_string());
        }
        let upgrade = request
            .headers
            .get("upgrade")
            .is_some_and(|value| value.eq_ignore_ascii_case("websocket"));
        let connection = request
            .headers
            .get("connection")
            .is_some_and(|value| value.to_ascii_lowercase().contains("upgrade"));
        let Some(key) = request.headers.get("sec-websocket-key") else {
            return WireHttpResponse::text(400, "Missing Sec-WebSocket-Key".to_string());
        };
        if !upgrade || !connection {
            return WireHttpResponse::text(400, "Invalid WebSocket upgrade".to_string());
        }
        WireHttpResponse::switching_protocols(BTreeMap::from([
            ("Upgrade".to_string(), "websocket".to_string()),
            ("Connection".to_string(), "Upgrade".to_string()),
            (
                "Sec-WebSocket-Accept".to_string(),
                websocket_accept_key(key),
            ),
        ]))
    }

    fn is_valid_websocket_upgrade(&self, request: &WireHttpRequest) -> bool {
        request.path == "/addin"
            && request.method == HttpMethod::Get
            && request.headers.get("origin") == Some(&self.config.addin_origin)
            && request
                .headers
                .get("upgrade")
                .is_some_and(|value| value.eq_ignore_ascii_case("websocket"))
            && request
                .headers
                .get("connection")
                .is_some_and(|value| value.to_ascii_lowercase().contains("upgrade"))
            && request.headers.contains_key("sec-websocket-key")
    }

    fn static_response_service(&self) -> StaticResponseService {
        StaticResponseService::new(self.config.addin_public_dir.clone())
    }

    fn ui_http_service(&self) -> UiHttpService {
        UiHttpService::from_config(&self.config)
    }

    fn route_request(
        frontend: &mut McpHttpFrontend,
        ui_state: &mut UiStateStore,
        registry: &SessionRegistry,
        shared_state: &Arc<RuntimeSharedState>,
        remote_addr: Option<String>,
        request: WireHttpRequest,
    ) -> WireHttpResponse {
        if request.path == "/healthz" && request.method == HttpMethod::Get {
            return WireHttpResponse::json(200, BTreeMap::new(), "{\"ok\":true}".to_string());
        }
        if request.path != "/mcp" {
            return WireHttpResponse::text(404, "Not found".to_string());
        }
        let body_bytes = request.body.len();
        let is_initialize = request.is_initialize();
        let body = request.body;
        let decision = frontend.handle_request(
            ui_state,
            &McpHttpRequest {
                method: request.method,
                headers: request.headers,
                remote_addr,
                body_bytes,
                is_initialize,
            },
            SystemTime::now(),
        );
        McpHttpResponseService::runtime_response(decision, registry, ui_state, shared_state, &body)
    }
}

impl Default for RuntimeServer {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
#[path = "server_tests.rs"]
mod server_tests;
