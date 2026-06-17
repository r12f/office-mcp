use crate::addin_mgr::websocket_accept_key;
use crate::addin_mgr::{AddinChannelServer, HeartbeatDecision};
use crate::addin_mgr::{
    AddinConnectionHub, AddinConnectionHubError, CommandRouter, ImageFetcher, ToolCallRequest,
    ToolResponse,
};
use crate::addin_mgr::{SessionDescriptorView, SessionRegistry};
use crate::addin_mgr::{WebSocketCodec, WebSocketCodecError, WebSocketFrame};
use crate::api::{
    CommandFailure, UiSnapshotEndpoints, UiSnapshotService, UiStateOptions, UiStateStore,
};
use crate::common::DaemonConfig;
use crate::common::{AuditLog, AuditRecord};
use crate::mcp::{
    ExcelToolCatalog, HttpMethod, McpHttpDecision, McpHttpFrontend, McpHttpRequest,
    ResourceReadRequest, WORD_V1_TOOLS, prompt_catalog_json, prompt_description, prompt_messages,
    resource_request_from_uri, tool_catalog_json, tool_failure, tool_failure_from_command,
    tool_success, word_resource_catalog_for_session, word_resource_templates,
};
use crate::runtime::addin_rpc::AddinJsonRpcRuntime;
use crate::runtime::http_wire::{WireHttpRequest, WireHttpResponse};
use crate::runtime::json_rpc;
use crate::runtime::static_response::StaticResponseService;
pub use crate::runtime::{RuntimeServerConfig, RuntimeServerError, default_pfx_path};
use crate::ui::UiRuntimeFile;
use native_tls::{TlsAcceptor, TlsStream};
use serde_json::{Value, json};
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
        if matches!(request.path.as_str(), "/ui" | "/ui/" | "/ui/index.html")
            && request.method == HttpMethod::Get
        {
            return self.static_response_service().serve_ui_asset("index.html");
        }
        if request.path == "/ui/app.css" && request.method == HttpMethod::Get {
            return self.static_response_service().serve_ui_asset("app.css");
        }
        if request.path == "/ui/app.js" && request.method == HttpMethod::Get {
            return self.static_response_service().serve_ui_asset("app.js");
        }
        if request.path == "/ui/state" && request.method == HttpMethod::Get {
            if let Some(origin) = request.headers.get("origin")
                && origin != &self.config.addin_origin
            {
                return WireHttpResponse::text(403, "Forbidden origin".to_string());
            }
            return WireHttpResponse::json(
                200,
                BTreeMap::new(),
                render_ui_snapshot(ui_state, registry, &self.config),
            );
        }
        if request.path == "/ui/events" && request.method == HttpMethod::Get {
            if let Some(origin) = request.headers.get("origin")
                && origin != &self.config.addin_origin
            {
                return WireHttpResponse::text(403, "Forbidden origin".to_string());
            }
            return WireHttpResponse::binary(
                200,
                "text/event-stream; charset=utf-8",
                format!(
                    "event: snapshot\ndata: {}\n\n",
                    render_ui_snapshot(ui_state, registry, &self.config)
                )
                .into_bytes(),
                BTreeMap::from([
                    ("Cache-Control".to_string(), "no-store".to_string()),
                    ("X-Accel-Buffering".to_string(), "no".to_string()),
                ]),
            );
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
        runtime_response(decision, registry, ui_state, shared_state, &body)
    }
}

impl Default for RuntimeServer {
    fn default() -> Self {
        Self::new()
    }
}

#[derive(Debug, Clone)]
struct RuntimeSharedState {
    registry: Arc<Mutex<SessionRegistry>>,
    addin_channel: Arc<Mutex<AddinChannelServer>>,
    connection_hub: Arc<AddinConnectionHub>,
    command_router: Arc<Mutex<CommandRouter>>,
    audit_log: AuditLog,
    image_fetcher: ImageFetcher,
}

struct McpDispatchContext<'a> {
    registry: &'a SessionRegistry,
    ui_state: &'a mut UiStateStore,
    addin_channel: &'a Arc<Mutex<AddinChannelServer>>,
    connection_hub: &'a Arc<AddinConnectionHub>,
    command_router: &'a Arc<Mutex<CommandRouter>>,
    audit_log: &'a AuditLog,
    image_fetcher: &'a ImageFetcher,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HeartbeatLoopDecision {
    KeepOpen,
    Close,
}

struct McpJsonRpcRuntime;

impl McpJsonRpcRuntime {
    fn handle_body(context: &mut McpDispatchContext<'_>, body: &[u8]) -> String {
        let Ok(value) = serde_json::from_slice::<Value>(body) else {
            return json_rpc::error(&Value::Null, -32700, "Parse error");
        };
        let id = value.get("id").cloned().unwrap_or(Value::Null);
        let Some(method) = value.get("method").and_then(Value::as_str) else {
            return json_rpc::error(&id, -32600, "Invalid Request");
        };
        match method {
            "tools/list" => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": { "tools": tool_catalog_json() }
            })
            .to_string(),
            "tools/call" => Self::handle_tools_call(context, &id, &value),
            "resources/list" => Self::handle_resources_list(context.registry, &id),
            "resources/templates/list" => Self::handle_resource_templates_list(&id),
            "resources/read" => Self::handle_resources_read(context, &id, &value),
            "prompts/list" => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": { "prompts": prompt_catalog_json() }
            })
            .to_string(),
            "prompts/get" => Self::handle_prompts_get(&id, &value),
            _ => json_rpc::error(&id, -32601, &format!("Unknown method {method}")),
        }
    }

    fn handle_resources_list(registry: &SessionRegistry, id: &Value) -> String {
        let mut resources = vec![json!({
            "uri": "office://sessions",
            "name": "office.sessions",
            "title": "Office Sessions",
            "mimeType": "application/json"
        })];
        for session in registry.list_sessions() {
            if session.app == "word" {
                resources.extend(word_resource_catalog_for_session(&session.session_id));
            }
        }
        json!({ "jsonrpc": "2.0", "id": id, "result": { "resources": resources } }).to_string()
    }

    fn handle_resource_templates_list(id: &Value) -> String {
        json!({
            "jsonrpc": "2.0",
            "id": id,
            "result": { "resourceTemplates": word_resource_templates() }
        })
        .to_string()
    }

    fn handle_resources_read(
        context: &mut McpDispatchContext<'_>,
        id: &Value,
        value: &Value,
    ) -> String {
        let Some(uri) = value
            .get("params")
            .and_then(|params| params.get("uri"))
            .and_then(Value::as_str)
        else {
            return json_rpc::error(id, -32602, "resources/read requires params.uri");
        };
        match resource_request_from_uri(context.registry, uri) {
            Ok(ResourceReadRequest::Sessions) => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": {
                    "contents": [{
                        "uri": "office://sessions",
                        "mimeType": "application/json",
                        "text": json!({
                            "sessions": context
                                .registry
                                .list_sessions()
                                .iter()
                                .map(|session| SessionDescriptorView::new(session).to_json())
                                .collect::<Vec<_>>()
                        }).to_string()
                    }]
                }
            })
            .to_string(),
            Ok(ResourceReadRequest::Forwarded {
                uri,
                tool,
                arguments,
                check_capability,
            }) => {
                let result = Self::call_forwarded_tool_with_capability(
                    context,
                    value,
                    tool,
                    &arguments,
                    check_capability,
                );
                json!({
                    "jsonrpc": "2.0",
                    "id": id,
                    "result": {
                        "contents": [{
                            "uri": uri,
                            "mimeType": "application/json",
                            "text": result.get("structuredContent").cloned().unwrap_or(result).to_string()
                        }]
                    }
                })
                .to_string()
            }
            Err(message) => json_rpc::error(id, -32602, &message),
        }
    }

    fn handle_prompts_get(id: &Value, value: &Value) -> String {
        let Some(name) = value
            .get("params")
            .and_then(|params| params.get("name"))
            .and_then(Value::as_str)
        else {
            return json_rpc::error(id, -32602, "prompts/get requires params.name");
        };
        let arguments = value
            .get("params")
            .and_then(|params| params.get("arguments"));
        match prompt_messages(name, arguments) {
            Some(messages) => json!({
                "jsonrpc": "2.0",
                "id": id,
                "result": { "description": prompt_description(name), "messages": messages }
            })
            .to_string(),
            None => json_rpc::error(id, -32602, &format!("Unknown prompt {name}")),
        }
    }
    fn handle_tools_call(
        context: &mut McpDispatchContext<'_>,
        id: &Value,
        value: &Value,
    ) -> String {
        let Some(params) = value.get("params") else {
            return json_rpc::error(id, -32602, "Missing tool call params");
        };
        let Some(name) = params.get("name").and_then(Value::as_str) else {
            return json_rpc::error(id, -32602, "Missing tool name");
        };
        let arguments = params.get("arguments").unwrap_or(&Value::Null);
        let result = match name {
            "office.list_sessions" => tool_success(&json!({
                "sessions": context
                    .registry
                    .list_sessions()
                    .iter()
                    .map(|session| SessionDescriptorView::new(session).to_json())
                    .collect::<Vec<_>>()
            })),
            "office.get_session_info" => Self::get_session_info(context.registry, arguments),
            tool if WORD_V1_TOOLS.contains(&tool) => {
                Self::call_forwarded_tool(context, value, tool, arguments)
            }
            tool if ExcelToolCatalog::contains(tool) => {
                Self::call_forwarded_tool(context, value, tool, arguments)
            }
            _ => tool_failure(
                "UNKNOWN_TOOL",
                &format!("Unknown tool {name}."),
                Some(name),
                false,
            ),
        };
        json!({ "jsonrpc": "2.0", "id": id, "result": result }).to_string()
    }

    fn call_forwarded_tool(
        context: &mut McpDispatchContext<'_>,
        request_value: &Value,
        tool: &str,
        arguments: &Value,
    ) -> Value {
        Self::call_forwarded_tool_with_capability(context, request_value, tool, arguments, true)
    }

    fn call_forwarded_tool_with_capability(
        context: &mut McpDispatchContext<'_>,
        request_value: &Value,
        tool: &str,
        arguments: &Value,
        check_capability: bool,
    ) -> Value {
        let Some(session_id) = arguments.get("session_id").and_then(Value::as_str) else {
            return tool_failure(
                "INVALID_ARGUMENTS",
                "Forwarded Office tools require session_id.",
                Some(tool),
                false,
            );
        };
        let audit_started_at = SystemTime::now();
        let request_id = request_value
            .get("id")
            .and_then(Value::as_str)
            .map(str::to_string);
        let arguments = match preprocess_tool_arguments(context.image_fetcher, tool, arguments) {
            Ok(arguments) => arguments,
            Err(failure) => return tool_failure_from_command(&failure),
        };
        let arguments_json = arguments.to_string();
        let queued = {
            let mut router = context.command_router.lock().expect("command router lock");
            match router.enqueue(
                context.registry,
                context.ui_state,
                ToolCallRequest {
                    request_id,
                    command_id: None,
                    client_id: None,
                    client_name: None,
                    session_id: session_id.to_string(),
                    tool: tool.to_string(),
                    arguments_json,
                    user_intent: None,
                    timeout: None,
                    check_capability,
                },
                SystemTime::now(),
            ) {
                Ok(queued) => queued,
                Err(error) => {
                    let failure = error.as_command_failure(tool);
                    record_failure_audit(
                        context.audit_log,
                        tool,
                        session_id,
                        &failure,
                        audit_started_at,
                        SystemTime::now(),
                    );
                    return tool_failure_from_command(&failure);
                }
            }
        };
        let tool_response = Self::invoke_queued_tool(context, &queued, tool);
        Self::complete_queued_tool(context, &queued, tool_response, tool, audit_started_at)
    }

    fn invoke_queued_tool(
        context: &McpDispatchContext<'_>,
        queued: &crate::addin_mgr::QueuedCommand,
        tool: &str,
    ) -> ToolResponse {
        let payload = {
            let addin_channel = context.addin_channel.lock().expect("addin channel lock");
            json_rpc::envelope_to_text(&addin_channel.tool_invoke_payload(queued))
        };
        match context.connection_hub.invoke(
            &queued.instance_id,
            &queued.request_id,
            payload,
            queued.timeout,
        ) {
            Ok(response) => addin_response_to_tool_response(&response),
            Err(AddinConnectionHubError::NoConnection) => ToolResponse::Failure(CommandFailure {
                office_mcp_code: "SESSION_LOST".to_string(),
                message: format!("Session {} lost its add-in connection.", queued.session_id),
                tool: Some(tool.to_string()),
                retriable: true,
                partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
            }),
            Err(AddinConnectionHubError::Timeout) => {
                Self::send_timeout_cancel(context, queued);
                ToolResponse::Failure(CommandFailure {
                    office_mcp_code: "TIMEOUT".to_string(),
                    message: format!(
                        "Tool {tool} timed out after {}ms.",
                        queued.timeout.as_millis()
                    ),
                    tool: Some(tool.to_string()),
                    retriable: true,
                    partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
                })
            }
        }
    }

    fn send_timeout_cancel(
        context: &McpDispatchContext<'_>,
        queued: &crate::addin_mgr::QueuedCommand,
    ) {
        let cancel_payload = {
            let addin_channel = context.addin_channel.lock().expect("addin channel lock");
            json_rpc::envelope_to_text(&addin_channel.tool_cancel_payload(
                &crate::addin_mgr::CancelCommand {
                    request_id: queued.request_id.clone(),
                    reason: "deadline_expired".to_string(),
                },
            ))
        };
        context
            .connection_hub
            .send_to_instance(&queued.instance_id, cancel_payload);
    }

    fn complete_queued_tool(
        context: &mut McpDispatchContext<'_>,
        queued: &crate::addin_mgr::QueuedCommand,
        tool_response: ToolResponse,
        tool: &str,
        audit_started_at: SystemTime,
    ) -> Value {
        let completed = {
            let mut router = context.command_router.lock().expect("command router lock");
            router.complete(
                context.ui_state,
                &queued.session_id,
                &queued.request_id,
                tool_response,
                SystemTime::now(),
            )
        };
        record_tool_audit(
            context.audit_log,
            tool,
            &queued.session_id,
            &completed,
            audit_started_at,
            SystemTime::now(),
        );
        match completed {
            Ok(ToolResponse::Success { json }) => serde_json::from_str::<Value>(&json)
                .map_or_else(|_| tool_success(&json!(json)), |value| tool_success(&value)),
            Ok(ToolResponse::Failure(failure)) => tool_failure_from_command(&failure),
            Err(error) => tool_failure_from_command(&error.as_command_failure(tool)),
        }
    }
    fn get_session_info(registry: &SessionRegistry, arguments: &Value) -> Value {
        let Some(session_id) = arguments.get("session_id").and_then(Value::as_str) else {
            return tool_failure(
                "INVALID_ARGUMENTS",
                "office.get_session_info requires session_id.",
                Some("office.get_session_info"),
                false,
            );
        };
        let Some(info) = registry.get_session_info(session_id) else {
            return tool_failure(
                "SESSION_NOT_FOUND",
                &format!("Session {session_id} is not registered."),
                Some("office.get_session_info"),
                false,
            );
        };
        tool_success(&json!({
            "descriptor": SessionDescriptorView::new(&info.descriptor).to_json(),
            "available_tools": info.available_tools
        }))
    }
}

fn runtime_response(
    decision: McpHttpDecision,
    registry: &SessionRegistry,
    ui_state: &mut UiStateStore,
    shared_state: &Arc<RuntimeSharedState>,
    body: &[u8],
) -> WireHttpResponse {
    let request_id = json_rpc_request_id(body);
    match decision {
        McpHttpDecision::InitializeTransport { session_id } => {
            let initialize_body = json!({
                "jsonrpc": "2.0",
                "id": request_id.unwrap_or(Value::Null),
                "result": {
                    "protocolVersion": "2025-06-18",
                    "capabilities": { "tools": {}, "resources": {}, "prompts": {} },
                    "serverInfo": { "name": "office-mcp", "version": "0.1.0" }
                }
            })
            .to_string();
            WireHttpResponse::json(
                200,
                BTreeMap::from([
                    ("MCP-Session-Id".to_string(), session_id),
                    ("MCP-Protocol-Version".to_string(), "2025-06-18".to_string()),
                ]),
                initialize_body,
            )
        }
        McpHttpDecision::ForwardToTransport { .. } => {
            let mut context = McpDispatchContext {
                registry,
                ui_state,
                addin_channel: &shared_state.addin_channel,
                connection_hub: &shared_state.connection_hub,
                command_router: &shared_state.command_router,
                audit_log: &shared_state.audit_log,
                image_fetcher: &shared_state.image_fetcher,
            };
            let body = McpJsonRpcRuntime::handle_body(&mut context, body);
            WireHttpResponse::json(200, BTreeMap::new(), body)
        }
        McpHttpDecision::Reject {
            status,
            body,
            headers,
        } => {
            let mut response = WireHttpResponse::text(status, body);
            response.headers.extend(headers);
            response
        }
        McpHttpDecision::JsonRpcError {
            status,
            code,
            message,
        } => WireHttpResponse::json(
            status,
            BTreeMap::new(),
            json!({
                "jsonrpc": "2.0",
                "id": request_id.unwrap_or(Value::Null),
                "error": { "code": code, "message": message }
            })
            .to_string(),
        ),
    }
}

fn json_rpc_request_id(body: &[u8]) -> Option<Value> {
    serde_json::from_slice::<Value>(body)
        .ok()
        .and_then(|value| value.get("id").cloned())
}

fn preprocess_tool_arguments(
    image_fetcher: &ImageFetcher,
    tool: &str,
    arguments: &Value,
) -> Result<Value, CommandFailure> {
    if tool != "word.insert_image" {
        return Ok(arguments.clone());
    }
    let Some(image) = arguments.get("image") else {
        return Ok(arguments.clone());
    };
    let processed = if let Some(base64) = image.get("base64").and_then(Value::as_str) {
        image_fetcher.validate_base64(base64)
    } else if let Some(url) = image.get("url").and_then(Value::as_str) {
        image_fetcher.fetch_url(url)
    } else {
        return Ok(arguments.clone());
    };
    let fetched = processed.map_err(|error| CommandFailure {
        office_mcp_code: "IMAGE_FETCH_FAILED".to_string(),
        message: error.to_string(),
        tool: Some(tool.to_string()),
        retriable: false,
        partial_effect: Some(crate::addin_mgr::PartialEffect::None),
    })?;
    let mut updated = arguments.clone();
    if let Some(object) = updated.as_object_mut() {
        object.insert(
            "image".to_string(),
            json!({
                "base64": fetched.base64,
                "mime_type": fetched.mime_type.as_str(),
                "byte_length": fetched.byte_length
            }),
        );
    }
    Ok(updated)
}

fn record_tool_audit(
    audit_log: &AuditLog,
    tool: &str,
    session_id: &str,
    completed: &Result<ToolResponse, crate::addin_mgr::CommandRouterError>,
    started_at: SystemTime,
    completed_at: SystemTime,
) {
    let duration_ms = duration_millis(started_at, completed_at);
    let record = match completed {
        Ok(ToolResponse::Success { .. }) => {
            AuditRecord::success(SystemTime::now(), tool, Some(session_id), duration_ms)
        }
        Ok(ToolResponse::Failure(failure)) => AuditRecord::failure(
            SystemTime::now(),
            tool,
            Some(session_id),
            duration_ms,
            &failure.office_mcp_code,
            &failure.message,
        ),
        Err(error) => {
            let failure = error.as_command_failure(tool);
            AuditRecord::failure(
                SystemTime::now(),
                tool,
                Some(session_id),
                duration_ms,
                &failure.office_mcp_code,
                &failure.message,
            )
        }
    };
    if let Err(error) = audit_log.record(&record) {
        tracing::error!(%error, "failed to write audit record");
        eprintln!("office-mcp-daemon failed to write audit record: {error}");
    }
}

fn record_failure_audit(
    audit_log: &AuditLog,
    tool: &str,
    session_id: &str,
    failure: &CommandFailure,
    started_at: SystemTime,
    completed_at: SystemTime,
) {
    let record = AuditRecord::failure(
        SystemTime::now(),
        tool,
        Some(session_id),
        duration_millis(started_at, completed_at),
        &failure.office_mcp_code,
        &failure.message,
    );
    if let Err(error) = audit_log.record(&record) {
        tracing::error!(%error, "failed to write audit record");
        eprintln!("office-mcp-daemon failed to write audit record: {error}");
    }
}

fn duration_millis(started_at: SystemTime, completed_at: SystemTime) -> u64 {
    completed_at
        .duration_since(started_at)
        .unwrap_or_default()
        .as_millis()
        .try_into()
        .unwrap_or(u64::MAX)
}

fn addin_response_to_tool_response(response: &Value) -> ToolResponse {
    if let Some(error) = response.get("error") {
        return ToolResponse::Failure(CommandFailure {
            office_mcp_code: error
                .get("office_mcp_code")
                .or_else(|| error.get("code"))
                .and_then(Value::as_str)
                .unwrap_or("ADDIN_ERROR")
                .to_string(),
            message: error
                .get("message")
                .and_then(Value::as_str)
                .unwrap_or("Add-in tool call failed.")
                .to_string(),
            tool: error
                .get("tool")
                .and_then(Value::as_str)
                .map(str::to_string),
            retriable: error
                .get("retriable")
                .and_then(Value::as_bool)
                .unwrap_or(false),
            partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
        });
    }
    let result = response.get("result").cloned().unwrap_or(Value::Null);
    if result.get("ok").and_then(Value::as_bool) == Some(false) {
        let error = result.get("error").cloned().unwrap_or(Value::Null);
        return ToolResponse::Failure(CommandFailure {
            office_mcp_code: error
                .get("office_mcp_code")
                .and_then(Value::as_str)
                .unwrap_or("ADDIN_ERROR")
                .to_string(),
            message: error
                .get("message")
                .and_then(Value::as_str)
                .unwrap_or("Add-in tool call failed.")
                .to_string(),
            tool: error
                .get("tool")
                .and_then(Value::as_str)
                .map(str::to_string),
            retriable: error
                .get("retriable")
                .and_then(Value::as_bool)
                .unwrap_or(false),
            partial_effect: Some(crate::addin_mgr::PartialEffect::Unknown),
        });
    }
    let data = result.get("data").cloned().unwrap_or(result);
    ToolResponse::Success {
        json: data.to_string(),
    }
}

fn render_ui_snapshot(
    ui_state: &Arc<Mutex<UiStateStore>>,
    registry: &Arc<Mutex<SessionRegistry>>,
    config: &RuntimeServerConfig,
) -> String {
    UiSnapshotService::new().render_runtime_snapshot(
        ui_state,
        registry,
        &UiSnapshotEndpoints {
            mcp_endpoint: format!("http://{}:{}/mcp", config.mcp_host, config.mcp_port),
            addin_endpoint: format!("{}/addin", config.addin_origin),
        },
    )
}

#[cfg(test)]
#[path = "server_tests.rs"]
mod server_tests;
