use crate::addin_mgr::SessionRegistry;
use crate::api::{UiStateOptions, UiStateStore};
use crate::common::DaemonConfig;
use crate::mcp::McpHttpFrontend;
use crate::runtime::mcp_response::RuntimeSharedState;
use crate::runtime::runtime_shared_state_factory::RuntimeSharedStateFactory;
use crate::runtime::server_connection::RuntimeConnectionHandler;
use crate::runtime::session_cleanup_service::SessionCleanupService;
pub use crate::runtime::{RuntimeServerConfig, RuntimeServerError, default_pfx_path};
use crate::ui::UiRuntimeFile;
use native_tls::TlsAcceptor;
#[cfg(test)]
use native_tls::TlsStream;
use std::net::{TcpListener, TcpStream};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::SystemTime;

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

    pub(crate) fn ui_state_options(&self) -> UiStateOptions {
        UiStateOptions {
            mcp_endpoint: format!(
                "http://{}:{}/mcp",
                self.config.mcp_host, self.config.mcp_port
            ),
            addin_endpoint: format!("{}/addin", self.config.addin_origin),
            config_path: self.config.config_path.clone(),
            log_path: self.config.log_path.clone(),
            tool_access_policy: self.config.tool_access_policy.clone(),
            now: SystemTime::now(),
            ..UiStateOptions::default()
        }
    }

    fn ui_state_store(&self) -> UiStateStore {
        UiStateStore::with_options(self.ui_state_options())
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
        let shared_state = RuntimeSharedStateFactory::with_registry(&self.config, seed_registry);
        SessionCleanupService::for_session_grace(self.config.session_grace)
            .spawn(Arc::clone(&shared_state));
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
        let shared_state = RuntimeSharedStateFactory::with_registry(&self.config, registry.clone());
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
        shared_state.prune_stale_sessions(SystemTime::now());
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
                server.connection_handler().handle_addin_tls_stream(
                    &mut stream,
                    &ui_state,
                    &shared_state,
                )
            })();
            if let Err(error) = result {
                tracing::warn!(%error, "ignored malformed add-in TLS client connection");
                eprintln!("office-mcp-daemon ignored malformed add-in client connection: {error}");
            }
        });
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
        self.connection_handler().handle_mcp_stream(
            stream,
            frontend,
            ui_state,
            registry,
            shared_state,
            remote_addr,
        )
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
        self.connection_handler()
            .handle_addin_stream(stream, ui_state)
    }

    #[cfg(test)]
    fn handle_addin_tls_stream(
        &self,
        stream: &mut TlsStream<TcpStream>,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        self.connection_handler()
            .handle_addin_tls_stream(stream, ui_state, shared_state)
    }

    fn connection_handler(&self) -> RuntimeConnectionHandler<'_> {
        RuntimeConnectionHandler::new(&self.config)
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
