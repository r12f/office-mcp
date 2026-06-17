use crate::addin_mgr::SessionRegistry;
use crate::api::UiStateStore;
use crate::mcp::McpHttpFrontend;
use crate::runtime::addin_http::AddinHttpService;
use crate::runtime::http_wire::WireHttpRequest;
use crate::runtime::mcp_response::RuntimeSharedState;
use crate::runtime::runtime_request_router::RuntimeRequestRouter;
use crate::runtime::websocket_session::RuntimeWebSocketSession;
use crate::runtime::{RuntimeServerConfig, RuntimeServerError};
use native_tls::TlsStream;
use std::io::Write;
use std::net::TcpStream;
use std::sync::{Arc, Mutex};

pub(crate) struct RuntimeConnectionHandler<'a> {
    config: &'a RuntimeServerConfig,
}

impl<'a> RuntimeConnectionHandler<'a> {
    #[must_use]
    pub(crate) const fn new(config: &'a RuntimeServerConfig) -> Self {
        Self { config }
    }

    /// Handles a single add-in/static/UI HTTP request on a plain stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the request cannot be read or the response cannot
    /// be written.
    pub(crate) fn handle_addin_stream(
        &self,
        stream: &mut TcpStream,
        ui_state: &UiStateStore,
    ) -> Result<(), RuntimeServerError> {
        let request = WireHttpRequest::read_from(stream, self.config.max_request_bytes)?;
        let registry = Arc::new(Mutex::new(SessionRegistry::new()));
        let ui_state = Arc::new(Mutex::new(ui_state.clone()));
        let response = self
            .addin_http_service()
            .route_request(&ui_state, &registry, &request);
        stream.write_all(&response.to_bytes())?;
        stream.flush()?;
        Ok(())
    }

    pub(crate) fn handle_addin_tls_stream(
        &self,
        stream: &mut TlsStream<TcpStream>,
        ui_state: &Arc<Mutex<UiStateStore>>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        let request = WireHttpRequest::read_from(stream, self.config.max_request_bytes)?;
        let addin_http = self.addin_http_service();
        let websocket_upgrade = addin_http.is_valid_websocket_upgrade(&request);
        let response = addin_http.route_request(ui_state, &shared_state.registry, &request);
        stream.write_all(&response.to_bytes())?;
        stream.flush()?;
        if websocket_upgrade && response.status == 101 {
            self.handle_websocket_messages(stream, shared_state)?;
        }
        Ok(())
    }

    /// Handles a single MCP HTTP request/response exchange on a stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the request cannot be read or the response cannot
    /// be written.
    pub(crate) fn handle_mcp_stream(
        &self,
        stream: &mut TcpStream,
        frontend: &mut McpHttpFrontend,
        ui_state: &mut UiStateStore,
        registry: &SessionRegistry,
        shared_state: &Arc<RuntimeSharedState>,
        remote_addr: Option<String>,
    ) -> Result<(), RuntimeServerError> {
        let request = WireHttpRequest::read_from(stream, self.config.max_request_bytes)?;
        let response = RuntimeRequestRouter::route(
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

    fn handle_websocket_messages(
        &self,
        stream: &mut TlsStream<TcpStream>,
        shared_state: &Arc<RuntimeSharedState>,
    ) -> Result<(), RuntimeServerError> {
        RuntimeWebSocketSession::from_config(self.config).handle(stream, shared_state)
    }

    fn addin_http_service(&self) -> AddinHttpService {
        AddinHttpService::from_config(self.config)
    }
}

#[cfg(test)]
#[path = "server_connection_tests.rs"]
mod server_connection_tests;
