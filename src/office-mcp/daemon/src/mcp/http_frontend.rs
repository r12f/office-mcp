use crate::api::{RegisterClientInput, UiClientTransport, UiStateStore};
use crate::mcp::{
    HttpMethod, McpClientSession, McpHttpConfig, McpHttpDecision, McpHttpError, McpHttpRequest,
    McpHttpRequestClass, RateLimitWindow,
};
use std::collections::BTreeMap;
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpHttpFrontend {
    config: McpHttpConfig,
    sessions: BTreeMap<String, McpClientSession>,
    rate_limits: BTreeMap<String, RateLimitWindow>,
    next_session_id: u64,
}

impl McpHttpFrontend {
    #[must_use]
    pub fn new() -> Self {
        Self::with_config(McpHttpConfig::default())
    }

    #[must_use]
    pub fn with_config(config: McpHttpConfig) -> Self {
        Self {
            config,
            sessions: BTreeMap::new(),
            rate_limits: BTreeMap::new(),
            next_session_id: 1,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        "owns Streamable HTTP request validation client tracking and MCP errors"
    }

    pub fn handle_request(
        &mut self,
        ui_state: &mut UiStateStore,
        request: &McpHttpRequest,
        now: SystemTime,
    ) -> McpHttpDecision {
        if let Err(error) = self.validate_origin(request.header("origin")) {
            tracing::warn!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                origin = ?request.header("origin"),
                error = %error,
                "rejected MCP HTTP request origin"
            );
            return error.into();
        }
        if let Err(error) = self.validate_request_size(request.body_bytes) {
            tracing::warn!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                body_bytes = request.body_bytes,
                error = %error,
                "rejected oversized MCP HTTP request"
            );
            return error.into();
        }
        if let Err(error) = self.check_rate_limit(request.rate_limit_key(), request.class, now) {
            tracing::warn!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                error = %error,
                "rate limited MCP HTTP request"
            );
            return error.into();
        }
        match request.method {
            HttpMethod::Post => self.handle_post(ui_state, request, now),
            HttpMethod::Get | HttpMethod::Delete => {
                self.handle_session_request(ui_state, request, now)
            }
            _ => {
                tracing::warn!(
                    component = "mcp_http_frontend",
                    client = %request.client_key(),
                    method = ?request.method,
                    "rejected unsupported MCP HTTP method"
                );
                McpHttpDecision::reject(405, "Method not allowed")
            }
        }
    }

    pub fn close_session(&mut self, ui_state: &mut UiStateStore, session_id: &str) -> bool {
        let removed = self.sessions.remove(session_id).is_some();
        if removed {
            ui_state.unregister_client(session_id);
            tracing::info!(
                component = "mcp_http_frontend",
                session_id = %session_id,
                active_sessions = self.sessions.len(),
                "closed MCP HTTP session"
            );
        }
        removed
    }

    #[must_use]
    pub fn active_session_count(&self) -> usize {
        self.sessions.len()
    }

    fn handle_post(
        &mut self,
        ui_state: &mut UiStateStore,
        request: &McpHttpRequest,
        now: SystemTime,
    ) -> McpHttpDecision {
        if let Some(session_id) = request.header("mcp-session-id") {
            if !self.sessions.contains_key(session_id) {
                tracing::warn!(
                    component = "mcp_http_frontend",
                    client = %request.client_key(),
                    session_id = %session_id,
                    "rejected unknown MCP HTTP session"
                );
                return McpHttpDecision::json_rpc_error(404, -32000, "Unknown MCP session ID.");
            }
            ui_state.touch_client(session_id, now);
            tracing::debug!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                session_id = %session_id,
                "forwarding MCP HTTP POST to transport"
            );
            return McpHttpDecision::ForwardToTransport {
                session_id: Some(session_id.to_string()),
            };
        }
        if !request.is_initialize {
            tracing::warn!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                "rejected MCP HTTP POST without session"
            );
            return McpHttpDecision::json_rpc_error(
                400,
                -32000,
                "Bad Request: missing MCP session ID.",
            );
        }
        let session_id = self.next_client_session_id();
        let client_name = request.client_name();
        self.sessions.insert(
            session_id.clone(),
            McpClientSession {
                session_id: session_id.clone(),
                source: request.client_key(),
                initialized_at: now,
                last_activity_at: now,
            },
        );
        ui_state.register_client(RegisterClientInput {
            client_id: Some(session_id.clone()),
            transport: UiClientTransport::Http,
            name: client_name,
        });
        tracing::info!(
            component = "mcp_http_frontend",
            client = %request.client_key(),
            session_id = %session_id,
            active_sessions = self.sessions.len(),
            "initialized MCP HTTP session"
        );
        McpHttpDecision::InitializeTransport { session_id }
    }

    fn handle_session_request(
        &mut self,
        ui_state: &mut UiStateStore,
        request: &McpHttpRequest,
        now: SystemTime,
    ) -> McpHttpDecision {
        let Some(session_id) = request.header("mcp-session-id") else {
            tracing::warn!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                method = ?request.method,
                "rejected MCP HTTP session request without session"
            );
            return McpHttpDecision::reject(400, "Invalid or missing MCP session ID");
        };
        let Some(session) = self.sessions.get_mut(session_id) else {
            tracing::warn!(
                component = "mcp_http_frontend",
                client = %request.client_key(),
                method = ?request.method,
                session_id = %session_id,
                "rejected unknown MCP HTTP session request"
            );
            return McpHttpDecision::reject(400, "Invalid or missing MCP session ID");
        };
        session.last_activity_at = now;
        ui_state.touch_client(session_id, now);
        tracing::debug!(
            component = "mcp_http_frontend",
            client = %request.client_key(),
            method = ?request.method,
            session_id = %session_id,
            "forwarding MCP HTTP session request to transport"
        );
        McpHttpDecision::ForwardToTransport {
            session_id: Some(session_id.to_string()),
        }
    }

    fn validate_origin(&self, origin: Option<&str>) -> Result<(), McpHttpError> {
        let Some(origin) = origin else {
            return Ok(());
        };
        if self.config.allowed_origins().contains(origin) {
            Ok(())
        } else {
            Err(McpHttpError::ForbiddenOrigin)
        }
    }

    fn validate_request_size(&self, body_bytes: usize) -> Result<(), McpHttpError> {
        if body_bytes <= self.config.max_request_bytes {
            Ok(())
        } else {
            Err(McpHttpError::RequestTooLarge {
                max_request_bytes: self.config.max_request_bytes,
            })
        }
    }

    fn check_rate_limit(
        &mut self,
        key: String,
        class: McpHttpRequestClass,
        now: SystemTime,
    ) -> Result<(), McpHttpError> {
        let window = self
            .rate_limits
            .entry(key)
            .or_insert_with(|| RateLimitWindow {
                window_started: now,
                count: 0,
            });
        if now
            .duration_since(window.window_started)
            .unwrap_or_default()
            >= Duration::from_mins(1)
        {
            window.window_started = now;
            window.count = 0;
        }
        window.count += 1;
        if window.count <= self.config.limit_for(class) {
            Ok(())
        } else {
            Err(McpHttpError::RateLimited)
        }
    }

    fn next_client_session_id(&mut self) -> String {
        let value = format!("mcp-session-{}", self.next_session_id);
        self.next_session_id += 1;
        value
    }
}

impl Default for McpHttpFrontend {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
#[path = "http_frontend_tests.rs"]
mod http_frontend_tests;
