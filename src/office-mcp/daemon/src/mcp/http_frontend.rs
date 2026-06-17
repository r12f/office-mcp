use crate::api::{RegisterClientInput, UiClientTransport, UiStateStore};
use std::collections::{BTreeMap, BTreeSet};
use std::fmt::{Display, Formatter};
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
            return error.into();
        }
        if let Err(error) = self.validate_request_size(request.body_bytes) {
            return error.into();
        }
        if let Err(error) = self.check_rate_limit(request.client_key(), now) {
            return error.into();
        }
        match request.method {
            HttpMethod::Post => self.handle_post(ui_state, request, now),
            HttpMethod::Get | HttpMethod::Delete => {
                self.handle_session_request(ui_state, request, now)
            }
            _ => McpHttpDecision::reject(405, "Method not allowed"),
        }
    }

    pub fn close_session(&mut self, ui_state: &mut UiStateStore, session_id: &str) -> bool {
        let removed = self.sessions.remove(session_id).is_some();
        if removed {
            ui_state.unregister_client(session_id);
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
                return McpHttpDecision::json_rpc_error(404, -32000, "Unknown MCP session ID.");
            }
            ui_state.touch_client(session_id, now);
            return McpHttpDecision::ForwardToTransport {
                session_id: Some(session_id.to_string()),
            };
        }
        if !request.is_initialize {
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
        McpHttpDecision::InitializeTransport { session_id }
    }

    fn handle_session_request(
        &mut self,
        ui_state: &mut UiStateStore,
        request: &McpHttpRequest,
        now: SystemTime,
    ) -> McpHttpDecision {
        let Some(session_id) = request.header("mcp-session-id") else {
            return McpHttpDecision::reject(400, "Invalid or missing MCP session ID");
        };
        let Some(session) = self.sessions.get_mut(session_id) else {
            return McpHttpDecision::reject(400, "Invalid or missing MCP session ID");
        };
        session.last_activity_at = now;
        ui_state.touch_client(session_id, now);
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

    fn check_rate_limit(&mut self, key: String, now: SystemTime) -> Result<(), McpHttpError> {
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
        if window.count <= self.config.requests_per_minute {
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpHttpConfig {
    pub host: String,
    pub port: u16,
    pub max_request_bytes: usize,
    pub requests_per_minute: u64,
}

impl McpHttpConfig {
    fn allowed_origins(&self) -> BTreeSet<String> {
        BTreeSet::from([
            format!("http://{}:{}", self.host, self.port),
            format!("http://localhost:{}", self.port),
            format!("http://127.0.0.1:{}", self.port),
        ])
    }
}

impl Default for McpHttpConfig {
    fn default() -> Self {
        Self {
            host: "127.0.0.1".to_string(),
            port: 8800,
            max_request_bytes: 16 * 1024 * 1024,
            requests_per_minute: 120,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpHttpRequest {
    pub method: HttpMethod,
    pub headers: BTreeMap<String, String>,
    pub remote_addr: Option<String>,
    pub body_bytes: usize,
    pub is_initialize: bool,
}

impl McpHttpRequest {
    #[must_use]
    pub fn header(&self, name: &str) -> Option<&str> {
        self.headers
            .get(&name.to_ascii_lowercase())
            .map(String::as_str)
    }

    #[must_use]
    pub fn client_key(&self) -> String {
        self.header("x-forwarded-for")
            .and_then(|value| value.split(',').next())
            .map(str::trim)
            .filter(|value| !value.is_empty())
            .map(str::to_string)
            .or_else(|| self.remote_addr.clone())
            .unwrap_or_else(|| "unknown".to_string())
    }

    #[must_use]
    pub fn client_name(&self) -> Option<String> {
        self.header("x-office-mcp-client")
            .or_else(|| self.header("user-agent"))
            .map(str::to_string)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HttpMethod {
    Get,
    Post,
    Delete,
    Put,
    Patch,
    Options,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpClientSession {
    pub session_id: String,
    pub source: String,
    pub initialized_at: SystemTime,
    pub last_activity_at: SystemTime,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RateLimitWindow {
    window_started: SystemTime,
    count: u64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum McpHttpDecision {
    InitializeTransport {
        session_id: String,
    },
    ForwardToTransport {
        session_id: Option<String>,
    },
    Reject {
        status: u16,
        body: String,
        headers: BTreeMap<String, String>,
    },
    JsonRpcError {
        status: u16,
        code: i64,
        message: String,
    },
}

impl McpHttpDecision {
    fn reject(status: u16, body: &str) -> Self {
        Self::Reject {
            status,
            body: body.to_string(),
            headers: BTreeMap::new(),
        }
    }

    fn reject_with_header(status: u16, body: &str, name: &str, value: &str) -> Self {
        Self::Reject {
            status,
            body: body.to_string(),
            headers: BTreeMap::from([(name.to_string(), value.to_string())]),
        }
    }

    fn json_rpc_error(status: u16, code: i64, message: &str) -> Self {
        Self::JsonRpcError {
            status,
            code,
            message: message.to_string(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum McpHttpError {
    ForbiddenOrigin,
    RateLimited,
    RequestTooLarge { max_request_bytes: usize },
}

impl From<McpHttpError> for McpHttpDecision {
    fn from(error: McpHttpError) -> Self {
        match error {
            McpHttpError::ForbiddenOrigin => Self::reject(403, "Forbidden origin"),
            McpHttpError::RateLimited => {
                Self::reject_with_header(429, "Rate limit exceeded", "Retry-After", "60")
            }
            McpHttpError::RequestTooLarge { max_request_bytes } => Self::json_rpc_error(
                413,
                -32000,
                &format!("Request body exceeds {max_request_bytes} bytes."),
            ),
        }
    }
}

impl Display for McpHttpError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::ForbiddenOrigin => formatter.write_str("Forbidden origin"),
            Self::RateLimited => formatter.write_str("Rate limit exceeded"),
            Self::RequestTooLarge { max_request_bytes } => {
                write!(formatter, "Request body exceeds {max_request_bytes} bytes.")
            }
        }
    }
}

impl std::error::Error for McpHttpError {}

#[cfg(test)]
#[path = "http_frontend_tests.rs"]
mod http_frontend_tests;
