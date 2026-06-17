use std::collections::{BTreeMap, BTreeSet};
use std::fmt::{Display, Formatter};
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpHttpConfig {
    pub host: String,
    pub port: u16,
    pub max_request_bytes: usize,
    pub requests_per_minute: u64,
}

impl McpHttpConfig {
    #[must_use]
    pub(crate) fn allowed_origins(&self) -> BTreeSet<String> {
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
pub(crate) struct RateLimitWindow {
    pub(crate) window_started: SystemTime,
    pub(crate) count: u64,
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
    pub(crate) fn reject(status: u16, body: &str) -> Self {
        Self::Reject {
            status,
            body: body.to_string(),
            headers: BTreeMap::new(),
        }
    }

    pub(crate) fn reject_with_header(status: u16, body: &str, name: &str, value: &str) -> Self {
        Self::Reject {
            status,
            body: body.to_string(),
            headers: BTreeMap::from([(name.to_string(), value.to_string())]),
        }
    }

    pub(crate) fn json_rpc_error(status: u16, code: i64, message: &str) -> Self {
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
#[path = "http_frontend_model_tests.rs"]
mod http_frontend_model_tests;
