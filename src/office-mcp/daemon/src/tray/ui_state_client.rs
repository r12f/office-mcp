use crate::tray::TrayPlatformError;
use native_tls::TlsConnector;
use serde_json::Value;
use std::io::{Read, Write};
use std::net::TcpStream;
use std::time::Duration;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TrayUiStateClient {
    url: String,
}

impl TrayUiStateClient {
    #[must_use]
    pub(crate) fn new(url: impl Into<String>) -> Self {
        Self { url: url.into() }
    }

    /// Fetches the daemon UI state from a loopback HTTP(S) runtime URL.
    ///
    /// # Errors
    ///
    /// Returns an error when the runtime URL is not loopback, the daemon cannot
    /// be reached, or the response is not valid JSON.
    pub(crate) fn fetch_json(&self) -> Result<Value, TrayPlatformError> {
        let target = LoopbackHttpTarget::parse(&self.url)?;
        let body = if target.scheme == "https" {
            fetch_https(&target)?
        } else {
            fetch_http(&target)?
        };
        serde_json::from_str(&body).map_err(|error| TrayPlatformError::new(error.to_string()))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct LoopbackHttpTarget {
    scheme: String,
    host: String,
    port: u16,
    path: String,
}

impl LoopbackHttpTarget {
    fn parse(url: &str) -> Result<Self, TrayPlatformError> {
        let (scheme, rest) = url
            .split_once("://")
            .ok_or_else(|| TrayPlatformError::new("UI state URL must include a scheme"))?;
        if scheme != "https" && scheme != "http" {
            return Err(TrayPlatformError::new("UI state URL must be HTTP or HTTPS"));
        }
        let (authority, path) = rest
            .split_once('/')
            .map_or((rest, "/"), |(authority, path)| (authority, path));
        let (host, port) = parse_authority(scheme, authority)?;
        if !is_loopback_host(&host) {
            return Err(TrayPlatformError::new(
                "UI state URL must target a loopback host",
            ));
        }
        Ok(Self {
            scheme: scheme.to_string(),
            host,
            port,
            path: format!("/{path}"),
        })
    }

    fn host_header(&self) -> String {
        format!("{}:{}", self.host, self.port)
    }
}

fn parse_authority(scheme: &str, authority: &str) -> Result<(String, u16), TrayPlatformError> {
    let default_port = if scheme == "https" { 443 } else { 80 };
    if authority.is_empty() {
        return Err(TrayPlatformError::new("UI state URL must include a host"));
    }
    if let Some((host, port)) = authority.rsplit_once(':') {
        let port = port
            .parse::<u16>()
            .map_err(|error| TrayPlatformError::new(error.to_string()))?;
        return Ok((host.to_string(), port));
    }
    Ok((authority.to_string(), default_port))
}

fn is_loopback_host(host: &str) -> bool {
    matches!(host, "localhost" | "127.0.0.1" | "[::1]" | "::1")
}

fn fetch_https(target: &LoopbackHttpTarget) -> Result<String, TrayPlatformError> {
    let stream = connect_tcp(target)?;
    let connector = TlsConnector::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    let mut stream = connector
        .connect(target.host.as_str(), stream)
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    fetch_over_stream(&mut stream, target)
}

fn fetch_http(target: &LoopbackHttpTarget) -> Result<String, TrayPlatformError> {
    let mut stream = connect_tcp(target)?;
    fetch_over_stream(&mut stream, target)
}

fn connect_tcp(target: &LoopbackHttpTarget) -> Result<TcpStream, TrayPlatformError> {
    let stream = TcpStream::connect((target.host.as_str(), target.port))
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    stream
        .set_read_timeout(Some(Duration::from_secs(3)))
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    stream
        .set_write_timeout(Some(Duration::from_secs(3)))
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    Ok(stream)
}

fn fetch_over_stream(
    stream: &mut (impl Read + Write),
    target: &LoopbackHttpTarget,
) -> Result<String, TrayPlatformError> {
    let request = format!(
        "GET {} HTTP/1.1\r\nHost: {}\r\nConnection: close\r\nAccept: application/json\r\n\r\n",
        target.path,
        target.host_header()
    );
    stream
        .write_all(request.as_bytes())
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    stream
        .flush()
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    let mut response = String::new();
    stream
        .read_to_string(&mut response)
        .map_err(|error| TrayPlatformError::new(error.to_string()))?;
    let (head, body) = response
        .split_once("\r\n\r\n")
        .ok_or_else(|| TrayPlatformError::new("UI state response is missing headers"))?;
    if !head.starts_with("HTTP/1.1 200") && !head.starts_with("HTTP/1.0 200") {
        return Err(TrayPlatformError::new(
            "UI state endpoint did not return 200",
        ));
    }
    Ok(body.to_string())
}

#[cfg(test)]
#[path = "ui_state_client_tests.rs"]
mod ui_state_client_tests;
