use crate::common::DaemonConfig;
use std::collections::BTreeMap;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::fmt::{Display, Formatter};
use std::io::{BufRead, BufReader, Read, Write};
use std::net::TcpStream;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StdioBridge {
    endpoint: McpEndpoint,
    session_id: Option<String>,
}

impl StdioBridge {
    #[must_use]
    pub fn from_config(config: &DaemonConfig) -> Self {
        Self::new(McpEndpoint::from_config(config))
    }

    #[must_use]
    pub const fn new(endpoint: McpEndpoint) -> Self {
        Self {
            endpoint,
            session_id: None,
        }
    }

    /// Bridges newline-delimited JSON-RPC messages from stdin to the daemon.
    ///
    /// # Errors
    ///
    /// Returns an error when stdin/stdout fails or the daemon cannot be reached.
    pub fn run(&mut self) -> Result<(), StdioBridgeError> {
        let stdin = std::io::stdin();
        let stdout = std::io::stdout();
        self.run_with(stdin.lock(), stdout.lock())
    }

    /// Bridges newline-delimited JSON-RPC messages from `input` to `output`.
    ///
    /// # Errors
    ///
    /// Returns an error when input/output fails or the daemon cannot be reached.
    pub fn run_with(
        &mut self,
        input: impl Read,
        mut output: impl Write,
    ) -> Result<(), StdioBridgeError> {
        let mut input = BufReader::new(input);
        while let Some(message) = read_stdio_message(&mut input)? {
            let response = self.forward_json_rpc(&message.body)?;
            write_stdio_response(&mut output, &response.body, message.framing)?;
        }
        Ok(())
    }

    fn forward_json_rpc(&mut self, message: &str) -> Result<McpHttpResponse, StdioBridgeError> {
        let request = McpHttpPost::new(&self.endpoint, self.session_id.as_deref(), message);
        let response = request.send()?;
        if response.status >= 400 {
            return Err(StdioBridgeError::Daemon(format!(
                "daemon returned HTTP {}: {}",
                response.status, response.body
            )));
        }
        if self.session_id.is_none() && is_initialize(message) {
            self.session_id = response.header("mcp-session-id").map(ToString::to_string);
        }
        Ok(response)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpEndpoint {
    pub host: String,
    pub port: u16,
    pub path: String,
}

impl McpEndpoint {
    #[must_use]
    pub fn from_config(config: &DaemonConfig) -> Self {
        Self {
            host: config.mcp.host.clone(),
            port: u16::try_from(config.mcp.port).unwrap_or(8800),
            path: "/mcp".to_string(),
        }
    }

    #[must_use]
    pub fn url(&self) -> String {
        format!("http://{}:{}{}", self.host, self.port, self.path)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct McpHttpPost<'a> {
    endpoint: &'a McpEndpoint,
    session_id: Option<&'a str>,
    body: &'a str,
}

impl<'a> McpHttpPost<'a> {
    const fn new(endpoint: &'a McpEndpoint, session_id: Option<&'a str>, body: &'a str) -> Self {
        Self {
            endpoint,
            session_id,
            body,
        }
    }

    fn send(&self) -> Result<McpHttpResponse, StdioBridgeError> {
        let mut stream = TcpStream::connect((&self.endpoint.host[..], self.endpoint.port))
            .map_err(StdioBridgeError::Io)?;
        let mut request = format!(
            concat!(
                "POST {} HTTP/1.1\r\n",
                "Host: {}:{}\r\n",
                "Content-Type: application/json\r\n",
                "Accept: application/json\r\n",
                "X-Office-Mcp-Client: office-mcp-stdio-bridge/0.1.0\r\n",
                "Content-Length: {}\r\n"
            ),
            self.endpoint.path,
            self.endpoint.host,
            self.endpoint.port,
            self.body.len()
        );
        if let Some(session_id) = self.session_id {
            let _ = write!(request, "MCP-Session-Id: {session_id}\r\n");
        }
        request.push_str("\r\n");
        stream
            .write_all(request.as_bytes())
            .and_then(|()| stream.write_all(self.body.as_bytes()))
            .map_err(StdioBridgeError::Io)?;
        McpHttpResponse::read(stream)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct McpHttpResponse {
    status: u16,
    headers: BTreeMap<String, String>,
    body: String,
}

impl McpHttpResponse {
    fn read(mut stream: TcpStream) -> Result<Self, StdioBridgeError> {
        let bytes = read_http_response_bytes(&mut stream)?;
        let Some(head_end) = find_header_end(&bytes) else {
            return Err(StdioBridgeError::Protocol(
                "daemon returned malformed HTTP response".to_string(),
            ));
        };
        let head = std::str::from_utf8(&bytes[..head_end]).map_err(|_| {
            StdioBridgeError::Protocol("daemon returned non-UTF-8 headers".to_string())
        })?;
        let mut lines = head.split("\r\n");
        let status = parse_status(lines.next().unwrap_or_default())?;
        let mut headers = BTreeMap::new();
        for line in lines.filter(|line| !line.is_empty()) {
            if let Some((name, value)) = line.split_once(':') {
                headers.insert(name.trim().to_ascii_lowercase(), value.trim().to_string());
            }
        }
        let body = String::from_utf8(bytes[head_end + 4..].to_vec()).map_err(|_| {
            StdioBridgeError::Protocol("daemon returned non-UTF-8 JSON body".to_string())
        })?;
        Ok(Self {
            status,
            headers,
            body,
        })
    }

    fn header(&self, name: &str) -> Option<&str> {
        self.headers
            .get(&name.to_ascii_lowercase())
            .map(String::as_str)
    }
}

#[derive(Debug)]
pub enum StdioBridgeError {
    Io(std::io::Error),
    Daemon(String),
    Protocol(String),
}

impl Display for StdioBridgeError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "stdio bridge I/O error: {error}"),
            Self::Daemon(message) | Self::Protocol(message) => formatter.write_str(message),
        }
    }
}

impl Error for StdioBridgeError {}

fn is_initialize(message: &str) -> bool {
    message.contains("\"method\"") && message.contains("\"initialize\"")
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StdioFraming {
    ContentLength,
    JsonLine,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StdioMessage {
    body: String,
    framing: StdioFraming,
}

fn read_stdio_message(reader: &mut impl BufRead) -> Result<Option<StdioMessage>, StdioBridgeError> {
    loop {
        let buffer = reader.fill_buf().map_err(StdioBridgeError::Io)?;
        if buffer.is_empty() {
            return Ok(None);
        }
        if buffer.starts_with(b"\r\n") {
            reader.consume(2);
            continue;
        }
        if buffer.starts_with(b"\n") {
            reader.consume(1);
            continue;
        }
        if starts_with_ignore_ascii_case(buffer, b"content-length:") {
            return read_content_length_message(reader).map(Some);
        }
        return read_json_line_message(reader).map(Some);
    }
}

fn read_content_length_message(
    reader: &mut impl BufRead,
) -> Result<StdioMessage, StdioBridgeError> {
    let mut content_length = None;
    let mut line = String::new();
    loop {
        line.clear();
        let bytes_read = reader.read_line(&mut line).map_err(StdioBridgeError::Io)?;
        if bytes_read == 0 {
            return Err(StdioBridgeError::Protocol(
                "incomplete MCP stdio frame headers".to_string(),
            ));
        }
        let header = line.trim_end_matches(['\r', '\n']);
        if header.is_empty() {
            break;
        }
        if let Some((name, value)) = header.split_once(':')
            && name.eq_ignore_ascii_case("content-length")
        {
            content_length = Some(value.trim().parse::<usize>().map_err(|_| {
                StdioBridgeError::Protocol("invalid MCP stdio Content-Length".to_string())
            })?);
        }
    }
    let Some(content_length) = content_length else {
        return Err(StdioBridgeError::Protocol(
            "missing MCP stdio Content-Length".to_string(),
        ));
    };
    let mut body = vec![0_u8; content_length];
    reader.read_exact(&mut body).map_err(StdioBridgeError::Io)?;
    let body = String::from_utf8(body).map_err(|_| {
        StdioBridgeError::Protocol("MCP stdio frame body is not UTF-8 JSON".to_string())
    })?;
    Ok(StdioMessage {
        body,
        framing: StdioFraming::ContentLength,
    })
}

fn read_json_line_message(reader: &mut impl BufRead) -> Result<StdioMessage, StdioBridgeError> {
    let mut line = String::new();
    reader.read_line(&mut line).map_err(StdioBridgeError::Io)?;
    Ok(StdioMessage {
        body: line.trim().to_string(),
        framing: StdioFraming::JsonLine,
    })
}

fn write_stdio_response(
    output: &mut impl Write,
    body: &str,
    framing: StdioFraming,
) -> Result<(), StdioBridgeError> {
    match framing {
        StdioFraming::ContentLength => {
            write!(output, "Content-Length: {}\r\n\r\n{}", body.len(), body)
        }
        StdioFraming::JsonLine => writeln!(output, "{body}"),
    }
    .and_then(|()| output.flush())
    .map_err(StdioBridgeError::Io)
}

fn starts_with_ignore_ascii_case(value: &[u8], prefix: &[u8]) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix))
}

fn parse_status(status_line: &str) -> Result<u16, StdioBridgeError> {
    status_line
        .split_whitespace()
        .nth(1)
        .and_then(|value| value.parse::<u16>().ok())
        .ok_or_else(|| StdioBridgeError::Protocol("missing HTTP status code".to_string()))
}

fn read_http_response_bytes(stream: &mut TcpStream) -> Result<Vec<u8>, StdioBridgeError> {
    let mut bytes = Vec::new();
    let mut buffer = [0_u8; 1024];
    loop {
        let size = stream.read(&mut buffer).map_err(StdioBridgeError::Io)?;
        if size == 0 {
            return Ok(bytes);
        }
        bytes.extend_from_slice(&buffer[..size]);
        let Some(header_end) = find_header_end(&bytes) else {
            continue;
        };
        let head = std::str::from_utf8(&bytes[..header_end]).map_err(|_| {
            StdioBridgeError::Protocol("daemon returned non-UTF-8 headers".to_string())
        })?;
        let Some(content_length) = content_length(head) else {
            return Ok(bytes);
        };
        if bytes.len() >= header_end + 4 + content_length {
            return Ok(bytes);
        }
    }
}

fn content_length(head: &str) -> Option<usize> {
    head.lines().find_map(|line| {
        let (name, value) = line.split_once(':')?;
        name.eq_ignore_ascii_case("content-length")
            .then(|| value.trim().parse::<usize>().ok())?
    })
}

fn find_header_end(bytes: &[u8]) -> Option<usize> {
    bytes.windows(4).position(|window| window == b"\r\n\r\n")
}

#[cfg(test)]
#[path = "stdio_bridge_tests.rs"]
mod stdio_bridge_tests;
