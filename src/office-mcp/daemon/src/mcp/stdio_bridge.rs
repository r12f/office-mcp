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
        for line in BufReader::new(input).lines() {
            let line = line.map_err(StdioBridgeError::Io)?;
            let message = line.trim();
            if message.is_empty() {
                continue;
            }
            let response = self.forward_json_rpc(message)?;
            output
                .write_all(response.body.as_bytes())
                .and_then(|()| output.write_all(b"\n"))
                .and_then(|()| output.flush())
                .map_err(StdioBridgeError::Io)?;
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
mod tests {
    use super::{McpEndpoint, StdioBridge};
    use std::io::{Read, Write};
    use std::net::{TcpListener, TcpStream};
    use std::sync::mpsc;
    use std::thread;

    #[test]
    fn bridge_forwards_initialize_and_reuses_session_id() {
        let listener = TcpListener::bind("127.0.0.1:0").expect("listener");
        let port = listener.local_addr().expect("addr").port();
        let (sender, receiver) = mpsc::channel();
        let server = thread::spawn(move || {
            let mut first_stream = listener.accept().expect("first accept").0;
            let first = read_request(&mut first_stream);
            assert!(first.contains("initialize"));
            respond(
                &mut first_stream,
                "MCP-Session-Id: mcp-session-1\r\n",
                r#"{"jsonrpc":"2.0","id":1,"result":{}}"#,
            );

            let mut second_stream = listener.accept().expect("second accept").0;
            let second = read_request(&mut second_stream);
            sender.send(second).expect("send second request");
            respond(
                &mut second_stream,
                "",
                r#"{"jsonrpc":"2.0","id":2,"result":{"tools":[]}}"#,
            );
        });

        let mut bridge = StdioBridge::new(McpEndpoint {
            host: "127.0.0.1".to_string(),
            port,
            path: "/mcp".to_string(),
        });
        let input = concat!(
            "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{}}\n",
            "{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/list\",\"params\":{}}\n"
        );
        let mut output = Vec::new();
        bridge
            .run_with(input.as_bytes(), &mut output)
            .expect("bridge runs");
        server.join().expect("server joins");

        let second = receiver.recv().expect("second request captured");
        assert!(second.contains("MCP-Session-Id: mcp-session-1"));
        let output = String::from_utf8(output).expect("output utf8");
        assert!(output.contains("tools"));
    }

    #[test]
    fn endpoint_formats_config_url() {
        let endpoint = McpEndpoint {
            host: "127.0.0.1".to_string(),
            port: 8800,
            path: "/mcp".to_string(),
        };
        assert_eq!(endpoint.url(), "http://127.0.0.1:8800/mcp");
    }

    fn read_request(stream: &mut TcpStream) -> String {
        let mut bytes = Vec::new();
        let mut buffer = [0_u8; 1024];
        loop {
            let size = stream.read(&mut buffer).expect("request read");
            assert_ne!(size, 0, "request closed before headers");
            bytes.extend_from_slice(&buffer[..size]);
            if let Some(header_end) = bytes.windows(4).position(|window| window == b"\r\n\r\n") {
                let head = String::from_utf8_lossy(&bytes[..header_end]).to_string();
                let content_length = head
                    .lines()
                    .find_map(|line| line.strip_prefix("Content-Length: "))
                    .and_then(|value| value.parse::<usize>().ok())
                    .unwrap_or(0);
                let body_start = header_end + 4;
                while bytes.len() < body_start + content_length {
                    let size = stream.read(&mut buffer).expect("body read");
                    assert_ne!(size, 0, "request closed before body");
                    bytes.extend_from_slice(&buffer[..size]);
                }
                return String::from_utf8_lossy(&bytes).to_string();
            }
        }
    }

    fn respond(stream: &mut TcpStream, extra_headers: &str, body: &str) {
        let response = format!(
            "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\n{extra_headers}Content-Length: {}\r\nConnection: close\r\n\r\n{}",
            body.len(),
            body
        );
        stream
            .write_all(response.as_bytes())
            .expect("response write");
    }
}
