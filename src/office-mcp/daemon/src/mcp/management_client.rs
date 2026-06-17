use crate::common::DaemonConfig;
use std::collections::BTreeMap;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::io::{Read, Write};
use std::net::TcpStream;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct McpManagementClient {
    host: String,
    port: u16,
}

impl McpManagementClient {
    #[must_use]
    pub fn from_config(config: &DaemonConfig) -> Self {
        Self {
            host: config.mcp.host.clone(),
            port: u16::try_from(config.mcp.port).unwrap_or(8800),
        }
    }

    /// Calls the daemon's `office.list_sessions` management tool.
    ///
    /// # Errors
    ///
    /// Returns an error when the daemon cannot be reached or returns invalid HTTP.
    pub fn list_sessions(&self) -> Result<String, McpManagementError> {
        let initialize = self.post(
            None,
            r#"{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}"#,
        )?;
        if initialize.status >= 400 {
            return Err(McpManagementError::Daemon(initialize.body));
        }
        let session_id = initialize
            .header("mcp-session-id")
            .ok_or_else(|| McpManagementError::Protocol("missing MCP session id".to_string()))?;
        let sessions = self.post(
            Some(session_id),
            r#"{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"office.list_sessions","arguments":{}}}"#,
        )?;
        if sessions.status >= 400 {
            return Err(McpManagementError::Daemon(sessions.body));
        }
        Ok(sessions.body)
    }

    fn post(
        &self,
        session_id: Option<&str>,
        body: &str,
    ) -> Result<HttpResponse, McpManagementError> {
        let mut stream =
            TcpStream::connect((&self.host[..], self.port)).map_err(McpManagementError::Io)?;
        let mut request = format!(
            concat!(
                "POST /mcp HTTP/1.1\r\n",
                "Host: {}:{}\r\n",
                "Content-Type: application/json\r\n",
                "Accept: application/json\r\n",
                "X-Office-Mcp-Client: office-mcp-cli/0.1.0\r\n",
                "Content-Length: {}\r\n"
            ),
            self.host,
            self.port,
            body.len()
        );
        if let Some(session_id) = session_id {
            request.push_str("MCP-Session-Id: ");
            request.push_str(session_id);
            request.push_str("\r\n");
        }
        request.push_str("\r\n");
        stream
            .write_all(request.as_bytes())
            .and_then(|()| stream.write_all(body.as_bytes()))
            .map_err(McpManagementError::Io)?;
        HttpResponse::read(&mut stream)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct HttpResponse {
    status: u16,
    headers: BTreeMap<String, String>,
    body: String,
}

impl HttpResponse {
    fn read(stream: &mut TcpStream) -> Result<Self, McpManagementError> {
        let bytes = read_http_response_bytes(stream)?;
        let Some(header_end) = find_header_end(&bytes) else {
            return Err(McpManagementError::Protocol(
                "malformed HTTP response".to_string(),
            ));
        };
        let head = std::str::from_utf8(&bytes[..header_end])
            .map_err(|_| McpManagementError::Protocol("non-UTF-8 headers".to_string()))?;
        let mut lines = head.split("\r\n");
        let status = lines
            .next()
            .and_then(|line| line.split_whitespace().nth(1))
            .and_then(|value| value.parse::<u16>().ok())
            .ok_or_else(|| McpManagementError::Protocol("missing HTTP status".to_string()))?;
        let mut headers = BTreeMap::new();
        for line in lines.filter(|line| !line.is_empty()) {
            if let Some((name, value)) = line.split_once(':') {
                headers.insert(name.trim().to_ascii_lowercase(), value.trim().to_string());
            }
        }
        let body = String::from_utf8(bytes[header_end + 4..].to_vec())
            .map_err(|_| McpManagementError::Protocol("non-UTF-8 body".to_string()))?;
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
pub enum McpManagementError {
    Io(std::io::Error),
    Daemon(String),
    Protocol(String),
}

impl Display for McpManagementError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "MCP management client I/O error: {error}"),
            Self::Daemon(message) | Self::Protocol(message) => formatter.write_str(message),
        }
    }
}

impl Error for McpManagementError {}

fn read_http_response_bytes(stream: &mut TcpStream) -> Result<Vec<u8>, McpManagementError> {
    let mut bytes = Vec::new();
    let mut buffer = [0_u8; 1024];
    loop {
        let size = stream.read(&mut buffer).map_err(McpManagementError::Io)?;
        if size == 0 {
            return Ok(bytes);
        }
        bytes.extend_from_slice(&buffer[..size]);
        let Some(header_end) = find_header_end(&bytes) else {
            continue;
        };
        let head = std::str::from_utf8(&bytes[..header_end])
            .map_err(|_| McpManagementError::Protocol("non-UTF-8 headers".to_string()))?;
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
    use super::McpManagementClient;
    use std::io::{Read, Write};
    use std::net::{TcpListener, TcpStream};
    use std::sync::mpsc;
    use std::thread;

    #[test]
    fn list_sessions_initializes_and_calls_tool() {
        let listener = TcpListener::bind("127.0.0.1:0").expect("listener");
        let port = listener.local_addr().expect("addr").port();
        let (sender, receiver) = mpsc::channel();
        let server = thread::spawn(move || {
            let mut first = listener.accept().expect("first").0;
            let first_request = read_request(&mut first);
            assert!(first_request.contains("initialize"));
            respond(
                &mut first,
                "MCP-Session-Id: mcp-session-1\r\n",
                r#"{"jsonrpc":"2.0","id":1,"result":{}}"#,
            );

            let mut second = listener.accept().expect("second").0;
            let second_request = read_request(&mut second);
            sender.send(second_request).expect("send second");
            respond(
                &mut second,
                "",
                r#"{"jsonrpc":"2.0","id":2,"result":{"content":[]}}"#,
            );
        });
        let client = McpManagementClient {
            host: "127.0.0.1".to_string(),
            port,
        };

        let body = client.list_sessions().expect("list sessions");

        server.join().expect("server joins");
        let second = receiver.recv().expect("second request");
        assert!(second.contains("MCP-Session-Id: mcp-session-1"));
        assert!(second.contains("office.list_sessions"));
        assert!(body.contains("content"));
    }

    fn read_request(stream: &mut TcpStream) -> String {
        let mut bytes = Vec::new();
        let mut buffer = [0_u8; 1024];
        loop {
            let size = stream.read(&mut buffer).expect("read");
            assert_ne!(size, 0);
            bytes.extend_from_slice(&buffer[..size]);
            if let Some(header_end) = bytes.windows(4).position(|window| window == b"\r\n\r\n") {
                let head = String::from_utf8_lossy(&bytes[..header_end]);
                let content_length = head
                    .lines()
                    .find_map(|line| line.strip_prefix("Content-Length: "))
                    .and_then(|value| value.parse::<usize>().ok())
                    .unwrap_or(0);
                if bytes.len() >= header_end + 4 + content_length {
                    return String::from_utf8_lossy(&bytes).to_string();
                }
            }
        }
    }

    fn respond(stream: &mut TcpStream, extra_headers: &str, body: &str) {
        let response = format!(
            "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\n{extra_headers}Content-Length: {}\r\nConnection: close\r\n\r\n{}",
            body.len(),
            body
        );
        stream.write_all(response.as_bytes()).expect("write");
    }
}
