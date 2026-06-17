use crate::mcp::HttpMethod;
use crate::runtime_server::RuntimeServerError;
use std::collections::BTreeMap;
use std::io::Read;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WireHttpRequest {
    pub(crate) method: HttpMethod,
    pub(crate) path: String,
    pub(crate) headers: BTreeMap<String, String>,
    pub(crate) body: Vec<u8>,
}

impl WireHttpRequest {
    pub(crate) fn read_from(
        stream: &mut impl Read,
        max_request_bytes: usize,
    ) -> Result<Self, RuntimeServerError> {
        let mut buffer = Vec::new();
        let mut chunk = [0_u8; 4096];
        let header_end = loop {
            let read = stream.read(&mut chunk)?;
            if read == 0 {
                return Err(RuntimeServerError::BadRequest(
                    "Client closed before HTTP headers completed.".to_string(),
                ));
            }
            buffer.extend_from_slice(&chunk[..read]);
            if buffer.len() > max_request_bytes {
                return Err(RuntimeServerError::BadRequest(
                    "HTTP request exceeds configured byte limit.".to_string(),
                ));
            }
            if let Some(index) = find_header_end(&buffer) {
                break index;
            }
        };
        let (head, body_start) = buffer.split_at(header_end);
        let mut request = Self::parse_head(head)?;
        request.body.extend_from_slice(body_start);
        let content_length = request
            .headers
            .get("content-length")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(0);
        while request.body.len() < content_length {
            let read = stream.read(&mut chunk)?;
            if read == 0 {
                return Err(RuntimeServerError::BadRequest(
                    "Client closed before HTTP body completed.".to_string(),
                ));
            }
            request.body.extend_from_slice(&chunk[..read]);
            if request.body.len() + head.len() > max_request_bytes {
                return Err(RuntimeServerError::BadRequest(
                    "HTTP request exceeds configured byte limit.".to_string(),
                ));
            }
        }
        request.body.truncate(content_length);
        Ok(request)
    }

    fn parse_head(head: &[u8]) -> Result<Self, RuntimeServerError> {
        let text = std::str::from_utf8(head).map_err(|_| {
            RuntimeServerError::BadRequest("HTTP headers must be UTF-8.".to_string())
        })?;
        let mut lines = text.split("\r\n");
        let request_line = lines
            .next()
            .ok_or_else(|| RuntimeServerError::BadRequest("Missing request line.".to_string()))?;
        let mut parts = request_line.split_whitespace();
        let method = parse_method(parts.next().unwrap_or_default())?;
        let target = parts.next().unwrap_or_default();
        let path = target
            .split_once('?')
            .map_or(target, |(path, _query)| path)
            .to_string();
        if path.is_empty() {
            return Err(RuntimeServerError::BadRequest(
                "Missing request path.".to_string(),
            ));
        }
        let mut headers = BTreeMap::new();
        for line in lines.filter(|line| !line.is_empty()) {
            let Some((name, value)) = line.split_once(':') else {
                return Err(RuntimeServerError::BadRequest(
                    "Malformed HTTP header.".to_string(),
                ));
            };
            headers.insert(name.trim().to_ascii_lowercase(), value.trim().to_string());
        }
        Ok(Self {
            method,
            path,
            headers,
            body: Vec::new(),
        })
    }

    pub(crate) fn is_initialize(&self) -> bool {
        std::str::from_utf8(&self.body)
            .is_ok_and(|body| body.contains("initialize") && body.contains("method"))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WireHttpResponse {
    pub(crate) status: u16,
    reason: &'static str,
    pub(crate) headers: BTreeMap<String, String>,
    body: Vec<u8>,
}

impl WireHttpResponse {
    pub(crate) fn text(status: u16, body: String) -> Self {
        Self::new(status, "text/plain; charset=utf-8", body)
    }

    pub(crate) fn json(status: u16, headers: BTreeMap<String, String>, body: String) -> Self {
        let mut response = Self::new(status, "application/json", body);
        response.headers.extend(headers);
        response
    }

    pub(crate) fn binary(
        status: u16,
        content_type: &str,
        body: Vec<u8>,
        headers: BTreeMap<String, String>,
    ) -> Self {
        let mut response_headers = BTreeMap::from([
            ("Content-Type".to_string(), content_type.to_string()),
            ("Content-Length".to_string(), body.len().to_string()),
            ("Connection".to_string(), "close".to_string()),
        ]);
        response_headers.extend(headers);
        Self {
            status,
            reason: reason_phrase(status),
            headers: response_headers,
            body,
        }
    }

    fn new(status: u16, content_type: &str, body: String) -> Self {
        let mut headers = BTreeMap::from([
            ("Content-Type".to_string(), content_type.to_string()),
            ("Connection".to_string(), "close".to_string()),
        ]);
        headers.insert("Content-Length".to_string(), body.len().to_string());
        Self {
            status,
            reason: reason_phrase(status),
            headers,
            body: body.into_bytes(),
        }
    }

    pub(crate) fn switching_protocols(headers: BTreeMap<String, String>) -> Self {
        Self {
            status: 101,
            reason: reason_phrase(101),
            headers,
            body: Vec::new(),
        }
    }

    pub(crate) fn to_bytes(&self) -> Vec<u8> {
        let mut response = format!("HTTP/1.1 {} {}\r\n", self.status, self.reason).into_bytes();
        for (name, value) in &self.headers {
            response.extend_from_slice(format!("{name}: {value}\r\n").as_bytes());
        }
        response.extend_from_slice(b"\r\n");
        response.extend_from_slice(&self.body);
        response
    }
}

fn parse_method(value: &str) -> Result<HttpMethod, RuntimeServerError> {
    match value {
        "GET" => Ok(HttpMethod::Get),
        "POST" => Ok(HttpMethod::Post),
        "DELETE" => Ok(HttpMethod::Delete),
        "PUT" => Ok(HttpMethod::Put),
        "PATCH" => Ok(HttpMethod::Patch),
        "OPTIONS" => Ok(HttpMethod::Options),
        other => Err(RuntimeServerError::BadRequest(format!(
            "Unsupported HTTP method {other}."
        ))),
    }
}

fn find_header_end(buffer: &[u8]) -> Option<usize> {
    buffer
        .windows(4)
        .position(|window| window == b"\r\n\r\n")
        .map(|index| index + 4)
}

fn reason_phrase(status: u16) -> &'static str {
    match status {
        200 => "OK",
        101 => "Switching Protocols",
        400 => "Bad Request",
        403 => "Forbidden",
        404 => "Not Found",
        405 => "Method Not Allowed",
        413 => "Payload Too Large",
        429 => "Too Many Requests",
        501 => "Not Implemented",
        _ => "Error",
    }
}

#[cfg(test)]
#[path = "http_wire_tests.rs"]
mod http_wire_tests;
