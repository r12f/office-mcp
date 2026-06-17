use super::{WireHttpRequest, WireHttpResponse};
use crate::mcp::HttpMethod;
use crate::runtime::RuntimeServerError;
use std::collections::BTreeMap;

#[test]
fn reads_http_request_with_headers_query_and_body() {
    let body = br#"{"jsonrpc":"2.0","method":"initialize"}"#;
    let request_text = format!(
        "POST /mcp?session=abc HTTP/1.1\r\nHost: localhost\r\nContent-Length: {}\r\nX-Test: yes\r\n\r\n{}",
        body.len(),
        std::str::from_utf8(body).expect("body utf8")
    );

    let request = WireHttpRequest::read_from(&mut request_text.as_bytes(), 4096).expect("request");

    assert_eq!(request.method, HttpMethod::Post);
    assert_eq!(request.path, "/mcp");
    assert_eq!(request.headers.get("host"), Some(&"localhost".to_string()));
    assert_eq!(request.headers.get("x-test"), Some(&"yes".to_string()));
    assert_eq!(request.body, body);
    assert!(request.is_initialize());
}

#[test]
fn rejects_oversized_http_request() {
    let request = b"GET /healthz HTTP/1.1\r\nHost: localhost\r\n\r\n";
    let error = WireHttpRequest::read_from(&mut request.as_slice(), 8)
        .expect_err("oversized request rejected");

    assert_bad_request(error, "exceeds");
}

#[test]
fn rejects_malformed_headers() {
    let request = b"GET /healthz HTTP/1.1\r\nBrokenHeader\r\n\r\n";
    let error = WireHttpRequest::read_from(&mut request.as_slice(), 4096)
        .expect_err("malformed header rejected");

    assert_bad_request(error, "Malformed");
}

#[test]
fn writes_json_and_switching_protocol_responses() {
    let response = WireHttpResponse::json(
        200,
        BTreeMap::from([("X-Test".to_string(), "yes".to_string())]),
        "{\"ok\":true}".to_string(),
    );
    let response_text = String::from_utf8(response.to_bytes()).expect("response utf8");

    assert!(response_text.starts_with("HTTP/1.1 200 OK\r\n"));
    assert!(response_text.contains("Content-Type: application/json\r\n"));
    assert!(response_text.contains("X-Test: yes\r\n"));
    assert!(response_text.ends_with("\r\n{\"ok\":true}"));

    let upgrade = WireHttpResponse::switching_protocols(BTreeMap::from([(
        "Upgrade".to_string(),
        "websocket".to_string(),
    )]));
    let upgrade_text = String::from_utf8(upgrade.to_bytes()).expect("upgrade utf8");

    assert!(upgrade_text.starts_with("HTTP/1.1 101 Switching Protocols\r\n"));
    assert!(upgrade_text.contains("Upgrade: websocket\r\n"));
}

fn assert_bad_request(error: RuntimeServerError, expected: &str) {
    match error {
        RuntimeServerError::BadRequest(message) => assert!(
            message.contains(expected),
            "expected `{message}` to contain `{expected}`"
        ),
        other => panic!("expected bad request, got {other:?}"),
    }
}
