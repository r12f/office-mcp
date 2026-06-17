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
