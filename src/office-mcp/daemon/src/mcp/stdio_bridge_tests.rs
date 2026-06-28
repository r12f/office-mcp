use super::{McpEndpoint, StdioBridge};
use std::io::{Read, Write};
use std::net::{TcpListener, TcpStream};
use std::sync::mpsc;
use std::thread;
use std::time::{Duration, Instant};

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
fn bridge_uses_standard_mcp_stdio_framing_and_exposes_daemon_tools() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("listener");
    listener
        .set_nonblocking(true)
        .expect("listener nonblocking");
    let port = listener.local_addr().expect("addr").port();
    let (sender, receiver) = mpsc::channel();
    let server = thread::spawn(move || {
        let mut captured_bodies = Vec::new();
        let deadline = Instant::now() + Duration::from_secs(2);
        while captured_bodies.len() < 4 && Instant::now() < deadline {
            match listener.accept() {
                Ok((mut stream, _addr)) => {
                    let request = read_request(&mut stream);
                    let body = request_body(&request).to_string();
                    let is_initialize = body.contains(r#""method":"initialize""#);
                    captured_bodies.push(body);
                    if is_initialize {
                        respond(
                            &mut stream,
                            "MCP-Session-Id: mcp-session-1\r\n",
                            r#"{"jsonrpc":"2.0","id":1,"result":{}}"#,
                        );
                    } else {
                        respond(
                            &mut stream,
                            "",
                            r#"{"jsonrpc":"2.0","id":2,"result":{"tools":[{"name":"excel.read_range"},{"name":"powerpoint.list_slides"}]}}"#,
                        );
                    }
                }
                Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                    thread::sleep(Duration::from_millis(10));
                }
                Err(error) => panic!("accept failed: {error}"),
            }
        }
        sender.send(captured_bodies).expect("send captured bodies");
    });

    let mut bridge = StdioBridge::new(McpEndpoint {
        host: "127.0.0.1".to_string(),
        port,
        path: "/mcp".to_string(),
    });
    let initialize = r#"{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}"#;
    let tools_list = r#"{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}"#;
    let input = format!(
        "Content-Length: {}\r\n\r\n{}Content-Length: {}\r\n\r\n{}",
        initialize.len(),
        initialize,
        tools_list.len(),
        tools_list
    );
    let mut output = Vec::new();
    bridge
        .run_with(input.as_bytes(), &mut output)
        .expect("bridge runs");
    server.join().expect("server joins");

    let captured_bodies = receiver.recv().expect("captured bodies");
    assert_eq!(captured_bodies, vec![initialize, tools_list]);
    let output = String::from_utf8(output).expect("output utf8");
    assert!(output.contains("Content-Length:"));
    assert!(output.contains("excel.read_range"));
    assert!(output.contains("powerpoint.list_slides"));
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

fn request_body(request: &str) -> &str {
    request
        .split_once("\r\n\r\n")
        .map_or("", |(_head, body)| body)
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
