use super::RuntimeConnectionHandler;
use crate::api::UiStateStore;
use crate::runtime::RuntimeServerConfig;
use std::io::{Read, Write};
use std::net::{TcpListener, TcpStream};
use std::thread;

#[test]
fn plain_addin_stream_serves_ui_state() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("listener addr").port();
    let request_thread = thread::spawn(move || {
        let mut stream = TcpStream::connect(("127.0.0.1", port)).expect("connect");
        stream
            .write_all(b"GET /ui/state HTTP/1.1\r\nHost: localhost\r\n\r\n")
            .expect("write request");
        let mut response = String::new();
        stream.read_to_string(&mut response).expect("read response");
        response
    });

    let (mut stream, _) = listener.accept().expect("accept");
    RuntimeConnectionHandler::new(&RuntimeServerConfig::default())
        .handle_addin_stream(&mut stream, &UiStateStore::new())
        .expect("handle addin stream");
    drop(stream);

    let response = request_thread.join().expect("request thread");
    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("\"status\":\"up\""));
}
