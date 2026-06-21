use super::{RuntimeServer, RuntimeServerConfig};
use crate::addin_mgr::AddinChannelServer;
use crate::addin_mgr::AddinConnectionHub;
use crate::addin_mgr::CommandRouter;
use crate::addin_mgr::ImageFetcher;
use crate::addin_mgr::SessionRegistry;
use crate::addin_mgr::websocket_accept_key;
use crate::addin_mgr::{AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::common::{
    AddinConfig, AuditConfig, ConfigLogLevel, DaemonConfig, LimitsConfig, LoggingConfig, McpConfig,
    ToolAccessConfig,
};
use crate::mcp::{AccessMode, McpHttpFrontend, ToolAccessPolicy};
use crate::runtime::mcp_response::RuntimeSharedState;
use crate::runtime::mcp_rpc::{McpDispatchContext, McpJsonRpcRuntime};
use native_tls::TlsConnector;
use serde_json::Value;
use std::io::{Read, Write};
use std::net::{TcpListener, TcpStream};
use std::sync::{Arc, Mutex};
use std::thread;

#[test]
fn serves_healthz_over_real_loopback_socket() {
    let response = roundtrip("GET /healthz HTTP/1.1\r\nHost: localhost\r\n\r\n");

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("{\"ok\":true}"));
}

#[test]
fn initializes_mcp_session_over_real_loopback_socket() {
    let body = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{}}";
    let response = roundtrip(&format!(
        "POST /mcp HTTP/1.1\r\nHost: localhost\r\nOrigin: http://127.0.0.1:8800\r\nContent-Length: {}\r\n\r\n{}",
        body.len(),
        body
    ));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("MCP-Session-Id: mcp-session-1"));
    let reply: Value = serde_json::from_str(http_body(&response)).expect("initialize json");
    assert_eq!(reply["id"], 1);
    assert_eq!(reply["result"]["serverInfo"]["name"], "office-mcp");
    assert!(reply["result"]["capabilities"]["tools"].is_object());
}

#[test]
fn serves_tools_list_after_mcp_session_initialization() {
    let body = r#"{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}"#;
    let response = roundtrip_with_frontend(
        &format!(
            "POST /mcp HTTP/1.1\r\nHost: localhost\r\nMCP-Session-Id: mcp-session-1\r\nContent-Length: {}\r\n\r\n{}",
            body.len(),
            body
        ),
        |frontend, ui_state| {
            let initialize = crate::mcp::McpHttpRequest {
                method: crate::mcp::HttpMethod::Post,
                headers: std::collections::BTreeMap::new(),
                remote_addr: Some("127.0.0.1".to_string()),
                body_bytes: 0,
                is_initialize: true,
            };
            frontend.handle_request(ui_state, &initialize, std::time::SystemTime::UNIX_EPOCH);
        },
    );

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("office.list_sessions"));
    assert!(response.contains("word.get_text"));
    assert!(!response.contains("not wired"));
}

#[test]
fn rejects_foreign_browser_origin_over_socket() {
    let response =
        roundtrip("GET /mcp HTTP/1.1\r\nHost: localhost\r\nOrigin: https://evil.example\r\n\r\n");

    assert!(response.starts_with("HTTP/1.1 403 Forbidden"));
    assert!(response.contains("Forbidden origin"));
}

#[test]
fn serves_addin_taskpane_over_tls_socket() {
    let response = addin_tls_roundtrip(
        "GET /taskpane.html HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n",
    );

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("taskpane-shell"));
}

#[test]
fn serves_redacted_ui_state_over_addin_listener() {
    let response = addin_roundtrip(
        "GET /ui/state HTTP/1.1\r\nHost: localhost\r\nOrigin: https://localhost:8765\r\n\r\n",
    );

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("\"status\":\"up\""));
    assert!(response.contains("\"clients\":[]"));
    assert!(response.contains("\"documents\""));
    assert!(response.contains("\"recent_commands\""));
}

#[test]
fn ui_state_request_prunes_expired_stale_sessions() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        addin_host: "127.0.0.1".to_string(),
        addin_port: port,
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        certificate_path: super::default_pfx_path(),
        session_grace: std::time::Duration::from_millis(1),
        ..RuntimeServerConfig::default()
    });
    let acceptor = server.config.tls_acceptor().expect("tls acceptor");
    let stale_since = std::time::SystemTime::now() - std::time::Duration::from_secs(1);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(runtime("instance-prune", stale_since));
    registry.add_session(session("session-prune", "instance-prune"), stale_since);
    assert!(registry.remove_runtime("instance-prune", stale_since));
    let shared_state = Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(registry)),
        session_grace: std::time::Duration::from_millis(1),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::new())),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
        tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
    });
    let server_shared_state = Arc::clone(&shared_state);
    let server_handle = thread::spawn(move || {
        let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
        let (stream, _) = listener.accept().expect("accept");
        let mut stream = acceptor.accept(stream).expect("accept tls");
        server
            .handle_addin_tls_stream(&mut stream, &ui_state, &server_shared_state)
            .expect("handle addin tls");
    });

    let response = addin_tls_request(
        port,
        "GET /ui/state HTTP/1.1\r\nHost: localhost\r\nOrigin: https://localhost:8765\r\nConnection: close\r\n\r\n",
    );

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(!response.contains("session-prune"));
    assert!(
        shared_state
            .registry
            .lock()
            .expect("registry lock")
            .get_session_info("session-prune")
            .is_none()
    );
    server_handle.join().expect("server thread");
}

#[test]
fn ui_state_exposes_configured_log_path_for_debugging() {
    let mut config = daemon_config_with_log_path("C:\\logs\\office-mcp.log");
    config.tool_access = ToolAccessConfig {
        access_mode: AccessMode::Read,
        disabled_apps: vec!["powerpoint".to_string()],
        disabled_categories: vec![("excel".to_string(), "Range".to_string())],
        disabled_tools: vec!["word.update_table".to_string()],
    };
    let server = RuntimeServer::from_daemon_config(&config).expect("server config");
    let ui_state = server.ui_state_store();
    let snapshot = ui_state.snapshot(&[], std::time::SystemTime::now());

    assert_eq!(
        snapshot.daemon.log_path.as_deref(),
        Some("C:\\logs\\office-mcp.log")
    );
    assert_eq!(
        snapshot.daemon.tool_access_policy.access_mode,
        AccessMode::Read
    );
    assert_eq!(
        snapshot.daemon.tool_access_policy.disabled_apps,
        vec!["powerpoint"]
    );
    assert_eq!(
        snapshot.daemon.tool_access_policy.disabled_categories,
        vec![("excel".to_string(), "Range".to_string())]
    );
    assert_eq!(
        snapshot.daemon.tool_access_policy.disabled_tools,
        vec!["word.update_table"]
    );
}

#[test]
fn serves_daemon_ui_assets_over_addin_listener() {
    let html = addin_roundtrip("GET /ui/ HTTP/1.1\r\nHost: localhost\r\n\r\n");
    assert!(html.starts_with("HTTP/1.1 200 OK"));
    assert!(html.contains("Content-Type: text/html; charset=utf-8"));
    assert!(html.contains("Office MCP"));
    assert!(html.contains("id=\"currentTasks\""));
    assert!(html.contains("id=\"clients\""));
    assert!(html.contains("id=\"daemonVersion\""));
    assert!(html.contains("id=\"daemonUptime\""));
    assert!(html.contains("data-copy=\"mcpEndpoint\""));
    assert!(html.contains("id=\"resultFilter\""));

    let css = addin_roundtrip("GET /ui/app.css HTTP/1.1\r\nHost: localhost\r\n\r\n");
    assert!(css.starts_with("HTTP/1.1 200 OK"));
    assert!(css.contains("prefers-color-scheme"));
    assert!(css.contains("forced-colors"));
    assert!(css.contains("prefers-reduced-motion"));

    let js = addin_roundtrip("GET /ui/app.js HTTP/1.1\r\nHost: localhost\r\n\r\n");
    assert!(js.starts_with("HTTP/1.1 200 OK"));
    assert!(js.contains("/ui/state"));
    assert!(js.contains("renderDocuments"));
    assert!(js.contains("document_command_history"));
    assert!(js.contains("RelativeTimeFormat"));
    assert!(js.contains("config_path"));
    assert!(js.contains("last_error"));
    assert!(js.contains("emptyState('No documents connected'"));
    assert!(js.contains("fallbackCopy"));
}

#[test]
fn production_bound_daemon_exposes_ui_state_and_events() {
    let mcp_listener = TcpListener::bind("127.0.0.1:0").expect("bind mcp listener");
    let mcp_port = mcp_listener.local_addr().expect("mcp addr").port();
    let addin_listener = TcpListener::bind("127.0.0.1:0").expect("bind addin listener");
    let addin_port = addin_listener.local_addr().expect("addin addr").port();
    let addin_origin = format!("https://localhost:{addin_port}");
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        mcp_host: "127.0.0.1".to_string(),
        mcp_port,
        addin_host: "127.0.0.1".to_string(),
        addin_port,
        addin_origin: addin_origin.clone(),
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        certificate_path: super::default_pfx_path(),
        ..RuntimeServerConfig::default()
    });

    thread::spawn(move || {
        let _ = server.serve_bound_with_state_forever(
            &mcp_listener,
            &addin_listener,
            UiStateStore::new(),
            SessionRegistry::new(),
        );
    });

    let html = addin_tls_request(
        addin_port,
        "GET /ui/ HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n",
    );
    assert!(html.starts_with("HTTP/1.1 200 OK"));
    assert!(html.contains("Office MCP"));

    let state = addin_tls_request(
        addin_port,
        &format!(
            "GET /ui/state HTTP/1.1\r\nHost: localhost\r\nOrigin: {addin_origin}\r\nConnection: close\r\n\r\n"
        ),
    );
    assert!(state.starts_with("HTTP/1.1 200 OK"));
    assert!(state.contains("\"status\":\"up\""));
    assert!(state.contains(&format!("http://127.0.0.1:{mcp_port}/mcp")));
    assert!(state.contains(&format!("https://localhost:{addin_port}/addin")));

    let events = addin_tls_request(
        addin_port,
        &format!(
            "GET /ui/events HTTP/1.1\r\nHost: localhost\r\nOrigin: {addin_origin}\r\nConnection: close\r\n\r\n"
        ),
    );
    assert!(events.starts_with("HTTP/1.1 200 OK"));
    assert!(events.contains("Content-Type: text/event-stream; charset=utf-8"));
    assert!(events.contains("event: snapshot"));
    assert!(events.contains("data: {"));
}
#[test]
fn rejects_foreign_ui_state_origin_over_addin_listener() {
    let response = addin_roundtrip(
        "GET /ui/state HTTP/1.1\r\nHost: localhost\r\nOrigin: https://evil.example\r\n\r\n",
    );

    assert!(response.starts_with("HTTP/1.1 403 Forbidden"));
}

#[test]
fn accepts_addin_websocket_upgrade_with_exact_origin() {
    let response = addin_roundtrip(concat!(
        "GET /addin HTTP/1.1\r\n",
        "Host: localhost\r\n",
        "Origin: https://localhost:8765\r\n",
        "Upgrade: websocket\r\n",
        "Connection: Upgrade\r\n",
        "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n",
        "Sec-WebSocket-Version: 13\r\n",
        "\r\n"
    ));

    assert!(response.starts_with("HTTP/1.1 101 Switching Protocols"));
    assert!(response.contains("Sec-WebSocket-Accept: s3pPLMBiTxaQ9kYGzzhZRbK+xOo="));
}

#[test]
fn rejects_addin_websocket_upgrade_with_foreign_origin() {
    let response = addin_roundtrip(concat!(
        "GET /addin HTTP/1.1\r\n",
        "Host: localhost\r\n",
        "Origin: https://evil.example\r\n",
        "Upgrade: websocket\r\n",
        "Connection: Upgrade\r\n",
        "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n",
        "Sec-WebSocket-Version: 13\r\n",
        "\r\n"
    ));

    assert!(response.starts_with("HTTP/1.1 403 Forbidden"));
}

#[test]
fn computes_websocket_accept_key() {
    assert_eq!(
        websocket_accept_key("dGhlIHNhbXBsZSBub25jZQ=="),
        "s3pPLMBiTxaQ9kYGzzhZRbK+xOo="
    );
}

#[test]
fn real_tls_websocket_forwards_mcp_tool_call_and_returns_response() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        addin_host: "127.0.0.1".to_string(),
        addin_port: port,
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        certificate_path: super::default_pfx_path(),
        ..RuntimeServerConfig::default()
    });
    let acceptor = server.config.tls_acceptor().expect("tls acceptor");
    let shared_state = Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: std::time::Duration::from_secs(60),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::new())),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
        tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
    });
    let server_shared_state = Arc::clone(&shared_state);
    let server_handle = thread::spawn(move || {
        let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
        let (stream, _) = listener.accept().expect("accept");
        let mut stream = acceptor.accept(stream).expect("accept tls");
        server
            .handle_addin_tls_stream(&mut stream, &ui_state, &server_shared_state)
            .expect("handle addin websocket");
    });

    let connector = TlsConnector::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .expect("tls connector");
    let stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    let mut stream = connector.connect("localhost", stream).expect("connect tls");
    stream
        .write_all(
            concat!(
                "GET /addin HTTP/1.1\r\n",
                "Host: localhost\r\n",
                "Origin: https://localhost:8765\r\n",
                "Upgrade: websocket\r\n",
                "Connection: Upgrade\r\n",
                "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n",
                "Sec-WebSocket-Version: 13\r\n",
                "\r\n"
            )
            .as_bytes(),
        )
        .expect("write upgrade");
    let upgrade = read_http_response_head(&mut stream);
    assert!(upgrade.starts_with("HTTP/1.1 101 Switching Protocols"));

    write_client_ws_text(
        &mut stream,
        r#"{"jsonrpc":"2.0","id":"register-1","method":"register","params":{"instance_id":"instance-1","host":{"app":"word","version":"16.0","platform":"windows"},"add_in":{"version":"0.1.0","protocol_version":"1.0","supported_features":["doc.read"]}}}"#,
    );
    let register_reply = read_server_ws_text(&mut stream);
    assert!(register_reply.contains("assigned_instance_id"));
    write_client_ws_text(
        &mut stream,
        r#"{"jsonrpc":"2.0","method":"session.added","params":{"session_id":"session-1","instance_id":"instance-1","document":{"filename":"Live.docx"},"available_tools":["word.get_text"],"is_active":true}}"#,
    );
    wait_for_session(&shared_state.registry, "session-1");

    let mcp_shared_state = Arc::clone(&shared_state);
    let mcp_handle = thread::spawn(move || {
        let registry = mcp_shared_state.registry.lock().expect("registry").clone();
        let tool_access_policy = mcp_shared_state.tool_access_policy();
        let mut ui_state = UiStateStore::new();
        let mut context = McpDispatchContext {
            registry: &registry,
            ui_state: &mut ui_state,
            addin_channel: &mcp_shared_state.addin_channel,
            connection_hub: &mcp_shared_state.connection_hub,
            command_router: &mcp_shared_state.command_router,
            audit_log: &mcp_shared_state.audit_log,
            image_fetcher: &mcp_shared_state.image_fetcher,
            tool_access_policy: &tool_access_policy,
        };
        McpJsonRpcRuntime::handle_body(
            &mut context,
            br#"{"jsonrpc":"2.0","id":"call-1","method":"tools/call","params":{"name":"word.get_text","arguments":{"session_id":"session-1","offset":0,"limit":1}}}"#,
        )
    });

    let invoke = read_server_ws_tool_invoke(&mut stream);
    assert_eq!(invoke["method"], "tool.invoke");
    assert_eq!(invoke["params"]["session_id"], "session-1");
    let request_id = invoke["id"].as_str().expect("request id").to_string();
    write_client_ws_text(
        &mut stream,
        &format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"text":"live"}}}}}}"#
        ),
    );

    let mcp_reply = mcp_handle.join().expect("mcp thread");
    let mcp_reply: serde_json::Value = serde_json::from_str(&mcp_reply).expect("mcp json");
    assert_eq!(mcp_reply["result"]["structuredContent"]["text"], "live");
    drop(stream);
    server_handle.join().expect("server thread");
}

#[test]
fn real_tls_websocket_protocol_error_sends_close_frame() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        addin_host: "127.0.0.1".to_string(),
        addin_port: port,
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        certificate_path: super::default_pfx_path(),
        ..RuntimeServerConfig::default()
    });
    let acceptor = server.config.tls_acceptor().expect("tls acceptor");
    let shared_state = Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: std::time::Duration::from_secs(60),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::new())),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
        tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
    });
    let server_shared_state = Arc::clone(&shared_state);
    let server_handle = thread::spawn(move || {
        let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
        let (stream, _) = listener.accept().expect("accept");
        let mut stream = acceptor.accept(stream).expect("accept tls");
        server
            .handle_addin_tls_stream(&mut stream, &ui_state, &server_shared_state)
            .expect("handle addin websocket");
    });

    let connector = TlsConnector::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .expect("tls connector");
    let stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    let mut stream = connector.connect("localhost", stream).expect("connect tls");
    websocket_upgrade(&mut stream);
    let upgrade = read_http_response_head(&mut stream);
    assert!(upgrade.starts_with("HTTP/1.1 101 Switching Protocols"));

    stream
        .write_all(&[0x81, 0x02, b'h', b'i'])
        .expect("write unmasked frame");
    stream.flush().expect("flush unmasked frame");
    let (code, reason) = read_server_ws_close(&mut stream);
    assert_eq!(code, 1002);
    assert!(reason.contains("masked"));
    drop(stream);
    server_handle.join().expect("server thread");
}
#[test]
fn real_tls_websocket_heartbeat_ping_accepts_pong_response() {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        addin_host: "127.0.0.1".to_string(),
        addin_port: port,
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        certificate_path: super::default_pfx_path(),
        heartbeat_interval: std::time::Duration::from_millis(20),
        heartbeat_timeout: std::time::Duration::from_millis(200),
        ..RuntimeServerConfig::default()
    });
    let acceptor = server.config.tls_acceptor().expect("tls acceptor");
    let shared_state = Arc::new(RuntimeSharedState {
        registry: Arc::new(Mutex::new(SessionRegistry::new())),
        session_grace: std::time::Duration::from_secs(60),
        addin_channel: Arc::new(Mutex::new(AddinChannelServer::with_config(
            server.config.addin_channel_config(),
        ))),
        connection_hub: Arc::new(AddinConnectionHub::new()),
        command_router: Arc::new(Mutex::new(CommandRouter::new())),
        audit_log: AuditLog::new(),
        image_fetcher: ImageFetcher::new(),
        tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
    });
    let server_shared_state = Arc::clone(&shared_state);
    let server_handle = thread::spawn(move || {
        let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
        let (stream, _) = listener.accept().expect("accept");
        let mut stream = acceptor.accept(stream).expect("accept tls");
        server
            .handle_addin_tls_stream(&mut stream, &ui_state, &server_shared_state)
            .expect("handle addin websocket");
    });

    let connector = TlsConnector::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .expect("tls connector");
    let stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    let mut stream = connector.connect("localhost", stream).expect("connect tls");
    websocket_upgrade(&mut stream);
    let upgrade = read_http_response_head(&mut stream);
    assert!(upgrade.starts_with("HTTP/1.1 101 Switching Protocols"));
    write_client_ws_text(
        &mut stream,
        r#"{"jsonrpc":"2.0","id":"register-1","method":"register","params":{"instance_id":"instance-1","host":{"app":"word","version":"16.0","platform":"windows"},"add_in":{"version":"0.1.0","protocol_version":"1.0","supported_features":["doc.read"]}}}"#,
    );
    assert!(read_server_ws_text(&mut stream).contains("assigned_instance_id"));

    let ping = read_server_ws_text(&mut stream);
    let ping: serde_json::Value = serde_json::from_str(&ping).expect("ping json");
    assert_eq!(ping["method"], "ping");
    let ping_id = ping["id"].as_str().expect("ping id");
    write_client_ws_text(
        &mut stream,
        &format!(r#"{{"jsonrpc":"2.0","id":"{ping_id}","result":{{}}}}"#),
    );
    drop(stream);
    server_handle.join().expect("server thread");
}

fn roundtrip(request: &str) -> String {
    roundtrip_with_frontend(request, |_frontend, _ui_state| {})
}

fn http_body(response: &str) -> &str {
    response.split_once("\r\n\r\n").map_or("", |(_, body)| body)
}

fn roundtrip_with_frontend(
    request: &str,
    setup: impl FnOnce(&mut McpHttpFrontend, &mut UiStateStore) + Send + 'static,
) -> String {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        mcp_port: 8800,
        ..RuntimeServerConfig::default()
    });
    let handle = thread::spawn(move || {
        let mut frontend = McpHttpFrontend::new();
        let mut ui_state = UiStateStore::new();
        setup(&mut frontend, &mut ui_state);
        server
            .serve_next(&listener, &mut frontend, &mut ui_state)
            .expect("serve next");
    });
    let mut stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    stream.write_all(request.as_bytes()).expect("write request");
    stream
        .shutdown(std::net::Shutdown::Write)
        .expect("shutdown");
    let mut response = String::new();
    stream.read_to_string(&mut response).expect("read response");
    handle.join().expect("server thread");
    response
}

fn daemon_config_with_log_path(log_path: &str) -> DaemonConfig {
    DaemonConfig {
        config_path: "C:\\office-mcp\\config.toml".to_string(),
        addin: AddinConfig {
            host: "localhost".to_string(),
            port: 8765,
            origin: "https://localhost:8765".to_string(),
            pfx_path: "cert.pfx".to_string(),
            pfx_passphrase: "office-mcp-localhost".to_string(),
            heartbeat_interval_sec: 30,
            heartbeat_timeout_sec: 10,
            session_grace_sec: 60,
            max_pending_per_session: 4,
        },
        mcp: McpConfig {
            host: "127.0.0.1".to_string(),
            port: 8800,
        },
        limits: LimitsConfig {
            max_response_bytes: 1024 * 1024,
            max_request_bytes: 16 * 1024 * 1024,
            max_ws_frame_bytes: 16 * 1024 * 1024,
            default_tool_timeout_ms: 30_000,
            requests_per_minute: 1000,
        },
        audit: AuditConfig {
            enabled: false,
            path: String::new(),
        },
        logging: LoggingConfig {
            level: ConfigLogLevel::Info,
            file: log_path.to_string(),
        },
        tool_access: ToolAccessConfig::default(),
    }
}

fn addin_roundtrip(request: &str) -> String {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        addin_port: port,
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        ..RuntimeServerConfig::default()
    });
    let handle = thread::spawn(move || {
        let ui_state = UiStateStore::new();
        let (mut stream, _) = listener.accept().expect("accept");
        server
            .handle_addin_stream(&mut stream, &ui_state)
            .expect("handle addin");
    });
    let mut stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    stream.write_all(request.as_bytes()).expect("write request");
    stream
        .shutdown(std::net::Shutdown::Write)
        .expect("shutdown");
    let mut response = String::new();
    stream.read_to_string(&mut response).expect("read response");
    handle.join().expect("server thread");
    response
}

fn addin_tls_request(port: u16, request: &str) -> String {
    let connector = TlsConnector::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .expect("tls connector");
    let stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    let mut stream = connector.connect("localhost", stream).expect("connect tls");
    stream.write_all(request.as_bytes()).expect("write request");
    stream.flush().expect("flush request");
    let mut response = String::new();
    stream.read_to_string(&mut response).expect("read response");
    response
}
fn addin_tls_roundtrip(request: &str) -> String {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
    let port = listener.local_addr().expect("local addr").port();
    let server = RuntimeServer::with_config(RuntimeServerConfig {
        addin_host: "127.0.0.1".to_string(),
        addin_port: port,
        addin_public_dir: crate::addin_mgr::default_addin_public_dir(),
        certificate_path: super::default_pfx_path(),
        ..RuntimeServerConfig::default()
    });
    let acceptor = server.config.tls_acceptor().expect("tls acceptor");
    let handle = thread::spawn(move || {
        let ui_state = Arc::new(Mutex::new(UiStateStore::new()));
        let registry = Arc::new(Mutex::new(SessionRegistry::new()));
        let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
        let connection_hub = Arc::new(AddinConnectionHub::new());
        let shared_state = Arc::new(RuntimeSharedState {
            registry,
            session_grace: std::time::Duration::from_secs(60),
            addin_channel,
            connection_hub,
            command_router: Arc::new(Mutex::new(CommandRouter::new())),
            audit_log: AuditLog::new(),
            image_fetcher: ImageFetcher::new(),
            tool_access_policy: Arc::new(Mutex::new(ToolAccessPolicy::default())),
        });
        let (stream, _) = listener.accept().expect("accept");
        let mut stream = acceptor.accept(stream).expect("accept tls");
        server
            .handle_addin_tls_stream(&mut stream, &ui_state, &shared_state)
            .expect("handle addin tls");
    });
    let connector = TlsConnector::builder()
        .danger_accept_invalid_certs(true)
        .build()
        .expect("tls connector");
    let stream = TcpStream::connect(("127.0.0.1", port)).expect("connect client");
    let mut stream = connector.connect("localhost", stream).expect("connect tls");
    stream.write_all(request.as_bytes()).expect("write request");
    stream.flush().expect("flush request");
    let mut response = String::new();
    stream.read_to_string(&mut response).expect("read response");
    handle.join().expect("server thread");
    response
}

fn read_http_response_head(stream: &mut impl Read) -> String {
    let mut buffer = Vec::new();
    let mut byte = [0_u8; 1];
    loop {
        stream.read_exact(&mut byte).expect("read response byte");
        buffer.push(byte[0]);
        if buffer.ends_with(b"\r\n\r\n") {
            break;
        }
    }
    String::from_utf8(buffer).expect("response utf8")
}

fn websocket_upgrade(stream: &mut impl Write) {
    stream
        .write_all(
            concat!(
                "GET /addin HTTP/1.1\r\n",
                "Host: localhost\r\n",
                "Origin: https://localhost:8765\r\n",
                "Upgrade: websocket\r\n",
                "Connection: Upgrade\r\n",
                "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n",
                "Sec-WebSocket-Version: 13\r\n",
                "\r\n"
            )
            .as_bytes(),
        )
        .expect("write upgrade");
}

fn write_client_ws_text(stream: &mut impl Write, text: &str) {
    stream
        .write_all(&masked_text_frame(text))
        .expect("write websocket frame");
    stream.flush().expect("flush websocket frame");
}

fn read_server_ws_close(stream: &mut impl Read) -> (u16, String) {
    let mut header = [0_u8; 2];
    stream.read_exact(&mut header).expect("read close header");
    assert_eq!(header[0] & 0x0f, 0x8, "expected close frame");
    assert_eq!(header[0] & 0x80, 0x80, "close frame must be final");
    assert_eq!(header[1] & 0x80, 0, "server close must not be masked");
    let length = usize::from(header[1] & 0x7f);
    let mut payload = vec![0_u8; length];
    stream.read_exact(&mut payload).expect("read close payload");
    let code = u16::from_be_bytes([payload[0], payload[1]]);
    let reason = String::from_utf8(payload[2..].to_vec()).expect("close reason utf8");
    (code, reason)
}
fn read_server_ws_text(stream: &mut impl Read) -> String {
    let mut header = [0_u8; 2];
    stream.read_exact(&mut header).expect("read frame header");
    let opcode = header[0] & 0x0f;
    assert_eq!(opcode, 0x1, "expected text frame");
    let masked = header[1] & 0x80 != 0;
    assert!(!masked, "server frames must not be masked");
    let mut length = usize::from(header[1] & 0x7f);
    if length == 126 {
        let mut extended = [0_u8; 2];
        stream.read_exact(&mut extended).expect("read extended len");
        length = usize::from(u16::from_be_bytes(extended));
    } else if length == 127 {
        let mut extended = [0_u8; 8];
        stream.read_exact(&mut extended).expect("read extended len");
        length = usize::try_from(u64::from_be_bytes(extended)).expect("frame len");
    }
    let mut payload = vec![0_u8; length];
    stream.read_exact(&mut payload).expect("read frame payload");
    String::from_utf8(payload).expect("text utf8")
}

fn read_server_ws_tool_invoke(stream: &mut (impl Read + Write)) -> serde_json::Value {
    for _ in 0..8 {
        let message = read_server_ws_text(stream);
        let value: serde_json::Value = serde_json::from_str(&message).expect("websocket json");
        if value["method"] == "tool.invoke" {
            return value;
        }
        if value["method"] == "ping" {
            let ping_id = value["id"].as_str().expect("ping id");
            write_client_ws_text(
                stream,
                &format!(r#"{{"jsonrpc":"2.0","id":"{ping_id}","result":{{}}}}"#),
            );
            continue;
        }
        panic!("unexpected websocket message before tool.invoke: {value}");
    }
    panic!("tool.invoke was not received before heartbeat guard limit");
}

fn wait_for_session(registry: &Arc<Mutex<SessionRegistry>>, session_id: &str) {
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(5);
    while std::time::Instant::now() < deadline {
        if registry
            .lock()
            .expect("registry")
            .get_session_info(session_id)
            .is_some()
        {
            return;
        }
        thread::sleep(std::time::Duration::from_millis(10));
    }
    panic!("session was not registered before MCP tool call: {session_id}");
}

fn runtime(instance_id: &str, registered_at: std::time::SystemTime) -> RuntimeInfo {
    RuntimeInfo {
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["doc.read".to_string()],
        },
        registered_at,
    }
}

fn session(session_id: &str, instance_id: &str) -> NewSessionInfo {
    NewSessionInfo {
        session_id: session_id.to_string(),
        instance_id: instance_id.to_string(),
        document: DocumentInfo {
            filename: Some("Closed.docx".to_string()),
            ..DocumentInfo::default()
        },
        available_tools: vec!["word.get_text".to_string()],
        is_active: Some(true),
    }
}

fn masked_text_frame(text: &str) -> Vec<u8> {
    let mask = [1_u8, 2, 3, 4];
    let mut frame = vec![0x81];
    if text.len() < 126 {
        frame.push(0x80 | u8::try_from(text.len()).expect("short text"));
    } else if let Ok(length) = u16::try_from(text.len()) {
        frame.push(0x80 | 0x7e);
        frame.extend_from_slice(&length.to_be_bytes());
    } else {
        frame.push(0x80 | 127);
        frame.extend_from_slice(&(text.len() as u64).to_be_bytes());
    }
    frame.extend_from_slice(&mask);
    for (index, byte) in text.as_bytes().iter().enumerate() {
        frame.push(byte ^ mask[index % 4]);
    }
    frame
}
