use super::{McpDispatchContext, McpJsonRpcRuntime};
use crate::addin_mgr::{
    AddInInfo, AddinChannelServer, AddinConnectionHub, CommandRouter, DocumentInfo, HostInfo,
    ImageFetcher, NewSessionInfo, RuntimeInfo, SessionRegistry,
};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use std::sync::{Arc, Mutex};
use std::thread;
#[test]
fn mcp_json_rpc_lists_tools_and_connected_sessions() {
    let registry = registry_with_word_session();

    let tools = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}"#,
    );
    let tools: serde_json::Value = serde_json::from_str(&tools).expect("tools json");
    let mut names = tools["result"]["tools"]
        .as_array()
        .expect("tools")
        .iter()
        .filter_map(|tool| tool["name"].as_str())
        .collect::<Vec<_>>();
    names.sort_unstable();
    assert_eq!(
        names,
        vec![
            "excel.add_sheet",
            "excel.apply_filter",
            "excel.clear_range",
            "excel.create_chart",
            "excel.create_pivot_table",
            "excel.create_table",
            "excel.delete_sheet",
            "excel.find_replace_cells",
            "excel.format_range",
            "excel.get_used_range",
            "excel.get_workbook_info",
            "excel.list_sheets",
            "excel.read_range",
            "excel.set_formula",
            "excel.sort_range",
            "excel.update_chart",
            "excel.update_pivot_table",
            "excel.update_sheet",
            "excel.update_table",
            "excel.write_range",
            "office.get_session_info",
            "office.list_sessions",
            "powerpoint.add_shape",
            "powerpoint.add_slide",
            "powerpoint.add_table",
            "powerpoint.add_text_box",
            "powerpoint.apply_layout",
            "powerpoint.delete_slide",
            "powerpoint.export_file",
            "powerpoint.export_slide",
            "powerpoint.format_text",
            "powerpoint.get_active_view",
            "powerpoint.get_presentation_info",
            "powerpoint.get_selection",
            "powerpoint.insert_image",
            "powerpoint.list_layouts",
            "powerpoint.list_shapes",
            "powerpoint.list_slides",
            "powerpoint.move_slide",
            "powerpoint.read_table",
            "powerpoint.read_text",
            "powerpoint.replace_text",
            "powerpoint.set_selection",
            "powerpoint.update_shape",
            "powerpoint.update_slide",
            "powerpoint.update_table",
            "powerpoint.update_tags",
            "word.add_comment",
            "word.apply_formatting",
            "word.apply_style",
            "word.delete_content_control",
            "word.delete_range",
            "word.find_text",
            "word.get_outline",
            "word.get_paragraph",
            "word.get_selection",
            "word.get_text",
            "word.insert_content_control",
            "word.insert_image",
            "word.insert_list",
            "word.insert_page_break",
            "word.insert_paragraph",
            "word.insert_table",
            "word.list_content_controls",
            "word.read_table",
            "word.replace_text",
            "word.resolve_comment",
            "word.save",
            "word.update_content_control",
            "word.update_paragraph",
            "word.update_table",
            "word.update_tracked_change",
        ]
    );

    let sessions = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"call-1","method":"tools/call","params":{"name":"office.list_sessions","arguments":{}}}"#,
    );
    let sessions: serde_json::Value = serde_json::from_str(&sessions).expect("sessions json");
    assert_eq!(
        sessions["result"]["structuredContent"]["sessions"][0]["session_id"],
        "session-1"
    );
    assert_eq!(
        sessions["result"]["structuredContent"]["sessions"][0]["document"]["title"],
        "Draft.docx"
    );

    let info = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"call-2","method":"tools/call","params":{"name":"office.get_session_info","arguments":{"session_id":"session-1"}}}"#,
    );
    let info: serde_json::Value = serde_json::from_str(&info).expect("info json");
    assert_eq!(
        info["result"]["structuredContent"]["descriptor"]["session_id"],
        "session-1"
    );
    assert_eq!(
        info["result"]["structuredContent"]["available_tools"][0],
        "word.get_text"
    );
}

#[test]
fn mcp_json_rpc_lists_resources_and_prompts() {
    let registry = registry_with_word_session();

    let resources = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"resources-1","method":"resources/list","params":{}}"#,
    );
    let resources: serde_json::Value = serde_json::from_str(&resources).expect("resources json");
    let uris = resources["result"]["resources"]
        .as_array()
        .expect("resources")
        .iter()
        .filter_map(|resource| resource["uri"].as_str())
        .collect::<Vec<_>>();
    assert!(uris.contains(&"office://sessions"));
    assert!(uris.contains(&"office://word/session-1/document?offset=0&limit=200"));
    assert!(uris.contains(&"office://word/session-1/comments"));

    let templates = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"templates-1","method":"resources/templates/list","params":{}}"#,
    );
    let templates: serde_json::Value = serde_json::from_str(&templates).expect("templates json");
    let uri_templates = templates["result"]["resourceTemplates"]
        .as_array()
        .expect("resource templates")
        .iter()
        .filter_map(|template| template["uriTemplate"].as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        uri_templates,
        vec![
            "office://word/{session_id}/comments",
            "office://word/{session_id}/document{?offset,limit}",
            "office://word/{session_id}/paragraph/{index}",
            "office://word/{session_id}/selection",
            "office://word/{session_id}/structure",
            "office://word/{session_id}/track_changes",
        ]
    );

    let prompts = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"prompts-1","method":"prompts/list","params":{}}"#,
    );
    let prompts: serde_json::Value = serde_json::from_str(&prompts).expect("prompts json");
    let names = prompts["result"]["prompts"]
        .as_array()
        .expect("prompts")
        .iter()
        .filter_map(|prompt| prompt["name"].as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        names,
        vec![
            "summarize_document",
            "polish_section",
            "extract_action_items"
        ]
    );

    let prompt = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"prompt-1","method":"prompts/get","params":{"name":"polish_section","arguments":{"session_id":"session-1","heading":"Scope"}}}"#,
    );
    let prompt: serde_json::Value = serde_json::from_str(&prompt).expect("prompt json");
    let prompt_text = prompt["result"]["messages"][0]["content"]["text"]
        .as_str()
        .expect("prompt text");
    assert!(prompt_text.contains("Scope"));
    assert!(prompt_text.contains("explicit approval"));

    let summary = mcp_handle_body(
        &registry,
        br#"{"jsonrpc":"2.0","id":"prompt-2","method":"prompts/get","params":{"name":"summarize_document","arguments":{"session_id":"session-1"}}}"#,
    );
    let summary: serde_json::Value = serde_json::from_str(&summary).expect("summary prompt");
    let summary_text = summary["result"]["messages"][0]["content"]["text"]
        .as_str()
        .expect("summary prompt text");
    assert!(summary_text.contains("office://word/session-1/document"));
    assert!(summary_text.contains("word.add_comment"));
}

#[test]
fn mcp_json_rpc_reads_resources_through_addin_connection() {
    let registry = registry_with_word_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("connection-1");
    connection_hub.bind_instance("connection-1", "instance-1");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("connection-1");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        assert_eq!(invoke["method"], "tool.invoke");
        assert_eq!(invoke["params"]["tool"], "word._get_comments");
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"comments":[]}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"read-1","method":"resources/read","params":{"uri":"office://word/session-1/comments"}}"#,
    );
    response_thread.join().expect("response thread");
    let reply: serde_json::Value = serde_json::from_str(&reply).expect("resource reply json");
    assert_eq!(
        reply["result"]["contents"][0]["uri"],
        "office://word/session-1/comments"
    );
    assert!(
        reply["result"]["contents"][0]["text"]
            .as_str()
            .expect("resource text")
            .contains("comments")
    );
}

#[test]
fn mcp_json_rpc_structure_resource_routes_to_full_structure_tool() {
    let registry = registry_with_word_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("connection-1");
    connection_hub.bind_instance("connection-1", "instance-1");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("connection-1");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        assert_eq!(invoke["method"], "tool.invoke");
        assert_eq!(invoke["params"]["tool"], "word._get_structure");
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"outline":[],"headings":[],"lists":[],"tables":[]}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"read-structure","method":"resources/read","params":{"uri":"office://word/session-1/structure"}}"#,
    );
    response_thread.join().expect("response thread");
    let reply: serde_json::Value = serde_json::from_str(&reply).expect("reply json");
    assert_eq!(
        reply["result"]["contents"][0]["uri"],
        "office://word/session-1/structure"
    );
    assert_eq!(
        reply["result"]["contents"][0]["mimeType"],
        "application/json"
    );
}
#[test]
fn mcp_json_rpc_forwarded_word_tool_invokes_addin_connection() {
    let registry = registry_with_word_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("connection-1");
    connection_hub.bind_instance("connection-1", "instance-1");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("connection-1");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        assert_eq!(outbound.len(), 1);
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        assert_eq!(invoke["method"], "tool.invoke");
        assert_eq!(invoke["params"]["session_id"], "session-1");
        assert_eq!(invoke["params"]["tool"], "word.get_text");
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"text":"hello"}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"call-3","method":"tools/call","params":{"name":"word.get_text","arguments":{"session_id":"session-1","offset":0,"limit":1}}}"#,
    );

    response_thread.join().expect("response thread");
    let reply: serde_json::Value = serde_json::from_str(&reply).expect("reply json");
    assert_eq!(reply["result"]["structuredContent"]["text"], "hello");
    assert!(
        !reply["result"]
            .get("isError")
            .and_then(serde_json::Value::as_bool)
            .unwrap_or(false)
    );
    assert!(
        ui_state
            .snapshot(&registry.list_sessions(), std::time::SystemTime::UNIX_EPOCH)
            .current_tasks
            .is_empty()
    );
}

#[test]
fn mcp_json_rpc_forwarded_excel_tool_invokes_addin_connection() {
    let registry = registry_with_excel_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("excel-connection");
    connection_hub.bind_instance("excel-connection", "excel-instance");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("excel-connection");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        assert_eq!(outbound.len(), 1);
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        assert_eq!(invoke["method"], "tool.invoke");
        assert_eq!(invoke["params"]["session_id"], "excel-session");
        assert_eq!(invoke["params"]["tool"], "excel.create_table");
        assert_eq!(invoke["params"]["args"]["address"], "A1:B2");
        assert_eq!(invoke["params"]["args"]["has_headers"], true);
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"table":"Table1","address":"A1:B2","has_headers":true}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"excel-call","method":"tools/call","params":{"name":"excel.create_table","arguments":{"session_id":"excel-session","address":"A1:B2","has_headers":true}}}"#,
    );

    response_thread.join().expect("response thread");
    let reply: serde_json::Value = serde_json::from_str(&reply).expect("reply json");
    assert_eq!(reply["result"]["structuredContent"]["table"], "Table1");
    assert_eq!(reply["result"]["structuredContent"]["has_headers"], true);
}

#[test]
fn mcp_json_rpc_forwarded_powerpoint_tool_invokes_addin_connection() {
    let registry = registry_with_powerpoint_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("powerpoint-connection");
    connection_hub.bind_instance("powerpoint-connection", "powerpoint-instance");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("powerpoint-connection");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        assert_eq!(outbound.len(), 1);
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        assert_eq!(invoke["method"], "tool.invoke");
        assert_eq!(invoke["params"]["session_id"], "powerpoint-session");
        assert_eq!(invoke["params"]["tool"], "powerpoint.add_slide");
        assert_eq!(invoke["params"]["args"]["layout"], "TitleOnly");
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"slide_id":"slide-2","layout":"TitleOnly"}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"powerpoint-call","method":"tools/call","params":{"name":"powerpoint.add_slide","arguments":{"session_id":"powerpoint-session","layout":"TitleOnly"}}}"#,
    );

    response_thread.join().expect("response thread");
    let reply: serde_json::Value = serde_json::from_str(&reply).expect("reply json");
    assert_eq!(reply["result"]["structuredContent"]["slide_id"], "slide-2");
    assert_eq!(reply["result"]["structuredContent"]["layout"], "TitleOnly");
}

#[test]
fn mcp_json_rpc_insert_image_base64_is_validated_before_forwarding() {
    let registry = registry_with_word_session_with_tools(vec!["word.insert_image"]);
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("connection-1");
    connection_hub.bind_instance("connection-1", "instance-1");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("connection-1");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        assert_eq!(invoke["params"]["tool"], "word.insert_image");
        let args = invoke["params"]["args"].as_str().map_or_else(
            || invoke["params"]["args"].clone(),
            |raw| serde_json::from_str(raw).expect("parsed args"),
        );
        assert_eq!(args["image"]["mime_type"], "image/png");
        assert_eq!(args["image"]["byte_length"], 9);
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"inserted":true}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"call-image","method":"tools/call","params":{"name":"word.insert_image","arguments":{"session_id":"session-1","anchor":{"kind":"end_of_document"},"image":{"base64":"iVBORw0KGgoA"}}}}"#,
    );
    response_thread.join().expect("response thread");
    let reply: serde_json::Value = serde_json::from_str(&reply).expect("reply json");
    assert_eq!(reply["result"]["structuredContent"]["inserted"], true);
}

#[test]
fn mcp_json_rpc_forwarded_word_tool_writes_audit_records() {
    let audit_dir = std::env::temp_dir().join(format!(
        "office-mcp-runtime-audit-{}-{}",
        std::process::id(),
        unique_suffix()
    ));
    let audit_path = audit_dir.join("audit.jsonl");
    let registry = registry_with_word_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("connection-1");
    connection_hub.bind_instance("connection-1", "instance-1");
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let audit_log = AuditLog::enabled(&audit_path);
    let response_hub = Arc::clone(&connection_hub);
    let response_thread = thread::spawn(move || {
        let outbound = loop {
            let outbound = response_hub.take_outbound("connection-1");
            if !outbound.is_empty() {
                break outbound;
            }
            thread::sleep(std::time::Duration::from_millis(5));
        };
        let invoke: serde_json::Value = serde_json::from_str(&outbound[0]).expect("invoke json");
        let request_id = invoke["id"].as_str().expect("request id");
        assert!(response_hub.complete_from_text(&format!(
            r#"{{"jsonrpc":"2.0","id":"{request_id}","result":{{"ok":true,"data":{{"text":"document body"}}}}}}"#
        )));
    });

    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &audit_log,
        image_fetcher: &ImageFetcher::new(),
    };
    let _reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"call-audit","method":"tools/call","params":{"name":"word.get_text","arguments":{"session_id":"session-1","offset":0,"limit":10}}}"#,
    );
    response_thread.join().expect("response thread");

    let contents = std::fs::read_to_string(&audit_path).expect("audit file");
    assert!(contents.contains("\"tool\":\"word.get_text\""));
    assert!(contents.contains("\"session_id\":\"session-1\""));
    assert!(contents.contains("\"ok\":true"));
    assert!(!contents.contains("document body"));
    let _ = std::fs::remove_dir_all(audit_dir);
}

#[test]
fn mcp_json_rpc_preflight_failure_writes_redacted_audit_record() {
    let audit_dir = std::env::temp_dir().join(format!(
        "office-mcp-runtime-audit-{}-{}",
        std::process::id(),
        unique_suffix()
    ));
    let audit_path = audit_dir.join("audit.jsonl");
    let registry = registry_with_word_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let audit_log = AuditLog::enabled(&audit_path);
    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &audit_log,
        image_fetcher: &ImageFetcher::new(),
    };

    let _reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"call-audit-failure","method":"tools/call","params":{"name":"word.insert_paragraph","arguments":{"session_id":"session-1","text":"secret body","anchor":{"kind":"end_of_document"}}}}"#,
    );

    let contents = std::fs::read_to_string(&audit_path).expect("audit file");
    assert!(contents.contains("\"tool\":\"word.insert_paragraph\""));
    assert!(contents.contains("HOST_CAPABILITY_UNAVAILABLE"));
    assert!(contents.contains("\"ok\":false"));
    assert!(!contents.contains("secret body"));
    let _ = std::fs::remove_dir_all(audit_dir);
}

#[test]
fn mcp_json_rpc_forwarded_word_tool_sends_cancel_on_timeout() {
    let registry = registry_with_word_session();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    connection_hub.register_connection("connection-1");
    connection_hub.bind_instance("connection-1", "instance-1");
    let command_router = Arc::new(Mutex::new(CommandRouter::with_limits(
        1024 * 1024,
        std::time::Duration::from_millis(10),
    )));
    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };

    let reply = McpJsonRpcRuntime::handle_body(
        &mut context,
        br#"{"jsonrpc":"2.0","id":"call-timeout","method":"tools/call","params":{"name":"word.get_text","arguments":{"session_id":"session-1"}}}"#,
    );

    let reply: serde_json::Value = serde_json::from_str(&reply).expect("reply json");
    assert_eq!(
        reply["result"]["structuredContent"]["error"]["office_mcp_code"],
        "TIMEOUT"
    );
    let outbound = connection_hub.take_outbound("connection-1");
    assert_eq!(outbound.len(), 2);
    let cancel: serde_json::Value = serde_json::from_str(&outbound[1]).expect("cancel json");
    assert_eq!(cancel["method"], "tool.cancel");
    assert_eq!(cancel["params"]["reason"], "deadline_expired");
}

fn mcp_handle_body(registry: &SessionRegistry, body: &[u8]) -> String {
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let mut context = McpDispatchContext {
        registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &AuditLog::new(),
        image_fetcher: &ImageFetcher::new(),
    };
    McpJsonRpcRuntime::handle_body(&mut context, body)
}

fn registry_with_word_session() -> SessionRegistry {
    registry_with_word_session_with_tools(vec![
        "word.get_text",
        "word.get_outline",
        "word.get_paragraph",
        "word.get_selection",
        "word.save",
    ])
}

fn registry_with_word_session_with_tools(tools: Vec<&str>) -> SessionRegistry {
    let now = std::time::SystemTime::UNIX_EPOCH + std::time::Duration::from_secs(10);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(RuntimeInfo {
        instance_id: "instance-1".to_string(),
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
        registered_at: now,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: "session-1".to_string(),
            instance_id: "instance-1".to_string(),
            document: DocumentInfo {
                filename: Some("Draft.docx".to_string()),
                ..DocumentInfo::default()
            },
            available_tools: tools.into_iter().map(str::to_string).collect(),
            is_active: Some(true),
        },
        now,
    );
    registry
}

fn registry_with_excel_session() -> SessionRegistry {
    let now = std::time::SystemTime::UNIX_EPOCH + std::time::Duration::from_secs(10);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(RuntimeInfo {
        instance_id: "excel-instance".to_string(),
        host: HostInfo {
            app: "excel".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.6".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["workbook.session".to_string()],
        },
        registered_at: now,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: "excel-session".to_string(),
            instance_id: "excel-instance".to_string(),
            document: DocumentInfo {
                filename: Some("Budget.xlsx".to_string()),
                ..DocumentInfo::default()
            },
            available_tools: crate::mcp::ExcelToolCatalog::tools()
                .iter()
                .map(|tool| tool.name.to_string())
                .collect(),
            is_active: Some(true),
        },
        now,
    );
    registry
}

fn registry_with_powerpoint_session() -> SessionRegistry {
    let now = std::time::SystemTime::UNIX_EPOCH + std::time::Duration::from_secs(10);
    let mut registry = SessionRegistry::new();
    registry.register_runtime(RuntimeInfo {
        instance_id: "powerpoint-instance".to_string(),
        host: HostInfo {
            app: "powerpoint".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["presentation.session".to_string()],
        },
        registered_at: now,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: "powerpoint-session".to_string(),
            instance_id: "powerpoint-instance".to_string(),
            document: DocumentInfo {
                filename: Some("Deck.pptx".to_string()),
                ..DocumentInfo::default()
            },
            available_tools: crate::mcp::PowerPointToolCatalog::tools()
                .iter()
                .map(|tool| tool.name.to_string())
                .collect(),
            is_active: Some(true),
        },
        now,
    );
    registry
}

fn unique_suffix() -> u128 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("time")
        .as_nanos()
}
