use super::McpCatalogResponder;
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionRegistry,
};
use crate::mcp::{AccessMode, ToolAccessPolicy};
use serde_json::{Value, json};

#[test]
fn tools_list_contains_office_word_and_excel_tools() {
    let response = parse(&McpCatalogResponder::tools_list(
        &json!(1),
        &ToolAccessPolicy::default(),
    ));
    let names = response["result"]["tools"]
        .as_array()
        .expect("tools")
        .iter()
        .filter_map(|tool| tool["name"].as_str())
        .collect::<Vec<_>>();

    assert!(names.contains(&"office.list_sessions"));
    assert!(names.contains(&"word.get_text"));
    assert!(names.contains(&"excel.read_range"));
    assert!(names.contains(&"powerpoint.add_slide"));
    assert!(names.contains(&"powerpoint.replace_text"));
}

#[test]
fn tools_list_filters_daemon_disabled_tools() {
    let policy = ToolAccessPolicy::default().with_disabled_tool("word.get_text");
    let response = parse(&McpCatalogResponder::tools_list(&json!(1), &policy));
    let names = response["result"]["tools"]
        .as_array()
        .expect("tools")
        .iter()
        .filter_map(|tool| tool["name"].as_str())
        .collect::<Vec<_>>();

    assert!(!names.contains(&"word.get_text"));
    assert!(names.contains(&"word.get_outline"));
    assert!(names.contains(&"office.get_session_info"));
}

#[test]
fn tools_list_filters_by_daemon_access_mode() {
    let policy = ToolAccessPolicy::default().with_access_mode(AccessMode::Read);
    let response = parse(&McpCatalogResponder::tools_list(&json!(1), &policy));
    let names = response["result"]["tools"]
        .as_array()
        .expect("tools")
        .iter()
        .filter_map(|tool| tool["name"].as_str())
        .collect::<Vec<_>>();

    assert!(names.contains(&"office.list_sessions"));
    assert!(names.contains(&"word.get_text"));
    assert!(names.contains(&"excel.read_range"));
    assert!(names.contains(&"powerpoint.list_slides"));
    assert!(!names.contains(&"word.insert_paragraph"));
    assert!(!names.contains(&"excel.write_range"));
    assert!(!names.contains(&"powerpoint.delete_slide"));
}

#[test]
fn resources_list_includes_sessions_and_word_resources() {
    let registry = registry_with_word_session();
    let response = parse(&McpCatalogResponder::resources_list(
        &registry,
        &json!("r1"),
    ));
    let uris = response["result"]["resources"]
        .as_array()
        .expect("resources")
        .iter()
        .filter_map(|resource| resource["uri"].as_str())
        .collect::<Vec<_>>();

    assert!(uris.contains(&"office://sessions"));
    assert!(uris.contains(&"office://word/session-1/document?offset=0&limit=200"));
    assert!(uris.contains(&"office://word/session-1/structure"));
}

#[test]
fn resource_templates_list_uses_word_templates() {
    let response = parse(&McpCatalogResponder::resource_templates_list(&json!("t1")));
    let templates = response["result"]["resourceTemplates"]
        .as_array()
        .expect("templates")
        .iter()
        .filter_map(|template| template["uriTemplate"].as_str())
        .collect::<Vec<_>>();

    assert!(templates.contains(&"office://word/{session_id}/document{?offset,limit}"));
    assert!(templates.contains(&"office://word/{session_id}/track_changes"));
}

#[test]
fn prompts_get_validates_name_and_renders_messages() {
    let missing = parse(&McpCatalogResponder::prompts_get(&json!("p0"), &json!({})));
    assert_eq!(missing["error"]["code"], -32602);

    let prompt = parse(&McpCatalogResponder::prompts_get(
        &json!("p1"),
        &json!({
            "params": {
                "name": "polish_section",
                "arguments": { "session_id": "session-1", "heading": "Scope" }
            }
        }),
    ));

    assert!(
        prompt["result"]["description"]
            .as_str()
            .expect("description")
            .contains("propose edits")
    );
    assert!(
        prompt["result"]["messages"][0]["content"]["text"]
            .as_str()
            .expect("prompt text")
            .contains("Scope")
    );
}

fn parse(text: &str) -> Value {
    serde_json::from_str(text).expect("json response")
}

fn registry_with_word_session() -> SessionRegistry {
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
            available_tools: vec!["word.get_text".to_string()],
            is_active: Some(true),
        },
        now,
    );
    registry
}
