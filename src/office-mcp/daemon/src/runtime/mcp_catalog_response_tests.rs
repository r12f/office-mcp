use super::McpCatalogResponder;
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionRegistry,
};
use crate::mcp::{AccessMode, ToolAccessPolicy, resource_request_from_uri};
use serde_json::{Value, json};
use std::collections::HashSet;

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
    assert!(names.contains(&"office.describe_tools"));
    assert!(!names.contains(&"office.describe_tool"));
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
    assert!(names.contains(&"office.describe_tools"));
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
    assert!(names.contains(&"office.describe_tools"));
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
fn resources_list_includes_powerpoint_resources_for_powerpoint_sessions() {
    let registry = registry_with_powerpoint_session();
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
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/presentation"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slides"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slide/0/text"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slide/0/shapes"));
}

#[test]
fn resources_list_includes_excel_resources_for_excel_sessions() {
    let registry = registry_with_excel_session();
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
    assert!(uris.contains(&"office://excel/excel-session/workbook"));
    assert!(uris.contains(&"office://excel/excel-session/sheets"));
    assert!(uris.contains(&"office://excel/excel-session/used-range"));
    assert!(uris.contains(&"office://excel/excel-session/range/A1"));
}

#[test]
fn resource_templates_list_uses_word_and_powerpoint_templates() {
    let response = parse(&McpCatalogResponder::resource_templates_list(&json!("t1")));
    let templates = response["result"]["resourceTemplates"]
        .as_array()
        .expect("templates")
        .iter()
        .filter_map(|template| template["uriTemplate"].as_str())
        .collect::<Vec<_>>();

    assert!(templates.contains(&"office://word/{session_id}/document{?offset,limit}"));
    assert!(templates.contains(&"office://word/{session_id}/track_changes"));
    assert!(templates.contains(&"office://excel/{session_id}/workbook"));
    assert!(templates.contains(&"office://excel/{session_id}/sheets"));
    assert!(templates.contains(&"office://excel/{session_id}/used-range{?sheet}"));
    assert!(templates.contains(&"office://excel/{session_id}/range/{address}{?sheet}"));
    assert!(templates.contains(&"office://powerpoint/{session_id}/presentation"));
    assert!(templates.contains(&"office://powerpoint/{session_id}/slide/{index}/text"));
}

#[test]
fn resource_discovery_covers_every_readable_resource_route() {
    let registry = registry_with_all_resource_sessions();
    let resources_response = parse(&McpCatalogResponder::resources_list(
        &registry,
        &json!("resources"),
    ));
    let template_response = parse(&McpCatalogResponder::resource_templates_list(&json!(
        "templates"
    )));
    let resources = resource_uris(&resources_response);
    let templates = resource_template_uris(&template_response);

    for route in concrete_resource_routes() {
        assert!(
            resource_request_from_uri(&registry, route.read_uri).is_ok(),
            "{uri} must be accepted by resources/read",
            uri = route.read_uri
        );
        assert!(
            resources.contains(route.discovered_uri),
            "{uri} must be exposed by resources/list as {discovered}",
            uri = route.read_uri,
            discovered = route.discovered_uri
        );
    }

    for route in dynamic_resource_routes() {
        assert!(
            resource_request_from_uri(&registry, route.read_uri).is_ok(),
            "{uri} must be accepted by resources/read",
            uri = route.read_uri
        );
        assert!(
            templates.contains(route.template_uri),
            "{uri} must be exposed by resources/templates/list as {template}",
            uri = route.read_uri,
            template = route.template_uri
        );
    }
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

#[derive(Debug, Clone, Copy)]
struct ConcreteResourceRoute {
    read_uri: &'static str,
    discovered_uri: &'static str,
}

#[derive(Debug, Clone, Copy)]
struct DynamicResourceRoute {
    read_uri: &'static str,
    template_uri: &'static str,
}

fn concrete_resource_routes() -> &'static [ConcreteResourceRoute] {
    &[
        ConcreteResourceRoute {
            read_uri: "office://sessions",
            discovered_uri: "office://sessions",
        },
        ConcreteResourceRoute {
            read_uri: "office://word/session-1/document?offset=0&limit=200",
            discovered_uri: "office://word/session-1/document?offset=0&limit=200",
        },
        ConcreteResourceRoute {
            read_uri: "office://word/session-1/structure",
            discovered_uri: "office://word/session-1/structure",
        },
        ConcreteResourceRoute {
            read_uri: "office://word/session-1/comments",
            discovered_uri: "office://word/session-1/comments",
        },
        ConcreteResourceRoute {
            read_uri: "office://word/session-1/track_changes",
            discovered_uri: "office://word/session-1/track_changes",
        },
        ConcreteResourceRoute {
            read_uri: "office://word/session-1/selection",
            discovered_uri: "office://word/session-1/selection",
        },
        ConcreteResourceRoute {
            read_uri: "office://excel/excel-session/workbook",
            discovered_uri: "office://excel/excel-session/workbook",
        },
        ConcreteResourceRoute {
            read_uri: "office://excel/excel-session/sheets",
            discovered_uri: "office://excel/excel-session/sheets",
        },
        ConcreteResourceRoute {
            read_uri: "office://excel/excel-session/used-range",
            discovered_uri: "office://excel/excel-session/used-range",
        },
        ConcreteResourceRoute {
            read_uri: "office://excel/excel-session/range/A1",
            discovered_uri: "office://excel/excel-session/range/A1",
        },
        ConcreteResourceRoute {
            read_uri: "office://powerpoint/powerpoint-session/presentation",
            discovered_uri: "office://powerpoint/powerpoint-session/presentation",
        },
        ConcreteResourceRoute {
            read_uri: "office://powerpoint/powerpoint-session/slides",
            discovered_uri: "office://powerpoint/powerpoint-session/slides",
        },
        ConcreteResourceRoute {
            read_uri: "office://powerpoint/powerpoint-session/slide/0/text",
            discovered_uri: "office://powerpoint/powerpoint-session/slide/0/text",
        },
        ConcreteResourceRoute {
            read_uri: "office://powerpoint/powerpoint-session/slide/0/shapes",
            discovered_uri: "office://powerpoint/powerpoint-session/slide/0/shapes",
        },
    ]
}

fn dynamic_resource_routes() -> &'static [DynamicResourceRoute] {
    &[
        DynamicResourceRoute {
            read_uri: "office://word/session-1/document?offset=10&limit=20",
            template_uri: "office://word/{session_id}/document{?offset,limit}",
        },
        DynamicResourceRoute {
            read_uri: "office://word/session-1/paragraph/3",
            template_uri: "office://word/{session_id}/paragraph/{index}",
        },
        DynamicResourceRoute {
            read_uri: "office://excel/excel-session/used-range?sheet=Data",
            template_uri: "office://excel/{session_id}/used-range{?sheet}",
        },
        DynamicResourceRoute {
            read_uri: "office://excel/excel-session/range/A1:B2?sheet=Data",
            template_uri: "office://excel/{session_id}/range/{address}{?sheet}",
        },
        DynamicResourceRoute {
            read_uri: "office://powerpoint/powerpoint-session/slide/2/text",
            template_uri: "office://powerpoint/{session_id}/slide/{index}/text",
        },
        DynamicResourceRoute {
            read_uri: "office://powerpoint/powerpoint-session/slide/2/shapes",
            template_uri: "office://powerpoint/{session_id}/slide/{index}/shapes",
        },
    ]
}

fn resource_uris(response: &Value) -> HashSet<&str> {
    response["result"]["resources"]
        .as_array()
        .expect("resources")
        .iter()
        .filter_map(|resource| resource["uri"].as_str())
        .collect()
}

fn resource_template_uris(response: &Value) -> HashSet<&str> {
    response["result"]["resourceTemplates"]
        .as_array()
        .expect("resource templates")
        .iter()
        .filter_map(|template| template["uriTemplate"].as_str())
        .collect()
}

fn registry_with_all_resource_sessions() -> SessionRegistry {
    let now = std::time::SystemTime::UNIX_EPOCH + std::time::Duration::from_secs(10);
    let mut registry = SessionRegistry::new();
    register_runtime_and_session(
        &mut registry,
        now,
        "word-instance",
        "word",
        "session-1",
        "Draft.docx",
        &["word.get_text", "word.get_selection"],
    );
    register_runtime_and_session(
        &mut registry,
        now,
        "excel-instance",
        "excel",
        "excel-session",
        "Budget.xlsx",
        &[
            "excel.get_workbook_info",
            "excel.list_sheets",
            "excel.get_used_range",
            "excel.read_range",
        ],
    );
    register_runtime_and_session(
        &mut registry,
        now,
        "powerpoint-instance",
        "powerpoint",
        "powerpoint-session",
        "Deck.pptx",
        &[
            "powerpoint.get_presentation_info",
            "powerpoint.list_slides",
            "powerpoint.read_text",
            "powerpoint.list_shapes",
        ],
    );
    registry
}

fn register_runtime_and_session(
    registry: &mut SessionRegistry,
    now: std::time::SystemTime,
    instance_id: &str,
    app: &str,
    session_id: &str,
    filename: &str,
    available_tools: &[&str],
) {
    registry.register_runtime(RuntimeInfo {
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: app.to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec![format!("{app}.session")],
        },
        registered_at: now,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: session_id.to_string(),
            instance_id: instance_id.to_string(),
            document: DocumentInfo {
                filename: Some(filename.to_string()),
                ..DocumentInfo::default()
            },
            available_tools: available_tools
                .iter()
                .map(|tool| (*tool).to_string())
                .collect(),
            is_active: Some(true),
        },
        now,
    );
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
            version: "0.1.0".to_string(),
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
            available_tools: vec!["excel.read_range".to_string()],
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
            available_tools: vec!["powerpoint.read_text".to_string()],
            is_active: Some(true),
        },
        now,
    );
    registry
}
