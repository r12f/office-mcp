use super::{ResourceReadRequest, resource_request_from_uri};
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionRegistry,
};
use serde_json::json;
use std::time::SystemTime;

#[test]
fn parses_sessions_resource() {
    let registry = registry();

    assert!(matches!(
        resource_request_from_uri(&registry, "office://sessions"),
        Ok(ResourceReadRequest::Sessions)
    ));
}

#[test]
fn parses_document_resource_with_default_and_custom_ranges() {
    let registry = registry();

    let default = resource_request_from_uri(&registry, "office://word/session-1/document")
        .expect("default document request");
    assert_eq!(
        default,
        ResourceReadRequest::Forwarded {
            uri: "office://word/session-1/document".to_string(),
            tool: "word.get_text",
            arguments: json!({ "session_id": "session-1", "offset": 0, "limit": 200 }),
            check_capability: true,
        }
    );

    let ranged = resource_request_from_uri(
        &registry,
        "office://word/session-1/document?offset=10&limit=20",
    )
    .expect("ranged document request");
    assert_eq!(
        ranged,
        ResourceReadRequest::Forwarded {
            uri: "office://word/session-1/document?offset=10&limit=20".to_string(),
            tool: "word.get_text",
            arguments: json!({ "session_id": "session-1", "offset": 10, "limit": 20 }),
            check_capability: true,
        }
    );
}

#[test]
fn parses_structure_and_paragraph_resources() {
    let registry = registry();

    assert_eq!(
        resource_request_from_uri(&registry, "office://word/session-1/structure")
            .expect("structure request"),
        ResourceReadRequest::Forwarded {
            uri: "office://word/session-1/structure".to_string(),
            tool: "word._get_structure",
            arguments: json!({ "session_id": "session-1" }),
            check_capability: false,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://word/session-1/paragraph/3")
            .expect("paragraph request"),
        ResourceReadRequest::Forwarded {
            uri: "office://word/session-1/paragraph/3".to_string(),
            tool: "word.get_paragraph",
            arguments: json!({ "session_id": "session-1", "index": 3 }),
            check_capability: true,
        }
    );
}

#[test]
fn parses_powerpoint_read_only_resources() {
    let registry = registry_with_app("powerpoint");

    assert_eq!(
        resource_request_from_uri(&registry, "office://powerpoint/session-1/presentation")
            .expect("presentation request"),
        ResourceReadRequest::Forwarded {
            uri: "office://powerpoint/session-1/presentation".to_string(),
            tool: "powerpoint.get_presentation_info",
            arguments: json!({ "session_id": "session-1" }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://powerpoint/session-1/slides")
            .expect("slides request"),
        ResourceReadRequest::Forwarded {
            uri: "office://powerpoint/session-1/slides".to_string(),
            tool: "powerpoint.list_slides",
            arguments: json!({ "session_id": "session-1" }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(
            &registry,
            "office://powerpoint/session-1/slides/text?start=2&end=5",
        )
        .expect("slides text request"),
        ResourceReadRequest::Forwarded {
            uri: "office://powerpoint/session-1/slides/text?start=2&end=5".to_string(),
            tool: "powerpoint.read_text",
            arguments: json!({
                "session_id": "session-1",
                "start": 2,
                "end": 5,
            }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://powerpoint/session-1/slides/text")
            .expect("slides text default range request"),
        ResourceReadRequest::Forwarded {
            uri: "office://powerpoint/session-1/slides/text".to_string(),
            tool: "powerpoint.read_text",
            arguments: json!({
                "session_id": "session-1",
                "start": 0,
            }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(
            &registry,
            "office://powerpoint/session-1/slide/2/text?offset=10&limit=20",
        )
        .expect("slide text request"),
        ResourceReadRequest::Forwarded {
            uri: "office://powerpoint/session-1/slide/2/text?offset=10&limit=20".to_string(),
            tool: "powerpoint.read_text",
            arguments: json!({
                "session_id": "session-1",
                "slide_index": 2,
                "offset": 10,
                "limit": 20,
            }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://powerpoint/session-1/slide/2/shapes")
            .expect("slide shapes request"),
        ResourceReadRequest::Forwarded {
            uri: "office://powerpoint/session-1/slide/2/shapes".to_string(),
            tool: "powerpoint.list_shapes",
            arguments: json!({ "session_id": "session-1", "slide_index": 2 }),
            check_capability: true,
        }
    );
}

#[test]
fn parses_excel_read_only_resources() {
    let registry = registry_with_app("excel");

    assert_eq!(
        resource_request_from_uri(&registry, "office://excel/session-1/workbook")
            .expect("workbook request"),
        ResourceReadRequest::Forwarded {
            uri: "office://excel/session-1/workbook".to_string(),
            tool: "excel.get_workbook_info",
            arguments: json!({ "session_id": "session-1" }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://excel/session-1/sheets")
            .expect("sheets request"),
        ResourceReadRequest::Forwarded {
            uri: "office://excel/session-1/sheets".to_string(),
            tool: "excel.list_sheets",
            arguments: json!({ "session_id": "session-1" }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://excel/session-1/used-range?sheet=Data")
            .expect("used range request"),
        ResourceReadRequest::Forwarded {
            uri: "office://excel/session-1/used-range?sheet=Data".to_string(),
            tool: "excel.get_used_range",
            arguments: json!({ "session_id": "session-1", "sheet": "Data" }),
            check_capability: true,
        }
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://excel/session-1/range/A1:B2?sheet=Data")
            .expect("range request"),
        ResourceReadRequest::Forwarded {
            uri: "office://excel/session-1/range/A1:B2?sheet=Data".to_string(),
            tool: "excel.read_range",
            arguments: json!({ "session_id": "session-1", "address": "A1:B2", "sheet": "Data" }),
            check_capability: true,
        }
    );
}

#[test]
fn rejects_bad_excel_resource_uris() {
    let registry = registry_with_app("excel");

    assert_eq!(
        resource_request_from_uri(&registry, "office://excel/missing/workbook")
            .expect_err("missing session"),
        "Session missing is not registered."
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://excel/session-1/range/")
            .expect_err("missing range address"),
        "Excel range resource requires a non-empty address."
    );
}

#[test]
fn rejects_bad_powerpoint_slide_index() {
    let registry = registry_with_app("powerpoint");

    assert_eq!(
        resource_request_from_uri(&registry, "office://powerpoint/session-1/slide/nope/text")
            .expect_err("bad slide index"),
        "slide index must be a non-negative integer."
    );
}

#[test]
fn rejects_unknown_sessions_and_bad_numbers() {
    let registry = registry();

    assert_eq!(
        resource_request_from_uri(&registry, "office://word/missing/document")
            .expect_err("missing session"),
        "Session missing is not registered."
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://word/session-1/document?limit=bad")
            .expect_err("bad limit"),
        "limit must be a non-negative integer."
    );
    assert_eq!(
        resource_request_from_uri(&registry, "office://word/session-1/paragraph/nope")
            .expect_err("bad paragraph"),
        "paragraph index must be a non-negative integer."
    );
}

fn registry() -> SessionRegistry {
    registry_with_app("word")
}

fn registry_with_app(app: &str) -> SessionRegistry {
    let mut registry = SessionRegistry::new();
    registry.register_runtime(RuntimeInfo {
        instance_id: "instance-1".to_string(),
        host: HostInfo {
            app: app.to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: Vec::new(),
        },
        registered_at: SystemTime::UNIX_EPOCH,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: "session-1".to_string(),
            instance_id: "instance-1".to_string(),
            document: DocumentInfo::default(),
            available_tools: Vec::new(),
            is_active: Some(true),
        },
        SystemTime::UNIX_EPOCH,
    );
    registry
}
