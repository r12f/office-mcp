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
