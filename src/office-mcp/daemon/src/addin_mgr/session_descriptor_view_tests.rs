use super::SessionDescriptorView;
use crate::addin_mgr::{DocumentDescriptor, HostDescriptor, SessionDescriptor, SessionStatus};
use crate::mcp::WORD_V1_TOOLS;
use std::time::{Duration, SystemTime};

#[test]
fn renders_active_session_descriptor_json() {
    let descriptor = descriptor(SessionStatus::Active);
    let rendered = SessionDescriptorView::new(&descriptor).to_json();

    assert_eq!(rendered["session_id"], "session-1");
    assert_eq!(rendered["instance_id"], "instance-1");
    assert_eq!(rendered["app"], "word");
    assert_eq!(rendered["host"]["version"], "16.0");
    assert_eq!(rendered["document"]["filename"], "Doc.docx");
    assert_eq!(rendered["document"]["is_dirty"], true);
    assert_eq!(rendered["capability_tiers"][0], "core");
    assert_eq!(rendered["available_tools"][0], "word.get_text");
    assert_eq!(rendered["available_tool_count"], 26);
    assert_eq!(rendered["queue_depth"], 2);
    assert_eq!(rendered["registered_at"], "unix:5");
    assert_eq!(rendered["status"], "active");
}

#[test]
fn renders_stale_session_status() {
    let descriptor = descriptor(SessionStatus::Stale);
    let rendered = SessionDescriptorView::new(&descriptor).to_json();

    assert_eq!(rendered["status"], "stale");
}

fn descriptor(status: SessionStatus) -> SessionDescriptor {
    SessionDescriptor {
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        app: "word".to_string(),
        host: HostDescriptor {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("pc".to_string()),
            build: Some("Desktop".to_string()),
        },
        document: DocumentDescriptor {
            title: Some("Doc".to_string()),
            url: None,
            filename: Some("Doc.docx".to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(false),
            protection_kind: None,
            rights: Some(vec!["write".to_string()]),
            rights_source: Some("host".to_string()),
        },
        is_active: Some(true),
        capability_tiers: vec!["core".to_string()],
        available_tools: vec!["word.get_text".to_string()],
        available_tool_count: WORD_V1_TOOLS.len(),
        queue_depth: 2,
        registered_at: SystemTime::UNIX_EPOCH + Duration::from_secs(5),
        status,
    }
}
