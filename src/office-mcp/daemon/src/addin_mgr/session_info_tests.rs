use super::SessionInfo;
use crate::addin_mgr::{DocumentInfo, NewSessionInfo, SessionPatch, SessionStatus};
use std::time::{Duration, SystemTime};

#[test]
fn new_session_info_starts_active_with_registered_metadata() {
    let registered_at = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let session = SessionInfo::new(new_session(), registered_at);

    assert_eq!(session.session_id, "session-1");
    assert_eq!(session.instance_id, "instance-1");
    assert_eq!(session.status, SessionStatus::Active);
    assert_eq!(session.registered_at, registered_at);
    assert_eq!(session.stale_since, None);
}

#[test]
fn patch_replaces_only_supplied_fields() {
    let registered_at = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut session = SessionInfo::new(new_session(), registered_at);

    session.apply_patch(SessionPatch {
        document: Some(DocumentInfo {
            title: Some("Final".to_string()),
            ..DocumentInfo::default()
        }),
        available_tools: Some(vec!["word.get_text".to_string()]),
        is_active: Some(Some(false)),
    });

    assert_eq!(session.document.title.as_deref(), Some("Final"));
    assert_eq!(session.available_tools, vec!["word.get_text"]);
    assert_eq!(session.is_active, Some(false));
    assert_eq!(session.registered_at, registered_at);
}

#[test]
fn stale_check_requires_stale_status_and_elapsed_grace() {
    let stale_since = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut session = SessionInfo::new(new_session(), stale_since);

    assert!(!session.is_stale_past(stale_since + Duration::from_mins(2), Duration::from_mins(1)));

    session.status = SessionStatus::Stale;
    session.stale_since = Some(stale_since);

    assert!(!session.is_stale_past(stale_since + Duration::from_mins(1), Duration::from_mins(1)));
    assert!(session.is_stale_past(
        stale_since + Duration::from_secs(61),
        Duration::from_mins(1)
    ));
}

fn new_session() -> NewSessionInfo {
    NewSessionInfo {
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        document: DocumentInfo {
            filename: Some("Draft.docx".to_string()),
            ..DocumentInfo::default()
        },
        available_tools: vec!["word.get_text".to_string(), "word.add_comment".to_string()],
        is_active: Some(true),
    }
}
