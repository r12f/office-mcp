use super::{empty_state, seeded_state};
use crate::api::UiStateOptions;
use std::time::SystemTime;

#[test]
fn seeded_ui_fixture_is_redacted_and_grouped() {
    let seed = seeded_state(options());
    let sessions = seed.registry.list_sessions();
    let snapshot = seed.ui_state.snapshot(&sessions, SystemTime::UNIX_EPOCH);

    assert_eq!(snapshot.clients.len(), 1);
    assert_eq!(snapshot.documents["word"].len(), 1);
    assert_eq!(snapshot.documents["excel"].len(), 1);
    assert_eq!(snapshot.current_tasks.len(), 1);
    assert_eq!(snapshot.recent_commands.len(), 10);
    assert_eq!(
        snapshot.daemon.config_path.as_deref(),
        Some("C:\\office-mcp\\config.toml")
    );
    let debug = format!("{snapshot:?}");
    assert!(!debug.contains("secret-value"));
    assert!(!debug.contains("base64,QUJDREVGRw"));
    assert!(debug.contains("certificate_passphrase=[redacted]"));
}

#[test]
fn empty_ui_fixture_exercises_zero_state() {
    let seed = empty_state(options());
    let sessions = seed.registry.list_sessions();
    let snapshot = seed.ui_state.snapshot(&sessions, SystemTime::UNIX_EPOCH);

    assert_eq!(snapshot.clients.len(), 0);
    assert_eq!(snapshot.documents.values().map(Vec::len).sum::<usize>(), 0);
    assert_eq!(snapshot.current_tasks.len(), 0);
    assert_eq!(snapshot.recent_commands.len(), 0);
}

fn options() -> UiStateOptions {
    UiStateOptions {
        config_path: Some("C:\\office-mcp\\config.toml".to_string()),
        now: SystemTime::UNIX_EPOCH,
        ..UiStateOptions::default()
    }
}
