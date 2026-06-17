use super::{empty_state, seeded_state};
use std::time::SystemTime;

#[test]
fn seeded_ui_fixture_is_redacted_and_grouped() {
    let seed = seeded_state(SystemTime::UNIX_EPOCH);
    let sessions = seed.registry.list_sessions();
    let snapshot = seed.ui_state.snapshot(&sessions, SystemTime::UNIX_EPOCH);

    assert_eq!(snapshot.clients.len(), 1);
    assert_eq!(snapshot.documents["word"].len(), 1);
    assert_eq!(snapshot.documents["excel"].len(), 1);
    assert_eq!(snapshot.current_tasks.len(), 1);
    assert_eq!(snapshot.recent_commands.len(), 10);
    let debug = format!("{snapshot:?}");
    assert!(!debug.contains("secret-value"));
    assert!(!debug.contains("base64,QUJDREVGRw"));
    assert!(debug.contains("certificate_passphrase=[redacted]"));
}

#[test]
fn empty_ui_fixture_exercises_zero_state() {
    let seed = empty_state();
    let sessions = seed.registry.list_sessions();
    let snapshot = seed.ui_state.snapshot(&sessions, SystemTime::UNIX_EPOCH);

    assert_eq!(snapshot.clients.len(), 0);
    assert_eq!(snapshot.documents.values().map(Vec::len).sum::<usize>(), 0);
    assert_eq!(snapshot.current_tasks.len(), 0);
    assert_eq!(snapshot.recent_commands.len(), 0);
}
