use super::SessionCommandQueue;
use crate::addin_mgr::QueuedCommand;
use std::time::{Duration, SystemTime};

#[test]
fn assigns_monotonic_sequences() {
    let mut queue = SessionCommandQueue::default();

    assert_eq!(queue.next_sequence(), 0);
    assert_eq!(queue.next_sequence(), 1);
    assert_eq!(queue.next_sequence(), 2);
}

#[test]
fn pushes_and_removes_commands_by_request_id() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut queue = SessionCommandQueue::default();
    queue.push(command("first", now, Duration::from_secs(30)));
    queue.push(command("second", now, Duration::from_secs(30)));

    let removed = queue.remove("first").expect("first command removed");

    assert_eq!(removed.request_id, "first");
    assert_eq!(queue.len(), 1);
    assert!(queue.remove("missing").is_none());
    assert_eq!(
        queue.remove("second").expect("second removed").request_id,
        "second"
    );
    assert!(queue.is_empty());
}

#[test]
fn marks_only_matching_command_dispatched() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut queue = SessionCommandQueue::default();
    queue.push(command("request-1", now, Duration::from_secs(30)));

    assert!(queue.mark_dispatched("request-1"));
    assert!(!queue.mark_dispatched("missing"));

    let command = queue.remove("request-1").expect("command remains queued");
    assert!(command.dispatched);
}

#[test]
fn returns_expired_request_ids_without_removing_commands() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    let mut queue = SessionCommandQueue::default();
    queue.push(command("expired", now, Duration::from_secs(5)));
    queue.push(command("active", now, Duration::from_secs(30)));

    let expired = queue.expired_request_ids(now + Duration::from_secs(6));

    assert_eq!(expired, vec!["expired".to_string()]);
    assert_eq!(queue.len(), 2);
}

fn command(request_id: &str, enqueued_at: SystemTime, timeout: Duration) -> QueuedCommand {
    QueuedCommand {
        command_id: format!("command-{request_id}"),
        request_id: request_id.to_string(),
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        tool: "word.get_text".to_string(),
        arguments_json: "{}".to_string(),
        timeout,
        enqueued_at,
        deadline_at: enqueued_at + timeout,
        sequence: 0,
        dispatched: false,
    }
}
