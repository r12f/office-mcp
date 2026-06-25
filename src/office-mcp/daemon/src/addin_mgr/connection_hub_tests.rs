use super::{AddinConnectionHub, AddinConnectionHubError};
use std::sync::Arc;
use std::thread;
use std::time::Duration;

#[test]
fn queues_outbound_payloads_by_bound_instance() {
    let hub = AddinConnectionHub::new();
    hub.register_connection("connection-1");
    hub.bind_instance("connection-1", "instance-1");

    assert!(hub.send_to_instance("instance-1", "payload-1".to_string()));
    assert_eq!(hub.take_outbound("connection-1"), vec!["payload-1"]);
    assert!(hub.take_outbound("connection-1").is_empty());
}

#[test]
fn invoke_waits_for_matching_response_text() {
    let hub = Arc::new(AddinConnectionHub::new());
    hub.register_connection("connection-1");
    hub.bind_instance("connection-1", "instance-1");
    let response_hub = Arc::clone(&hub);

    let handle = thread::spawn(move || {
        let outbound = wait_for_outbound(&response_hub, "connection-1");
        assert_eq!(outbound, vec!["request".to_string()]);
        assert!(
            response_hub
                .complete_from_text(r#"{"jsonrpc":"2.0","id":"request-1","result":{"ok":true}}"#)
        );
    });

    let response = hub
        .invoke(
            "instance-1",
            "request-1",
            "request".to_string(),
            Duration::from_secs(1),
        )
        .expect("invoke response");
    handle.join().expect("response thread");

    assert_eq!(response["result"]["ok"], true);
}

fn wait_for_outbound(hub: &AddinConnectionHub, connection_id: &str) -> Vec<String> {
    for _ in 0..100 {
        let outbound = hub.take_outbound(connection_id);
        if !outbound.is_empty() {
            return outbound;
        }
        thread::sleep(Duration::from_millis(1));
    }
    Vec::new()
}

#[test]
fn invoke_reports_missing_connection_and_timeout() {
    let hub = AddinConnectionHub::new();
    assert_eq!(
        hub.invoke(
            "missing-instance",
            "request-1",
            "request".to_string(),
            Duration::from_millis(1),
        ),
        Err(AddinConnectionHubError::NoConnection)
    );

    hub.register_connection("connection-1");
    hub.bind_instance("connection-1", "instance-1");
    assert_eq!(
        hub.invoke(
            "instance-1",
            "request-1",
            "request".to_string(),
            Duration::from_millis(1),
        ),
        Err(AddinConnectionHubError::Timeout)
    );
}

#[test]
fn remove_connection_unbinds_instance() {
    let hub = AddinConnectionHub::new();
    hub.register_connection("connection-1");
    hub.bind_instance("connection-1", "instance-1");
    hub.remove_connection("connection-1");

    assert!(!hub.send_to_instance("instance-1", "payload".to_string()));
    assert_eq!(hub.take_outbound("connection-1"), Vec::<String>::new());
}
