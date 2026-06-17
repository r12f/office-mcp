use super::{LoopbackHttpTarget, TrayUiStateClient};

#[test]
fn parses_loopback_https_state_url() {
    let target = LoopbackHttpTarget::parse("https://localhost:8765/ui/state")
        .expect("parse loopback state URL");

    assert_eq!(target.scheme, "https");
    assert_eq!(target.host, "localhost");
    assert_eq!(target.port, 8765);
    assert_eq!(target.path, "/ui/state");
}

#[test]
fn rejects_non_loopback_state_url() {
    let error = LoopbackHttpTarget::parse("https://example.invalid/ui/state")
        .expect_err("non-loopback rejected");

    assert!(error.to_string().contains("loopback"));
}

#[test]
fn client_keeps_runtime_url_for_debugging() {
    let client = TrayUiStateClient::new("https://localhost:8765/ui/state");

    assert_eq!(client.url, "https://localhost:8765/ui/state");
}
