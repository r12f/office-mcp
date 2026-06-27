use super::{RecordingShutdownController, TrayQuitRequest};

#[test]
fn request_records_controller_stop() {
    let controller = RecordingShutdownController::default();
    let request = TrayQuitRequest::new("native_tray_menu");

    request
        .shutdown_with(&controller)
        .expect("shutdown succeeds");

    assert_eq!(controller.stop_count(), 1);
    assert_eq!(request.action(), "quit");
    assert_eq!(request.source(), "native_tray_menu");
}

#[test]
fn request_failure_message_contains_action_source_pid_and_controller_error() {
    let controller = RecordingShutdownController::failing("controller refused stop");

    let error = TrayQuitRequest::new("native_tray_menu")
        .shutdown_with(&controller)
        .expect_err("controller error should propagate");
    let message = error.to_string();

    assert!(message.contains("quit"));
    assert!(message.contains("native_tray_menu"));
    assert!(message.contains(&format!("pid={}", std::process::id())));
    assert!(message.contains("controller refused stop"));
}
