use super::{TrayAction, TrayController, TrayHealth, TrayMenuItem, TrayPlatform, TrayStatusInput};

#[test]
fn tray_snapshot_matches_required_menu_order() {
    let controller = TrayController::with_platform(TrayPlatform::WindowsNotificationArea);
    let snapshot = controller.snapshot(TrayStatusInput {
        health: TrayHealth::Up,
        client_count: 2,
        document_count: 3,
        running_task_count: 1,
    });

    assert_eq!(snapshot.platform, TrayPlatform::WindowsNotificationArea);
    assert_eq!(
        snapshot.tooltip,
        "Office MCP - Up - 2 clients - 3 documents"
    );
    assert_eq!(snapshot.menu[0].label(), Some("Status: Up"));
    assert_eq!(snapshot.menu[1].label(), Some("Clients: 2"));
    assert_eq!(snapshot.menu[2].label(), Some("Documents: 3"));
    assert_eq!(snapshot.menu[3], TrayMenuItem::Separator);
    assert_eq!(
        snapshot.menu[4],
        TrayMenuItem::Action {
            action: TrayAction::ShowUi,
            label: "Show Office MCP".to_string()
        }
    );
    assert_eq!(snapshot.menu[5].label(), Some("Quit Office MCP"));
}

#[test]
fn quit_confirmation_includes_current_counts_and_required_actions() {
    let snapshot =
        TrayController::with_platform(TrayPlatform::MacOsMenuBar).snapshot(TrayStatusInput {
            health: TrayHealth::Degraded,
            client_count: 1,
            document_count: 4,
            running_task_count: 2,
        });

    assert_eq!(snapshot.menu[0].label(), Some("Status: Degraded"));
    assert!(snapshot.quit_confirmation.body.contains("1 clients"));
    assert!(snapshot.quit_confirmation.body.contains("4 documents"));
    assert!(snapshot.quit_confirmation.body.contains("2 running tasks"));
    assert_eq!(snapshot.quit_confirmation.primary_action, "Quit Office MCP");
    assert_eq!(snapshot.quit_confirmation.secondary_action, "Keep Running");
}

#[test]
fn down_status_defaults_to_zero_counts() {
    let snapshot = TrayController::with_platform(TrayPlatform::LinuxStatusNotifier)
        .snapshot(TrayStatusInput::down());

    assert_eq!(snapshot.menu[0].label(), Some("Status: Down"));
    assert_eq!(snapshot.menu[1].label(), Some("Clients: 0"));
    assert_eq!(snapshot.menu[2].label(), Some("Documents: 0"));
}

#[test]
fn tray_model_supports_all_desktop_platform_targets() {
    for platform in [
        TrayPlatform::WindowsNotificationArea,
        TrayPlatform::MacOsMenuBar,
        TrayPlatform::LinuxStatusNotifier,
    ] {
        let snapshot = TrayController::with_platform(platform).snapshot(TrayStatusInput {
            health: TrayHealth::Up,
            client_count: 1,
            document_count: 2,
            running_task_count: 0,
        });

        assert_eq!(snapshot.platform, platform);
        assert_eq!(snapshot.menu[4].label(), Some("Show Office MCP"));
        assert_eq!(snapshot.menu[5].label(), Some("Quit Office MCP"));
        assert!(!snapshot.platform.label().is_empty());
    }
}

#[test]
fn tray_status_can_be_derived_from_redacted_ui_state() {
    let ui_state = serde_json::json!({
        "daemon": { "status": "degraded" },
        "clients": [{ "client_id": "client-1" }],
        "documents": { "word": [{ "session_id": "word-1" }], "excel": [{ "session_id": "excel-1" }], "other": [] },
        "current_tasks": [{ "command_id": "command-1" }]
    });

    let snapshot = TrayController::with_platform(TrayPlatform::WindowsNotificationArea)
        .snapshot_from_ui_state(Some(&ui_state));

    assert_eq!(snapshot.menu[0].label(), Some("Status: Degraded"));
    assert_eq!(snapshot.menu[1].label(), Some("Clients: 1"));
    assert_eq!(snapshot.menu[2].label(), Some("Documents: 2"));
    assert!(snapshot.quit_confirmation.body.contains("1 running tasks"));
    assert_eq!(snapshot.probe_json()["menu_items"][3], "---");
}
