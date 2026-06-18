use super::{TrayAction, TrayHealth, TrayMenuItem, TrayPlatform, TraySnapshot, TrayStatusInput};

#[test]
fn snapshot_matches_required_menu_order() {
    let snapshot = TraySnapshot::from_input(
        TrayPlatform::WindowsNotificationArea,
        TrayStatusInput {
            health: TrayHealth::Up,
            client_count: 2,
            document_count: 3,
            running_task_count: 1,
        },
    );

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
    let snapshot = TraySnapshot::from_input(
        TrayPlatform::MacOsMenuBar,
        TrayStatusInput {
            health: TrayHealth::Degraded,
            client_count: 1,
            document_count: 4,
            running_task_count: 2,
        },
    );

    assert_eq!(snapshot.menu[0].label(), Some("Status: Degraded"));
    assert!(snapshot.quit_confirmation.body.contains("1 clients"));
    assert!(snapshot.quit_confirmation.body.contains("4 documents"));
    assert!(snapshot.quit_confirmation.body.contains("2 running tasks"));
    assert_eq!(snapshot.quit_confirmation.primary_action, "Quit Office MCP");
    assert_eq!(snapshot.quit_confirmation.secondary_action, "Keep Running");
}

#[test]
fn tray_status_can_be_derived_from_redacted_ui_state() {
    let ui_state = serde_json::json!({
        "daemon": { "status": "degraded" },
        "clients": [{ "client_id": "client-1" }],
        "documents": { "word": [{ "session_id": "word-1" }], "excel": [{ "session_id": "excel-1" }], "other": [] },
        "current_tasks": [{ "command_id": "command-1" }]
    });

    let input = TrayStatusInput::from_ui_state(Some(&ui_state));
    let snapshot = TraySnapshot::from_input(TrayPlatform::WindowsNotificationArea, input);

    assert_eq!(snapshot.menu[0].label(), Some("Status: Degraded"));
    assert_eq!(snapshot.menu[1].label(), Some("Clients: 1"));
    assert_eq!(snapshot.menu[2].label(), Some("Documents: 2"));
    assert!(snapshot.quit_confirmation.body.contains("1 running tasks"));
    let probe = snapshot.probe_json();
    assert_eq!(probe["menu_items"][3], "---");
    assert_eq!(
        probe["tooltip"],
        "Office MCP - Degraded - 1 clients - 2 documents"
    );
    assert_eq!(probe["menu"][0]["kind"], "read_only");
    assert_eq!(probe["menu"][0]["enabled"], false);
    assert_eq!(probe["menu"][3]["kind"], "separator");
    assert_eq!(probe["menu"][4]["kind"], "action");
    assert_eq!(probe["menu"][4]["action"], "show_ui");
    assert_eq!(probe["menu"][5]["action"], "quit");
}

#[test]
fn tray_product_surface_has_no_scaffold_or_debug_labels() {
    let snapshot = TraySnapshot::from_input(
        TrayPlatform::WindowsNotificationArea,
        TrayStatusInput {
            health: TrayHealth::Up,
            client_count: 0,
            document_count: 0,
            running_task_count: 0,
        },
    );
    let probe = snapshot.probe_json();
    let rendered = serde_json::to_string(&probe).expect("serialize tray probe");

    assert!(rendered.contains("Office MCP"));
    assert!(rendered.contains("Show Office MCP"));
    assert!(rendered.contains("Quit Office MCP"));
    assert!(rendered.contains("Keep Running"));
    assert!(!rendered.to_ascii_lowercase().contains("debug"));
    assert!(!rendered.to_ascii_lowercase().contains("prototype"));
    assert!(!rendered.to_ascii_lowercase().contains("placeholder"));
    assert!(!rendered.to_ascii_lowercase().contains("test tray"));
    assert_eq!(
        probe["tooltip"],
        "Office MCP - Up - 0 clients - 0 documents"
    );
    assert_eq!(probe["menu"][0]["kind"], "read_only");
    assert_eq!(probe["menu"][0]["enabled"], false);
    assert_eq!(probe["menu"][3]["kind"], "separator");
    assert_eq!(probe["menu"][4]["kind"], "action");
    assert_eq!(probe["menu"][4]["enabled"], true);
    assert_eq!(probe["menu"][5]["kind"], "action");
    assert_eq!(probe["menu"][5]["enabled"], true);
    assert_eq!(probe["quit_confirmation"]["title"], "Quit Office MCP");
    assert_eq!(
        probe["quit_confirmation"]["primary_action"],
        "Quit Office MCP"
    );
    assert_eq!(
        probe["quit_confirmation"]["secondary_action"],
        "Keep Running"
    );
}

#[test]
fn down_status_defaults_to_zero_counts() {
    let snapshot =
        TraySnapshot::from_input(TrayPlatform::LinuxStatusNotifier, TrayStatusInput::down());

    assert_eq!(snapshot.menu[0].label(), Some("Status: Down"));
    assert_eq!(snapshot.menu[1].label(), Some("Clients: 0"));
    assert_eq!(snapshot.menu[2].label(), Some("Documents: 0"));
}
