use super::{CommandFailure, CommandResult, UiCommandStatus, UiStateOptions};
use crate::addin_mgr::PartialEffect;
use crate::mcp::AccessMode;
use std::time::SystemTime;

#[test]
fn default_options_match_local_daemon_ui_contract() {
    let options = UiStateOptions::default();

    assert_eq!(options.version, "0.1.0");
    assert_eq!(options.mcp_endpoint, "http://127.0.0.1:8800/mcp");
    assert_eq!(options.addin_endpoint, "https://localhost:8765/addin");
    let policy = options.tool_access_policy.snapshot();
    assert_eq!(policy.access_mode, AccessMode::All);
    assert!(policy.disabled_tools.is_empty());
    assert_eq!(options.now, SystemTime::UNIX_EPOCH);
}

#[test]
fn command_result_maps_timeout_cancel_and_thrown_statuses() {
    assert_eq!(
        CommandResult::Failure(CommandFailure {
            office_mcp_code: "TIMEOUT".to_string(),
            message: "body=secret".to_string(),
            tool: Some("word.get_text".to_string()),
            retriable: true,
            partial_effect: Some(PartialEffect::None),
            debug: None,
        })
        .into_status()
        .status,
        UiCommandStatus::Timeout
    );
    assert_eq!(
        CommandResult::Failure(CommandFailure {
            office_mcp_code: "CANCELLED".to_string(),
            message: "cancelled".to_string(),
            tool: None,
            retriable: false,
            partial_effect: None,
            debug: None,
        })
        .into_status()
        .status,
        UiCommandStatus::Cancelled
    );
    let thrown = CommandResult::Thrown("password=open-sesame".to_string()).into_status();
    assert_eq!(thrown.status, UiCommandStatus::Failure);
    assert_eq!(
        thrown.error.expect("thrown error").message,
        "password=[redacted]"
    );
}
