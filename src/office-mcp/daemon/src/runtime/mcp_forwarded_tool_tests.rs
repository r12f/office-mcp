use super::McpForwardedToolInvoker;
use crate::addin_mgr::{
    AddinChannelServer, AddinConnectionHub, CommandRouter, ImageFetcher, SessionRegistry,
};
use crate::api::UiStateStore;
use crate::common::AuditLog;
use crate::runtime::mcp_rpc::McpDispatchContext;
use serde_json::json;
use std::sync::{Arc, Mutex};

#[test]
fn forwarded_tools_require_session_id() {
    let registry = SessionRegistry::new();
    let mut ui_state = UiStateStore::new();
    let addin_channel = Arc::new(Mutex::new(AddinChannelServer::new()));
    let connection_hub = Arc::new(AddinConnectionHub::new());
    let command_router = Arc::new(Mutex::new(CommandRouter::new()));
    let audit_log = AuditLog::new();
    let image_fetcher = ImageFetcher::new();
    let mut context = McpDispatchContext {
        registry: &registry,
        ui_state: &mut ui_state,
        addin_channel: &addin_channel,
        connection_hub: &connection_hub,
        command_router: &command_router,
        audit_log: &audit_log,
        image_fetcher: &image_fetcher,
    };

    let result = McpForwardedToolInvoker::call(
        &mut context,
        &json!({ "id": "call-1" }),
        "word.get_text",
        &json!({}),
        true,
    );

    assert_eq!(
        result["structuredContent"]["error"]["office_mcp_code"],
        "INVALID_ARGUMENTS"
    );
    assert_eq!(result["isError"], true);
}
