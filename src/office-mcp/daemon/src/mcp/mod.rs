pub mod catalog;
pub mod http_frontend;
pub mod http_frontend_model;
pub mod management_client;
pub mod prompt_catalog;
pub mod resource_request;
pub mod stdio_bridge;
pub mod tool_access_policy;
pub mod tool_metadata;
pub mod tool_result;

pub use catalog::{
    ExcelToolCatalog, PowerPointToolCatalog, WORD_V1_TOOLS, tool_catalog_json,
    tool_catalog_json_for_policy, word_resource_catalog_for_session, word_resource_templates,
};
pub use http_frontend::McpHttpFrontend;
pub(crate) use http_frontend_model::RateLimitWindow;
pub use http_frontend_model::{
    HttpMethod, McpClientSession, McpHttpConfig, McpHttpDecision, McpHttpError, McpHttpRequest,
};
pub use management_client::McpManagementClient;
pub use prompt_catalog::{prompt_catalog_json, prompt_description, prompt_messages};
pub use resource_request::{ResourceReadRequest, resource_request_from_uri};
pub use stdio_bridge::{StdioBridge, StdioBridgeError};
pub use tool_access_policy::ToolAccessPolicy;
pub use tool_metadata::{
    AccessMode, ToolMetadata, ToolSideEffect, UiToolAccessPolicySnapshot, tool_metadata,
};
pub use tool_result::{
    tool_failure, tool_failure_from_command, tool_not_available_by_policy, tool_success,
};
