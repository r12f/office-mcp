pub mod catalog;
pub mod http_frontend;
pub mod management_client;
pub mod resource_request;
pub mod stdio_bridge;

pub use catalog::{
    ExcelToolCatalog, WORD_V1_TOOLS, prompt_catalog_json, prompt_description, prompt_messages,
    tool_catalog_json, word_resource_catalog_for_session, word_resource_templates,
};
pub use http_frontend::{
    HttpMethod, McpHttpConfig, McpHttpDecision, McpHttpFrontend, McpHttpRequest,
};
pub use management_client::McpManagementClient;
pub use resource_request::{ResourceReadRequest, resource_request_from_uri};
pub use stdio_bridge::{StdioBridge, StdioBridgeError};
