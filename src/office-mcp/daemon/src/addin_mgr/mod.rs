pub mod addin_channel;
pub mod addin_channel_config;
pub mod addin_channel_error;
pub mod addin_channel_model;
pub(crate) mod addin_connection_state;
pub mod addin_heartbeat;
pub mod addin_json_rpc;
pub mod command_model;
pub(crate) mod command_queue;
pub mod command_router;
pub mod command_router_error;
pub mod connection_hub;
pub mod document_info;
pub mod image_fetcher;
pub mod session_descriptor_view;
pub mod session_info;
pub mod session_registry;
pub mod static_assets;
pub mod websocket_codec;
pub mod websocket_handshake;

pub use addin_channel::AddinChannelServer;
pub use addin_channel_config::AddinChannelConfig;
pub use addin_channel_error::AddinChannelError;
pub use addin_channel_model::{
    HeartbeatDecision, RegisterRequest, SessionAddedEvent, SessionRemovedEvent,
    SessionRemovedReason, SessionUpdatedEvent,
};
pub(crate) use addin_connection_state::AddinConnectionState;
pub use addin_json_rpc::{JsonRpcEnvelope, JsonRpcId, RegisterResult};
pub use command_model::{CancelCommand, QueuedCommand, ToolCallRequest, ToolResponse};
pub(crate) use command_queue::SessionCommandQueue;
pub use command_router::CommandRouter;
pub use command_router_error::CommandRouterError;
pub use connection_hub::{AddinConnectionHub, AddinConnectionHubError};
pub use document_info::{DocumentDescriptor, DocumentInfo, ProtectionInfo};
pub use image_fetcher::{FetchedImage, ImageFetchError, ImageFetcher};
pub use session_descriptor_view::SessionDescriptorView;
pub use session_info::SessionInfo;
pub use session_registry::{
    AddInInfo, HostDescriptor, HostInfo, InvocationPermit, NewSessionInfo, OfficeMcpCode,
    PartialEffect, RegistrationOutcome, RuntimeInfo, SessionDescriptor, SessionDetails,
    SessionPatch, SessionRegistry, SessionStatus, ToolFailure, ToolInvocationError,
};
pub use static_assets::{
    default_addin_public_dir, default_office_ctl_common_dir, default_office_ctl_host_public_dir,
    find_addin_public_dir_from, static_asset_content_type,
};
pub use websocket_codec::{
    WebSocketCodec, WebSocketCodecError, WebSocketFrame, WebSocketProtocolError,
};
pub use websocket_handshake::websocket_accept_key;
