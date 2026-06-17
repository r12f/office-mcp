pub mod addin_channel;
pub mod addin_heartbeat;
pub mod addin_json_rpc;
pub mod command_model;
pub(crate) mod command_queue;
pub mod command_router;
pub mod command_router_error;
pub mod connection_hub;
pub mod image_fetcher;
pub mod session_descriptor_view;
pub mod session_registry;
pub mod static_assets;
pub mod websocket_codec;
pub mod websocket_handshake;

pub use addin_channel::{
    AddinChannelConfig, AddinChannelError, AddinChannelServer, HeartbeatDecision, RegisterRequest,
    SessionAddedEvent, SessionRemovedEvent, SessionRemovedReason, SessionUpdatedEvent,
};
pub use addin_json_rpc::{JsonRpcEnvelope, JsonRpcId, RegisterResult};
pub use command_model::{CancelCommand, QueuedCommand, ToolCallRequest, ToolResponse};
pub(crate) use command_queue::SessionCommandQueue;
pub use command_router::CommandRouter;
pub use command_router_error::CommandRouterError;
pub use connection_hub::{AddinConnectionHub, AddinConnectionHubError};
pub use image_fetcher::{FetchedImage, ImageFetchError, ImageFetcher};
pub use session_descriptor_view::SessionDescriptorView;
pub use session_registry::{
    AddInInfo, DocumentDescriptor, DocumentInfo, HostDescriptor, HostInfo, InvocationPermit,
    NewSessionInfo, OfficeMcpCode, PartialEffect, ProtectionInfo, RegistrationOutcome, RuntimeInfo,
    SessionDescriptor, SessionDetails, SessionInfo, SessionPatch, SessionRegistry, SessionStatus,
    ToolFailure, ToolInvocationError,
};
pub use static_assets::{
    default_addin_public_dir, default_office_ctl_common_dir, default_office_ctl_host_public_dir,
    find_addin_public_dir_from, static_asset_content_type,
};
pub use websocket_codec::{
    WebSocketCodec, WebSocketCodecError, WebSocketFrame, WebSocketProtocolError,
};
pub use websocket_handshake::websocket_accept_key;
