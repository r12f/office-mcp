pub mod addin_channel;
pub mod command_router;
pub mod connection_hub;
pub mod image_fetcher;
pub mod session_registry;
pub mod websocket_codec;
pub mod websocket_handshake;

pub use addin_channel::{
    AddinChannelConfig, AddinChannelError, AddinChannelServer, HeartbeatDecision, JsonRpcEnvelope,
    JsonRpcId, RegisterRequest, RegisterResult, SessionAddedEvent, SessionRemovedEvent,
    SessionRemovedReason, SessionUpdatedEvent,
};
pub use command_router::{
    CancelCommand, CommandRouter, CommandRouterError, QueuedCommand, ToolCallRequest, ToolResponse,
};
pub use connection_hub::{AddinConnectionHub, AddinConnectionHubError};
pub use image_fetcher::{FetchedImage, ImageFetchError, ImageFetcher};
pub use session_registry::{
    AddInInfo, DocumentDescriptor, DocumentInfo, HostDescriptor, HostInfo, InvocationPermit,
    NewSessionInfo, OfficeMcpCode, PartialEffect, ProtectionInfo, RegistrationOutcome, RuntimeInfo,
    SessionDescriptor, SessionDetails, SessionInfo, SessionPatch, SessionRegistry, SessionStatus,
    ToolFailure, ToolInvocationError,
};
pub use websocket_codec::{
    WebSocketCodec, WebSocketCodecError, WebSocketFrame, WebSocketProtocolError,
};
pub use websocket_handshake::websocket_accept_key;
