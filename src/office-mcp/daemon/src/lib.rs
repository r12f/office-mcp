pub mod addin_channel;
pub mod audit_log;
pub mod client_config;
pub mod command_router;
pub mod config_service;
pub mod daemon;
pub mod daemon_control;
pub mod evidence_fixture;
pub mod image_fetcher;
pub mod logger;
pub mod mcp;
pub mod parity;
pub mod runtime_server;
pub mod session_registry;
pub mod tray;
pub mod ui;

pub use daemon::OfficeMcpDaemon;
pub use parity::{ParityGate, ParityPlan};
