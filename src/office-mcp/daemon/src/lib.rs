pub mod addin_mgr;
pub mod api;
pub mod client_config;
pub mod common;
pub mod daemon;
pub mod daemon_control;
pub mod evidence_fixture;
pub mod mcp;
pub mod parity;
pub mod runtime_server;
pub mod tray;
pub mod ui;

pub use daemon::OfficeMcpDaemon;
pub use parity::{ParityGate, ParityPlan};
