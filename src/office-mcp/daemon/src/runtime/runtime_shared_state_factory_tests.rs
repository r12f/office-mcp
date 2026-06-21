use super::RuntimeSharedStateFactory;
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, OfficeMcpCode, RuntimeInfo, SessionRegistry,
};
use crate::mcp::{AccessMode, ToolAccessPolicy};
use crate::runtime::RuntimeServerConfig;
use std::time::{Duration, SystemTime};

#[test]
fn factory_applies_addin_channel_origin_from_runtime_config() {
    let config = RuntimeServerConfig {
        addin_origin: "https://localhost:9001".to_string(),
        ..RuntimeServerConfig::default()
    };
    let shared_state = RuntimeSharedStateFactory::with_registry(
        &config,
        SessionRegistry::with_limits(config.max_pending_per_session),
    );
    let addin_channel = shared_state
        .addin_channel
        .lock()
        .expect("addin channel lock");

    assert!(
        addin_channel
            .validate_upgrade("/addin", Some("https://localhost:9001"))
            .is_ok()
    );
    assert!(
        addin_channel
            .validate_upgrade("/addin", Some("https://localhost:8765"))
            .is_err()
    );
}

#[test]
fn factory_applies_registry_pending_limit_from_runtime_config() {
    let config = RuntimeServerConfig {
        max_pending_per_session: 1,
        ..RuntimeServerConfig::default()
    };
    let shared_state = RuntimeSharedStateFactory::with_registry(
        &config,
        SessionRegistry::with_limits(config.max_pending_per_session),
    );
    let mut registry = shared_state.registry.lock().expect("registry lock");
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    registry.register_runtime(runtime(now));
    registry.add_session(session(), now);
    registry.set_connection_pending("instance-1", 1);

    let error = registry
        .prepare_invocation("session-1", "word.get_text", true)
        .expect_err("pending limit reached");

    assert_eq!(
        error.failure.office_mcp_code,
        OfficeMcpCode::MaxPendingExceeded
    );
}

#[test]
fn shared_state_prunes_stale_sessions_after_configured_grace() {
    let config = RuntimeServerConfig {
        session_grace: Duration::from_secs(300),
        ..RuntimeServerConfig::default()
    };
    let shared_state = RuntimeSharedStateFactory::with_registry(
        &config,
        SessionRegistry::with_limits(config.max_pending_per_session),
    );
    let stale_since = SystemTime::UNIX_EPOCH + Duration::from_secs(10);
    {
        let mut registry = shared_state.registry.lock().expect("registry lock");
        registry.register_runtime(runtime(stale_since));
        registry.add_session(session(), stale_since);
        assert!(registry.remove_runtime("instance-1", stale_since));
    }

    assert_eq!(
        shared_state.prune_stale_sessions(stale_since + Duration::from_secs(299)),
        0
    );
    assert!(
        shared_state
            .registry
            .lock()
            .expect("registry lock")
            .get_session_info("session-1")
            .is_some()
    );
    assert_eq!(
        shared_state.prune_stale_sessions(stale_since + Duration::from_secs(301)),
        1
    );
    assert!(
        shared_state
            .registry
            .lock()
            .expect("registry lock")
            .get_session_info("session-1")
            .is_none()
    );
}

#[test]
fn factory_applies_runtime_tool_access_policy() {
    let config = RuntimeServerConfig {
        tool_access_policy: ToolAccessPolicy::default().with_access_mode(AccessMode::Read),
        ..RuntimeServerConfig::default()
    };
    let shared_state = RuntimeSharedStateFactory::with_registry(
        &config,
        SessionRegistry::with_limits(config.max_pending_per_session),
    );

    let policy = shared_state.tool_access_policy();

    assert!(policy.allows_tool("word.get_text"));
    assert!(!policy.allows_tool("word.insert_paragraph"));
}

fn runtime(registered_at: SystemTime) -> RuntimeInfo {
    RuntimeInfo {
        instance_id: "instance-1".to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["doc.read".to_string()],
        },
        registered_at,
    }
}

fn session() -> NewSessionInfo {
    NewSessionInfo {
        session_id: "session-1".to_string(),
        instance_id: "instance-1".to_string(),
        document: DocumentInfo::default(),
        available_tools: vec!["word.get_text".to_string()],
        is_active: Some(true),
    }
}
