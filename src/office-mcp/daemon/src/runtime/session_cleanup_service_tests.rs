use super::SessionCleanupService;
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionRegistry,
};
use crate::runtime::RuntimeServerConfig;
use crate::runtime::runtime_shared_state_factory::RuntimeSharedStateFactory;
use std::time::{Duration, SystemTime};

#[test]
fn run_once_prunes_expired_stale_sessions_without_client_request() {
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

    let service = SessionCleanupService::for_session_grace(config.session_grace);

    assert_eq!(
        service.run_once(&shared_state, stale_since + Duration::from_secs(299)),
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
        service.run_once(&shared_state, stale_since + Duration::from_secs(301)),
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
        document: DocumentInfo {
            filename: Some("Doc.docx".to_string()),
            title: Some("Doc.docx".to_string()),
            ..DocumentInfo::default()
        },
        available_tools: vec!["word.get_text".to_string()],
        is_active: Some(true),
    }
}
