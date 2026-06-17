use super::{UiSnapshotEndpoints, UiSnapshotService};
use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, RuntimeInfo, SessionRegistry,
};
use crate::api::{UiStateOptions, UiStateStore};
use crate::common::{Logger, LoggerLogLevel};
use std::fs::{read_to_string, remove_dir_all};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

#[test]
fn service_renders_runtime_endpoints_and_registry_sessions() {
    let ui_state = Arc::new(Mutex::new(UiStateStore::with_options(UiStateOptions {
        mcp_endpoint: "http://stale/mcp".to_string(),
        addin_endpoint: "https://stale/addin".to_string(),
        now: SystemTime::UNIX_EPOCH,
        ..UiStateOptions::default()
    })));
    let registry = Arc::new(Mutex::new(registry_with_session()));
    let endpoints = UiSnapshotEndpoints {
        mcp_endpoint: "http://127.0.0.1:8800/mcp".to_string(),
        addin_endpoint: "https://localhost:8765/addin".to_string(),
    };

    let text = UiSnapshotService::new().render_runtime_snapshot(&ui_state, &registry, &endpoints);
    let rendered: serde_json::Value = serde_json::from_str(&text).expect("snapshot json");

    assert_eq!(rendered["daemon"]["mcp_endpoint"], endpoints.mcp_endpoint);
    assert_eq!(
        rendered["daemon"]["addin_endpoint"],
        endpoints.addin_endpoint
    );
    assert_eq!(rendered["documents"]["word"][0]["session_id"], "session-1");
}

#[test]
fn writes_structured_tracing_event_for_ui_api_snapshot() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-ui-snapshot-service-log-{}",
        std::process::id()
    ));
    let path = dir.join("office-mcp.log");
    let (subscriber, guard) =
        Logger::tracing_file_default(LoggerLogLevel::Debug, &path).expect("init tracing");
    let ui_state = Arc::new(Mutex::new(UiStateStore::with_options(UiStateOptions {
        mcp_endpoint: "http://stale/mcp".to_string(),
        addin_endpoint: "https://stale/addin".to_string(),
        now: SystemTime::UNIX_EPOCH,
        ..UiStateOptions::default()
    })));
    let registry = Arc::new(Mutex::new(registry_with_session()));
    let endpoints = UiSnapshotEndpoints {
        mcp_endpoint: "http://127.0.0.1:8800/mcp".to_string(),
        addin_endpoint: "https://localhost:8765/addin".to_string(),
    };

    tracing::subscriber::with_default(subscriber, || {
        let text =
            UiSnapshotService::new().render_runtime_snapshot(&ui_state, &registry, &endpoints);
        assert!(text.contains("session-1"));
    });
    drop(guard);

    let contents = read_to_string(&path).expect("read tracing log file");
    assert!(contents.contains("rendered UI API snapshot"));
    assert!(contents.contains("\"component\":\"ui_snapshot_service\""));
    assert!(contents.contains("\"documents\":1"));
    assert!(contents.contains("\"ui_state_lock_failed\":false"));
    let _ = remove_dir_all(dir);
}

fn registry_with_session() -> SessionRegistry {
    let mut registry = SessionRegistry::new();
    registry.register_runtime(RuntimeInfo {
        instance_id: "instance-1".to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: None,
            platform: None,
            build: None,
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: Vec::new(),
        },
        registered_at: SystemTime::UNIX_EPOCH,
    });
    registry.add_session(
        NewSessionInfo {
            session_id: "session-1".to_string(),
            instance_id: "instance-1".to_string(),
            document: DocumentInfo {
                title: Some("Doc.docx".to_string()),
                url: None,
                filename: Some("Doc.docx".to_string()),
                is_dirty: None,
                is_read_only: None,
                is_protected: None,
                protection: None,
            },
            available_tools: Vec::new(),
            is_active: Some(true),
        },
        SystemTime::UNIX_EPOCH,
    );
    registry
}
