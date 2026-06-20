use super::OfficeMcpDaemon;
use std::collections::BTreeSet;

#[test]
fn daemon_owns_all_required_domain_objects() {
    let daemon = OfficeMcpDaemon::new();
    let components = daemon.component_descriptions();
    let names = components
        .iter()
        .map(super::ComponentDescription::name)
        .collect::<Vec<_>>();

    assert_eq!(names.len(), 10);
    assert!(names.contains(&"DaemonConfigService"));
    assert!(names.contains(&"McpHttpFrontend"));
    assert!(names.contains(&"AddinChannelServer"));
    assert!(names.contains(&"SessionRegistry"));
    assert!(names.contains(&"CommandRouter"));
    assert!(names.contains(&"UiStateStore"));
    assert!(names.contains(&"TrayController"));
    assert!(names.contains(&"AuditLog"));
    assert!(names.contains(&"Logger"));
    assert!(names.contains(&"RuntimeServer"));
}

#[test]
fn daemon_tracks_rust_runtime_readiness_gates() {
    let daemon = OfficeMcpDaemon::new();

    assert_eq!(daemon.parity_plan().gates().len(), 8);
}

#[test]
fn daemon_src_root_only_contains_composition_and_transitional_files() {
    let src_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("src");
    let allowed_root_files = BTreeSet::from(["daemon.rs", "lib.rs", "main.rs", "parity.rs"]);
    let forbidden_service_files = BTreeSet::from([
        "addin_channel.rs",
        "audit_log.rs",
        "client_config.rs",
        "command_router.rs",
        "config_service.rs",
        "daemon_control.rs",
        "image_fetcher.rs",
        "logger.rs",
        "mcp_http_frontend.rs",
        "mcp_management_client.rs",
        "session_registry.rs",
        "state_store.rs",
        "stdio_bridge.rs",
    ]);

    let root_files = std::fs::read_dir(&src_dir)
        .expect("read daemon src dir")
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().is_ok_and(|file_type| file_type.is_file()))
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect::<BTreeSet<_>>();

    for file in &forbidden_service_files {
        assert!(
            !root_files.contains(*file),
            "service module {file} must stay under its functional module directory"
        );
    }

    for file in &root_files {
        if file.ends_with("_tests.rs") {
            continue;
        }

        assert!(
            allowed_root_files.contains(file.as_str()),
            "unexpected daemon src root file {file}; place service code under common, ui, api, mcp, addin_mgr, or tray"
        );
    }
}
