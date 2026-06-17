use super::ClaudeDesktopConfigBuilder;
use std::collections::BTreeMap;
use std::path::PathBuf;

#[test]
fn development_config_uses_rust_stdio_command() {
    let json = ClaudeDesktopConfigBuilder::development()
        .with_current_exe(PathBuf::from("C:\\office-mcp\\office-mcp-daemon.exe"))
        .to_json();

    assert!(json.contains("office-mcp-daemon.exe"));
    assert!(json.contains("\"stdio\""));
    assert!(!json.contains("dist/src/cli.js"));
    assert!(!json.contains("node"));
}

#[test]
fn development_config_omits_legacy_reference_node_paths() {
    let json = ClaudeDesktopConfigBuilder::development()
        .with_current_exe(PathBuf::from(
            "C:\\Code\\office-mcp\\target\\debug\\office-mcp-daemon.exe",
        ))
        .to_json();

    assert!(!json.contains("reference-node"));
    assert!(!json.contains("dist/src/cli.js"));
    assert!(!json.contains("cargo run"));
}

#[test]
fn installed_config_uses_launcher() {
    let json = ClaudeDesktopConfigBuilder::installed(Some(PathBuf::from("D:\\Apps\\office-mcp")))
        .to_json();

    assert!(json.contains("powershell.exe"));
    assert!(json.contains("D:\\\\Apps\\\\office-mcp\\\\office-mcp.ps1"));
    assert!(json.contains("\"stdio\""));
}

#[test]
fn installed_config_without_root_uses_office_mcp_install_root() {
    let json = ClaudeDesktopConfigBuilder::installed(None)
        .with_env(BTreeMap::from([(
            "OFFICE_MCP_INSTALL_ROOT".to_string(),
            "E:\\OfficeMcp".to_string(),
        )]))
        .to_json();

    assert!(json.contains("E:\\\\OfficeMcp\\\\office-mcp.ps1"));
}
