use super::ConfigPathResolver;
use std::collections::BTreeMap;

#[test]
fn config_path_uses_platform_config_root() {
    let env = env_map(&[
        ("APPDATA", "C:\\Users\\Ada\\AppData\\Roaming"),
        ("HOME", "/Users/ada"),
        ("XDG_CONFIG_HOME", "/home/ada/.config"),
    ]);
    let path = ConfigPathResolver::new(&env).config_path();

    assert!(
        path.to_string_lossy()
            .replace('\\', "/")
            .ends_with("office-mcp/config.toml")
    );
}

#[test]
fn state_paths_use_platform_state_root() {
    let env = env_map(&[
        ("LOCALAPPDATA", "C:\\Users\\Ada\\AppData\\Local"),
        ("HOME", "/Users/ada"),
        ("XDG_STATE_HOME", "/home/ada/.local/state"),
    ]);
    let resolver = ConfigPathResolver::new(&env);

    assert!(
        resolver
            .audit_path()
            .replace('\\', "/")
            .ends_with("office-mcp/audit.jsonl")
    );
    assert!(
        resolver
            .log_path()
            .replace('\\', "/")
            .ends_with("office-mcp/office-mcp.log")
    );
}

#[test]
fn pfx_path_uses_current_working_directory() {
    let path = ConfigPathResolver::pfx_path();

    assert!(path.ends_with(".office-mcp-localhost.pfx"));
}

fn env_map(values: &[(&str, &str)]) -> BTreeMap<String, String> {
    values
        .iter()
        .map(|(key, value)| ((*key).to_string(), (*value).to_string()))
        .collect()
}
