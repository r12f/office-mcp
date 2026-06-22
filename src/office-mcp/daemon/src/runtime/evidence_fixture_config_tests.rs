use super::{UiFixtureState, default_addin_public_dir, fixture_config};
use std::path::PathBuf;

#[test]
fn fixture_state_defaults_to_seeded_unless_empty_is_requested() {
    assert_eq!(UiFixtureState::from_value(None), UiFixtureState::Seeded);
    assert_eq!(
        UiFixtureState::from_value(Some("seeded")),
        UiFixtureState::Seeded
    );
    assert_eq!(
        UiFixtureState::from_value(Some("empty")),
        UiFixtureState::Empty
    );
    assert_eq!(
        UiFixtureState::from_value(Some("EMPTY")),
        UiFixtureState::Empty
    );
}

#[test]
fn fixture_config_uses_allocated_ports_and_certificate() {
    let config = fixture_config(
        8801,
        8766,
        &PathBuf::from("C:/certs/localhost.pfx"),
        "passphrase".to_string(),
    );

    assert_eq!(config.mcp.port, 8801);
    assert_eq!(config.addin.port, 8766);
    assert_eq!(config.addin.origin, "https://localhost:8766");
    assert_eq!(config.addin.pfx_path, "C:/certs/localhost.pfx");
    assert_eq!(config.addin.pfx_passphrase, "passphrase");
    assert!(config.logging.file.ends_with("office-mcp-ui-fixture.log"));
    assert!(!config.audit.enabled);
}

#[test]
fn default_addin_public_dir_has_repo_relative_fallback() {
    let path = default_addin_public_dir();

    assert!(path.ends_with("src/office-ctl/word/public") || path.is_dir());
}
