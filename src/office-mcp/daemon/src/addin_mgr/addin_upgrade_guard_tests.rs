use super::AddinUpgradeGuard;
use crate::addin_mgr::{AddinChannelConfig, AddinChannelError};

#[test]
fn guard_accepts_exact_addin_path_and_origin() {
    let guard = AddinUpgradeGuard::new(&AddinChannelConfig::default());

    assert!(
        guard
            .validate("/addin", Some("https://localhost:8765"))
            .is_ok()
    );
}

#[test]
fn guard_rejects_wrong_path_and_foreign_origin() {
    let guard = AddinUpgradeGuard::new(&AddinChannelConfig::default());

    assert!(matches!(
        guard.validate("/wrong", Some("https://localhost:8765")),
        Err(AddinChannelError::InvalidUpgradePath(_))
    ));
    assert!(matches!(
        guard.validate("/addin", Some("https://example.invalid")),
        Err(AddinChannelError::ForbiddenOrigin(_))
    ));
}

#[test]
fn guard_uses_configured_origin() {
    let guard = AddinUpgradeGuard::new(&AddinChannelConfig {
        origin: "https://localhost:9443".to_string(),
        ..AddinChannelConfig::default()
    });

    assert!(
        guard
            .validate("/addin", Some("https://localhost:9443"))
            .is_ok()
    );
    assert!(matches!(
        guard.validate("/addin", Some("https://localhost:8765")),
        Err(AddinChannelError::ForbiddenOrigin(_))
    ));
}
