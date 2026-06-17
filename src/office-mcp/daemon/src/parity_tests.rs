use super::{ParityGate, ParityPlan};

#[test]
fn rust_runtime_readiness_covers_protocol_runtime_ui_and_packaging() {
    let plan = ParityPlan::rust_runtime_readiness();

    assert!(plan.gates().contains(&ParityGate::McpTransport));
    assert!(plan.gates().contains(&ParityGate::AddinRegistration));
    assert!(plan.gates().contains(&ParityGate::RuntimeEvidence));
    assert!(plan.gates().contains(&ParityGate::UiStateRedaction));
    assert!(plan.gates().contains(&ParityGate::TrayEvidence));
    assert!(plan.gates().contains(&ParityGate::PackagingSmoke));
}
