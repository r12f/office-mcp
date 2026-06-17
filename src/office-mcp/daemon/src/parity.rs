#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParityGate {
    McpTransport,
    AddinRegistration,
    WordToolSurface,
    ErrorModel,
    UiStateRedaction,
    RuntimeEvidence,
    TrayEvidence,
    PackagingSmoke,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParityPlan {
    gates: Vec<ParityGate>,
}

impl ParityPlan {
    #[must_use]
    pub fn rust_runtime_readiness() -> Self {
        Self {
            gates: vec![
                ParityGate::McpTransport,
                ParityGate::AddinRegistration,
                ParityGate::WordToolSurface,
                ParityGate::ErrorModel,
                ParityGate::UiStateRedaction,
                ParityGate::RuntimeEvidence,
                ParityGate::TrayEvidence,
                ParityGate::PackagingSmoke,
            ],
        }
    }

    #[must_use]
    pub fn gates(&self) -> &[ParityGate] {
        &self.gates
    }
}

#[cfg(test)]
mod tests {
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
}
