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
#[path = "parity_tests.rs"]
mod parity_tests;
