use std::collections::BTreeSet;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ToolAccessPolicy {
    disabled_tools: BTreeSet<String>,
}

impl ToolAccessPolicy {
    #[must_use]
    pub fn with_disabled_tool(mut self, tool: &str) -> Self {
        self.disabled_tools.insert(tool.to_string());
        self
    }

    #[must_use]
    pub fn allows_tool(&self, tool: &str) -> bool {
        !self.disabled_tools.contains(tool)
    }
}

#[cfg(test)]
#[path = "tool_access_policy_tests.rs"]
mod tool_access_policy_tests;
