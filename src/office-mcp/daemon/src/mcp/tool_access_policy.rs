use crate::mcp::{AccessMode, UiToolAccessPolicySnapshot, tool_metadata};
use std::collections::BTreeSet;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ToolAccessPolicy {
    access_mode: AccessMode,
    disabled_apps: BTreeSet<String>,
    disabled_categories: BTreeSet<(String, String)>,
    disabled_tools: BTreeSet<String>,
}

impl ToolAccessPolicy {
    #[must_use]
    pub fn with_access_mode(mut self, mode: AccessMode) -> Self {
        self.access_mode = mode;
        self
    }

    #[must_use]
    pub fn with_disabled_app(mut self, app: &str) -> Self {
        self.disabled_apps.insert(app.to_string());
        self
    }

    #[must_use]
    pub fn with_disabled_category(mut self, app: &str, category: &str) -> Self {
        self.disabled_categories
            .insert((app.to_string(), category.to_string()));
        self
    }

    #[must_use]
    pub fn with_disabled_tool(mut self, tool: &str) -> Self {
        self.disabled_tools.insert(tool.to_string());
        self
    }

    #[must_use]
    pub fn allows_tool(&self, tool: &str) -> bool {
        if self.disabled_tools.contains(tool) {
            return false;
        }

        let Some(metadata) = tool_metadata(tool) else {
            return true;
        };

        self.access_mode.allows(metadata.side_effect)
            && !self.disabled_apps.contains(metadata.app)
            && !self
                .disabled_categories
                .contains(&(metadata.app.to_string(), metadata.category.to_string()))
    }

    #[must_use]
    pub fn snapshot(&self) -> UiToolAccessPolicySnapshot {
        UiToolAccessPolicySnapshot {
            access_mode: self.access_mode,
            disabled_apps: self.disabled_apps.iter().cloned().collect(),
            disabled_categories: self.disabled_categories.iter().cloned().collect(),
            disabled_tools: self.disabled_tools.iter().cloned().collect(),
        }
    }
}

#[cfg(test)]
#[path = "tool_access_policy_tests.rs"]
mod tool_access_policy_tests;
