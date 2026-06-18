#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrayPlatform {
    WindowsNotificationArea,
    MacOsMenuBar,
    LinuxStatusNotifier,
    Unsupported,
}

impl TrayPlatform {
    #[must_use]
    pub const fn detect() -> Self {
        if cfg!(windows) {
            Self::WindowsNotificationArea
        } else if cfg!(target_os = "macos") {
            Self::MacOsMenuBar
        } else if cfg!(target_os = "linux") {
            Self::LinuxStatusNotifier
        } else {
            Self::Unsupported
        }
    }

    #[must_use]
    pub const fn label(self) -> &'static str {
        match self {
            Self::WindowsNotificationArea => "windows-notification-area",
            Self::MacOsMenuBar => "macos-menu-bar",
            Self::LinuxStatusNotifier => "linux-status-notifier",
            Self::Unsupported => "unsupported",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrayHealth {
    Up,
    Degraded,
    Down,
}

impl TrayHealth {
    #[must_use]
    pub const fn label(self) -> &'static str {
        match self {
            Self::Up => "Up",
            Self::Degraded => "Degraded",
            Self::Down => "Down",
        }
    }

    fn from_json(value: Option<&serde_json::Value>) -> Self {
        match value.and_then(serde_json::Value::as_str) {
            Some(value) if value.eq_ignore_ascii_case("up") => Self::Up,
            Some(value) if value.eq_ignore_ascii_case("degraded") => Self::Degraded,
            _ => Self::Down,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TrayStatusInput {
    pub health: TrayHealth,
    pub client_count: usize,
    pub document_count: usize,
    pub running_task_count: usize,
}

impl TrayStatusInput {
    #[must_use]
    pub const fn down() -> Self {
        Self {
            health: TrayHealth::Down,
            client_count: 0,
            document_count: 0,
            running_task_count: 0,
        }
    }

    #[must_use]
    pub fn from_ui_state(ui_state: Option<&serde_json::Value>) -> Self {
        let Some(value) = ui_state else {
            return Self::down();
        };
        Self {
            health: TrayHealth::from_json(value.pointer("/daemon/status")),
            client_count: value
                .get("clients")
                .and_then(serde_json::Value::as_array)
                .map_or(0, Vec::len),
            document_count: document_count(value.get("documents")),
            running_task_count: value
                .get("current_tasks")
                .and_then(serde_json::Value::as_array)
                .map_or(0, Vec::len),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TraySnapshot {
    pub platform: TrayPlatform,
    pub tooltip: String,
    pub menu: Vec<TrayMenuItem>,
    pub quit_confirmation: QuitConfirmation,
}

impl TraySnapshot {
    #[must_use]
    pub fn from_input(platform: TrayPlatform, input: TrayStatusInput) -> Self {
        let tooltip = format!(
            "Office MCP Control - {} - {} clients - {} documents",
            input.health.label(),
            input.client_count,
            input.document_count
        );
        Self {
            platform,
            tooltip,
            menu: vec![
                TrayMenuItem::read_only(format!("Status: {}", input.health.label())),
                TrayMenuItem::read_only(format!("Clients: {}", input.client_count)),
                TrayMenuItem::read_only(format!("Documents: {}", input.document_count)),
                TrayMenuItem::Separator,
                TrayMenuItem::action(TrayAction::ShowUi, "Show Office MCP Control"),
                TrayMenuItem::action(TrayAction::Quit, "Quit Office MCP Control"),
            ],
            quit_confirmation: QuitConfirmation::from_input(&input),
        }
    }

    #[must_use]
    pub fn probe_json(&self) -> serde_json::Value {
        serde_json::json!({
            "platform": self.platform.label(),
            "tooltip": self.tooltip,
            "menu_items": self.menu.iter().map(TrayMenuItem::probe_label).collect::<Vec<_>>(),
            "menu": self.menu.iter().map(TrayMenuItem::probe_item_json).collect::<Vec<_>>(),
            "quit_confirmation": {
                "title": self.quit_confirmation.title,
                "body": self.quit_confirmation.body,
                "primary_action": self.quit_confirmation.primary_action,
                "secondary_action": self.quit_confirmation.secondary_action
            }
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TrayMenuItem {
    ReadOnly { label: String },
    Action { action: TrayAction, label: String },
    Separator,
}

impl TrayMenuItem {
    #[must_use]
    pub fn read_only(label: String) -> Self {
        Self::ReadOnly { label }
    }

    #[must_use]
    pub fn action(action: TrayAction, label: &str) -> Self {
        Self::Action {
            action,
            label: label.to_string(),
        }
    }

    #[must_use]
    pub fn label(&self) -> Option<&str> {
        match self {
            Self::ReadOnly { label } | Self::Action { label, .. } => Some(label.as_str()),
            Self::Separator => None,
        }
    }

    #[must_use]
    pub fn probe_label(&self) -> String {
        self.label()
            .map_or_else(|| "---".to_string(), str::to_string)
    }

    #[must_use]
    pub fn probe_item_json(&self) -> serde_json::Value {
        match self {
            Self::ReadOnly { label } => serde_json::json!({
                "kind": "read_only",
                "label": label,
                "enabled": false
            }),
            Self::Action { action, label } => serde_json::json!({
                "kind": "action",
                "label": label,
                "action": action.id(),
                "enabled": true
            }),
            Self::Separator => serde_json::json!({
                "kind": "separator",
                "label": "---",
                "enabled": false
            }),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrayAction {
    ShowUi,
    Quit,
}

impl TrayAction {
    #[must_use]
    pub const fn id(self) -> &'static str {
        match self {
            Self::ShowUi => "show_ui",
            Self::Quit => "quit",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QuitConfirmation {
    pub title: String,
    pub body: String,
    pub primary_action: String,
    pub secondary_action: String,
}

impl QuitConfirmation {
    #[must_use]
    pub fn from_input(input: &TrayStatusInput) -> Self {
        Self {
            title: "Quit Office MCP Control".to_string(),
            body: format!(
                "Quit Office MCP Control and disconnect {} clients, {} documents, and {} running tasks?",
                input.client_count, input.document_count, input.running_task_count
            ),
            primary_action: "Quit Office MCP Control".to_string(),
            secondary_action: "Keep Running".to_string(),
        }
    }
}

fn document_count(value: Option<&serde_json::Value>) -> usize {
    value
        .and_then(serde_json::Value::as_object)
        .map_or(0, |groups| {
            groups
                .values()
                .filter_map(serde_json::Value::as_array)
                .map(|sessions| {
                    sessions
                        .iter()
                        .filter(|session| document_session_counts_as_connected(session))
                        .count()
                })
                .sum()
        })
}

fn document_session_counts_as_connected(session: &serde_json::Value) -> bool {
    session
        .get("status")
        .and_then(serde_json::Value::as_str)
        .is_none_or(|status| status.eq_ignore_ascii_case("active"))
}

#[cfg(test)]
#[path = "model_tests.rs"]
mod model_tests;
