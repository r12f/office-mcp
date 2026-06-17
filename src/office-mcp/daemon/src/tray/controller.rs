#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayController {
    description: &'static str,
    platform: TrayPlatform,
}

impl TrayController {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            description: "owns native tray menu bar status menu show UI and graceful quit",
            platform: TrayPlatform::detect(),
        }
    }

    #[must_use]
    pub const fn with_platform(platform: TrayPlatform) -> Self {
        Self {
            description: "owns native tray menu bar status menu show UI and graceful quit",
            platform,
        }
    }

    #[must_use]
    pub const fn description(&self) -> &'static str {
        self.description
    }

    #[must_use]
    pub const fn platform(&self) -> TrayPlatform {
        self.platform
    }

    #[must_use]
    pub fn snapshot(&self, input: TrayStatusInput) -> TraySnapshot {
        TraySnapshot::from_input(self.platform, input)
    }

    #[must_use]
    pub fn snapshot_from_ui_state(&self, ui_state: Option<&serde_json::Value>) -> TraySnapshot {
        self.snapshot(TrayStatusInput::from_ui_state(ui_state))
    }
}

impl Default for TrayController {
    fn default() -> Self {
        Self::new()
    }
}

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
            "Office MCP - {} - {} clients - {} documents",
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
                TrayMenuItem::action(TrayAction::ShowUi, "Show Office MCP"),
                TrayMenuItem::action(TrayAction::Quit, "Quit Office MCP"),
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
            "quit_confirmation": {
                "title": self.quit_confirmation.title,
                "body": self.quit_confirmation.body,
                "primary_action": self.quit_confirmation.primary_action,
                "secondary_action": self.quit_confirmation.secondary_action
            }
        })
    }
}

impl TrayPlatform {
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

impl TrayHealth {
    fn from_json(value: Option<&serde_json::Value>) -> Self {
        match value.and_then(serde_json::Value::as_str) {
            Some(value) if value.eq_ignore_ascii_case("up") => Self::Up,
            Some(value) if value.eq_ignore_ascii_case("degraded") => Self::Degraded,
            _ => Self::Down,
        }
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
}

fn document_count(value: Option<&serde_json::Value>) -> usize {
    value
        .and_then(serde_json::Value::as_object)
        .map_or(0, |groups| {
            groups
                .values()
                .filter_map(serde_json::Value::as_array)
                .map(Vec::len)
                .sum()
        })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrayAction {
    ShowUi,
    Quit,
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
            title: "Quit Office MCP".to_string(),
            body: format!(
                "Quit Office MCP and disconnect {} clients, {} documents, and {} running tasks?",
                input.client_count, input.document_count, input.running_task_count
            ),
            primary_action: "Quit Office MCP".to_string(),
            secondary_action: "Keep Running".to_string(),
        }
    }
}

pub trait TrayPlatformAdapter {
    /// Applies a tray snapshot to the platform-specific tray surface.
    ///
    /// # Errors
    ///
    /// Returns an error when the platform tray surface cannot be created or updated.
    fn apply_snapshot(&mut self, snapshot: &TraySnapshot) -> Result<(), TrayPlatformError>;

    /// Opens or focuses the daemon UI.
    ///
    /// # Errors
    ///
    /// Returns an error when the UI URL cannot be opened.
    fn show_ui(&mut self) -> Result<(), TrayPlatformError>;

    /// Requests graceful daemon shutdown.
    ///
    /// # Errors
    ///
    /// Returns an error when the daemon cannot be asked to stop.
    fn quit(&mut self) -> Result<(), TrayPlatformError>;
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TrayPlatformError {
    message: String,
}

impl TrayPlatformError {
    #[must_use]
    pub fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

impl std::fmt::Display for TrayPlatformError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for TrayPlatformError {}

#[cfg(test)]
#[path = "controller_tests.rs"]
mod controller_tests;
