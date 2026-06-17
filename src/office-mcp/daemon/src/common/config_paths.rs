use std::collections::BTreeMap;
use std::env;
use std::path::PathBuf;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ConfigPathResolver<'a> {
    env: &'a BTreeMap<String, String>,
}

impl<'a> ConfigPathResolver<'a> {
    #[must_use]
    pub(crate) const fn new(env: &'a BTreeMap<String, String>) -> Self {
        Self { env }
    }

    #[must_use]
    pub(crate) fn config_path(&self) -> PathBuf {
        if cfg!(windows) {
            return PathBuf::from(
                self.env
                    .get("APPDATA")
                    .cloned()
                    .or_else(|| {
                        self.env
                            .get("USERPROFILE")
                            .map(|path| format!("{path}\\AppData\\Roaming"))
                    })
                    .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Roaming".to_string()),
            )
            .join("office-mcp")
            .join("config.toml");
        }
        if cfg!(target_os = "macos") {
            return PathBuf::from(
                self.env
                    .get("HOME")
                    .cloned()
                    .unwrap_or_else(|| ".".to_string()),
            )
            .join("Library")
            .join("Application Support")
            .join("office-mcp")
            .join("config.toml");
        }
        PathBuf::from(self.env.get("XDG_CONFIG_HOME").cloned().unwrap_or_else(|| {
            PathBuf::from(
                self.env
                    .get("HOME")
                    .cloned()
                    .unwrap_or_else(|| ".".to_string()),
            )
            .join(".config")
            .display()
            .to_string()
        }))
        .join("office-mcp")
        .join("config.toml")
    }

    #[must_use]
    pub(crate) fn pfx_path() -> PathBuf {
        env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join(".office-mcp-localhost.pfx")
    }

    #[must_use]
    pub(crate) fn audit_path(&self) -> String {
        self.state_path("audit.jsonl")
    }

    #[must_use]
    pub(crate) fn log_path(&self) -> String {
        self.state_path("office-mcp.log")
    }

    fn state_path(&self, filename: &str) -> String {
        if cfg!(windows) {
            return PathBuf::from(
                self.env
                    .get("LOCALAPPDATA")
                    .cloned()
                    .or_else(|| {
                        self.env
                            .get("USERPROFILE")
                            .map(|path| format!("{path}\\AppData\\Local"))
                    })
                    .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Local".to_string()),
            )
            .join("office-mcp")
            .join(filename)
            .display()
            .to_string();
        }
        if cfg!(target_os = "macos") {
            return PathBuf::from(
                self.env
                    .get("HOME")
                    .cloned()
                    .unwrap_or_else(|| ".".to_string()),
            )
            .join("Library")
            .join("Logs")
            .join("office-mcp")
            .join(filename)
            .display()
            .to_string();
        }
        PathBuf::from(self.env.get("XDG_STATE_HOME").cloned().unwrap_or_else(|| {
            PathBuf::from(
                self.env
                    .get("HOME")
                    .cloned()
                    .unwrap_or_else(|| ".".to_string()),
            )
            .join(".local")
            .join("state")
            .display()
            .to_string()
        }))
        .join("office-mcp")
        .join(filename)
        .display()
        .to_string()
    }
}

#[cfg(test)]
#[path = "config_paths_tests.rs"]
mod config_paths_tests;
