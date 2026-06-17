use crate::common::DaemonConfig;
use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiRuntimeFile {
    path: PathBuf,
    info: UiRuntimeInfo,
}

impl UiRuntimeFile {
    #[must_use]
    pub fn from_config(config: &DaemonConfig) -> Self {
        Self::with_path(Self::default_path(), UiRuntimeInfo::from_config(config))
    }

    #[must_use]
    pub const fn with_path(path: PathBuf, info: UiRuntimeInfo) -> Self {
        Self { path, info }
    }

    #[must_use]
    pub fn default_path() -> PathBuf {
        if let Some(path) = env::var_os("OFFICE_MCP_UI_RUNTIME_PATH") {
            return PathBuf::from(path);
        }
        default_path_from_env(&env::vars().collect())
    }

    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    #[must_use]
    pub const fn info(&self) -> &UiRuntimeInfo {
        &self.info
    }

    /// Writes the runtime file for tray and browser launchers.
    ///
    /// # Errors
    ///
    /// Returns an error when the parent directory cannot be created or the file
    /// cannot be written.
    pub fn write(&self) -> Result<(), UiRuntimeError> {
        if let Some(parent) = self.path.parent() {
            fs::create_dir_all(parent).map_err(UiRuntimeError::Io)?;
        }
        fs::write(&self.path, self.info.to_json()).map_err(UiRuntimeError::Io)
    }

    /// Removes the runtime file if it exists.
    ///
    /// # Errors
    ///
    /// Returns an error when the file exists but cannot be removed.
    pub fn remove(&self) -> Result<(), UiRuntimeError> {
        match fs::remove_file(&self.path) {
            Ok(()) => Ok(()),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(error) => Err(UiRuntimeError::Io(error)),
        }
    }

    /// Reads a runtime file written by the daemon.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be read or does not contain the
    /// required runtime URLs.
    pub fn read_path(path: &Path) -> Result<UiRuntimeInfo, UiRuntimeError> {
        let body = fs::read_to_string(path).map_err(UiRuntimeError::Io)?;
        UiRuntimeInfo::from_json(&body)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiRuntimeInfo {
    pub origin: String,
    pub state_url: String,
    pub ui_url: String,
    pub log_path: Option<String>,
    pub pid: u32,
    pub created_at: String,
}

impl UiRuntimeInfo {
    #[must_use]
    pub fn from_config(config: &DaemonConfig) -> Self {
        Self::with_origin_and_log_path(
            config.addin.origin.clone(),
            Some(config.logging.file.clone()),
        )
    }

    #[must_use]
    pub fn with_origin(origin: String) -> Self {
        Self::with_origin_and_log_path(origin, None)
    }

    #[must_use]
    pub fn with_origin_and_log_path(origin: String, log_path: Option<String>) -> Self {
        Self {
            state_url: format!("{origin}/ui/state"),
            ui_url: format!("{origin}/ui/"),
            origin,
            log_path,
            pid: std::process::id(),
            created_at: current_timestamp(),
        }
    }

    #[must_use]
    pub fn to_json(&self) -> String {
        format!(
            concat!(
                "{{\n",
                "  \"origin\": \"{}\",\n",
                "  \"stateUrl\": \"{}\",\n",
                "  \"uiUrl\": \"{}\",\n",
                "  \"logPath\": {},\n",
                "  \"pid\": {},\n",
                "  \"createdAt\": \"{}\"\n",
                "}}\n"
            ),
            json_escape(&self.origin),
            json_escape(&self.state_url),
            json_escape(&self.ui_url),
            self.log_path.as_ref().map_or_else(
                || "null".to_string(),
                |path| format!("\"{}\"", json_escape(path))
            ),
            self.pid,
            json_escape(&self.created_at)
        )
    }

    /// Parses a runtime file body written by [`UiRuntimeInfo::to_json`].
    ///
    /// # Errors
    ///
    /// Returns an error when the JSON is malformed or missing required fields.
    pub fn from_json(body: &str) -> Result<Self, UiRuntimeError> {
        let value = serde_json::from_str::<serde_json::Value>(body)
            .map_err(|error| UiRuntimeError::Parse(error.to_string()))?;
        Ok(Self {
            origin: required_string(&value, "origin")?,
            state_url: required_string(&value, "stateUrl")?,
            ui_url: required_string(&value, "uiUrl")?,
            log_path: value
                .get("logPath")
                .and_then(serde_json::Value::as_str)
                .map(ToString::to_string),
            pid: value
                .get("pid")
                .and_then(serde_json::Value::as_u64)
                .and_then(|value| value.try_into().ok())
                .unwrap_or(0),
            created_at: value
                .get("createdAt")
                .and_then(serde_json::Value::as_str)
                .unwrap_or_default()
                .to_string(),
        })
    }
}

#[derive(Debug)]
pub enum UiRuntimeError {
    Io(std::io::Error),
    Parse(String),
}

impl Display for UiRuntimeError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "UI runtime file error: {error}"),
            Self::Parse(error) => write!(formatter, "UI runtime file parse error: {error}"),
        }
    }
}

impl Error for UiRuntimeError {}

#[must_use]
pub fn default_path_from_env(env: &BTreeMap<String, String>) -> PathBuf {
    if cfg!(windows) {
        return PathBuf::from(
            env.get("LOCALAPPDATA")
                .cloned()
                .or_else(|| {
                    env.get("USERPROFILE")
                        .map(|path| format!("{path}\\AppData\\Local"))
                })
                .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Local".to_string()),
        )
        .join("office-mcp")
        .join("ui-runtime.json");
    }
    if cfg!(target_os = "macos") {
        return PathBuf::from(env.get("HOME").cloned().unwrap_or_else(|| ".".to_string()))
            .join("Library")
            .join("Application Support")
            .join("office-mcp")
            .join("ui-runtime.json");
    }
    PathBuf::from(env.get("XDG_RUNTIME_DIR").cloned().unwrap_or_else(|| {
        PathBuf::from(env.get("HOME").cloned().unwrap_or_else(|| ".".to_string()))
            .join(".local")
            .join("state")
            .join("office-mcp")
            .display()
            .to_string()
    }))
    .join("ui-runtime.json")
}

fn current_timestamp() -> String {
    let seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("{seconds}")
}

fn json_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

fn required_string(value: &serde_json::Value, key: &str) -> Result<String, UiRuntimeError> {
    value
        .get(key)
        .and_then(serde_json::Value::as_str)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
        .ok_or_else(|| UiRuntimeError::Parse(format!("missing {key}")))
}

#[cfg(test)]
#[path = "runtime_tests.rs"]
mod runtime_tests;
