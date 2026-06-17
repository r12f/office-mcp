use crate::config_service::DaemonConfig;
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
    pub pid: u32,
    pub created_at: String,
}

impl UiRuntimeInfo {
    #[must_use]
    pub fn from_config(config: &DaemonConfig) -> Self {
        Self::with_origin(config.addin.origin.clone())
    }

    #[must_use]
    pub fn with_origin(origin: String) -> Self {
        Self {
            state_url: format!("{origin}/ui/state"),
            ui_url: format!("{origin}/ui/"),
            origin,
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
                "  \"pid\": {},\n",
                "  \"createdAt\": \"{}\"\n",
                "}}\n"
            ),
            json_escape(&self.origin),
            json_escape(&self.state_url),
            json_escape(&self.ui_url),
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
mod tests {
    use super::{UiRuntimeFile, UiRuntimeInfo, default_path_from_env};
    use crate::config_service::{
        AddinConfig, AuditConfig, DaemonConfig, LimitsConfig, LogLevel, LoggingConfig, McpConfig,
    };
    use std::collections::BTreeMap;
    use std::fs;

    #[test]
    fn runtime_info_uses_ui_urls_without_credentials() {
        let info = UiRuntimeInfo::with_origin("https://localhost:8765".to_string());
        let json = info.to_json();

        assert!(json.contains("\"origin\": \"https://localhost:8765\""));
        assert!(json.contains("\"stateUrl\": \"https://localhost:8765/ui/state\""));
        assert!(json.contains("\"uiUrl\": \"https://localhost:8765/ui/\""));
        assert!(!json.contains("token"));
        assert!(!json.contains("secret"));
    }

    #[test]
    fn writes_and_removes_runtime_file() {
        let dir =
            std::env::temp_dir().join(format!("office-mcp-ui-runtime-test-{}", std::process::id()));
        let path = dir.join("ui-runtime.json");
        let file = UiRuntimeFile::with_path(
            path.clone(),
            UiRuntimeInfo::with_origin("https://localhost:8765".to_string()),
        );

        file.write().expect("runtime file writes");
        let body = fs::read_to_string(&path).expect("runtime file readable");
        assert!(body.contains("https://localhost:8765/ui/"));

        file.remove().expect("runtime file removes");
        assert!(!path.exists());
        let _ = fs::remove_dir_all(dir);
    }

    #[test]
    fn reads_runtime_file_body() {
        let info = UiRuntimeInfo::with_origin("https://localhost:8765".to_string());
        let parsed = UiRuntimeInfo::from_json(&info.to_json()).expect("parse runtime info");

        assert_eq!(parsed.origin, info.origin);
        assert_eq!(parsed.state_url, info.state_url);
        assert_eq!(parsed.ui_url, info.ui_url);
    }

    #[test]
    fn default_runtime_path_uses_local_app_data_on_windows() {
        if !cfg!(windows) {
            return;
        }
        let env = BTreeMap::from([("LOCALAPPDATA".to_string(), "C:\\Local".to_string())]);
        assert_eq!(
            default_path_from_env(&env),
            std::path::PathBuf::from("C:\\Local")
                .join("office-mcp")
                .join("ui-runtime.json")
        );
    }

    #[test]
    fn runtime_file_can_be_built_from_config() {
        let config = DaemonConfig {
            addin: AddinConfig {
                host: "localhost".to_string(),
                port: 8765,
                origin: "https://localhost:8765".to_string(),
                pfx_path: String::new(),
                pfx_passphrase: String::new(),
                heartbeat_interval_sec: 30,
                heartbeat_timeout_sec: 10,
                session_grace_sec: 60,
                max_pending_per_session: 4,
            },
            mcp: McpConfig {
                host: "127.0.0.1".to_string(),
                port: 8800,
            },
            limits: LimitsConfig {
                max_response_bytes: 1,
                max_request_bytes: 1,
                max_ws_frame_bytes: 1,
                default_tool_timeout_ms: 1,
                requests_per_minute: 1,
            },
            audit: AuditConfig {
                enabled: false,
                path: String::new(),
            },
            logging: LoggingConfig {
                level: LogLevel::Info,
                file: String::new(),
            },
        };

        let file = UiRuntimeFile::from_config(&config);
        assert_eq!(file.info().origin, "https://localhost:8765");
        assert!(file.path().ends_with("ui-runtime.json"));
    }
}
