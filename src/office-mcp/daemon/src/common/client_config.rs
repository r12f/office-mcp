use std::collections::BTreeMap;
use std::env;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ClaudeDesktopConfigBuilder {
    mode: ClaudeDesktopMode,
    install_root: Option<PathBuf>,
    current_exe: PathBuf,
    env: BTreeMap<String, String>,
}

impl ClaudeDesktopConfigBuilder {
    #[must_use]
    pub fn development() -> Self {
        Self {
            mode: ClaudeDesktopMode::Development,
            install_root: None,
            current_exe: env::current_exe().unwrap_or_else(|_| PathBuf::from("office-mcp-daemon")),
            env: env::vars().collect(),
        }
    }

    #[must_use]
    pub fn installed(install_root: Option<PathBuf>) -> Self {
        Self {
            mode: ClaudeDesktopMode::Installed,
            install_root,
            current_exe: env::current_exe().unwrap_or_else(|_| PathBuf::from("office-mcp-daemon")),
            env: env::vars().collect(),
        }
    }

    #[must_use]
    pub fn with_current_exe(mut self, current_exe: PathBuf) -> Self {
        self.current_exe = current_exe;
        self
    }

    #[must_use]
    pub fn with_env(mut self, env: BTreeMap<String, String>) -> Self {
        self.env = env;
        self
    }

    #[must_use]
    pub fn to_json(&self) -> String {
        match self.mode {
            ClaudeDesktopMode::Development => format!(
                concat!(
                    "{{\n",
                    "  \"mcpServers\": {{\n",
                    "    \"office-mcp\": {{\n",
                    "      \"command\": \"{}\",\n",
                    "      \"args\": [\"stdio\"]\n",
                    "    }}\n",
                    "  }}\n",
                    "}}"
                ),
                json_escape(&self.current_exe.display().to_string())
            ),
            ClaudeDesktopMode::Installed => {
                let root = self
                    .install_root
                    .clone()
                    .unwrap_or_else(|| default_windows_install_root(&self.env));
                let daemon = windows_child_path(&root, "office-mcp-daemon.exe");
                let config_path = windows_child_path(&root, "config.toml");
                let certificate_path = windows_child_path(&root, ".office-mcp-localhost.pfx");
                format!(
                    concat!(
                        "{{\n",
                        "  \"mcpServers\": {{\n",
                        "    \"office-mcp\": {{\n",
                        "      \"command\": \"{}\",\n",
                        "      \"args\": [\"stdio\"],\n",
                        "      \"env\": {{\n",
                        "        \"OFFICE_MCP_INSTALL_ROOT\": \"{}\",\n",
                        "        \"OFFICE_MCP_CONFIG_PATH\": \"{}\",\n",
                        "        \"OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH\": \"{}\",\n",
                        "        \"OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE\": \"office-mcp-localhost\",\n",
                        "        \"OFFICE_MCP_ADDIN_CHANNEL__BIND\": \"localhost\",\n",
                        "        \"OFFICE_MCP_ADDIN_CHANNEL__PORT\": \"8765\",\n",
                        "        \"OFFICE_MCP_MCP_HTTP__BIND\": \"127.0.0.1\",\n",
                        "        \"OFFICE_MCP_MCP_HTTP__PORT\": \"8800\"\n",
                        "      }}\n",
                        "    }}\n",
                        "  }}\n",
                        "}}"
                    ),
                    json_escape(&daemon),
                    json_escape(&root.display().to_string()),
                    json_escape(&config_path),
                    json_escape(&certificate_path)
                )
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ClaudeDesktopMode {
    Development,
    Installed,
}

fn windows_child_path(root: &Path, leaf: &str) -> String {
    let root = root.display().to_string().replace('/', "\\");
    format!("{}\\{}", root.trim_end_matches('\\'), leaf)
}

fn default_windows_install_root(env: &BTreeMap<String, String>) -> PathBuf {
    if let Some(root) = env.get("OFFICE_MCP_INSTALL_ROOT") {
        return PathBuf::from(root);
    }
    PathBuf::from(
        env.get("LOCALAPPDATA")
            .cloned()
            .or_else(|| {
                env.get("USERPROFILE")
                    .map(|path| format!("{path}\\AppData\\Local"))
            })
            .unwrap_or_else(|| "C:\\Users\\Default\\AppData\\Local".to_string()),
    )
    .join("office-mcp")
}

fn json_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

#[cfg(test)]
#[path = "client_config_tests.rs"]
mod client_config_tests;
