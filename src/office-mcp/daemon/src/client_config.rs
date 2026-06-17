use std::collections::BTreeMap;
use std::env;
use std::path::PathBuf;

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
                let launcher = root.join("office-mcp.ps1");
                format!(
                    concat!(
                        "{{\n",
                        "  \"mcpServers\": {{\n",
                        "    \"office-mcp\": {{\n",
                        "      \"command\": \"powershell.exe\",\n",
                        "      \"args\": [\n",
                        "        \"-NoProfile\",\n",
                        "        \"-ExecutionPolicy\",\n",
                        "        \"Bypass\",\n",
                        "        \"-File\",\n",
                        "        \"{}\",\n",
                        "        \"stdio\"\n",
                        "      ]\n",
                        "    }}\n",
                        "  }}\n",
                        "}}"
                    ),
                    json_escape(&launcher.display().to_string())
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
mod tests {
    use super::ClaudeDesktopConfigBuilder;
    use std::collections::BTreeMap;
    use std::path::PathBuf;

    #[test]
    fn development_config_uses_rust_stdio_command() {
        let json = ClaudeDesktopConfigBuilder::development()
            .with_current_exe(PathBuf::from("C:\\office-mcp\\office-mcp-daemon.exe"))
            .to_json();

        assert!(json.contains("office-mcp-daemon.exe"));
        assert!(json.contains("\"stdio\""));
        assert!(!json.contains("dist/src/cli.js"));
        assert!(!json.contains("node"));
    }

    #[test]
    fn development_config_omits_legacy_reference_node_paths() {
        let json = ClaudeDesktopConfigBuilder::development()
            .with_current_exe(PathBuf::from(
                "C:\\Code\\office-mcp\\target\\debug\\office-mcp-daemon.exe",
            ))
            .to_json();

        assert!(!json.contains("reference-node"));
        assert!(!json.contains("dist/src/cli.js"));
        assert!(!json.contains("cargo run"));
    }

    #[test]
    fn installed_config_uses_launcher() {
        let json =
            ClaudeDesktopConfigBuilder::installed(Some(PathBuf::from("D:\\Apps\\office-mcp")))
                .to_json();

        assert!(json.contains("powershell.exe"));
        assert!(json.contains("D:\\\\Apps\\\\office-mcp\\\\office-mcp.ps1"));
        assert!(json.contains("\"stdio\""));
    }

    #[test]
    fn installed_config_without_root_uses_office_mcp_install_root() {
        let json = ClaudeDesktopConfigBuilder::installed(None)
            .with_env(BTreeMap::from([(
                "OFFICE_MCP_INSTALL_ROOT".to_string(),
                "E:\\OfficeMcp".to_string(),
            )]))
            .to_json();

        assert!(json.contains("E:\\\\OfficeMcp\\\\office-mcp.ps1"));
    }
}
