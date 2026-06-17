use office_mcp_daemon::OfficeMcpDaemon;
use office_mcp_daemon::client_config::ClaudeDesktopConfigBuilder;
use office_mcp_daemon::config_service::{DaemonConfig, DaemonConfigService, LoadConfigOptions};
use office_mcp_daemon::daemon_control::DaemonController;
use office_mcp_daemon::evidence_fixture::{UiFixtureOptions, run_ui_fixture};
use office_mcp_daemon::mcp_management_client::McpManagementClient;
use office_mcp_daemon::runtime_server::RuntimeServer;
use office_mcp_daemon::stdio_bridge::StdioBridge;
use office_mcp_daemon::tray::{TrayHost, TrayHostOptions};
use office_mcp_daemon::ui::{UiRuntimeError, UiRuntimeFile};
use std::path::PathBuf;

fn main() {
    let daemon = OfficeMcpDaemon::new();
    let args = std::env::args().skip(1).collect::<Vec<_>>();

    match args.as_slice() {
        [] => print_description(&daemon),
        [flag] if flag == "--describe" => print_description(&daemon),
        [flag] if flag == "--parity-gates" => print_parity_gates(&daemon),
        [command] if command == "serve" => serve_daemon(),
        [command] if command == "stdio" => run_stdio_bridge(),
        [command] if command == "sessions" => list_sessions(),
        [command, subcommand] if command == "daemon" && subcommand == "run" => serve_daemon(),
        [command, subcommand] if command == "daemon" && subcommand == "status" => {
            println!("{}", DaemonController::from_env().status_json());
        }
        [command, subcommand] if command == "daemon" && subcommand == "start" => {
            if let Err(error) = DaemonController::from_env().start() {
                exit_error(error);
            }
        }
        [command, subcommand] if command == "daemon" && subcommand == "stop" => {
            if let Err(error) = DaemonController::from_env().stop() {
                exit_error(error);
            }
        }
        [command] if command == "ui" => open_ui(),
        [command, rest @ ..] if command == "tray" => run_tray(rest),
        [command, subcommand] if command == "evidence" && subcommand == "ui-fixture" => {
            if let Err(error) = run_ui_fixture(UiFixtureOptions::from_env()) {
                exit_error(error);
            }
        }
        [command, subcommand] if command == "config" && subcommand == "endpoints" => {
            print_config_endpoints();
        }
        [command, subcommand] if command == "config" && subcommand == "show" => {
            print_config_show();
        }
        [command, subcommand, rest @ ..]
            if command == "config" && subcommand == "claude-desktop" =>
        {
            print_claude_desktop_config(rest);
        }
        _ => {
            eprintln!(
                "usage: office-mcp-daemon [--describe|--parity-gates|serve|stdio|sessions|ui|tray [--probe] [--runtime-path <path>] [--probe-state-path <path>]|daemon run|daemon status|daemon start|daemon stop|config endpoints|config show|config claude-desktop [--installed] [--install-root <path>]]"
            );
            std::process::exit(2);
        }
    }
}

fn list_sessions() {
    match load_config() {
        Ok(config) => match McpManagementClient::from_config(&config).list_sessions() {
            Ok(body) => println!("{body}"),
            Err(error) => exit_error(error),
        },
        Err(error) => exit_error(error),
    }
}

fn run_stdio_bridge() {
    match load_config() {
        Ok(config) => {
            let mut bridge = StdioBridge::from_config(&config);
            if let Err(error) = bridge.run() {
                exit_error(error);
            }
        }
        Err(error) => exit_error(error),
    }
}

fn serve_daemon() {
    match load_config().and_then(|config| {
        DaemonConfigService::assert_boundary_auth_config(&config)?;
        Ok(config)
    }) {
        Ok(config) => match RuntimeServer::from_daemon_config(&config) {
            Ok(server) => {
                let endpoints = config.endpoints();
                eprintln!("office-mcp-daemon MCP listening on {}", endpoints.mcp);
                eprintln!(
                    "office-mcp-daemon add-in listening on {}",
                    endpoints.addin_origin
                );
                let runtime_file = UiRuntimeFile::from_config(&config);
                if let Err(error) = server.serve_forever_with_runtime_file(&runtime_file) {
                    exit_error(error);
                }
            }
            Err(error) => exit_error(error),
        },
        Err(error) => exit_error(error),
    }
}

fn open_ui() {
    match ui_url_from_runtime_path(&UiRuntimeFile::default_path()) {
        Ok(url) => {
            if let Err(error) = open_url(&url) {
                exit_error(error);
            }
            println!("{url}");
        }
        Err(error) => exit_error(format!(
            "No running office-mcp daemon UI was found. Start the daemon with `office-mcp-daemon daemon run` and try again. {error}"
        )),
    }
}

fn ui_url_from_runtime_path(path: &std::path::Path) -> Result<String, UiRuntimeError> {
    Ok(UiRuntimeFile::read_path(path)?.ui_url)
}

fn run_tray(args: &[String]) {
    if let Err(error) = TrayHost::new(TrayHostOptions::from_args(args)).run() {
        exit_error(error);
    }
}

fn print_description(daemon: &OfficeMcpDaemon) {
    println!("office-mcp-daemon rust reference scaffold");
    for component in daemon.component_descriptions() {
        println!("{}: {}", component.name(), component.description());
    }
}

fn print_parity_gates(daemon: &OfficeMcpDaemon) {
    for gate in daemon.parity_plan().gates() {
        println!("{gate:?}");
    }
}

fn print_config_endpoints() {
    match load_config() {
        Ok(config) => {
            let endpoints = config.endpoints();
            println!(
                "{{\n  \"mcp\": \"{}\",\n  \"addin_origin\": \"{}\",\n  \"addin_wss\": \"{}\"\n}}",
                json_escape(&endpoints.mcp),
                json_escape(&endpoints.addin_origin),
                json_escape(&endpoints.addin_wss)
            );
        }
        Err(error) => exit_error(error),
    }
}

fn print_config_show() {
    match load_config() {
        Ok(config) => println!("{}", render_redacted_config(&config)),
        Err(error) => exit_error(error),
    }
}

fn print_claude_desktop_config(args: &[String]) {
    let installed = args.iter().any(|arg| arg == "--installed");
    let install_root = read_option(args, "--install-root").map(PathBuf::from);
    let builder = if installed {
        ClaudeDesktopConfigBuilder::installed(install_root)
    } else {
        ClaudeDesktopConfigBuilder::development()
    };
    println!("{}", builder.to_json());
}

fn load_config() -> Result<DaemonConfig, office_mcp_daemon::config_service::ConfigError> {
    DaemonConfigService::new().load(LoadConfigOptions::default())
}

fn render_redacted_config(config: &DaemonConfig) -> String {
    let redacted = DaemonConfigService::redacted(config);
    let endpoints = config.endpoints();
    format!(
        concat!(
            "{{\n",
            "  \"addin\": {{\n",
            "    \"host\": \"{}\",\n",
            "    \"port\": {},\n",
            "    \"origin\": \"{}\",\n",
            "    \"pfxPath\": \"{}\",\n",
            "    \"pfxPassphrase\": \"{}\"\n",
            "  }},\n",
            "  \"mcp\": {{\n",
            "    \"host\": \"{}\",\n",
            "    \"port\": {}\n",
            "  }},\n",
            "  \"endpoints\": {{\n",
            "    \"mcp\": \"{}\",\n",
            "    \"addin_origin\": \"{}\",\n",
            "    \"addin_wss\": \"{}\"\n",
            "  }}\n",
            "}}"
        ),
        json_escape(&redacted.addin.host),
        redacted.addin.port,
        json_escape(&redacted.addin.origin),
        json_escape(&redacted.addin.pfx_path),
        json_escape(&redacted.addin.pfx_passphrase),
        json_escape(&redacted.mcp.host),
        redacted.mcp.port,
        json_escape(&endpoints.mcp),
        json_escape(&endpoints.addin_origin),
        json_escape(&endpoints.addin_wss)
    )
}

fn json_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

fn read_option(args: &[String], name: &str) -> Option<String> {
    args.iter()
        .position(|arg| arg == name)
        .and_then(|index| args.get(index + 1))
        .filter(|value| !value.starts_with("--"))
        .cloned()
}

fn open_url(url: &str) -> Result<(), std::io::Error> {
    #[cfg(windows)]
    {
        std::process::Command::new("cmd")
            .args(["/C", "start", "", url])
            .spawn()?
            .wait()?;
        return Ok(());
    }
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("open")
            .arg(url)
            .spawn()?
            .wait()?;
        return Ok(());
    }
    #[cfg(all(unix, not(target_os = "macos")))]
    {
        std::process::Command::new("xdg-open")
            .arg(url)
            .spawn()?
            .wait()?;
        return Ok(());
    }
    #[allow(unreachable_code)]
    Ok(())
}

fn exit_error(error: impl std::fmt::Display) -> ! {
    eprintln!("{error}");
    std::process::exit(1);
}

#[cfg(test)]
mod tests {
    use super::{json_escape, render_redacted_config, ui_url_from_runtime_path};
    use office_mcp_daemon::config_service::{
        AddinConfig, AuditConfig, DaemonConfig, LimitsConfig, LogLevel, LoggingConfig, McpConfig,
    };

    #[test]
    fn redacted_config_json_hides_secrets_and_includes_endpoints() {
        let config = DaemonConfig {
            addin: AddinConfig {
                host: "localhost".to_string(),
                port: 8765,
                origin: "https://localhost:8765".to_string(),
                pfx_path: "C:\\cert.pfx".to_string(),
                pfx_passphrase: "secret".to_string(),
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

        let json = render_redacted_config(&config);

        assert!(json.contains("<redacted>"));
        assert!(!json.contains("secret"));
        assert!(json.contains("http://127.0.0.1:8800/mcp"));
        assert!(json.contains("wss://localhost:8765/addin"));
    }

    #[test]
    fn escapes_json_control_characters() {
        assert_eq!(json_escape("a\\b\"c\n"), "a\\\\b\\\"c\\n");
    }

    #[test]
    fn ui_command_reads_runtime_file_url_instead_of_config_defaults() {
        let dir =
            std::env::temp_dir().join(format!("office-mcp-ui-command-test-{}", std::process::id()));
        let path = dir.join("ui-runtime.json");
        std::fs::create_dir_all(&dir).expect("temp dir");
        std::fs::write(
            &path,
            concat!(
                "{",
                "\"origin\":\"https://localhost:8766\",",
                "\"stateUrl\":\"https://localhost:8766/ui/state\",",
                "\"uiUrl\":\"https://localhost:8766/ui/\",",
                "\"pid\":123,",
                "\"createdAt\":\"1\"",
                "}"
            ),
        )
        .expect("runtime file");

        let url = ui_url_from_runtime_path(&path).expect("ui url");

        assert_eq!(url, "https://localhost:8766/ui/");
        let _ = std::fs::remove_dir_all(dir);
    }
}
