use office_mcp_daemon::OfficeMcpDaemon;
use office_mcp_daemon::api::DaemonController;
use office_mcp_daemon::common::{
    ClaudeDesktopConfigBuilder, DaemonConfig, DaemonConfigService, LoadConfigOptions,
    Logger as DaemonLogger, LoggerLogLevel,
};
use office_mcp_daemon::evidence_fixture::{UiFixtureOptions, run_ui_fixture};
use office_mcp_daemon::mcp::McpManagementClient;
use office_mcp_daemon::mcp::StdioBridge;
use office_mcp_daemon::runtime_server::RuntimeServer;
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
        [command, subcommand] if command == "daemon" && subcommand == "run" => {
            serve_daemon_with_optional_tray(false)
        }
        [command, subcommand, flag]
            if command == "daemon" && subcommand == "run" && flag == "--with-tray" =>
        {
            serve_daemon_with_optional_tray(true)
        }
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
                "usage: office-mcp-daemon [--describe|--parity-gates|serve|stdio|sessions|ui|tray [--probe] [--runtime-path <path>] [--probe-state-path <path>]|daemon run [--with-tray]|daemon status|daemon start|daemon stop|config endpoints|config show|config claude-desktop [--installed] [--install-root <path>]]"
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
    serve_daemon_with_optional_tray(false);
}

fn serve_daemon_with_optional_tray(start_tray: bool) {
    if start_tray {
        start_tray_background();
    }
    match load_config().and_then(|config| {
        DaemonConfigService::assert_boundary_auth_config(&config)?;
        Ok(config)
    }) {
        Ok(config) => match RuntimeServer::from_daemon_config(&config) {
            Ok(server) => {
                let _log_guard = match DaemonLogger::init_tracing_file(
                    logger_level_from_config(config.logging.level),
                    &config.logging.file,
                ) {
                    Ok(guard) => Some(guard),
                    Err(error) => {
                        eprintln!("office-mcp-daemon failed to initialize tracing log: {error}");
                        None
                    }
                };
                let endpoints = config.endpoints();
                eprintln!("office-mcp-daemon MCP listening on {}", endpoints.mcp);
                eprintln!(
                    "office-mcp-daemon add-in listening on {}",
                    endpoints.addin_origin
                );
                tracing::info!(
                    mcp_endpoint = %endpoints.mcp,
                    addin_origin = %endpoints.addin_origin,
                    log_path = %config.logging.file,
                    "office-mcp-daemon started"
                );
                let runtime_file = UiRuntimeFile::from_config(&config);
                if let Err(error) = server.serve_forever_with_runtime_file(&runtime_file) {
                    tracing::error!(%error, "office-mcp-daemon stopped with error");
                    exit_error(error);
                }
            }
            Err(error) => exit_error(error),
        },
        Err(error) => exit_error(error),
    }
}

fn start_tray_background() {
    let _ = std::thread::Builder::new()
        .name("office-mcp-tray".to_string())
        .spawn(|| {
            if let Err(error) = TrayHost::new(TrayHostOptions::default()).run() {
                tracing::error!(%error, "office-mcp tray host stopped with error");
                eprintln!("office-mcp-daemon tray host stopped with error: {error}");
            }
        })
        .map_err(|error| {
            tracing::error!(%error, "office-mcp tray host thread failed to start");
            eprintln!("office-mcp-daemon failed to start tray host thread: {error}");
        });
}

const fn logger_level_from_config(
    level: office_mcp_daemon::common::ConfigLogLevel,
) -> LoggerLogLevel {
    match level {
        office_mcp_daemon::common::ConfigLogLevel::Trace => LoggerLogLevel::Trace,
        office_mcp_daemon::common::ConfigLogLevel::Debug => LoggerLogLevel::Debug,
        office_mcp_daemon::common::ConfigLogLevel::Info => LoggerLogLevel::Info,
        office_mcp_daemon::common::ConfigLogLevel::Warn => LoggerLogLevel::Warn,
        office_mcp_daemon::common::ConfigLogLevel::Error => LoggerLogLevel::Error,
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

fn load_config() -> Result<DaemonConfig, office_mcp_daemon::common::ConfigError> {
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
#[path = "main_tests.rs"]
mod main_tests;
