use super::{DaemonControlError, DaemonController, PowerShellExecutor};
use std::cell::RefCell;
use std::fs;
use std::path::PathBuf;

#[test]
fn start_uses_windows_scheduled_task_when_it_succeeds() {
    let calls = RefCell::new(Vec::new());
    let executor = FakePowerShellExecutor::new(&calls, 0);
    let controller = DaemonController {
        task_name: "office-mcp".to_string(),
        install_root: None,
        runtime_path: PathBuf::from("runtime.json"),
    };

    controller
        .start_with_executor(&executor)
        .expect("scheduled task start succeeds");

    let calls = calls.borrow();
    assert_eq!(calls.len(), 1);
    assert!(calls[0].contains("Start-ScheduledTask -TaskName 'office-mcp'"));
}

#[test]
fn start_falls_back_to_installed_launcher_when_scheduled_task_fails() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-daemon-control-install-{}-{}",
        std::process::id(),
        unique_suffix()
    ));
    fs::create_dir_all(&dir).expect("install dir");
    fs::write(dir.join("office-mcp-daemon.ps1"), "").expect("launcher");
    let calls = RefCell::new(Vec::new());
    let executor = FakePowerShellExecutor::new(&calls, 1);
    let controller = DaemonController {
        task_name: "office-mcp".to_string(),
        install_root: Some(dir.clone()),
        runtime_path: PathBuf::from("runtime.json"),
    };

    controller
        .start_with_executor(&executor)
        .expect("launcher fallback succeeds");

    let calls = calls.borrow();
    assert_eq!(calls.len(), 2);
    assert!(calls[0].contains("Start-ScheduledTask"));
    assert!(calls[1].contains("Start-Process -WindowStyle Hidden"));
    assert!(calls[1].contains("office-mcp-daemon.ps1"));
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn stop_falls_back_to_installed_launcher_process() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-daemon-control-install-{}-{}",
        std::process::id(),
        unique_suffix()
    ));
    fs::create_dir_all(&dir).expect("install dir");
    fs::write(dir.join("office-mcp-daemon.ps1"), "").expect("launcher");
    let calls = RefCell::new(Vec::new());
    let executor = FakePowerShellExecutor::new(&calls, 1);
    let controller = DaemonController {
        task_name: "office-mcp".to_string(),
        install_root: Some(dir.clone()),
        runtime_path: PathBuf::from("runtime.json"),
    };

    controller
        .stop_with_executor(&executor)
        .expect("launcher stop fallback succeeds");

    let calls = calls.borrow();
    assert_eq!(calls.len(), 2);
    assert!(calls[0].contains("Stop-ScheduledTask"));
    assert!(calls[1].contains("Get-CimInstance Win32_Process"));
    assert!(calls[1].contains("office-mcp-daemon.ps1"));
    assert!(calls[1].contains("Stop-Process"));
    let _ = fs::remove_dir_all(dir);
}

#[test]
fn status_reports_runtime_details_without_auth_material() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-daemon-control-test-{}",
        std::process::id()
    ));
    let path = dir.join("ui-runtime.json");
    fs::create_dir_all(&dir).expect("temp dir");
    fs::write(
        &path,
        "{\"pid\":0,\"uiUrl\":\"https://localhost:8765/ui/\",\"stateUrl\":\"https://localhost:8765/ui/state\",\"logPath\":\"C:\\\\logs\\\\office-mcp.log\"}",
    )
    .expect("runtime file");

    let json = DaemonController::with_runtime_path(path.clone()).status_json();

    assert!(json.contains("\"running\": false"));
    assert!(json.contains("https://localhost:8765/ui/"));
    assert!(json.contains("\"logPath\": \"C:\\\\logs\\\\office-mcp.log\""));
    assert!(!json.contains("token"));
    assert!(!json.contains("secret"));
    let _ = fs::remove_dir_all(dir);
}

struct FakePowerShellExecutor<'a> {
    calls: &'a RefCell<Vec<String>>,
    remaining_failures: RefCell<usize>,
}

impl<'a> FakePowerShellExecutor<'a> {
    const fn new(calls: &'a RefCell<Vec<String>>, failures: usize) -> Self {
        Self {
            calls,
            remaining_failures: RefCell::new(failures),
        }
    }
}

impl PowerShellExecutor for FakePowerShellExecutor<'_> {
    fn run(&self, command: &str) -> Result<(), DaemonControlError> {
        self.calls.borrow_mut().push(command.to_string());
        let mut remaining_failures = self.remaining_failures.borrow_mut();
        if *remaining_failures > 0 {
            *remaining_failures -= 1;
            return Err(DaemonControlError::CommandFailed(
                "task unavailable".to_string(),
            ));
        }
        Ok(())
    }
}

fn unique_suffix() -> u128 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("time")
        .as_nanos()
}
