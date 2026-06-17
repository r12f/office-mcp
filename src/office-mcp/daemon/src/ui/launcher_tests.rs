use super::UiLauncher;

#[test]
fn ui_launcher_reads_runtime_file_url_instead_of_config_defaults() {
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

    let url = UiLauncher::with_runtime_path(path)
        .ui_url()
        .expect("ui url");

    assert_eq!(url, "https://localhost:8766/ui/");
    let _ = std::fs::remove_dir_all(dir);
}

#[test]
fn ui_launch_error_explains_missing_running_daemon() {
    let dir = std::env::temp_dir().join(format!(
        "office-mcp-ui-missing-command-test-{}",
        std::process::id()
    ));
    let path = dir.join("missing-ui-runtime.json");

    let error = UiLauncher::with_runtime_path(path)
        .open()
        .expect_err("missing runtime file");

    assert!(
        error
            .to_string()
            .contains("No running office-mcp daemon UI was found")
    );
    assert!(error.to_string().contains("office-mcp-daemon daemon run"));
}
