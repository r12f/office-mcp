use super::StaticResponseService;

#[test]
fn serves_word_taskpane_static_asset() {
    let response = response_text(&service().serve_addin_asset("/taskpane.html"));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("Content-Type: text/html; charset=utf-8"));
    assert!(response.contains("Office MCP"));
    assert!(response.contains("taskpane-shell"));
    assert!(response.contains("/common/browser-ui.js"));
    assert!(response.contains("/common/addin-channel.js"));
    assert!(response.contains("/common/logger.js"));
    assert!(response.contains("/common/task-history.js"));
}

#[test]
fn serves_versioned_addin_static_assets_with_query_strings() {
    let word_js = response_text(&service().serve_addin_asset("/taskpane.js?v=0.1.6"));
    assert!(word_js.starts_with("HTTP/1.1 200 OK"));
    assert!(word_js.contains("__OFFICE_MCP_TASKPANE_READY__"));

    let common_js = response_text(&service().serve_addin_asset("/common/addin-channel.js?v=0.1.6"));
    assert!(common_js.starts_with("HTTP/1.1 200 OK"));
    assert!(common_js.contains("OfficeCtlAddinChannel"));

    let excel_js = response_text(&service().serve_addin_asset("/excel/taskpane.js?v=0.1.6"));
    assert!(excel_js.starts_with("HTTP/1.1 200 OK"));
    assert!(excel_js.contains("function isExcelHost"));
    assert!(excel_js.contains("Office.HostType?.Excel"));
}

#[test]
fn serves_excel_taskpane_static_assets() {
    let html = response_text(&service().serve_addin_asset("/excel/taskpane.html"));
    assert!(html.starts_with("HTTP/1.1 200 OK"));
    assert!(html.contains("Office MCP Excel"));
    assert!(html.contains("/excel/taskpane.js?v=0.1.6"));
    assert!(html.contains("/common/addin-channel.js?v=0.1.6"));

    let js = response_text(&service().serve_addin_asset("/excel/taskpane.js"));
    assert!(js.starts_with("HTTP/1.1 200 OK"));
    assert!(js.contains("function isExcelHost"));
    assert!(js.contains("Office.HostType?.Excel"));
    assert!(js.contains("sessionAddedNotification"));

    let css = response_text(&service().serve_addin_asset("/excel/taskpane.css"));
    assert!(css.starts_with("HTTP/1.1 200 OK"));
    assert!(css.contains("--excel: #217346"));
}

#[test]
fn serves_office_ctl_common_browser_asset() {
    let response = response_text(&service().serve_addin_asset("/common/browser-ui.js"));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("Content-Type: text/javascript; charset=utf-8"));
    assert!(response.contains("OfficeCtlCommon"));
    assert!(response.contains("redactText"));
}

#[test]
fn serves_office_ctl_common_channel_asset() {
    let response = response_text(&service().serve_addin_asset("/common/addin-channel.js"));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("Content-Type: text/javascript; charset=utf-8"));
    assert!(response.contains("OfficeCtlAddinChannel"));
    assert!(response.contains("sendJsonRpc"));
}

#[test]
fn serves_office_ctl_common_task_history_asset() {
    let response = response_text(&service().serve_addin_asset("/common/task-history.js"));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("Content-Type: text/javascript; charset=utf-8"));
    assert!(response.contains("OfficeCtlTaskHistory"));
    assert!(response.contains("TaskHistoryStore"));
}

#[test]
fn serves_office_ctl_common_logger_asset() {
    let response = response_text(&service().serve_addin_asset("/common/logger.js"));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("Content-Type: text/javascript; charset=utf-8"));
    assert!(response.contains("OfficeCtlLogger"));
    assert!(response.contains("AddinLogger"));
}

fn service() -> StaticResponseService {
    StaticResponseService::new(crate::addin_mgr::default_addin_public_dir())
}

fn response_text(response: &crate::runtime::http_wire::WireHttpResponse) -> String {
    String::from_utf8(response.to_bytes()).expect("response utf8")
}
