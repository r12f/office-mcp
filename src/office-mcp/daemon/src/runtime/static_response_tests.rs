use super::StaticResponseService;

#[test]
fn serves_word_taskpane_static_asset() {
    let response = response_text(&service().serve_addin_asset("/word/taskpane.html"));

    assert!(response.starts_with("HTTP/1.1 200 OK"));
    assert!(response.contains("Content-Type: text/html; charset=utf-8"));
    assert!(response.contains("Office MCP"));
    assert!(response.contains("taskpane-shell"));
    assert!(response.contains("/common/browser-ui.js"));
    assert!(response.contains("/common/addin-channel.js"));
    assert!(response.contains("/common/logger.js"));
    assert!(response.contains("/common/task-history.js"));
    assert!(response.contains("/common/main-ui.js"));
    assert!(response.contains("/common/taskpane.css"));
}

#[test]
fn serves_versioned_addin_static_assets_with_query_strings() {
    let word_js = response_text(&service().serve_addin_asset("/word/taskpane.js?v=0.1.12"));
    assert!(word_js.starts_with("HTTP/1.1 200 OK"));
    assert!(word_js.contains("__OFFICE_MCP_TASKPANE_READY__"));

    let common_js = response_text(&service().serve_addin_asset("/common/addin-channel.js?v=0.1.12"));
    assert!(common_js.starts_with("HTTP/1.1 200 OK"));
    assert!(common_js.contains("OfficeCtlAddinChannel"));

    let main_ui_js = response_text(&service().serve_addin_asset("/common/main-ui.js?v=0.1.12"));
    assert!(main_ui_js.starts_with("HTTP/1.1 200 OK"));
    assert!(main_ui_js.contains("OfficeCtlMainUi"));

    let excel_js = response_text(&service().serve_addin_asset("/excel/taskpane.js?v=0.1.10"));
    assert!(excel_js.starts_with("HTTP/1.1 200 OK"));
    assert!(excel_js.contains("function isExcelHost"));
    assert!(excel_js.contains("Office.HostType?.Excel"));
}

#[test]
fn serves_excel_taskpane_static_assets() {
    let html = response_text(&service().serve_addin_asset("/excel/taskpane.html"));
    assert!(html.starts_with("HTTP/1.1 200 OK"));
    assert!(html.contains("Office MCP Control"));
    assert!(html.contains("/excel/taskpane.js?v=0.1.10"));
    assert!(html.contains("/common/addin-channel.js?v=0.1.10"));

    let js = response_text(&service().serve_addin_asset("/excel/taskpane.js"));
    assert!(js.starts_with("HTTP/1.1 200 OK"));
    assert!(js.contains("function isExcelHost"));
    assert!(js.contains("Office.HostType?.Excel"));
    assert!(js.contains("sessionAddedNotification"));

    let css = response_text(&service().serve_addin_asset("/common/taskpane.css"));
    assert!(css.starts_with("HTTP/1.1 200 OK"));
    assert!(css.contains("--accent: #3b6478"));

    let excel_css = response_text(&service().serve_addin_asset("/excel/taskpane.css"));
    assert!(excel_css.starts_with("HTTP/1.1 200 OK"));
    assert!(excel_css.contains("--accent: #217346"));
}

#[test]
fn serves_powerpoint_taskpane_static_assets() {
    let html = response_text(&service().serve_addin_asset("/powerpoint/taskpane.html"));
    assert!(html.starts_with("HTTP/1.1 200 OK"));
    assert!(html.contains("Office MCP Control"));
    assert!(html.contains("/powerpoint/taskpane.js?v=0.1.1"));
    assert!(html.contains("/common/addin-channel.js?v=0.1.1"));

    let js = response_text(&service().serve_addin_asset("/powerpoint/taskpane.js"));
    assert!(js.starts_with("HTTP/1.1 200 OK"));
    assert!(js.contains("function isPowerPointHost"));
    assert!(js.contains("Office.HostType?.PowerPoint"));
    assert!(js.contains("sessionAddedNotification"));
    assert!(js.contains("available_tools: effectiveTools()"));
    assert!(js.contains("async function addSlide"));
    assert!(js.contains("async function exportPdf"));

    let css = response_text(&service().serve_addin_asset("/powerpoint/taskpane.css"));
    assert!(css.starts_with("HTTP/1.1 200 OK"));
    assert!(css.contains("--powerpoint: #b7472a"));
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

#[test]
fn serves_generated_brand_icons_instead_of_placeholder_pngs() {
    let icon_16 = response_bytes(&service().serve_addin_asset("/assets/icon-16.png"));
    let icon_32 = response_bytes(&service().serve_addin_asset("/assets/icon-32.png"));
    let icon_80 = response_bytes(&service().serve_addin_asset("/assets/icon-80.png"));

    assert!(starts_with_http_ok(&icon_16));
    assert!(String::from_utf8_lossy(&icon_16).contains("Content-Type: image/png"));
    assert!(png_body(&icon_16).len() > 100);
    assert_eq!(png_dimensions(png_body(&icon_16)), (16, 16));

    assert!(starts_with_http_ok(&icon_32));
    assert!(String::from_utf8_lossy(&icon_32).contains("Content-Type: image/png"));
    assert!(png_body(&icon_32).len() > 120);
    assert_eq!(png_dimensions(png_body(&icon_32)), (32, 32));

    assert!(starts_with_http_ok(&icon_80));
    assert!(String::from_utf8_lossy(&icon_80).contains("Content-Type: image/png"));
    assert!(png_body(&icon_80).len() > png_body(&icon_32).len());
    assert_eq!(png_dimensions(png_body(&icon_80)), (80, 80));
}

fn service() -> StaticResponseService {
    StaticResponseService::new(crate::addin_mgr::default_addin_public_dir())
}

fn response_text(response: &crate::runtime::http_wire::WireHttpResponse) -> String {
    String::from_utf8(response.to_bytes()).expect("response utf8")
}

fn response_bytes(response: &crate::runtime::http_wire::WireHttpResponse) -> Vec<u8> {
    response.to_bytes()
}

fn starts_with_http_ok(response: &[u8]) -> bool {
    response.starts_with(b"HTTP/1.1 200 OK")
}

fn png_body(response: &[u8]) -> &[u8] {
    let marker = b"\r\n\r\n";
    let start = response
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("response header terminator")
        + marker.len();
    &response[start..]
}

fn png_dimensions(png: &[u8]) -> (u32, u32) {
    assert!(png.starts_with(&[137, 80, 78, 71, 13, 10, 26, 10]));
    (
        u32::from_be_bytes(png[16..20].try_into().unwrap()),
        u32::from_be_bytes(png[20..24].try_into().unwrap()),
    )
}
