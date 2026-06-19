use super::{
    ExcelToolCatalog, PowerPointToolCatalog, WORD_V1_TOOLS, tool_catalog_json,
    word_resource_catalog_for_session, word_resource_templates,
};

#[test]
fn tool_catalog_includes_office_word_and_excel_tools() {
    let tools = tool_catalog_json();
    let names = tools
        .iter()
        .filter_map(|tool| tool["name"].as_str())
        .collect::<Vec<_>>();

    assert!(names.contains(&"office.list_sessions"));
    assert!(names.contains(&"office.get_session_info"));
    assert!(names.contains(&"word.get_text"));
    assert!(names.contains(&"excel.read_range"));
    assert!(names.contains(&"excel.sort_range"));
    assert!(names.contains(&"excel.apply_filter"));
    assert!(names.contains(&"excel.update_table"));
    assert!(names.contains(&"excel.update_chart"));
    assert!(names.contains(&"powerpoint.add_slide"));
    assert!(names.contains(&"powerpoint.replace_text"));
    assert_eq!(WORD_V1_TOOLS.len(), 27);
    assert_eq!(ExcelToolCatalog::tools().len(), 18);
    assert_eq!(PowerPointToolCatalog::tools().len(), 5);
    assert_eq!(tools.len(), 52);
}

#[test]
fn excel_tool_catalog_checks_supported_names() {
    assert!(ExcelToolCatalog::contains("excel.get_workbook_info"));
    assert!(ExcelToolCatalog::contains("excel.list_sheets"));
    assert!(ExcelToolCatalog::contains("excel.update_sheet"));
    assert!(ExcelToolCatalog::contains("excel.delete_sheet"));
    assert!(ExcelToolCatalog::contains("excel.get_used_range"));
    assert!(ExcelToolCatalog::contains("excel.clear_range"));
    assert!(ExcelToolCatalog::contains("excel.find_replace_cells"));
    assert!(ExcelToolCatalog::contains("excel.sort_range"));
    assert!(ExcelToolCatalog::contains("excel.apply_filter"));
    assert!(ExcelToolCatalog::contains("excel.update_table"));
    assert!(ExcelToolCatalog::contains("excel.update_chart"));
    assert!(ExcelToolCatalog::contains("excel.write_range"));
    assert!(!ExcelToolCatalog::contains("excel.unsupported"));
}

#[test]
fn powerpoint_tool_catalog_checks_supported_names() {
    assert!(PowerPointToolCatalog::contains("powerpoint.apply_layout"));
    assert!(PowerPointToolCatalog::contains("powerpoint.export_pdf"));
    assert!(!PowerPointToolCatalog::contains("powerpoint.unsupported"));
}

#[test]
fn word_resource_catalog_uses_session_scoped_uris() {
    let resources = word_resource_catalog_for_session("session-1");
    let uris = resources
        .iter()
        .filter_map(|resource| resource["uri"].as_str())
        .collect::<Vec<_>>();

    assert!(uris.contains(&"office://word/session-1/document?offset=0&limit=200"));
    assert!(uris.contains(&"office://word/session-1/comments"));
}

#[test]
fn word_resource_templates_include_document_and_paragraph_routes() {
    let templates = word_resource_templates();
    let names = templates
        .iter()
        .filter_map(|template| template["name"].as_str())
        .collect::<Vec<_>>();

    assert!(names.contains(&"word.document.template"));
    assert!(names.contains(&"word.paragraph.template"));
}
