use super::{
    ExcelToolCatalog, PowerPointToolCatalog, WORD_V1_TOOLS, tool_catalog_json,
    word_resource_catalog_for_session, word_resource_templates,
};
use serde_json::Value;

const POWERPOINT_V1_TOOL_NAMES: &[&str] = &[
    "powerpoint.get_presentation_info",
    "powerpoint.get_active_view",
    "powerpoint.export_file",
    "powerpoint.update_tags",
    "powerpoint.list_slides",
    "powerpoint.add_slide",
    "powerpoint.update_slide",
    "powerpoint.delete_slide",
    "powerpoint.move_slide",
    "powerpoint.export_slide",
    "powerpoint.list_layouts",
    "powerpoint.apply_layout",
    "powerpoint.get_selection",
    "powerpoint.set_selection",
    "powerpoint.list_shapes",
    "powerpoint.add_text_box",
    "powerpoint.add_shape",
    "powerpoint.insert_image",
    "powerpoint.update_shape",
    "powerpoint.read_text",
    "powerpoint.replace_text",
    "powerpoint.format_text",
    "powerpoint.add_table",
    "powerpoint.read_table",
    "powerpoint.update_table",
];

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
    assert!(names.contains(&"word.list_content_controls"));
    assert!(names.contains(&"word.insert_content_control"));
    assert!(names.contains(&"word.update_content_control"));
    assert!(names.contains(&"word.delete_content_control"));
    assert!(names.contains(&"word.resize_image"));
    assert!(names.contains(&"word.update_table"));
    assert!(names.contains(&"word.update_tracked_change"));
    assert!(!names.contains(&"word.insert_heading"));
    assert!(!names.contains(&"word.set_heading_level"));
    assert!(!names.contains(&"word.update_cell"));
    assert!(!names.contains(&"word.add_row"));
    assert!(!names.contains(&"word.add_column"));
    assert!(!names.contains(&"word.format_cell"));
    assert!(!names.contains(&"word.accept_change"));
    assert!(!names.contains(&"word.reject_change"));
    assert!(names.contains(&"excel.read_range"));
    assert!(names.contains(&"excel.sort_range"));
    assert!(names.contains(&"excel.apply_filter"));
    assert!(names.contains(&"excel.update_table"));
    assert!(names.contains(&"excel.update_chart"));
    assert!(names.contains(&"excel.create_pivot_table"));
    assert!(names.contains(&"excel.update_pivot_table"));
    for name in POWERPOINT_V1_TOOL_NAMES {
        assert!(names.contains(name), "missing PowerPoint tool {name}");
    }
    assert!(!names.contains(&"powerpoint.export_pdf"));
    assert!(!names.contains(&"powerpoint.duplicate_slide"));
    assert!(!names.contains(&"powerpoint.set_slide_background"));
    assert_eq!(WORD_V1_TOOLS.len(), 26);
    assert_eq!(ExcelToolCatalog::tools().len(), 20);
    assert_eq!(PowerPointToolCatalog::tools().len(), 25);
    assert_eq!(tools.len(), 73);
}

#[test]
fn every_tool_exposes_a_strict_input_schema() {
    for tool in tool_catalog_json() {
        let name = tool["name"].as_str().expect("tool name");
        let schema = &tool["inputSchema"];

        assert_eq!(schema["type"], "object", "{name} schema must be an object");
        assert_eq!(
            schema["additionalProperties"], false,
            "{name} schema must reject unknown arguments"
        );
        assert!(
            schema["properties"].is_object(),
            "{name} schema must define properties"
        );
        assert!(
            schema["required"].is_array(),
            "{name} schema must define required fields"
        );
    }
}

#[test]
fn representative_word_schemas_are_specific() {
    let paragraph = schema_for("word.get_paragraph");
    assert_required(&paragraph, &["session_id", "index"]);
    assert_eq!(paragraph["properties"]["index"]["type"], "integer");
    assert_eq!(paragraph["properties"]["index"]["minimum"], 0);
    assert!(paragraph["properties"].get("paragraph_index").is_none());

    let image = schema_for("word.insert_image");
    assert_required(&image, &["session_id", "anchor", "image"]);
    assert_eq!(
        image["properties"]["anchor"]["oneOf"]
            .as_array()
            .expect("anchor oneOf")
            .len(),
        6
    );
    assert_eq!(
        image["properties"]["placement"]["enum"]
            .as_array()
            .expect("placement enum")
            .len(),
        7
    );
    assert_eq!(image["properties"]["placement"]["default"], "inline");
    assert_eq!(
        image["properties"]["image"]["oneOf"]
            .as_array()
            .expect("image oneOf")
            .len(),
        2
    );
}

#[test]
fn word_anchor_schemas_advertise_per_tool_supported_kinds() {
    let paragraph = schema_for("word.insert_paragraph");
    let paragraph_kinds = anchor_kinds(&paragraph);
    assert!(paragraph_kinds.contains(&"after_paragraph_index"));
    assert!(paragraph_kinds.contains(&"heading"));
    assert!(paragraph_kinds.contains(&"bookmark"));

    let image = schema_for("word.insert_image");
    let image_kinds = anchor_kinds(&image);
    assert!(image_kinds.contains(&"paragraph_index"));
    assert!(image_kinds.contains(&"before_paragraph_index"));
    assert!(image_kinds.contains(&"before_text"));
    assert!(image_kinds.contains(&"after_text"));
    assert!(image_kinds.contains(&"heading"));
    assert!(image_kinds.contains(&"bookmark"));
    assert!(image_kinds.contains(&"after_paragraph_index"));
}

#[test]
fn representative_excel_and_powerpoint_schemas_are_specific() {
    let range = schema_for("excel.read_range");
    assert_required(&range, &["session_id", "address"]);
    assert_eq!(range["properties"]["sheet"]["type"], "string");
    assert!(range["properties"].get("range").is_none());

    let slide = schema_for("powerpoint.add_slide");
    assert_required(&slide, &["session_id"]);
    assert_eq!(slide["properties"]["layout"]["type"], "string");
    assert_eq!(
        slide["properties"]["title_box"]["additionalProperties"],
        false
    );
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
    assert!(ExcelToolCatalog::contains("excel.create_pivot_table"));
    assert!(ExcelToolCatalog::contains("excel.update_pivot_table"));
    assert!(ExcelToolCatalog::contains("excel.write_range"));
    assert!(!ExcelToolCatalog::contains("excel.unsupported"));
}

#[test]
fn powerpoint_tool_catalog_checks_supported_names() {
    assert!(PowerPointToolCatalog::contains(
        "powerpoint.get_presentation_info"
    ));
    assert!(PowerPointToolCatalog::contains("powerpoint.export_file"));
    assert!(PowerPointToolCatalog::contains("powerpoint.list_slides"));
    assert!(PowerPointToolCatalog::contains("powerpoint.apply_layout"));
    assert!(PowerPointToolCatalog::contains("powerpoint.update_table"));
    assert!(!PowerPointToolCatalog::contains("powerpoint.export_pdf"));
    assert!(!PowerPointToolCatalog::contains(
        "powerpoint.duplicate_slide"
    ));
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

fn schema_for(name: &str) -> Value {
    tool_catalog_json()
        .into_iter()
        .find(|tool| tool["name"] == name)
        .unwrap_or_else(|| panic!("missing tool {name}"))["inputSchema"]
        .clone()
}

fn assert_required(schema: &Value, required: &[&str]) {
    let actual = schema["required"]
        .as_array()
        .expect("required array")
        .iter()
        .filter_map(Value::as_str)
        .collect::<Vec<_>>();
    for field in required {
        assert!(actual.contains(field), "missing required field {field}");
    }
}

fn anchor_kinds(schema: &Value) -> Vec<&str> {
    schema["properties"]["anchor"]["oneOf"]
        .as_array()
        .expect("anchor oneOf")
        .iter()
        .flat_map(|variant| {
            let kind = &variant["properties"]["kind"];
            if let Some(value) = kind["const"].as_str() {
                vec![value]
            } else {
                kind["enum"]
                    .as_array()
                    .expect("kind enum")
                    .iter()
                    .map(|value| value.as_str().expect("kind string"))
                    .collect()
            }
        })
        .collect()
}
