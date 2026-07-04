use super::{
    ExcelToolCatalog, PowerPointToolCatalog, WORD_V1_TOOLS, all_office_tool_names,
    excel_resource_catalog_for_session, excel_resource_templates, office_tool_catalogs,
    powerpoint_resource_catalog_for_session, powerpoint_resource_templates, tool_catalog_json,
    word_resource_catalog_for_session, word_resource_templates,
};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};

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
    assert!(names.contains(&"office_list_sessions"));
    assert!(names.contains(&"office.get_session_info"));
    assert!(names.contains(&"office_get_session_info"));
    assert!(names.contains(&"office.describe_tools"));
    assert!(names.contains(&"office_describe_tools"));
    assert!(!names.contains(&"office.describe_tool"));
    assert!(names.contains(&"word.get_text"));
    assert!(names.contains(&"word_get_text"));
    assert!(names.contains(&"word.get_header_footer"));
    assert!(names.contains(&"word.update_header_footer"));
    assert!(names.contains(&"word.resolve_anchor"));
    assert!(names.contains(&"word.list_content_controls"));
    assert!(names.contains(&"word.insert_content_control"));
    assert!(names.contains(&"word.update_content_control"));
    assert!(names.contains(&"word.delete_content_control"));
    assert!(names.contains(&"word.resize_image"));
    assert!(names.contains(&"word.update_table"));
    assert!(names.contains(&"word_update_table"));
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
    assert!(names.contains(&"excel_read_range"));
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
    assert_eq!(WORD_V1_TOOLS.len(), 29);
    assert_eq!(ExcelToolCatalog::tools().len(), 20);
    assert_eq!(PowerPointToolCatalog::tools().len(), 25);
    assert_eq!(tools.len(), 154);
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
fn office_tools_expose_rich_contract_metadata_in_tools_list() {
    for name in [
        "word.insert_image",
        "word.update_table",
        "excel.update_table",
        "powerpoint.update_table",
    ] {
        let tool = tool_for(name);
        let meta = &tool["_meta"];

        assert_eq!(
            meta["com.office-mcp/app"],
            app_for(name),
            "{name} app metadata"
        );
        assert!(
            meta["com.office-mcp/category"].is_string(),
            "{name} category metadata"
        );
        assert!(
            ["read", "mutating", "destructive"].contains(
                &meta["com.office-mcp/side_effects"]
                    .as_str()
                    .expect("side effect")
            ),
            "{name} side effect metadata"
        );
        assert!(
            meta["com.office-mcp/common_errors"]
                .as_array()
                .expect("common errors")
                .iter()
                .any(|error| error["code"] == "INVALID_ARGUMENTS"),
            "{name} common errors metadata"
        );
        assert!(
            !meta["com.office-mcp/examples"]
                .as_array()
                .expect("examples")
                .is_empty(),
            "{name} complex tool examples metadata"
        );

        let annotations = &tool["annotations"];
        assert_eq!(
            annotations["openWorldHint"], false,
            "{name} should not be modeled as open-world"
        );
        assert!(
            annotations["readOnlyHint"].is_boolean(),
            "{name} read-only annotation"
        );
        assert!(
            annotations["destructiveHint"].is_boolean(),
            "{name} destructive annotation"
        );
    }
}

#[test]
fn tools_list_contract_metadata_matches_describe_tool_contract() {
    for name in [
        "word.insert_image",
        "excel.update_table",
        "powerpoint.update_table",
    ] {
        let listed = tool_for(name);
        let described = super::describe_tool_contract(name).expect("described tool");

        assert_eq!(listed["inputSchema"], described["input_schema"]);
        assert!(
            described["parameters"]
                .as_array()
                .expect("parameters")
                .iter()
                .any(|parameter| parameter["name"] == "session_id"
                    && parameter["required"] == true
                    && parameter["schema"]["type"] == "string")
        );
        assert_eq!(
            listed["_meta"]["com.office-mcp/examples"],
            described["examples"]
        );
        assert_eq!(
            listed["_meta"]["com.office-mcp/common_errors"],
            described["common_errors"]
        );
        assert_eq!(
            listed["_meta"]["com.office-mcp/side_effects"],
            described["side_effect"]
        );
    }
}

#[test]
fn safe_tool_aliases_share_canonical_contract_metadata() {
    let canonical = tool_for("word.update_table");
    let alias = tool_for("word_update_table");

    assert_eq!(
        alias["_meta"]["com.office-mcp/alias_for"],
        "word.update_table"
    );
    assert_eq!(
        alias["_meta"]["com.office-mcp/canonical_name"],
        "word.update_table"
    );
    assert_eq!(alias["inputSchema"], canonical["inputSchema"]);
    assert_eq!(alias["annotations"], canonical["annotations"]);
    assert_eq!(
        alias["_meta"]["com.office-mcp/side_effects"],
        canonical["_meta"]["com.office-mcp/side_effects"]
    );

    let described = super::describe_tool_contract("word_update_table").expect("alias contract");
    assert_eq!(described["name"], "word_update_table");
    assert_eq!(described["canonical_name"], "word.update_table");
    assert_eq!(described["alias_for"], "word.update_table");
    assert_eq!(described["input_schema"], canonical["inputSchema"]);
}

#[test]
fn office_tool_exposure_parity_covers_all_layers() {
    assert_office_tool_exposure_parity();
}

#[test]
fn shared_office_tool_catalog_path_covers_all_apps() {
    let catalogs = office_tool_catalogs();
    assert_eq!(catalogs.len(), 3);
    assert_eq!(catalogs[0].app(), "word");
    assert_eq!(catalogs[1].app(), "excel");
    assert_eq!(catalogs[2].app(), "powerpoint");

    let all_tools = all_office_tool_names().collect::<Vec<_>>();
    assert_eq!(all_tools.len(), 74);
    assert_eq!(all_tools.iter().copied().collect::<BTreeSet<_>>().len(), 74);
    assert!(all_tools.contains(&"word.update_table"));
    assert!(all_tools.contains(&"excel.write_range"));
    assert!(all_tools.contains(&"powerpoint.add_slide"));

    for catalog in catalogs {
        for tool in catalog.tool_names() {
            assert!(
                catalog.contains(tool),
                "{} should contain {tool}",
                catalog.app()
            );
        }
    }
}

#[test]
fn representative_word_schemas_are_specific() {
    let describe = schema_for("office.describe_tools");
    assert_required(&describe, &["tools"]);
    assert_eq!(describe["properties"]["tools"]["type"], "array");

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
    assert_eq!(image["properties"]["validate_only"]["type"], "boolean");
    assert_eq!(
        image["properties"]["image"]["oneOf"]
            .as_array()
            .expect("image oneOf")
            .len(),
        2
    );

    let resolve_anchor = schema_for("word.resolve_anchor");
    assert_required(&resolve_anchor, &["session_id", "anchor"]);
    assert_eq!(
        resolve_anchor["properties"]["include_text_preview"]["type"],
        "boolean"
    );
    assert_eq!(
        resolve_anchor["properties"]["anchor"]["oneOf"]
            .as_array()
            .expect("anchor oneOf")
            .len(),
        6
    );

    let get_header_footer = schema_for("word.get_header_footer");
    assert_required(&get_header_footer, &["session_id", "location"]);
    assert_eq!(
        get_header_footer["properties"]["location"]["enum"][0],
        "header"
    );
    assert_eq!(
        get_header_footer["properties"]["header_footer_type"]["default"],
        "primary"
    );
    assert_eq!(
        get_header_footer["properties"]["section_index"]["minimum"],
        0
    );

    let update_header_footer = schema_for("word.update_header_footer");
    assert_required(&update_header_footer, &["session_id", "location", "action"]);
    assert_eq!(
        update_header_footer["properties"]["action"]["enum"][0],
        "set_text"
    );
    assert_eq!(
        update_header_footer["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn word_validation_only_schemas_accept_validate_only_flag() {
    for tool in [
        "word.insert_image",
        "word.replace_text",
        "word.update_paragraph",
        "word.delete_range",
        "word.update_header_footer",
    ] {
        let schema = schema_for(tool);
        assert_eq!(
            schema["properties"]["validate_only"]["type"], "boolean",
            "{tool} must advertise validate_only"
        );
    }
}

#[test]
fn word_replace_text_schema_accepts_no_mutation_preview_options() {
    let schema = schema_for("word.replace_text");
    assert_eq!(schema["properties"]["dry_run"]["type"], "boolean");
    assert_eq!(schema["properties"]["validate_only"]["type"], "boolean");
    assert_eq!(schema["properties"]["partial_ok"]["type"], "boolean");
    assert_eq!(schema["properties"]["wildcards"]["type"], "boolean");
    assert_eq!(schema["properties"]["limit"]["type"], "integer");
    assert_eq!(
        schema["properties"]["scope"]["properties"]["selection_only"]["type"],
        "boolean"
    );
    assert_eq!(
        schema["properties"]["scope"]["properties"]["paragraph_range"]["minItems"],
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

#[test]
fn powerpoint_resource_catalog_uses_session_scoped_uris() {
    let resources = powerpoint_resource_catalog_for_session("powerpoint-session");
    let uris = resources
        .iter()
        .filter_map(|resource| resource["uri"].as_str())
        .collect::<Vec<_>>();

    assert!(uris.contains(&"office://powerpoint/powerpoint-session/presentation"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slides"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slides/text?start=0"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slide/0/text"));
    assert!(uris.contains(&"office://powerpoint/powerpoint-session/slide/0/shapes"));
}

#[test]
fn powerpoint_resource_templates_include_read_only_routes() {
    let templates = powerpoint_resource_templates();
    let names = templates
        .iter()
        .filter_map(|template| template["name"].as_str())
        .collect::<Vec<_>>();

    assert!(names.contains(&"powerpoint.presentation.template"));
    assert!(names.contains(&"powerpoint.slides.template"));
    assert!(names.contains(&"powerpoint.slides.text.template"));
    assert!(names.contains(&"powerpoint.slide.text.template"));
    assert!(names.contains(&"powerpoint.slide.shapes.template"));
}

#[test]
fn resource_catalogs_cover_every_parsed_resource_route() {
    let word_resource_catalog = word_resource_catalog_for_session("session-1");
    let word_resource_template_catalog = word_resource_templates();
    let word_resources = resource_uris(&word_resource_catalog);
    let word_templates = resource_template_uris(&word_resource_template_catalog);
    assert_resource_route_covered(
        &word_resources,
        &word_templates,
        "office://word/session-1/document?offset=0&limit=200",
        "office://word/{session_id}/document{?offset,limit}",
    );
    assert_resource_route_covered(
        &word_resources,
        &word_templates,
        "office://word/session-1/structure",
        "office://word/{session_id}/structure",
    );
    assert_resource_route_covered(
        &word_resources,
        &word_templates,
        "office://word/session-1/paragraph/0",
        "office://word/{session_id}/paragraph/{index}",
    );
    assert_resource_route_covered(
        &word_resources,
        &word_templates,
        "office://word/session-1/comments",
        "office://word/{session_id}/comments",
    );
    assert_resource_route_covered(
        &word_resources,
        &word_templates,
        "office://word/session-1/track_changes",
        "office://word/{session_id}/track_changes",
    );
    assert_resource_route_covered(
        &word_resources,
        &word_templates,
        "office://word/session-1/selection",
        "office://word/{session_id}/selection",
    );

    let excel_resource_catalog = excel_resource_catalog_for_session("session-1");
    let excel_resource_template_catalog = excel_resource_templates();
    let excel_resources = resource_uris(&excel_resource_catalog);
    let excel_templates = resource_template_uris(&excel_resource_template_catalog);
    assert_resource_route_covered(
        &excel_resources,
        &excel_templates,
        "office://excel/session-1/workbook",
        "office://excel/{session_id}/workbook",
    );
    assert_resource_route_covered(
        &excel_resources,
        &excel_templates,
        "office://excel/session-1/sheets",
        "office://excel/{session_id}/sheets",
    );
    assert_resource_route_covered(
        &excel_resources,
        &excel_templates,
        "office://excel/session-1/used-range",
        "office://excel/{session_id}/used-range{?sheet}",
    );
    assert_resource_route_covered(
        &excel_resources,
        &excel_templates,
        "office://excel/session-1/range/A1",
        "office://excel/{session_id}/range/{address}{?sheet}",
    );

    assert_powerpoint_resource_routes_covered();
}

fn assert_powerpoint_resource_routes_covered() {
    let resource_catalog = powerpoint_resource_catalog_for_session("session-1");
    let resource_template_catalog = powerpoint_resource_templates();
    let resources = resource_uris(&resource_catalog);
    let templates = resource_template_uris(&resource_template_catalog);
    assert_resource_route_covered(
        &resources,
        &templates,
        "office://powerpoint/session-1/presentation",
        "office://powerpoint/{session_id}/presentation",
    );
    assert_resource_route_covered(
        &resources,
        &templates,
        "office://powerpoint/session-1/slides",
        "office://powerpoint/{session_id}/slides",
    );
    assert_resource_route_covered(
        &resources,
        &templates,
        "office://powerpoint/session-1/slides/text?start=0",
        "office://powerpoint/{session_id}/slides/text{?start,end}",
    );
    assert_resource_route_covered(
        &resources,
        &templates,
        "office://powerpoint/session-1/slide/0/text",
        "office://powerpoint/{session_id}/slide/{index}/text",
    );
    assert_resource_route_covered(
        &resources,
        &templates,
        "office://powerpoint/session-1/slide/0/shapes",
        "office://powerpoint/{session_id}/slide/{index}/shapes",
    );
}

fn schema_for(name: &str) -> Value {
    tool_for(name)["inputSchema"].clone()
}

fn tool_for(name: &str) -> Value {
    tool_catalog_json()
        .into_iter()
        .find(|tool| tool["name"] == name)
        .unwrap_or_else(|| panic!("missing tool {name}"))
}

fn app_for(name: &str) -> &str {
    name.split_once('.')
        .map_or_else(|| panic!("missing app prefix for {name}"), |(app, _)| app)
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

fn resource_uris(resources: &[Value]) -> Vec<String> {
    resources
        .iter()
        .filter_map(|resource| resource["uri"].as_str().map(ToString::to_string))
        .collect()
}

fn resource_template_uris(templates: &[Value]) -> Vec<String> {
    templates
        .iter()
        .filter_map(|template| template["uriTemplate"].as_str().map(ToString::to_string))
        .collect()
}

fn assert_resource_route_covered(
    resources: &[String],
    templates: &[String],
    concrete_uri: &str,
    template_uri: &str,
) {
    assert!(
        resources.iter().any(|uri| uri == concrete_uri)
            || templates.iter().any(|uri| uri == template_uri),
        "missing resource discovery coverage for {concrete_uri} / {template_uri}"
    );
}

fn assert_office_tool_exposure_parity() {
    let app_layers = [
        (
            "word",
            OfficeToolLayers {
                taskpane: parse_taskpane_available_tools(include_str!(
                    "../../../../office-ctl/word/public/taskpane.js"
                )),
                daemon: WORD_V1_TOOLS.iter().map(ToString::to_string).collect(),
                metadata: metadata_tools_for_app("word"),
                implemented_spec: parse_implemented_capability_tools(include_str!(
                    "../../../../../doc/spec/04-word-capabilities.md"
                )),
                tools_list: tools_list_canonical_for_app("word"),
                aliases: tools_list_aliases_for_app("word"),
            },
        ),
        (
            "excel",
            OfficeToolLayers {
                taskpane: parse_taskpane_available_tools(include_str!(
                    "../../../../office-ctl/excel/public/taskpane.js"
                )),
                daemon: ExcelToolCatalog::tools()
                    .iter()
                    .copied()
                    .map(ToString::to_string)
                    .collect(),
                metadata: metadata_tools_for_app("excel"),
                implemented_spec: parse_implemented_capability_tools(include_str!(
                    "../../../../../doc/spec/04-excel-capabilities.md"
                )),
                tools_list: tools_list_canonical_for_app("excel"),
                aliases: tools_list_aliases_for_app("excel"),
            },
        ),
        (
            "powerpoint",
            OfficeToolLayers {
                taskpane: parse_taskpane_available_tools(include_str!(
                    "../../../../office-ctl/powerpoint/public/taskpane.js"
                )),
                daemon: PowerPointToolCatalog::tools()
                    .iter()
                    .copied()
                    .map(ToString::to_string)
                    .collect(),
                metadata: metadata_tools_for_app("powerpoint"),
                implemented_spec: parse_implemented_capability_tools(include_str!(
                    "../../../../../doc/spec/04-powerpoint-capabilities.md"
                )),
                tools_list: tools_list_canonical_for_app("powerpoint"),
                aliases: tools_list_aliases_for_app("powerpoint"),
            },
        ),
    ];

    for (app, layers) in app_layers {
        layers.assert_equal(app);
    }

    assert_tools_list_alias_integrity();
    for alias in [
        "word_update_table",
        "excel_write_range",
        "powerpoint_add_slide",
    ] {
        assert!(
            tool_catalog_json()
                .iter()
                .any(|tool| tool["name"].as_str() == Some(alias)),
            "missing representative mutating alias {alias}"
        );
    }
}

struct OfficeToolLayers {
    taskpane: BTreeSet<String>,
    daemon: BTreeSet<String>,
    metadata: BTreeSet<String>,
    implemented_spec: BTreeSet<String>,
    tools_list: BTreeSet<String>,
    aliases: BTreeSet<String>,
}

impl OfficeToolLayers {
    fn assert_equal(&self, app: &str) {
        let layers = [
            ("taskpane AVAILABLE_TOOLS", &self.taskpane),
            ("daemon catalog", &self.daemon),
            ("metadata catalog", &self.metadata),
            ("implemented spec", &self.implemented_spec),
            ("tools/list canonical", &self.tools_list),
            ("tools/list alias", &self.aliases),
        ];
        let expected = &self.daemon;
        let mut errors = Vec::new();
        for (layer_name, actual) in layers {
            let missing = expected
                .difference(actual)
                .map(String::as_str)
                .collect::<Vec<_>>();
            let extra = actual
                .difference(expected)
                .map(String::as_str)
                .collect::<Vec<_>>();
            if !missing.is_empty() {
                errors.push(format!(
                    "{app} {layer_name} missing tools: {}",
                    missing.join(", ")
                ));
            }
            if !extra.is_empty() {
                errors.push(format!(
                    "{app} {layer_name} extra tools: {}",
                    extra.join(", ")
                ));
            }
        }
        assert!(errors.is_empty(), "{}", errors.join("\n"));
    }
}

fn parse_taskpane_available_tools(source: &'static str) -> BTreeSet<String> {
    let start = source
        .find("const AVAILABLE_TOOLS = [")
        .expect("AVAILABLE_TOOLS declaration");
    let rest = &source[start..];
    let end = rest.find("];").expect("AVAILABLE_TOOLS terminator");
    parse_quoted_tool_names(&rest[..end])
}

fn parse_implemented_capability_tools(markdown: &'static str) -> BTreeSet<String> {
    markdown
        .lines()
        .filter_map(|line| {
            if !line.starts_with("| `") || !line.contains("| implemented |") {
                return None;
            }
            line.split('`').nth(1).map(ToString::to_string)
        })
        .collect()
}

fn parse_quoted_tool_names(source: &'static str) -> BTreeSet<String> {
    source
        .lines()
        .filter_map(|line| {
            let line = line.trim();
            line.strip_prefix('"')
                .and_then(|line| line.split_once('"').map(|(tool, _)| tool))
                .or_else(|| {
                    line.strip_prefix('\'')
                        .and_then(|line| line.split_once('\'').map(|(tool, _)| tool))
                })
        })
        .filter(|tool| tool.contains('.'))
        .map(ToString::to_string)
        .collect()
}

fn metadata_tools_for_app(app: &str) -> BTreeSet<String> {
    crate::mcp::tool_metadata_catalog()
        .iter()
        .filter(|metadata| metadata.app == app)
        .map(|metadata| metadata.name)
        .map(ToString::to_string)
        .collect()
}

fn tools_list_canonical_for_app(app: &str) -> BTreeSet<String> {
    tool_catalog_json()
        .iter()
        .filter_map(|tool| tool["name"].as_str())
        .filter(|name| name.starts_with(app) && name.contains('.'))
        .map(ToString::to_string)
        .collect()
}

fn tools_list_aliases_for_app(app: &str) -> BTreeSet<String> {
    tool_catalog_json()
        .iter()
        .filter_map(|tool| {
            let name = tool["name"].as_str()?;
            let canonical = tool["_meta"]["com.office-mcp/alias_for"].as_str()?;
            (canonical.starts_with(app) && !name.contains('.')).then(|| canonical.to_string())
        })
        .collect()
}

fn assert_tools_list_alias_integrity() {
    let mut aliases = BTreeMap::new();
    for tool in tool_catalog_json() {
        let Some(alias) = tool["name"].as_str().filter(|name| !name.contains('.')) else {
            continue;
        };
        let Some(canonical) = tool["_meta"]["com.office-mcp/alias_for"].as_str() else {
            continue;
        };
        let previous = aliases.insert(alias.to_string(), canonical.to_string());
        assert_eq!(previous, None, "alias collision for {alias}");
        assert_eq!(
            alias,
            super::mcp_safe_tool_alias(canonical),
            "alias {alias} must be derived from {canonical}"
        );
    }

    let canonical_office_tools = WORD_V1_TOOLS
        .iter()
        .copied()
        .chain(ExcelToolCatalog::tools().iter().copied())
        .chain(PowerPointToolCatalog::tools().iter().copied())
        .collect::<BTreeSet<_>>();
    for canonical in canonical_office_tools {
        let alias = super::mcp_safe_tool_alias(canonical);
        assert_eq!(
            aliases.get(&alias),
            Some(&canonical.to_string()),
            "missing or mismatched alias for {canonical}"
        );
    }
}
