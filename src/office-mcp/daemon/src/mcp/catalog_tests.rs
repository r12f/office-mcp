use super::{
    ExcelToolCatalog, PowerPointToolCatalog, WORD_V1_TOOLS, all_office_tool_names,
    excel_resource_catalog_for_session, excel_resource_templates, office_tool_catalogs,
    powerpoint_resource_catalog_for_session, powerpoint_resource_templates, tool_catalog_json,
    word_resource_catalog_for_session, word_resource_templates,
};
use crate::mcp::{ToolSideEffect, tool_metadata_catalog};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};

const POWERPOINT_V1_TOOL_NAMES: &[&str] = &[
    "powerpoint.get_presentation_info",
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
    assert!(!names.contains(&"word.get_paragraph"));
    assert!(!names.contains(&"word_get_paragraph"));
    assert!(names.contains(&"word.get_header_footer"));
    assert!(names.contains(&"word.update_header_footer"));
    assert!(names.contains(&"word.insert_break"));
    assert!(names.contains(&"word.list_sections"));
    assert!(names.contains(&"word.update_page_setup"));
    assert!(names.contains(&"word.resolve_anchor"));
    assert!(names.contains(&"word.set_selection"));
    assert!(names.contains(&"word.insert_hyperlink"));
    assert!(names.contains(&"word.list_hyperlinks"));
    assert!(names.contains(&"word.remove_hyperlink"));
    assert!(names.contains(&"word.insert_bookmark"));
    assert!(names.contains(&"word.list_bookmarks"));
    assert!(names.contains(&"word.delete_bookmark"));
    assert!(names.contains(&"word.list_content_controls"));
    assert!(names.contains(&"word.insert_content_control"));
    assert!(names.contains(&"word.update_content_control"));
    assert!(names.contains(&"word.delete_content_control"));
    assert!(names.contains(&"word.insert_note"));
    assert!(names.contains(&"word.list_notes"));
    assert!(names.contains(&"word.update_note"));
    assert!(names.contains(&"word.delete_note"));
    assert!(names.contains(&"word.list_fields"));
    assert!(names.contains(&"word.insert_field"));
    assert!(names.contains(&"word.update_field"));
    assert!(names.contains(&"word.delete_field"));
    assert!(names.contains(&"word.get_document_properties"));
    assert!(names.contains(&"word.update_document_properties"));
    assert!(names.contains(&"word.list_styles"));
    assert!(names.contains(&"word.create_style"));
    assert!(names.contains(&"word.update_style"));
    assert!(names.contains(&"word.list_images"));
    assert!(names.contains(&"word.get_image"));
    assert!(names.contains(&"word.update_image"));
    assert!(!names.contains(&"word.resize_image"));
    assert!(!names.contains(&"word_resize_image"));
    assert!(!names.contains(&"word.delete_image"));
    assert!(!names.contains(&"word_delete_image"));
    assert!(names.contains(&"word.update_table"));
    assert!(names.contains(&"word_update_table"));
    assert!(names.contains(&"word.set_change_tracking"));
    assert!(names.contains(&"word.update_tracked_change"));
    assert!(!names.contains(&"word.insert_heading"));
    assert!(!names.contains(&"word.set_heading_level"));
    assert!(!names.contains(&"word.update_cell"));
    assert!(!names.contains(&"word.add_row"));
    assert!(!names.contains(&"word.add_column"));
    assert!(!names.contains(&"word.format_cell"));
    assert!(!names.contains(&"word.accept_change"));
    assert!(!names.contains(&"word.reject_change"));
    assert!(!names.contains(&"word.insert_page_break"));
    assert!(names.contains(&"excel.read_range"));
    assert!(names.contains(&"excel_read_range"));
    assert!(names.contains(&"excel.save"));
    assert!(names.contains(&"excel.calculate"));
    assert!(names.contains(&"excel.list_named_items"));
    assert!(names.contains(&"excel.update_named_item"));
    assert!(names.contains(&"excel.get_document_properties"));
    assert!(names.contains(&"excel.update_document_properties"));
    assert!(names.contains(&"excel.add_comment"));
    assert!(names.contains(&"excel.list_comments"));
    assert!(names.contains(&"excel.update_comment"));
    assert!(names.contains(&"excel.sort_range"));
    assert!(names.contains(&"excel.apply_filter"));
    assert!(names.contains(&"excel.update_table"));
    assert!(names.contains(&"excel.update_chart"));
    assert!(names.contains(&"excel.create_pivot_table"));
    assert!(names.contains(&"excel.update_pivot_table"));
    assert!(names.contains(&"excel.insert_image"));
    assert!(names.contains(&"excel.list_shapes"));
    assert!(names.contains(&"excel.update_shape"));
    for name in POWERPOINT_V1_TOOL_NAMES {
        assert!(names.contains(name), "missing PowerPoint tool {name}");
    }
    assert!(!names.contains(&"powerpoint.add_text_box"));
    assert!(!names.contains(&"powerpoint_add_text_box"));
    assert!(!names.contains(&"powerpoint.export_pdf"));
    assert!(!names.contains(&"powerpoint.duplicate_slide"));
    assert!(!names.contains(&"powerpoint.set_slide_background"));
    assert_eq!(WORD_V1_TOOLS.len(), 62);
    assert_eq!(ExcelToolCatalog::tools().len(), 36);
    assert_eq!(PowerPointToolCatalog::tools().len(), 23);
    assert_eq!(tools.len(), 248);
}

#[test]
fn tools_list_exposes_action_side_effects_for_mixed_owner_tools() {
    let tools = tool_catalog_json();
    let word_update_table = tools
        .iter()
        .find(|tool| tool["name"] == "word.update_table")
        .expect("word.update_table tool");
    assert_eq!(
        word_update_table["_meta"]["com.office-mcp/action_side_effects"]["update_cell"],
        "mutating"
    );
    assert_eq!(
        word_update_table["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );

    let excel_update_table = tools
        .iter()
        .find(|tool| tool["name"] == "excel.update_table")
        .expect("excel.update_table tool");
    assert_eq!(
        excel_update_table["_meta"]["com.office-mcp/action_side_effects"]["read"],
        "read"
    );
    assert_eq!(
        excel_update_table["_meta"]["com.office-mcp/action_side_effects"]["add_rows"],
        "mutating"
    );

    let excel_update_comment = tools
        .iter()
        .find(|tool| tool["name"] == "excel.update_comment")
        .expect("excel.update_comment tool");
    assert_eq!(
        excel_update_comment["_meta"]["com.office-mcp/action_side_effects"]["reply"],
        "mutating"
    );
    assert_eq!(
        excel_update_comment["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );

    let excel_update_conditional_format = tools
        .iter()
        .find(|tool| tool["name"] == "excel.update_conditional_format")
        .expect("excel.update_conditional_format tool");
    assert_eq!(
        excel_update_conditional_format["_meta"]["com.office-mcp/action_side_effects"]["add"],
        "mutating"
    );
    assert_eq!(
        excel_update_conditional_format["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );
    assert_eq!(
        excel_update_conditional_format["_meta"]["com.office-mcp/action_side_effects"]["clear_range"],
        "destructive"
    );

    let excel_copy_range = tools
        .iter()
        .find(|tool| tool["name"] == "excel.copy_range")
        .expect("excel.copy_range tool");
    assert_eq!(
        excel_copy_range["_meta"]["com.office-mcp/action_side_effects"]["copy"],
        "mutating"
    );
    assert_eq!(
        excel_copy_range["_meta"]["com.office-mcp/action_side_effects"]["autofill"],
        "mutating"
    );

    let excel_update_shape = tools
        .iter()
        .find(|tool| tool["name"] == "excel.update_shape")
        .expect("excel.update_shape tool");
    assert_eq!(
        excel_update_shape["_meta"]["com.office-mcp/action_side_effects"]["move"],
        "mutating"
    );
    assert_eq!(
        excel_update_shape["_meta"]["com.office-mcp/action_side_effects"]["set_z_order"],
        "mutating"
    );
    assert_eq!(
        excel_update_shape["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );

    let powerpoint_update_tags = tools
        .iter()
        .find(|tool| tool["name"] == "powerpoint.update_tags")
        .expect("powerpoint.update_tags tool");
    assert_eq!(
        powerpoint_update_tags["_meta"]["com.office-mcp/action_side_effects"]["list"],
        "read"
    );
    assert_eq!(
        powerpoint_update_tags["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );
}

#[test]
fn action_side_effect_metadata_covers_advertised_action_enum() {
    for metadata in tool_metadata_catalog() {
        let Some(action_side_effects) = metadata.action_side_effects else {
            continue;
        };
        let schema = schema_for(metadata.name);
        let schema_actions = schema["properties"]["action"]["enum"]
            .as_array()
            .unwrap_or_else(|| panic!("{} must advertise action enum", metadata.name))
            .iter()
            .map(|action| action.as_str().expect("action string"))
            .collect::<BTreeSet<_>>();
        let metadata_actions = action_side_effects
            .iter()
            .map(|entry| entry.action)
            .collect::<BTreeSet<_>>();
        assert_eq!(
            schema_actions, metadata_actions,
            "{} action map",
            metadata.name
        );
    }
}

#[test]
fn word_document_property_tools_expose_expected_contracts() {
    let get_properties = tool_for("word.get_document_properties");
    assert_eq!(
        get_properties["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(
        get_properties["inputSchema"]["properties"]["include_custom"]["type"],
        "boolean"
    );
    assert_eq!(
        get_properties["inputSchema"]["properties"]["include_custom"]["default"],
        true
    );
    assert_eq!(
        get_properties["_meta"]["com.office-mcp/side_effects"],
        "read"
    );

    let update_properties = tool_for("word.update_document_properties");
    assert_eq!(
        update_properties["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(
        update_properties["inputSchema"]["anyOf"],
        serde_json::json!([
            { "required": ["title"] },
            { "required": ["subject"] },
            { "required": ["author"] },
            { "required": ["keywords"] },
            { "required": ["category"] },
            { "required": ["comments"] },
            { "required": ["company"] },
            { "required": ["manager"] },
            { "required": ["custom_set"] },
            { "required": ["custom_delete"] }
        ])
    );
    assert_eq!(
        update_properties["inputSchema"]["properties"]["custom_set"]["items"]["required"],
        serde_json::json!(["key", "value"])
    );
    assert_eq!(
        update_properties["inputSchema"]["properties"]["custom_set"]["items"]["properties"]["key"]
            ["minLength"],
        1
    );
    assert_eq!(
        update_properties["inputSchema"]["properties"]["custom_set"]["items"]["properties"]["value"]
            ["oneOf"][2]["type"],
        "boolean"
    );
    assert_eq!(
        update_properties["inputSchema"]["properties"]["custom_delete"]["items"]["minLength"],
        1
    );
    assert_eq!(
        update_properties["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );
}

#[test]
fn word_image_tools_expose_expected_contracts() {
    let list_images = tool_for("word.list_images");
    assert_eq!(
        list_images["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(list_images["_meta"]["com.office-mcp/side_effects"], "read");

    let get_image = tool_for("word.get_image");
    assert_eq!(
        get_image["inputSchema"]["required"],
        serde_json::json!(["session_id", "image"])
    );
    assert_eq!(
        get_image["inputSchema"]["properties"]["image"]["properties"]["kind"]["const"],
        "paragraph_index"
    );
    assert_eq!(
        get_image["inputSchema"]["properties"]["image"]["properties"]["image_index"]["default"],
        0
    );
    assert_eq!(get_image["_meta"]["com.office-mcp/side_effects"], "read");

    let update_image = tool_for("word.update_image");
    assert_eq!(
        update_image["inputSchema"]["required"],
        serde_json::json!(["session_id", "image", "action"])
    );
    assert_eq!(
        update_image["inputSchema"]["properties"]["action"]["enum"],
        serde_json::json!([
            "resize",
            "set_alt_text",
            "set_hyperlink",
            "replace",
            "delete"
        ])
    );
    assert_eq!(
        update_image["inputSchema"]["properties"]["base64"]["type"],
        "string"
    );
    assert_eq!(
        update_image["inputSchema"]["properties"]["width_pt"]["exclusiveMinimum"],
        0
    );
    assert_eq!(
        update_image["inputSchema"]["properties"]["validate_only"]["type"],
        "boolean"
    );
    assert_eq!(
        update_image["_meta"]["com.office-mcp/side_effects"],
        "destructive"
    );

    let list_images_schema = schema_for("word.list_images");
    assert_required(&list_images_schema, &["session_id"]);

    let get_image_schema = schema_for("word.get_image");
    assert_required(&get_image_schema, &["session_id", "image"]);
    assert_eq!(
        get_image_schema["properties"]["image"]["properties"]["kind"]["const"],
        "paragraph_index"
    );
    assert_eq!(
        get_image_schema["properties"]["image"]["properties"]["index"]["minimum"],
        0
    );
    assert_eq!(
        get_image_schema["properties"]["image"]["properties"]["image_index"]["default"],
        0
    );

    let update_image_schema = schema_for("word.update_image");
    assert_required(&update_image_schema, &["session_id", "image", "action"]);
    assert_eq!(
        update_image_schema["properties"]["action"]["enum"],
        serde_json::json!([
            "resize",
            "set_alt_text",
            "set_hyperlink",
            "replace",
            "delete"
        ])
    );
    assert_eq!(
        update_image_schema["properties"]["hyperlink"]["format"],
        "uri"
    );
    assert_eq!(
        update_image_schema["properties"]["base64"]["type"],
        "string"
    );
    assert_eq!(
        update_image_schema["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn word_shape_tools_expose_expected_contracts() {
    let list_shapes = tool_for("word.list_shapes");
    assert_eq!(
        list_shapes["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(list_shapes["_meta"]["com.office-mcp/side_effects"], "read");

    let insert_shape = tool_for("word.insert_shape");
    assert_eq!(
        insert_shape["inputSchema"]["required"],
        serde_json::json!(["session_id", "shape_type"])
    );
    assert_eq!(
        insert_shape["inputSchema"]["properties"]["shape_type"]["enum"],
        serde_json::json!([
            "text_box",
            "rectangle",
            "ellipse",
            "rounded_rectangle",
            "line",
            "picture"
        ])
    );
    assert_eq!(
        insert_shape["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );

    let update_shape = tool_for("word.update_shape");
    assert_eq!(
        update_shape["inputSchema"]["required"],
        serde_json::json!(["session_id", "shape_id", "action"])
    );
    assert_eq!(
        update_shape["inputSchema"]["properties"]["action"]["enum"],
        serde_json::json!([
            "move",
            "resize",
            "set_text",
            "set_alt_text",
            "set_fill",
            "set_line",
            "set_wrap",
            "set_visibility"
        ])
    );
    assert_eq!(
        update_shape["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );

    let delete_shape = tool_for("word.delete_shape");
    assert_eq!(
        delete_shape["inputSchema"]["required"],
        serde_json::json!(["session_id", "shape_id"])
    );
    assert_eq!(
        delete_shape["_meta"]["com.office-mcp/side_effects"],
        "destructive"
    );

    let list_shapes_schema = schema_for("word.list_shapes");
    assert_required(&list_shapes_schema, &["session_id"]);
    assert_eq!(
        list_shapes_schema["properties"]["scope"]["enum"],
        serde_json::json!(["body", "paragraph", "anchor"])
    );

    let insert_shape_schema = schema_for("word.insert_shape");
    assert_required(&insert_shape_schema, &["session_id", "shape_type"]);
    assert_eq!(
        insert_shape_schema["properties"]["image"]["oneOf"][0]["required"],
        serde_json::json!(["base64"])
    );

    let update_shape_schema = schema_for("word.update_shape");
    assert_required(&update_shape_schema, &["session_id", "shape_id", "action"]);
    assert_eq!(
        update_shape_schema["properties"]["wrap_type"]["enum"],
        serde_json::json!(["inline", "square", "tight", "behind", "front", "top_bottom"])
    );

    let delete_shape_schema = schema_for("word.delete_shape");
    assert_required(&delete_shape_schema, &["session_id", "shape_id"]);
    assert_eq!(
        delete_shape_schema["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn word_style_tools_expose_expected_contracts() {
    let list_styles = tool_for("word.list_styles");
    assert_eq!(
        list_styles["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(
        list_styles["inputSchema"]["properties"]["type"]["enum"],
        serde_json::json!(["paragraph", "character", "table", "list"])
    );
    assert_eq!(
        list_styles["inputSchema"]["properties"]["in_use_only"]["type"],
        "boolean"
    );
    assert_eq!(list_styles["_meta"]["com.office-mcp/side_effects"], "read");

    let create_style = tool_for("word.create_style");
    assert_eq!(
        create_style["inputSchema"]["required"],
        serde_json::json!(["session_id", "name", "type"])
    );
    assert_eq!(
        create_style["inputSchema"]["properties"]["name"]["minLength"],
        1
    );
    assert_eq!(
        create_style["inputSchema"]["properties"]["font"]["properties"]["bold"]["type"],
        "boolean"
    );
    assert_eq!(
        create_style["inputSchema"]["properties"]["paragraph"]["properties"]["alignment"]["enum"],
        serde_json::json!(["left", "center", "right", "justified"])
    );
    assert_eq!(
        create_style["inputSchema"]["properties"]["validate_only"]["type"],
        "boolean"
    );

    let update_style = tool_for("word.update_style");
    assert_eq!(
        update_style["inputSchema"]["required"],
        serde_json::json!(["session_id", "name"])
    );
    assert_eq!(
        update_style["inputSchema"]["anyOf"],
        serde_json::json!([
            { "required": ["base_style"] },
            { "required": ["font"] },
            { "required": ["paragraph"] }
        ])
    );
    assert_eq!(
        update_style["inputSchema"]["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn word_field_tools_expose_expected_contracts() {
    let list_fields = tool_for("word.list_fields");
    assert_eq!(
        list_fields["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(
        list_fields["inputSchema"]["properties"]["limit"]["minimum"],
        1
    );
    assert_eq!(
        list_fields["inputSchema"]["properties"]["limit"]["maximum"],
        200
    );
    assert_eq!(list_fields["_meta"]["com.office-mcp/side_effects"], "read");

    let insert_field = tool_for("word.insert_field");
    assert_eq!(
        insert_field["inputSchema"]["required"],
        serde_json::json!(["session_id", "anchor", "field_type"])
    );
    assert_eq!(
        insert_field["inputSchema"]["properties"]["field_type"]["enum"],
        serde_json::json!([
            "toc",
            "page",
            "num_pages",
            "date",
            "time",
            "ref",
            "hyperlink",
            "seq",
            "styleref"
        ])
    );
    assert_eq!(
        insert_field["inputSchema"]["properties"]["validate_only"]["type"],
        "boolean"
    );

    let update_field = tool_for("word.update_field");
    assert_eq!(
        update_field["inputSchema"]["required"],
        serde_json::json!(["session_id", "action"])
    );
    assert_eq!(
        update_field["inputSchema"]["properties"]["action"]["enum"],
        serde_json::json!(["refresh", "refresh_all", "lock", "unlock"])
    );
    assert_eq!(
        update_field["inputSchema"]["allOf"][0]["then"]["required"],
        serde_json::json!(["field_index"])
    );
    assert_eq!(
        update_field["inputSchema"]["allOf"][1]["then"]["required"],
        serde_json::json!(["expected_count"])
    );

    let delete_field = tool_for("word.delete_field");
    assert_eq!(
        delete_field["inputSchema"]["required"],
        serde_json::json!(["session_id", "field_index"])
    );
    assert_eq!(
        delete_field["_meta"]["com.office-mcp/side_effects"],
        "destructive"
    );
}

#[test]
fn word_change_tracking_tools_expose_expected_contracts() {
    let set_change_tracking = tool_for("word.set_change_tracking");
    assert_eq!(
        set_change_tracking["inputSchema"]["required"],
        serde_json::json!(["session_id", "mode"])
    );
    assert_eq!(
        set_change_tracking["inputSchema"]["properties"]["mode"]["enum"],
        serde_json::json!(["off", "track_all", "track_mine_only"])
    );
    assert_eq!(
        set_change_tracking["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );

    let update_tracked_change = tool_for("word.update_tracked_change");
    assert_eq!(
        update_tracked_change["inputSchema"]["required"],
        serde_json::json!(["session_id", "action"])
    );
    assert_eq!(
        update_tracked_change["inputSchema"]["properties"]["action"]["enum"],
        serde_json::json!(["accept", "reject", "accept_all", "reject_all"])
    );
    assert_eq!(
        update_tracked_change["inputSchema"]["properties"]["expected_count"]["minimum"],
        0
    );
    assert_eq!(
        update_tracked_change["inputSchema"]["allOf"][0]["then"]["required"],
        serde_json::json!(["change_index", "expected_fingerprint"])
    );
    assert_eq!(
        update_tracked_change["inputSchema"]["allOf"][1]["then"]["required"],
        serde_json::json!(["expected_count"])
    );
}

#[test]
fn word_section_tools_expose_expected_contracts() {
    let insert_break = tool_for("word.insert_break");
    assert_eq!(
        insert_break["inputSchema"]["required"],
        serde_json::json!(["session_id", "anchor"])
    );
    assert_eq!(
        insert_break["inputSchema"]["properties"]["break_type"]["enum"],
        serde_json::json!([
            "page",
            "line",
            "section_next",
            "section_continuous",
            "section_even",
            "section_odd"
        ])
    );
    assert_eq!(
        insert_break["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );

    let list_sections = tool_for("word.list_sections");
    assert_eq!(
        list_sections["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(
        list_sections["inputSchema"]["properties"]["include_page_setup"]["type"],
        "boolean"
    );
    assert_eq!(
        list_sections["_meta"]["com.office-mcp/side_effects"],
        "read"
    );

    let update_page_setup = tool_for("word.update_page_setup");
    assert_eq!(
        update_page_setup["inputSchema"]["required"],
        serde_json::json!(["session_id"])
    );
    assert_eq!(
        update_page_setup["inputSchema"]["properties"]["orientation"]["enum"],
        serde_json::json!(["portrait", "landscape"])
    );
    assert_eq!(
        update_page_setup["inputSchema"]["properties"]["margins_pt"]["properties"]["top"]["minimum"],
        0
    );
    assert_eq!(
        update_page_setup["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );
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
fn mutating_tools_advertise_validate_only_contract() {
    for name in mutating_tool_names() {
        let schema = schema_for(name);
        assert_eq!(
            schema["properties"]["validate_only"]["type"], "boolean",
            "{name} must accept validate_only"
        );

        let tool = tool_for(name);
        assert_eq!(
            tool["_meta"]["com.office-mcp/supports_validate_only"], true,
            "{name} tools/list metadata must advertise validate_only support"
        );

        let described = super::describe_tool_contract(name).expect("described tool");
        assert_eq!(
            described["supports_validate_only"], true,
            "{name} describe_tools metadata must advertise validate_only support"
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
    assert_eq!(all_tools.len(), 121);
    assert_eq!(
        all_tools.iter().copied().collect::<BTreeSet<_>>().len(),
        121
    );
    assert!(all_tools.contains(&"word.update_comment"));
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
#[allow(clippy::too_many_lines)]
fn representative_word_schemas_are_specific() {
    let describe = schema_for("office.describe_tools");
    assert_required(&describe, &["tools"]);
    assert_eq!(describe["properties"]["tools"]["type"], "array");

    let text = schema_for("word.get_text");
    assert_required(&text, &["session_id"]);
    assert_eq!(text["properties"]["offset"]["type"], "integer");
    assert_eq!(text["properties"]["limit"]["maximum"], 1000);
    assert_eq!(text["properties"]["include_metadata"]["type"], "boolean");
    assert_eq!(text["properties"]["include_formatting"]["type"], "boolean");
    assert!(text["properties"].get("paragraph_index").is_none());

    let apply_formatting = schema_for("word.apply_formatting");
    assert_required(&apply_formatting, &["session_id", "anchor"]);
    assert!(
        !apply_formatting["required"]
            .as_array()
            .expect("required")
            .iter()
            .any(|value| value == "formatting"),
        "word.apply_formatting should allow paragraph-only formatting"
    );
    assert_eq!(
        apply_formatting["properties"]["paragraph"]["properties"]["alignment"]["enum"][1],
        "center"
    );
    assert_eq!(
        apply_formatting["properties"]["paragraph"]["properties"]["line_spacing_pt"]["exclusiveMinimum"],
        0
    );
    assert!(
        apply_formatting
            .get("anyOf")
            .and_then(|value| value.as_array())
            .is_some_and(|items| items.len() == 2),
        "schema should require at least formatting or paragraph"
    );

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

    let update_comment = schema_for("word.update_comment");
    assert_required(&update_comment, &["session_id", "comment_id", "action"]);
    assert_eq!(update_comment["properties"]["comment_id"]["minLength"], 1);
    assert_eq!(update_comment["properties"]["reply_id"]["minLength"], 1);
    assert_eq!(update_comment["properties"]["action"]["enum"][0], "reply");
    assert_eq!(update_comment["properties"]["action"]["enum"][3], "reopen");
    assert_eq!(
        update_comment["properties"]["validate_only"]["type"],
        "boolean"
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

    let set_selection = schema_for("word.set_selection");
    assert_required(&set_selection, &["session_id", "anchor"]);
    assert_eq!(set_selection["properties"]["mode"]["default"], "select");
    assert_eq!(
        set_selection["properties"]["mode"]["enum"][1],
        "cursor_start"
    );
    assert!(set_selection["properties"].get("extent").is_some());
    assert_eq!(
        set_selection["properties"]["anchor"]["oneOf"]
            .as_array()
            .expect("anchor oneOf")
            .len(),
        6
    );

    let get_html = schema_for("word.get_html");
    assert_required(&get_html, &["session_id"]);
    assert!(get_html["properties"].get("anchor").is_some());
    assert!(get_html["properties"].get("extent").is_some());
    assert_eq!(
        get_html["properties"]["anchor"]["oneOf"]
            .as_array()
            .expect("anchor oneOf")
            .len(),
        6
    );

    let insert_html = schema_for("word.insert_html");
    assert_required(&insert_html, &["session_id", "anchor", "html"]);
    assert_eq!(insert_html["properties"]["html"]["minLength"], 1);
    assert_eq!(insert_html["properties"]["html"]["maxLength"], 1_000_000);
    assert_eq!(
        insert_html["properties"]["insert_location"]["enum"]
            .as_array()
            .expect("insert location enum")
            .len(),
        5
    );
    assert_eq!(
        insert_html["properties"]["insert_location"]["default"],
        "after"
    );
    assert_eq!(
        insert_html["properties"]["validate_only"]["type"],
        "boolean"
    );
    assert_eq!(
        insert_html["properties"]["anchor"]["oneOf"]
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
fn word_update_table_schema_covers_complete_table_mutations() {
    let update_table = schema_for("word.update_table");
    assert_required(&update_table, &["session_id", "table_index", "action"]);
    let actions = update_table["properties"]["action"]["enum"]
        .as_array()
        .expect("word.update_table action enum")
        .iter()
        .filter_map(|value| value.as_str())
        .collect::<BTreeSet<_>>();
    for action in [
        "delete_row",
        "delete_column",
        "merge_cells",
        "set_column_width",
        "distribute_columns",
        "set_borders",
        "set_header_row",
    ] {
        assert!(actions.contains(action), "missing table action {action}");
    }
    assert_eq!(update_table["properties"]["row_range"]["type"], "array");
    assert_eq!(update_table["properties"]["row_range"]["minItems"], 2);
    assert_eq!(update_table["properties"]["col_range"]["maxItems"], 2);
    assert_eq!(
        update_table["properties"]["width_pt"]["exclusiveMinimum"],
        0
    );
    assert_eq!(update_table["properties"]["header_row"]["type"], "boolean");
    assert_eq!(update_table["properties"]["borders"]["type"], "object");
    assert_eq!(
        update_table["properties"]["borders"]["properties"]["style"]["enum"][0],
        "single"
    );
}

#[test]
fn word_hyperlink_schemas_are_specific() {
    let insert_hyperlink = schema_for("word.insert_hyperlink");
    assert_required(&insert_hyperlink, &["session_id", "anchor", "url"]);
    assert_eq!(insert_hyperlink["properties"]["url"]["type"], "string");
    assert_eq!(insert_hyperlink["properties"]["text"]["type"], "string");
    assert_eq!(
        insert_hyperlink["properties"]["validate_only"]["type"],
        "boolean"
    );
    assert_eq!(
        insert_hyperlink["properties"]["anchor"]["oneOf"]
            .as_array()
            .expect("anchor oneOf")
            .len(),
        6
    );

    let list_hyperlinks = schema_for("word.list_hyperlinks");
    assert_required(&list_hyperlinks, &["session_id"]);
    assert_eq!(list_hyperlinks["properties"]["offset"]["minimum"], 0);
    assert_eq!(list_hyperlinks["properties"]["limit"]["minimum"], 1);
    assert_eq!(list_hyperlinks["properties"]["limit"]["maximum"], 200);

    let remove_hyperlink = schema_for("word.remove_hyperlink");
    assert_required(&remove_hyperlink, &["session_id", "anchor"]);
    assert_eq!(
        remove_hyperlink["properties"]["keep_text"]["type"],
        "boolean"
    );
    assert_eq!(remove_hyperlink["properties"]["keep_text"]["default"], true);
}

#[test]
fn word_bookmark_schemas_are_specific() {
    let insert_bookmark = schema_for("word.insert_bookmark");
    assert_required(&insert_bookmark, &["session_id", "name", "anchor"]);
    assert_eq!(
        insert_bookmark["properties"]["name"]["pattern"],
        "^[A-Za-z_][A-Za-z0-9_]{0,39}$"
    );
    assert_eq!(
        insert_bookmark["properties"]["overwrite"]["type"],
        "boolean"
    );
    assert!(anchor_kinds(&insert_bookmark).contains(&"bookmark"));

    let list_bookmarks = schema_for("word.list_bookmarks");
    assert_required(&list_bookmarks, &["session_id"]);
    assert_eq!(
        list_bookmarks["properties"]["include_hidden"]["type"],
        "boolean"
    );

    let delete_bookmark = schema_for("word.delete_bookmark");
    assert_required(&delete_bookmark, &["session_id", "name"]);
    assert_eq!(delete_bookmark["properties"]["name"]["minLength"], 1);
}

#[test]
fn word_content_control_schemas_cover_typed_form_controls() {
    let list_controls = schema_for("word.list_content_controls");
    assert_required(&list_controls, &["session_id"]);
    assert_eq!(list_controls["properties"]["type"]["enum"][2], "checkbox");
    assert_eq!(
        list_controls["properties"]["type"]["enum"][3],
        "dropdown_list"
    );
    assert_eq!(list_controls["properties"]["type"]["enum"][4], "combo_box");

    let insert_control = schema_for("word.insert_content_control");
    assert_required(&insert_control, &["session_id"]);
    assert_eq!(insert_control["properties"]["type"]["enum"][2], "checkbox");
    assert_eq!(insert_control["properties"]["checked"]["type"], "boolean");
    assert_eq!(insert_control["properties"]["list_items"]["minItems"], 1);
    assert_eq!(
        insert_control["properties"]["list_items"]["items"]["required"],
        serde_json::json!(["display_text"])
    );
    assert_eq!(
        insert_control["properties"]["list_items"]["items"]["properties"]["display_text"]["minLength"],
        1
    );

    let update_control = schema_for("word.update_content_control");
    assert_required(&update_control, &["session_id", "content_control_id"]);
    assert_eq!(update_control["properties"]["checked"]["type"], "boolean");
    assert_eq!(
        update_control["properties"]["selected_value"]["type"],
        "string"
    );
    assert_eq!(
        update_control["properties"]["list_items_add"]["items"]["properties"]["index"]["minimum"],
        0
    );
    assert_eq!(
        update_control["properties"]["list_items_delete"]["items"]["type"],
        "string"
    );
    assert_eq!(
        update_control["properties"]["list_items_clear"]["type"],
        "boolean"
    );
}

#[test]
fn word_note_schemas_are_specific() {
    let insert_note = schema_for("word.insert_note");
    assert_required(&insert_note, &["session_id", "anchor", "kind", "text"]);
    assert_eq!(insert_note["properties"]["kind"]["enum"][0], "footnote");
    assert_eq!(insert_note["properties"]["kind"]["enum"][1], "endnote");
    assert_eq!(insert_note["properties"]["text"]["minLength"], 1);
    assert_eq!(
        insert_note["properties"]["validate_only"]["type"],
        "boolean"
    );
    assert_eq!(
        insert_note["properties"]["anchor"]["oneOf"]
            .as_array()
            .expect("anchor oneOf")
            .len(),
        6
    );

    let list_notes = schema_for("word.list_notes");
    assert_required(&list_notes, &["session_id", "kind"]);
    assert_eq!(list_notes["properties"]["offset"]["minimum"], 0);
    assert_eq!(list_notes["properties"]["limit"]["minimum"], 1);
    assert_eq!(list_notes["properties"]["limit"]["maximum"], 200);

    let update_note = schema_for("word.update_note");
    assert_required(&update_note, &["session_id", "kind", "index", "text"]);
    assert_eq!(update_note["properties"]["index"]["minimum"], 0);
    assert_eq!(
        update_note["properties"]["validate_only"]["type"],
        "boolean"
    );

    let delete_note = schema_for("word.delete_note");
    assert_required(&delete_note, &["session_id", "kind", "index"]);
    assert_eq!(delete_note["properties"]["index"]["minimum"], 0);
    assert_eq!(
        delete_note["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn word_field_schemas_are_specific() {
    let list_fields = schema_for("word.list_fields");
    assert_required(&list_fields, &["session_id"]);
    assert_eq!(list_fields["properties"]["offset"]["minimum"], 0);
    assert_eq!(list_fields["properties"]["limit"]["minimum"], 1);
    assert_eq!(list_fields["properties"]["limit"]["maximum"], 200);
    assert_eq!(list_fields["properties"]["type"]["type"], "string");

    let insert_field = schema_for("word.insert_field");
    assert_required(&insert_field, &["session_id", "anchor", "field_type"]);
    assert_eq!(insert_field["properties"]["field_type"]["enum"][0], "toc");
    assert_eq!(insert_field["properties"]["code_options"]["type"], "string");
    assert_eq!(
        insert_field["properties"]["validate_only"]["type"],
        "boolean"
    );
    assert!(anchor_kinds(&insert_field).contains(&"bookmark"));

    let update_field = schema_for("word.update_field");
    assert_required(&update_field, &["session_id", "action"]);
    assert_eq!(update_field["properties"]["field_index"]["minimum"], 0);
    assert_eq!(update_field["properties"]["expected_count"]["minimum"], 0);
    assert_eq!(
        update_field["properties"]["validate_only"]["type"],
        "boolean"
    );

    let delete_field = schema_for("word.delete_field");
    assert_required(&delete_field, &["session_id", "field_index"]);
    assert_eq!(delete_field["properties"]["field_index"]["minimum"], 0);
    assert_eq!(
        delete_field["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn word_style_schemas_are_specific() {
    let list_styles = schema_for("word.list_styles");
    assert_required(&list_styles, &["session_id"]);
    assert_eq!(list_styles["properties"]["type"]["enum"][0], "paragraph");
    assert_eq!(list_styles["properties"]["built_in"]["type"], "boolean");
    assert_eq!(list_styles["properties"]["in_use_only"]["type"], "boolean");

    let create_style = schema_for("word.create_style");
    assert_required(&create_style, &["session_id", "name", "type"]);
    assert_eq!(create_style["properties"]["name"]["minLength"], 1);
    assert_eq!(create_style["properties"]["type"]["enum"][1], "character");
    assert_eq!(
        create_style["properties"]["font"]["properties"]["color"]["type"],
        "string"
    );
    assert_eq!(
        create_style["properties"]["paragraph"]["properties"]["outline_level"]["maximum"],
        9
    );

    let update_style = schema_for("word.update_style");
    assert_required(&update_style, &["session_id", "name"]);
    assert_eq!(update_style["properties"]["base_style"]["minLength"], 1);
    assert_eq!(
        update_style["anyOf"][0]["required"],
        serde_json::json!(["base_style"])
    );
}

#[test]
fn word_document_property_schemas_are_specific() {
    let get_properties = schema_for("word.get_document_properties");
    assert_required(&get_properties, &["session_id"]);
    assert_eq!(
        get_properties["properties"]["include_custom"]["type"],
        "boolean"
    );
    assert_eq!(
        get_properties["properties"]["include_custom"]["default"],
        true
    );

    let update_properties = schema_for("word.update_document_properties");
    assert_required(&update_properties, &["session_id"]);
    assert_eq!(update_properties["properties"]["title"]["type"], "string");
    assert_eq!(update_properties["properties"]["manager"]["type"], "string");
    assert!(update_properties["properties"].get("last_author").is_none());
    assert!(
        update_properties["properties"]
            .get("revision_number")
            .is_none()
    );
    assert_eq!(
        update_properties["properties"]["custom_set"]["items"]["additionalProperties"],
        false
    );
    assert_eq!(
        update_properties["properties"]["custom_set"]["items"]["properties"]["value"]["oneOf"][0]["type"],
        "string"
    );
    assert_eq!(
        update_properties["properties"]["custom_delete"]["items"]["minLength"],
        1
    );
    assert_eq!(
        update_properties["anyOf"][9]["required"],
        serde_json::json!(["custom_delete"])
    );
}

#[test]
fn word_validation_only_schemas_accept_validate_only_flag() {
    for tool in [
        "word.insert_image",
        "word.update_image",
        "word.insert_hyperlink",
        "word.insert_note",
        "word.replace_text",
        "word.update_paragraph",
        "word.update_note",
        "word.insert_field",
        "word.update_field",
        "word.create_style",
        "word.update_style",
        "word.delete_range",
        "word.delete_note",
        "word.delete_field",
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
    assert_eq!(range["properties"]["include_hyperlinks"]["type"], "boolean");

    let insert_range = schema_for("excel.insert_range");
    assert_required(&insert_range, &["session_id", "address", "shift"]);
    assert_eq!(
        insert_range["properties"]["shift"]["enum"],
        serde_json::json!(["down", "right"])
    );
    assert_eq!(insert_range["properties"]["count"]["minimum"], 1);
    assert_eq!(
        insert_range["properties"]["validate_only"]["type"],
        "boolean"
    );
    assert_eq!(range["properties"]["sheet"]["type"], "string");
    assert!(range["properties"].get("range").is_none());

    let workbook_info = schema_for("excel.get_workbook_info");
    assert_required(&workbook_info, &["session_id"]);
    assert!(workbook_info["properties"].get("include_styles").is_none());

    let save = schema_for("excel.save");
    assert_required(&save, &["session_id"]);
    assert_eq!(save["properties"]["validate_only"]["type"], "boolean");
    assert!(save["properties"].get("path").is_none());
    assert!(save["properties"].get("format").is_none());

    let calculate = schema_for("excel.calculate");
    assert_required(&calculate, &["session_id"]);
    assert_eq!(calculate["properties"]["type"]["default"], "recalculate");
    assert_eq!(
        calculate["properties"]["type"]["enum"],
        serde_json::json!(["recalculate", "full", "full_rebuild"])
    );
    assert_eq!(calculate["properties"]["validate_only"]["type"], "boolean");

    let list_named_items = schema_for("excel.list_named_items");
    assert_required(&list_named_items, &["session_id"]);
    assert_eq!(
        list_named_items["properties"]["scope"]["enum"],
        serde_json::json!(["workbook", "sheet", "all"])
    );
    assert_eq!(list_named_items["properties"]["scope"]["default"], "all");

    let update_named_item = schema_for("excel.update_named_item");
    assert_required(&update_named_item, &["session_id", "action", "name"]);
    assert_eq!(
        update_named_item["properties"]["action"]["enum"],
        serde_json::json!(["add", "edit", "delete"])
    );
    assert_eq!(
        update_named_item["properties"]["scope"]["enum"],
        serde_json::json!(["workbook", "sheet"])
    );
    assert_eq!(
        update_named_item["properties"]["scope"]["default"],
        "workbook"
    );
    assert_eq!(update_named_item["properties"]["name"]["minLength"], 1);
    assert_eq!(
        update_named_item["properties"]["validate_only"]["type"],
        "boolean"
    );
    let update_named_item_tool = tool_for("excel.update_named_item");
    assert_eq!(
        update_named_item_tool["_meta"]["com.office-mcp/action_side_effects"]["add"],
        "mutating"
    );
    assert_eq!(
        update_named_item_tool["_meta"]["com.office-mcp/action_side_effects"]["edit"],
        "mutating"
    );
    assert_eq!(
        update_named_item_tool["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );

    let slide = schema_for("powerpoint.add_slide");
    assert_required(&slide, &["session_id"]);
    assert_eq!(slide["properties"]["layout"]["type"], "string");
    assert_eq!(
        slide["properties"]["title_box"]["additionalProperties"],
        false
    );

    let shape = schema_for("powerpoint.add_shape");
    assert_required(&shape, &["session_id", "shape_type"]);
    assert_eq!(
        shape["properties"]["shape_type"]["enum"],
        serde_json::json!([
            "text_box",
            "rectangle",
            "ellipse",
            "rounded_rectangle",
            "line"
        ])
    );
    assert_eq!(shape["properties"]["text"]["type"], "string");

    let retired_text_box = tool_catalog_json()
        .into_iter()
        .find(|tool| tool["name"] == "powerpoint.add_text_box");
    assert!(retired_text_box.is_none());
}

#[test]
fn excel_hyperlink_schema_is_specific() {
    let set_hyperlink = schema_for("excel.set_hyperlink");
    assert_required(&set_hyperlink, &["session_id", "address", "action"]);
    assert_eq!(
        set_hyperlink["properties"]["action"]["enum"],
        serde_json::json!(["set", "clear"])
    );
    assert_eq!(set_hyperlink["properties"]["sheet"]["type"], "string");
    assert_eq!(set_hyperlink["properties"]["url"]["format"], "uri");
    assert_eq!(
        set_hyperlink["properties"]["document_reference"]["minLength"],
        1
    );
    assert_eq!(
        set_hyperlink["properties"]["text_to_display"]["type"],
        "string"
    );
    assert_eq!(set_hyperlink["properties"]["screen_tip"]["type"], "string");
    assert_eq!(
        set_hyperlink["properties"]["validate_only"]["type"],
        "boolean"
    );

    let set_hyperlink_tool = tool_for("excel.set_hyperlink");
    assert_eq!(
        set_hyperlink_tool["_meta"]["com.office-mcp/action_side_effects"]["set"],
        "mutating"
    );
    assert_eq!(
        set_hyperlink_tool["_meta"]["com.office-mcp/action_side_effects"]["clear"],
        "mutating"
    );
}

#[test]
fn excel_document_property_schemas_are_specific() {
    let get_properties = schema_for("excel.get_document_properties");
    assert_required(&get_properties, &["session_id"]);
    assert_eq!(
        get_properties["properties"]["include_custom"]["type"],
        "boolean"
    );
    assert_eq!(
        get_properties["properties"]["include_custom"]["default"],
        true
    );

    let update_properties = schema_for("excel.update_document_properties");
    assert_required(&update_properties, &["session_id"]);
    assert_eq!(update_properties["properties"]["title"]["type"], "string");
    assert_eq!(update_properties["properties"]["manager"]["type"], "string");
    assert!(update_properties["properties"].get("last_author").is_none());
    assert!(
        update_properties["properties"]
            .get("revision_number")
            .is_none()
    );
    assert_eq!(
        update_properties["properties"]["custom_set"]["items"]["additionalProperties"],
        false
    );
    assert_eq!(
        update_properties["properties"]["custom_set"]["items"]["properties"]["value"]["oneOf"][0]["type"],
        "string"
    );
    assert_eq!(
        update_properties["properties"]["custom_delete"]["items"]["minLength"],
        1
    );
    assert_eq!(
        update_properties["anyOf"][9]["required"],
        serde_json::json!(["custom_delete"])
    );
    assert_eq!(
        update_properties["properties"]["validate_only"]["type"],
        "boolean"
    );

    let update_tool = tool_for("excel.update_document_properties");
    assert_eq!(
        update_tool["_meta"]["com.office-mcp/action_side_effects"]["custom_set"],
        "mutating"
    );
    assert_eq!(
        update_tool["_meta"]["com.office-mcp/action_side_effects"]["custom_delete"],
        "mutating"
    );
}

#[test]
fn excel_format_range_schema_covers_layout_completion() {
    let format_range = schema_for("excel.format_range");
    assert_required(&format_range, &["session_id", "address"]);
    assert_eq!(
        format_range["properties"]["merge"]["enum"],
        serde_json::json!(["merge", "merge_across", "unmerge"])
    );
    assert_eq!(format_range["properties"]["column_width_pt"]["minimum"], 0);
    assert_eq!(format_range["properties"]["row_height_pt"]["minimum"], 0);
    assert_eq!(
        format_range["properties"]["hidden_columns"]["type"],
        "boolean"
    );
    assert_eq!(format_range["properties"]["hidden_rows"]["type"], "boolean");
    assert_eq!(format_range["properties"]["style"]["minLength"], 1);
    assert_eq!(
        format_range["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn excel_conditional_format_schemas_are_specific() {
    let list_formats = schema_for("excel.list_conditional_formats");
    assert_required(&list_formats, &["session_id"]);
    assert_eq!(list_formats["properties"]["sheet"]["type"], "string");
    assert_eq!(list_formats["properties"]["address"]["type"], "string");

    let update_format = schema_for("excel.update_conditional_format");
    assert_required(&update_format, &["session_id", "action"]);
    assert_eq!(
        update_format["properties"]["action"]["enum"],
        serde_json::json!(["add", "delete", "clear_range"])
    );
    assert_eq!(update_format["properties"]["id"]["minLength"], 1);
    assert_eq!(update_format["properties"]["priority"]["minimum"], 0);
    assert_eq!(
        update_format["properties"]["stop_if_true"]["type"],
        "boolean"
    );
    assert_eq!(
        update_format["properties"]["rule"]["properties"]["type"]["enum"],
        serde_json::json!([
            "cell_value",
            "color_scale",
            "data_bar",
            "icon_set",
            "top_bottom",
            "preset_criteria",
            "contains_text",
            "custom_formula"
        ])
    );
    assert_eq!(
        update_format["properties"]["rule"]["properties"]["colors"]["minItems"],
        2
    );
    assert_eq!(
        update_format["properties"]["rule"]["properties"]["colors"]["maxItems"],
        3
    );
    assert_eq!(
        update_format["properties"]["rule"]["properties"]["format"]["properties"]["fill_color"]["pattern"],
        "^#[0-9A-Fa-f]{6}$"
    );
    assert_eq!(
        update_format["properties"]["validate_only"]["type"],
        "boolean"
    );
}

#[test]
fn excel_data_validation_schemas_are_specific() {
    let read_range = schema_for("excel.read_range");
    assert_eq!(
        read_range["properties"]["include_validation"]["type"],
        "boolean"
    );

    let set_validation = schema_for("excel.set_data_validation");
    assert_required(&set_validation, &["session_id", "address", "action"]);
    assert_eq!(
        set_validation["properties"]["action"]["enum"],
        serde_json::json!(["set", "clear"])
    );
    assert_eq!(
        set_validation["properties"]["rule"]["properties"]["type"]["enum"],
        serde_json::json!([
            "list",
            "whole_number",
            "decimal",
            "date",
            "time",
            "text_length",
            "custom"
        ])
    );
    assert_eq!(
        set_validation["properties"]["rule"]["properties"]["operator"]["enum"],
        serde_json::json!([
            "between",
            "not_between",
            "equal_to",
            "not_equal_to",
            "greater_than",
            "less_than",
            "greater_than_or_equal_to",
            "less_than_or_equal_to"
        ])
    );
    assert_eq!(
        set_validation["properties"]["rule"]["properties"]["list_source"]["oneOf"][0]["minItems"],
        1
    );
    assert_eq!(
        set_validation["properties"]["error_alert"]["properties"]["style"]["enum"],
        serde_json::json!(["stop", "warning", "information"])
    );
    assert_eq!(
        set_validation["properties"]["validate_only"]["type"],
        "boolean"
    );

    let validation_tool = tool_for("excel.set_data_validation");
    assert_eq!(
        validation_tool["_meta"]["com.office-mcp/action_side_effects"]["set"],
        "mutating"
    );
    assert_eq!(
        validation_tool["_meta"]["com.office-mcp/action_side_effects"]["clear"],
        "destructive"
    );
}

#[test]
fn excel_copy_range_schema_is_specific() {
    let copy_range = schema_for("excel.copy_range");
    assert_required(&copy_range, &["session_id", "action"]);
    assert_eq!(
        copy_range["properties"]["action"]["enum"],
        serde_json::json!(["copy", "autofill"])
    );
    assert_eq!(copy_range["properties"]["source_sheet"]["type"], "string");
    assert_eq!(copy_range["properties"]["source_address"]["type"], "string");
    assert_eq!(
        copy_range["properties"]["destination_sheet"]["type"],
        "string"
    );
    assert_eq!(
        copy_range["properties"]["destination_address"]["type"],
        "string"
    );
    assert_eq!(
        copy_range["properties"]["copy_type"]["enum"],
        serde_json::json!(["all", "values", "formulas", "formats", "link"])
    );
    assert_eq!(
        copy_range["properties"]["autofill_type"]["enum"],
        serde_json::json!([
            "default",
            "copy",
            "series",
            "formats",
            "values",
            "flash_fill"
        ])
    );
    assert_eq!(copy_range["properties"]["skip_blanks"]["type"], "boolean");
    assert_eq!(copy_range["properties"]["transpose"]["type"], "boolean");
    assert_eq!(copy_range["properties"]["validate_only"]["type"], "boolean");

    let copy_tool = tool_for("excel.copy_range");
    assert_eq!(
        copy_tool["_meta"]["com.office-mcp/action_side_effects"]["copy"],
        "mutating"
    );
    assert_eq!(
        copy_tool["_meta"]["com.office-mcp/action_side_effects"]["autofill"],
        "mutating"
    );
}

#[test]
fn excel_comment_schemas_are_specific() {
    let add_comment = schema_for("excel.add_comment");
    assert_required(&add_comment, &["session_id", "cell", "text"]);
    assert_eq!(add_comment["properties"]["text"]["minLength"], 1);

    let list_comments = schema_for("excel.list_comments");
    assert_required(&list_comments, &["session_id"]);
    assert_eq!(list_comments["properties"]["resolved"]["type"], "boolean");

    let update_comment = schema_for("excel.update_comment");
    assert_required(&update_comment, &["session_id", "comment_id", "action"]);
    assert_eq!(
        update_comment["properties"]["action"]["enum"],
        serde_json::json!(["reply", "edit", "resolve", "reopen", "delete"])
    );
    assert_eq!(update_comment["properties"]["text"]["minLength"], 1);
    assert_eq!(
        update_comment["properties"]["validate_only"]["type"],
        "boolean"
    );

    let update_comment_tool = tool_for("excel.update_comment");
    assert_eq!(
        update_comment_tool["_meta"]["com.office-mcp/action_side_effects"]["delete"],
        "destructive"
    );
}

#[test]
fn excel_shape_tools_expose_expected_contracts() {
    let insert_image = schema_for("excel.insert_image");
    assert_required(&insert_image, &["session_id", "image"]);
    assert_eq!(
        insert_image["properties"]["image"]["oneOf"][0]["required"],
        serde_json::json!(["base64"])
    );
    assert_eq!(
        insert_image["properties"]["image"]["oneOf"][1]["required"],
        serde_json::json!(["url"])
    );
    assert_eq!(insert_image["properties"]["left_pt"]["type"], "number");
    assert_eq!(
        insert_image["properties"]["width_pt"]["exclusiveMinimum"],
        0
    );
    assert_eq!(
        insert_image["properties"]["validate_only"]["type"],
        "boolean"
    );

    let list_shapes = schema_for("excel.list_shapes");
    assert_required(&list_shapes, &["session_id"]);
    assert_eq!(list_shapes["properties"]["sheet"]["type"], "string");

    let update_shape = schema_for("excel.update_shape");
    assert_required(&update_shape, &["session_id", "shape_id", "action"]);
    assert_eq!(
        update_shape["properties"]["action"]["enum"],
        serde_json::json!([
            "move",
            "resize",
            "set_alt_text",
            "set_text",
            "set_z_order",
            "delete"
        ])
    );
    assert_eq!(update_shape["properties"]["shape_id"]["minLength"], 1);
    assert_eq!(
        update_shape["properties"]["z_order"]["enum"],
        serde_json::json!([
            "bring_forward",
            "send_backward",
            "bring_to_front",
            "send_to_back"
        ])
    );
    assert_eq!(
        update_shape["properties"]["validate_only"]["type"],
        "boolean"
    );

    let insert_tool = tool_for("excel.insert_image");
    assert_eq!(
        insert_tool["_meta"]["com.office-mcp/side_effects"],
        "mutating"
    );
    assert_eq!(insert_tool["_meta"]["com.office-mcp/category"], "Shapes");
    assert_eq!(
        insert_tool["_meta"]["com.office-mcp/supports_validate_only"],
        true
    );

    let list_tool = tool_for("excel.list_shapes");
    assert_eq!(list_tool["_meta"]["com.office-mcp/side_effects"], "read");
    assert_eq!(list_tool["_meta"]["com.office-mcp/category"], "Shapes");

    let update_tool = tool_for("excel.update_shape");
    assert_eq!(
        update_tool["_meta"]["com.office-mcp/side_effects"],
        "destructive"
    );
    assert_eq!(update_tool["_meta"]["com.office-mcp/category"], "Shapes");
    assert_eq!(
        update_tool["_meta"]["com.office-mcp/supports_validate_only"],
        true
    );
}

#[test]
fn excel_tool_catalog_checks_supported_names() {
    assert!(ExcelToolCatalog::contains("excel.get_workbook_info"));
    assert!(ExcelToolCatalog::contains("excel.save"));
    assert!(ExcelToolCatalog::contains("excel.calculate"));
    assert!(ExcelToolCatalog::contains("excel.list_named_items"));
    assert!(ExcelToolCatalog::contains("excel.update_named_item"));
    assert!(ExcelToolCatalog::contains("excel.list_sheets"));
    assert!(ExcelToolCatalog::contains("excel.update_sheet"));
    assert!(ExcelToolCatalog::contains("excel.delete_sheet"));
    assert!(ExcelToolCatalog::contains("excel.get_used_range"));
    assert!(ExcelToolCatalog::contains("excel.insert_range"));
    assert!(ExcelToolCatalog::contains("excel.clear_range"));
    assert!(ExcelToolCatalog::contains("excel.set_hyperlink"));
    assert!(ExcelToolCatalog::contains("excel.set_data_validation"));
    assert!(ExcelToolCatalog::contains("excel.copy_range"));
    assert!(ExcelToolCatalog::contains("excel.find_replace_cells"));
    assert!(ExcelToolCatalog::contains("excel.list_conditional_formats"));
    assert!(ExcelToolCatalog::contains(
        "excel.update_conditional_format"
    ));
    assert!(ExcelToolCatalog::contains("excel.sort_range"));
    assert!(ExcelToolCatalog::contains("excel.apply_filter"));
    assert!(ExcelToolCatalog::contains("excel.update_table"));
    assert!(ExcelToolCatalog::contains("excel.update_chart"));
    assert!(ExcelToolCatalog::contains("excel.create_pivot_table"));
    assert!(ExcelToolCatalog::contains("excel.update_pivot_table"));
    assert!(ExcelToolCatalog::contains("excel.insert_image"));
    assert!(ExcelToolCatalog::contains("excel.list_shapes"));
    assert!(ExcelToolCatalog::contains("excel.update_shape"));
    assert!(ExcelToolCatalog::contains("excel.write_range"));
    assert!(!ExcelToolCatalog::contains("excel.unsupported"));
}

#[test]
fn powerpoint_tool_catalog_checks_supported_names() {
    assert!(PowerPointToolCatalog::contains(
        "powerpoint.get_presentation_info"
    ));
    assert!(PowerPointToolCatalog::contains("powerpoint.export_file"));
    assert!(!PowerPointToolCatalog::contains(
        "powerpoint.get_active_view"
    ));
    assert!(!PowerPointToolCatalog::contains("powerpoint.add_text_box"));
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

fn mutating_tool_names() -> Vec<&'static str> {
    tool_metadata_catalog()
        .iter()
        .filter(|metadata| !matches!(metadata.side_effect, ToolSideEffect::Read))
        .map(|metadata| metadata.name)
        .collect()
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
