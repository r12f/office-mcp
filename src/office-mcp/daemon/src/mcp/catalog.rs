use crate::mcp::{ToolAccessPolicy, ToolSideEffect, tool_metadata};
use serde_json::{Value, json};

pub const WORD_V1_TOOLS: &[&str] = &[
    "word.add_comment",
    "word.apply_formatting",
    "word.apply_style",
    "word.delete_bookmark",
    "word.delete_content_control",
    "word.delete_field",
    "word.delete_shape",
    "word.delete_range",
    "word.find_text",
    "word.get_document_properties",
    "word.get_header_footer",
    "word.get_html",
    "word.get_outline",
    "word.get_selection",
    "word.set_selection",
    "word.get_text",
    "word.get_image",
    "word.insert_bookmark",
    "word.insert_content_control",
    "word.insert_break",
    "word.insert_field",
    "word.insert_html",
    "word.insert_hyperlink",
    "word.insert_image",
    "word.insert_list",
    "word.insert_note",
    "word.insert_paragraph",
    "word.insert_shape",
    "word.insert_table",
    "word.list_bookmarks",
    "word.list_content_controls",
    "word.list_fields",
    "word.list_hyperlinks",
    "word.list_lists",
    "word.list_images",
    "word.list_notes",
    "word.list_sections",
    "word.list_shapes",
    "word.list_styles",
    "word.read_table",
    "word.remove_hyperlink",
    "word.replace_text",
    "word.resolve_anchor",
    "word.update_comment",
    "word.save",
    "word.set_change_tracking",
    "word.update_header_footer",
    "word.update_content_control",
    "word.update_document_properties",
    "word.update_field",
    "word.update_image",
    "word.update_list",
    "word.update_note",
    "word.update_page_setup",
    "word.update_paragraph",
    "word.update_shape",
    "word.create_style",
    "word.update_style",
    "word.update_table",
    "word.update_tracked_change",
    "word.delete_note",
];

const EXCEL_V1_TOOLS: &[OfficeToolDefinition] = &[
    "excel.add_comment",
    "excel.add_sheet",
    "excel.apply_filter",
    "excel.clear_range",
    "excel.copy_range",
    "excel.create_chart",
    "excel.create_pivot_table",
    "excel.create_table",
    "excel.delete_sheet",
    "excel.find_replace_cells",
    "excel.format_range",
    "excel.list_conditional_formats",
    "excel.get_workbook_info",
    "excel.insert_image",
    "excel.set_data_validation",
    "excel.set_hyperlink",
    "excel.update_conditional_format",
    "excel.save",
    "excel.calculate",
    "excel.list_named_items",
    "excel.get_document_properties",
    "excel.list_comments",
    "excel.list_shapes",
    "excel.update_named_item",
    "excel.update_document_properties",
    "excel.list_sheets",
    "excel.read_range",
    "excel.set_formula",
    "excel.sort_range",
    "excel.update_chart",
    "excel.update_comment",
    "excel.update_pivot_table",
    "excel.update_range_structure",
    "excel.update_shape",
    "excel.update_table",
    "excel.update_sheet",
    "excel.write_range",
];

const POWERPOINT_V1_TOOLS: &[OfficeToolDefinition] = &[
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

pub type OfficeToolDefinition = &'static str;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OfficeToolCatalog {
    app: &'static str,
    tools: &'static [OfficeToolDefinition],
}

impl OfficeToolCatalog {
    const fn new(app: &'static str, tools: &'static [OfficeToolDefinition]) -> Self {
        Self { app, tools }
    }

    #[must_use]
    pub const fn app(self) -> &'static str {
        self.app
    }

    #[must_use]
    pub const fn tools(self) -> &'static [OfficeToolDefinition] {
        self.tools
    }

    pub fn tool_names(self) -> impl Iterator<Item = &'static str> {
        self.tools.iter().copied()
    }

    #[must_use]
    pub fn contains(self, name: &str) -> bool {
        self.tools.contains(&name)
    }
}

const WORD_TOOL_CATALOG: OfficeToolCatalog = OfficeToolCatalog::new("word", WORD_V1_TOOLS);
const EXCEL_TOOL_CATALOG: OfficeToolCatalog = OfficeToolCatalog::new("excel", EXCEL_V1_TOOLS);
const POWERPOINT_TOOL_CATALOG: OfficeToolCatalog =
    OfficeToolCatalog::new("powerpoint", POWERPOINT_V1_TOOLS);
const OFFICE_TOOL_CATALOGS: &[OfficeToolCatalog] = &[
    WORD_TOOL_CATALOG,
    EXCEL_TOOL_CATALOG,
    POWERPOINT_TOOL_CATALOG,
];

#[must_use]
pub const fn office_tool_catalogs() -> &'static [OfficeToolCatalog] {
    OFFICE_TOOL_CATALOGS
}

pub fn all_office_tool_names() -> impl Iterator<Item = &'static str> {
    office_tool_catalogs()
        .iter()
        .copied()
        .flat_map(OfficeToolCatalog::tool_names)
}

#[must_use]
pub fn is_office_tool(name: &str) -> bool {
    office_tool_catalogs()
        .iter()
        .copied()
        .any(|catalog| catalog.contains(name))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExcelToolCatalog;

impl ExcelToolCatalog {
    #[must_use]
    pub const fn tools() -> &'static [OfficeToolDefinition] {
        EXCEL_TOOL_CATALOG.tools()
    }

    #[must_use]
    pub fn contains(name: &str) -> bool {
        EXCEL_TOOL_CATALOG.contains(name)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointToolCatalog;

impl PowerPointToolCatalog {
    #[must_use]
    pub const fn tools() -> &'static [OfficeToolDefinition] {
        POWERPOINT_TOOL_CATALOG.tools()
    }

    #[must_use]
    pub fn contains(name: &str) -> bool {
        POWERPOINT_TOOL_CATALOG.contains(name)
    }
}

#[must_use]
pub fn tool_catalog_json() -> Vec<Value> {
    tool_catalog_json_for_policy(&ToolAccessPolicy::default())
}

#[must_use]
pub fn tool_catalog_json_for_policy(policy: &ToolAccessPolicy) -> Vec<Value> {
    let mut tools = Vec::new();
    push_tool_with_aliases(
        &mut tools,
        "office.list_sessions",
        "List Office Sessions",
        "List connected Office document sessions.",
    );
    push_tool_with_aliases(
        &mut tools,
        "office.get_session_info",
        "Get Office Session Info",
        "Return metadata and supported tools for one Office document session.",
    );
    push_tool_with_aliases(
        &mut tools,
        "office.describe_tools",
        "Describe Office Tools",
        "Return runtime contracts, examples, and common errors for multiple Office MCP tools.",
    );
    for catalog in office_tool_catalogs().iter().copied() {
        for tool in catalog.tool_names().filter(|tool| policy.allows_tool(tool)) {
            push_tool_with_aliases(
                &mut tools,
                tool,
                tool,
                forwarded_tool_description(catalog.app()),
            );
        }
    }
    tools
}

const fn forwarded_tool_description(app: &str) -> &'static str {
    match app.as_bytes() {
        b"word" => "Forward this Word tool call to the selected Office document session.",
        b"excel" => "Forward this Excel tool call to the selected Office workbook session.",
        b"powerpoint" => {
            "Forward this PowerPoint tool call to the selected Office presentation session."
        }
        _ => "Forward this Office tool call to the selected Office document session.",
    }
}

fn push_tool_with_aliases(tools: &mut Vec<Value>, name: &str, title: &str, description: &str) {
    tools.push(tool_json(name, title, description));
    let alias = mcp_safe_tool_alias(name);
    tools.push(tool_json(&alias, title, description));
}

#[must_use]
pub fn mcp_safe_tool_alias(tool: &str) -> String {
    tool.replace('.', "_")
}

#[must_use]
pub fn canonical_tool_name(tool: &str) -> &str {
    if tool.contains('.') {
        return tool;
    }
    for candidate in management_tool_names()
        .iter()
        .copied()
        .chain(all_office_tool_names())
    {
        if mcp_safe_tool_alias(candidate) == tool {
            return candidate;
        }
    }
    tool
}

const fn management_tool_names() -> &'static [&'static str] {
    &[
        "office.list_sessions",
        "office.get_session_info",
        "office.describe_tools",
    ]
}

#[must_use]
pub fn word_resource_catalog_for_session(session_id: &str) -> Vec<Value> {
    vec![
        resource_json(
            &format!("office://word/{session_id}/document?offset=0&limit=200"),
            "word.document",
            "Word Document Text",
        ),
        resource_json(
            &format!("office://word/{session_id}/structure"),
            "word.structure",
            "Word Structure",
        ),
        resource_json(
            &format!("office://word/{session_id}/comments"),
            "word.comments",
            "Word Comments",
        ),
        resource_json(
            &format!("office://word/{session_id}/track_changes"),
            "word.track_changes",
            "Word Tracked Changes",
        ),
        resource_json(
            &format!("office://word/{session_id}/selection"),
            "word.selection",
            "Word Selection",
        ),
    ]
}

#[must_use]
pub fn word_resource_templates() -> Vec<Value> {
    vec![
        resource_template_json(
            "office://word/{session_id}/comments",
            "word.comments.template",
            "Word Comments",
        ),
        resource_template_json(
            "office://word/{session_id}/document{?offset,limit}",
            "word.document.template",
            "Word Document Text",
        ),
        resource_template_json(
            "office://word/{session_id}/paragraph/{index}",
            "word.paragraph.template",
            "Word Paragraph",
        ),
        resource_template_json(
            "office://word/{session_id}/selection",
            "word.selection.template",
            "Word Selection",
        ),
        resource_template_json(
            "office://word/{session_id}/structure",
            "word.structure.template",
            "Word Structure",
        ),
        resource_template_json(
            "office://word/{session_id}/track_changes",
            "word.track_changes.template",
            "Word Tracked Changes",
        ),
    ]
}

#[must_use]
pub fn excel_resource_catalog_for_session(session_id: &str) -> Vec<Value> {
    vec![
        resource_json(
            &format!("office://excel/{session_id}/workbook"),
            "excel.workbook",
            "Excel Workbook Info",
        ),
        resource_json(
            &format!("office://excel/{session_id}/sheets"),
            "excel.sheets",
            "Excel Worksheets",
        ),
        resource_json(
            &format!("office://excel/{session_id}/used-range"),
            "excel.used_range",
            "Excel Used Range",
        ),
        resource_json(
            &format!("office://excel/{session_id}/range/A1"),
            "excel.range",
            "Excel Range",
        ),
    ]
}

#[must_use]
pub fn excel_resource_templates() -> Vec<Value> {
    vec![
        resource_template_json(
            "office://excel/{session_id}/workbook",
            "excel.workbook.template",
            "Excel Workbook Info",
        ),
        resource_template_json(
            "office://excel/{session_id}/sheets",
            "excel.sheets.template",
            "Excel Worksheets",
        ),
        resource_template_json(
            "office://excel/{session_id}/used-range{?sheet}",
            "excel.used_range.template",
            "Excel Used Range",
        ),
        resource_template_json(
            "office://excel/{session_id}/range/{address}{?sheet}",
            "excel.range.template",
            "Excel Range",
        ),
    ]
}

#[must_use]
pub fn powerpoint_resource_catalog_for_session(session_id: &str) -> Vec<Value> {
    vec![
        resource_json(
            &format!("office://powerpoint/{session_id}/presentation"),
            "powerpoint.presentation",
            "PowerPoint Presentation Info",
        ),
        resource_json(
            &format!("office://powerpoint/{session_id}/slides"),
            "powerpoint.slides",
            "PowerPoint Slides",
        ),
        resource_json(
            &format!("office://powerpoint/{session_id}/slides/text?start=0"),
            "powerpoint.slides.text",
            "PowerPoint Slides Text",
        ),
        resource_json(
            &format!("office://powerpoint/{session_id}/slide/0/text"),
            "powerpoint.slide.text",
            "PowerPoint Slide Text",
        ),
        resource_json(
            &format!("office://powerpoint/{session_id}/slide/0/shapes"),
            "powerpoint.slide.shapes",
            "PowerPoint Slide Shapes",
        ),
    ]
}

#[must_use]
pub fn powerpoint_resource_templates() -> Vec<Value> {
    vec![
        resource_template_json(
            "office://powerpoint/{session_id}/presentation",
            "powerpoint.presentation.template",
            "PowerPoint Presentation Info",
        ),
        resource_template_json(
            "office://powerpoint/{session_id}/slides",
            "powerpoint.slides.template",
            "PowerPoint Slides",
        ),
        resource_template_json(
            "office://powerpoint/{session_id}/slides/text{?start,end}",
            "powerpoint.slides.text.template",
            "PowerPoint Slides Text",
        ),
        resource_template_json(
            "office://powerpoint/{session_id}/slide/{index}/text",
            "powerpoint.slide.text.template",
            "PowerPoint Slide Text",
        ),
        resource_template_json(
            "office://powerpoint/{session_id}/slide/{index}/shapes",
            "powerpoint.slide.shapes.template",
            "PowerPoint Slide Shapes",
        ),
    ]
}

fn resource_json(uri: &str, name: &str, title: &str) -> Value {
    json!({
        "uri": uri,
        "name": name,
        "title": title,
        "mimeType": "application/json"
    })
}

fn resource_template_json(uri_template: &str, name: &str, title: &str) -> Value {
    json!({
        "uriTemplate": uri_template,
        "name": name,
        "title": title,
        "mimeType": "application/json"
    })
}

fn tool_json(name: &str, title: &str, description: &str) -> Value {
    let canonical_name = canonical_tool_name(name);
    let side_effect = tool_side_effect(canonical_name).unwrap_or("read");
    let mut tool = json!({
        "name": name,
        "title": title,
        "description": description,
        "inputSchema": input_schema_for_tool(canonical_name),
        "annotations": tool_annotations(side_effect),
        "_meta": {
            "com.office-mcp/canonical_name": canonical_name,
            "com.office-mcp/side_effects": side_effect,
            "com.office-mcp/common_errors": common_errors_for_tool(canonical_name),
            "com.office-mcp/examples": examples_for_tool(canonical_name)
        }
    });
    if name != canonical_name {
        tool["_meta"]["com.office-mcp/alias_for"] = json!(canonical_name);
    }
    if let Some(metadata) = tool_metadata(canonical_name) {
        tool["_meta"]["com.office-mcp/app"] = json!(metadata.app);
        tool["_meta"]["com.office-mcp/category"] = json!(metadata.category);
        if supports_validate_only(canonical_name) {
            tool["_meta"]["com.office-mcp/supports_validate_only"] = json!(true);
        }
        if let Some(action_side_effects) = action_side_effects_json(canonical_name) {
            tool["_meta"]["com.office-mcp/action_side_effects"] = action_side_effects;
        }
    }
    tool
}

fn tool_annotations(side_effect: &str) -> Value {
    json!({
        "readOnlyHint": side_effect == "read",
        "destructiveHint": side_effect == "destructive",
        "idempotentHint": false,
        "openWorldHint": false
    })
}

/// Validates that a tool call argument object only uses advertised top-level
/// fields and includes every advertised required field.
///
/// # Errors
///
/// Returns an error when `arguments` is not an object, includes an unknown
/// top-level field, or omits a required top-level field.
pub fn validate_tool_arguments(tool: &str, arguments: &Value) -> Result<(), String> {
    let tool = canonical_tool_name(tool);
    let schema = input_schema_for_tool(tool);
    let Some(arguments) = arguments.as_object() else {
        return Err(format!("{tool} arguments must be a JSON object."));
    };
    let Some(properties) = schema["properties"].as_object() else {
        return Err(format!("{tool} has an invalid input schema."));
    };
    for key in arguments.keys() {
        if !properties.contains_key(key) {
            return Err(format!("{tool} does not accept argument {key}."));
        }
        validate_enum_argument(tool, key, arguments, &properties[key])?;
    }
    let Some(required_fields) = schema["required"].as_array() else {
        return Err(format!("{tool} has an invalid input schema."));
    };
    for required in required_fields {
        let Some(field) = required.as_str() else {
            return Err(format!("{tool} has an invalid input schema."));
        };
        if !arguments.contains_key(field) {
            return Err(format!("{tool} requires argument {field}."));
        }
    }
    validate_anchor_argument(tool, arguments)?;
    Ok(())
}

fn validate_enum_argument(
    tool: &str,
    key: &str,
    arguments: &serde_json::Map<String, Value>,
    schema: &Value,
) -> Result<(), String> {
    let Some(allowed_values) = schema.get("enum").and_then(Value::as_array) else {
        return Ok(());
    };
    let Some(value) = arguments.get(key) else {
        return Ok(());
    };
    if allowed_values.iter().any(|allowed| allowed == value) {
        return Ok(());
    }
    let allowed = allowed_values
        .iter()
        .filter_map(Value::as_str)
        .collect::<Vec<_>>()
        .join(", ");
    Err(format!("{tool} argument {key} must be one of {allowed}."))
}

#[must_use]
pub fn input_schema_for_tool(tool: &str) -> Value {
    let spec = tool_input_spec(canonical_tool_name(tool));
    object_schema(tool, spec.required, spec.properties)
}

#[must_use]
pub fn describe_tool_contract(tool: &str) -> Option<Value> {
    let canonical_name = canonical_tool_name(tool);
    let side_effect = tool_side_effect(canonical_name)?;
    let input_schema = input_schema_for_tool(canonical_name);
    let mut contract = json!({
        "name": tool,
        "canonical_name": canonical_name,
        "input_schema": input_schema,
        "parameters": parameters_for_schema(&input_schema),
        "examples": examples_for_tool(canonical_name),
        "side_effect": side_effect,
        "common_errors": common_errors_for_tool(canonical_name)
    });
    if tool != canonical_name {
        contract["alias_for"] = json!(canonical_name);
    }
    if let Some(metadata) = tool_metadata(canonical_name) {
        contract["app"] = json!(metadata.app);
        contract["category"] = json!(metadata.category);
        if supports_validate_only(canonical_name) {
            contract["supports_validate_only"] = json!(true);
        }
        if let Some(action_side_effects) = action_side_effects_json(canonical_name) {
            contract["action_side_effects"] = action_side_effects;
        }
    }
    Some(contract)
}

#[must_use]
pub fn unknown_tool_contract(tool: &str) -> Value {
    json!({
        "name": tool,
        "error": {
            "office_mcp_code": "UNKNOWN_TOOL",
            "message": format!("Unknown tool {tool}.")
        }
    })
}

fn parameters_for_schema(schema: &Value) -> Vec<Value> {
    let required = schema["required"]
        .as_array()
        .map(|fields| fields.iter().filter_map(Value::as_str).collect::<Vec<_>>())
        .unwrap_or_default();
    let Some(properties) = schema["properties"].as_object() else {
        return Vec::new();
    };
    properties
        .iter()
        .map(|(name, schema)| {
            json!({
                "name": name,
                "required": required.contains(&name.as_str()),
                "schema": schema
            })
        })
        .collect()
}

fn tool_side_effect(tool: &str) -> Option<&'static str> {
    match tool {
        "office.list_sessions" | "office.get_session_info" | "office.describe_tools" => {
            Some("read")
        }
        _ => tool_metadata(tool).map(|metadata| side_effect_name(metadata.side_effect)),
    }
}

fn supports_validate_only(tool: &str) -> bool {
    tool_metadata(tool)
        .is_some_and(|metadata| !matches!(metadata.side_effect, ToolSideEffect::Read))
}

#[must_use]
pub fn action_side_effect_for_tool<'a>(
    tool: &str,
    arguments: &'a Value,
) -> Option<(&'a str, &'static str)> {
    let tool = canonical_tool_name(tool);
    let metadata = tool_metadata(tool)?;
    let action = action_argument_for_policy(tool, arguments)?;
    let side_effect = metadata.side_effect_for_action(action)?;
    Some((action, side_effect_name(side_effect)))
}

fn action_argument_for_policy<'a>(tool: &str, arguments: &'a Value) -> Option<&'a str> {
    if tool == "powerpoint.update_tags" && arguments.get("action").is_none() {
        return Some("list");
    }
    arguments.get("action").and_then(Value::as_str)
}

fn action_side_effects_json(tool: &str) -> Option<Value> {
    let metadata = tool_metadata(tool)?;
    let actions = metadata.action_side_effects?;
    let mut map = serde_json::Map::new();
    for action in actions {
        map.insert(
            action.action.to_string(),
            json!(side_effect_name(action.side_effect)),
        );
    }
    Some(Value::Object(map))
}

const fn side_effect_name(side_effect: ToolSideEffect) -> &'static str {
    match side_effect {
        ToolSideEffect::Read => "read",
        ToolSideEffect::Mutating => "mutating",
        ToolSideEffect::Destructive => "destructive",
    }
}

fn examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "office.describe_tools" => vec![json!({
            "description": "Inspect contracts for Word image insertion and Excel range reads.",
            "arguments": { "tools": ["word.insert_image", "excel.read_range"] }
        })],
        "word.insert_image" => vec![json!({
            "description": "Insert a PNG as a new paragraph after paragraph 2.",
            "arguments": {
                "session_id": "session-1",
                "anchor": { "kind": "after_paragraph_index", "index": 2 },
                "placement": "new_paragraph_after",
                "image": { "base64": "iVBORw0KGgo...", "mime_type": "image/png" },
                "alt_text": "Quarterly revenue chart",
                "width_pt": 420
            }
        })],
        "word.delete_range" => vec![json!({
            "description": "Delete the first occurrence of a phrase without relying on character offsets.",
            "arguments": {
                "session_id": "session-1",
                "anchor": { "kind": "before_text", "text": "Remove this paragraph", "occurrence": 1 }
            }
        })],
        "word.insert_hyperlink" | "word.list_hyperlinks" | "word.remove_hyperlink" => {
            word_hyperlink_examples(tool)
        }
        "word.insert_content_control"
        | "word.update_content_control"
        | "word.delete_content_control" => content_control_examples_for_tool(tool),
        "word.insert_bookmark" | "word.list_bookmarks" | "word.delete_bookmark" => {
            bookmark_examples_for_tool(tool)
        }
        "word.insert_note" | "word.list_notes" | "word.update_note" | "word.delete_note" => {
            note_examples_for_tool(tool)
        }
        "word.list_fields" | "word.insert_field" | "word.update_field" | "word.delete_field" => {
            field_examples_for_tool(tool)
        }
        "word.list_styles" | "word.create_style" | "word.update_style" => {
            style_examples_for_tool(tool)
        }
        "word.get_document_properties"
        | "word.update_document_properties"
        | "excel.get_document_properties"
        | "excel.update_document_properties" => document_property_examples_for_tool(tool),
        "word.update_table" => vec![json!({
            "description": "Merge the first row across two columns after validating table bounds.",
            "arguments": {
                "session_id": "session-1",
                "table_index": 0,
                "action": "merge_cells",
                "row_range": [0, 0],
                "col_range": [0, 1]
            }
        })],
        "excel.update_table" => vec![json!({
            "description": "Append rows to an Excel table.",
            "arguments": {
                "session_id": "session-1",
                "table": "Table1",
                "action": "add_rows",
                "values": [["North", 1200], ["South", 980]]
            }
        })],
        "powerpoint.update_table" => vec![json!({
            "description": "Update a PowerPoint table cell.",
            "arguments": {
                "session_id": "session-1",
                "slide_index": 0,
                "shape_id": "table-1",
                "action": "set_cell_text",
                "row_index": 0,
                "column_index": 1,
                "text": "Q4"
            }
        })],
        _ => Vec::new(),
    }
}

fn content_control_examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "word.insert_content_control" => vec![json!({
            "description": "Wrap inserted review text in a tagged content control.",
            "arguments": {
                "session_id": "session-1",
                "anchor": { "kind": "end_of_document" },
                "text": "Approved by legal.",
                "tag": "approval-note",
                "title": "Approval note"
            }
        })],
        "word.update_content_control" => vec![json!({
            "description": "Replace text and metadata for an existing content control.",
            "arguments": {
                "session_id": "session-1",
                "content_control_id": 42,
                "text": "Final approval received.",
                "tag": "approval-note-final"
            }
        })],
        "word.delete_content_control" => vec![json!({
            "description": "Remove a content control while keeping its contents in the document.",
            "arguments": {
                "session_id": "session-1",
                "content_control_id": 42,
                "delete_contents": false
            }
        })],
        _ => Vec::new(),
    }
}

fn document_property_examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "word.get_document_properties" | "excel.get_document_properties" => vec![json!({
            "description": "Read core and custom document metadata.",
            "arguments": {
                "session_id": "session-1",
                "include_custom": true
            }
        })],
        "word.update_document_properties" | "excel.update_document_properties" => vec![json!({
            "description": "Set a title and upsert a custom property.",
            "arguments": {
                "session_id": "session-1",
                "title": "Quarterly review",
                "custom_set": [{ "key": "Workflow", "value": "review" }]
            }
        })],
        _ => Vec::new(),
    }
}

fn style_examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "word.list_styles" => vec![json!({
            "description": "List paragraph styles available in the document.",
            "arguments": {
                "session_id": "session-1",
                "type": "paragraph"
            }
        })],
        "word.create_style" => vec![json!({
            "description": "Create a reusable paragraph style.",
            "arguments": {
                "session_id": "session-1",
                "name": "Review Heading",
                "type": "paragraph",
                "font": { "bold": true, "color": "#1F4E79" },
                "paragraph": { "alignment": "center" }
            }
        })],
        "word.update_style" => vec![json!({
            "description": "Update an existing style definition after listing styles.",
            "arguments": {
                "session_id": "session-1",
                "name": "Review Heading",
                "font": { "italic": true }
            }
        })],
        _ => Vec::new(),
    }
}

fn word_hyperlink_examples(tool: &str) -> Vec<Value> {
    match tool {
        "word.insert_hyperlink" => vec![json!({
            "description": "Insert linked text after a paragraph marker.",
            "arguments": {
                "session_id": "session-1",
                "anchor": { "kind": "after_text", "text": "See also:" },
                "text": "Project plan",
                "url": "https://example.com/project-plan"
            }
        })],
        "word.list_hyperlinks" => vec![json!({
            "description": "List the first page of document hyperlinks.",
            "arguments": {
                "session_id": "session-1",
                "offset": 0,
                "limit": 50
            }
        })],
        "word.remove_hyperlink" => vec![json!({
            "description": "Remove a hyperlink while keeping the linked text.",
            "arguments": {
                "session_id": "session-1",
                "anchor": { "kind": "after_text", "text": "Project plan" },
                "keep_text": true
            }
        })],
        _ => Vec::new(),
    }
}

fn bookmark_examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "word.insert_bookmark" => vec![json!({
            "description": "Create a bookmark around a located section heading.",
            "arguments": {
                "session_id": "session-1",
                "name": "ResultsSection",
                "anchor": { "kind": "heading", "text": "Results", "level": 2 }
            }
        })],
        "word.list_bookmarks" => vec![json!({
            "description": "List visible bookmarks and bounded previews.",
            "arguments": { "session_id": "session-1" }
        })],
        "word.delete_bookmark" => vec![json!({
            "description": "Delete a bookmark marker while keeping its text.",
            "arguments": {
                "session_id": "session-1",
                "name": "ResultsSection"
            }
        })],
        _ => Vec::new(),
    }
}

fn note_examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "word.insert_note" => vec![json!({
            "description": "Insert a footnote after a cited sentence.",
            "arguments": {
                "session_id": "session-1",
                "kind": "footnote",
                "anchor": { "kind": "after_text", "text": "Market data source" },
                "text": "Source: internal finance workbook."
            }
        })],
        "word.list_notes" => vec![json!({
            "description": "List the first page of footnotes.",
            "arguments": {
                "session_id": "session-1",
                "kind": "footnote",
                "offset": 0,
                "limit": 50
            }
        })],
        "word.update_note" => vec![json!({
            "description": "Replace the first endnote body after reviewing current note indices.",
            "arguments": {
                "session_id": "session-1",
                "kind": "endnote",
                "index": 0,
                "text": "Updated citation text."
            }
        })],
        "word.delete_note" => vec![json!({
            "description": "Delete the first footnote after re-reading current note indices.",
            "arguments": {
                "session_id": "session-1",
                "kind": "footnote",
                "index": 0
            }
        })],
        _ => Vec::new(),
    }
}

fn field_examples_for_tool(tool: &str) -> Vec<Value> {
    match tool {
        "word.list_fields" => vec![json!({
            "description": "List document fields with bounded previews.",
            "arguments": {
                "session_id": "session-1",
                "offset": 0,
                "limit": 50
            }
        })],
        "word.insert_field" => vec![json!({
            "description": "Insert a hyperlinkable table of contents at the document start.",
            "arguments": {
                "session_id": "session-1",
                "anchor": { "kind": "start_of_document" },
                "field_type": "toc"
            }
        })],
        "word.update_field" => vec![json!({
            "description": "Refresh all fields after confirming the current count.",
            "arguments": {
                "session_id": "session-1",
                "action": "refresh_all",
                "expected_count": 3
            }
        })],
        "word.delete_field" => vec![json!({
            "description": "Delete one field by current index after re-listing fields.",
            "arguments": {
                "session_id": "session-1",
                "field_index": 0
            }
        })],
        _ => Vec::new(),
    }
}

fn common_errors_for_tool(tool: &str) -> Vec<Value> {
    let mut errors = vec![json!({
        "code": "INVALID_ARGUMENTS",
        "cause": "The arguments do not match the advertised input schema."
    })];
    if tool_metadata(tool).is_some() {
        errors.push(json!({
            "code": "TOOL_NOT_ENABLED_FOR_DOCUMENT",
            "cause": "The selected Office session did not advertise this tool."
        }));
        errors.push(json!({
            "code": "TOOL_NOT_AVAILABLE",
            "cause": "The daemon access policy currently hides or disables this tool."
        }));
    }
    match tool {
        "word.insert_image" => errors.push(json!({
            "code": "INVALID_ARGUMENTS",
            "cause": "Image input must be base64 data or an HTTPS URL, and paragraph anchors may require an explicit placement."
        })),
        "word.delete_range" => errors.push(json!({
            "code": "INVALID_ARGUMENTS",
            "cause": "The anchor must resolve to a deletable range, not the whole document body."
        })),
        "word.insert_hyperlink" | "word.remove_hyperlink" => errors.push(json!({
            "code": "INVALID_ARGUMENTS",
            "cause": "The anchor must resolve to a range, and hyperlink URLs must use http, https, mailto, or an in-document bookmark target."
        })),
        "word.insert_bookmark" | "word.delete_bookmark" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Bookmark names must follow Word naming rules and duplicate names require overwrite=true."
            }));
        }
        "word.insert_content_control" | "word.update_content_control" | "word.delete_content_control" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Content control IDs are runtime identifiers; refresh the list before updating stale IDs."
            }));
        }
        "word.insert_note" | "word.update_note" | "word.delete_note" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Note kind must be footnote or endnote, and note indices must be refreshed after insertion or deletion."
            }));
        }
        "word.insert_field" | "word.update_field" | "word.delete_field" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Field type must be allowlisted, and field indices or expected counts must be refreshed after document edits."
            }));
        }
        "word.create_style" | "word.update_style" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Style names and base styles must resolve before mutation, and update_style requires at least one style property."
            }));
        }
        "word.update_table" | "excel.update_table" | "powerpoint.update_table" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Table action arguments must match the action being requested."
            }));
        }
        _ => {}
    }
    errors
}

#[derive(Clone, Copy)]
struct ToolInputSpec {
    required: &'static [&'static str],
    properties: &'static [&'static str],
}

impl ToolInputSpec {
    const EMPTY: Self = Self {
        required: &[],
        properties: &[],
    };
}

macro_rules! tool_spec {
    ($tool:literal, [$($required:literal),* $(,)?], [$($property:literal),* $(,)?]) => {
        (
            $tool,
            ToolInputSpec {
                required: &[$($required),*],
                properties: &[$($property),*],
            },
        )
    };
}

const TOOL_INPUT_SPECS: &[(&str, ToolInputSpec)] = &[
    tool_spec!("office.list_sessions", [], []),
    tool_spec!("office.get_session_info", ["session_id"], ["session_id"]),
    tool_spec!("office.describe_tools", ["tools"], ["tools"]),
    tool_spec!(
        "word.get_text",
        ["session_id"],
        [
            "session_id",
            "offset",
            "limit",
            "include_metadata",
            "include_formatting"
        ]
    ),
    tool_spec!(
        "word.get_outline",
        ["session_id"],
        ["session_id", "max_depth"]
    ),
    tool_spec!(
        "word.find_text",
        ["session_id", "query"],
        [
            "session_id",
            "query",
            "match_case",
            "whole_word",
            "occurrence",
            "limit"
        ]
    ),
    tool_spec!(
        "word.resolve_anchor",
        ["session_id", "anchor"],
        ["session_id", "anchor", "include_text_preview"]
    ),
    tool_spec!(
        "word.list_bookmarks",
        ["session_id"],
        ["session_id", "include_hidden"]
    ),
    tool_spec!("word.get_selection", ["session_id"], ["session_id"]),
    tool_spec!(
        "word.set_selection",
        ["session_id", "anchor"],
        ["session_id", "anchor", "extent", "mode"]
    ),
    tool_spec!(
        "word.get_html",
        ["session_id"],
        ["session_id", "anchor", "extent"]
    ),
    tool_spec!(
        "word.insert_html",
        ["session_id", "anchor", "html"],
        [
            "session_id",
            "anchor",
            "html",
            "insert_location",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.get_header_footer",
        ["session_id", "location"],
        [
            "session_id",
            "location",
            "header_footer_type",
            "section_index",
            "include_metadata"
        ]
    ),
    tool_spec!("word.save", ["session_id"], ["session_id"]),
    tool_spec!(
        "word.list_sections",
        ["session_id"],
        ["session_id", "include_page_setup"]
    ),
    tool_spec!(
        "word.list_fields",
        ["session_id"],
        ["session_id", "type", "offset", "limit"]
    ),
    tool_spec!(
        "word.list_styles",
        ["session_id"],
        ["session_id", "type", "built_in", "in_use_only"]
    ),
    tool_spec!(
        "word.get_document_properties",
        ["session_id"],
        ["session_id", "include_custom"]
    ),
    tool_spec!(
        "word.update_document_properties",
        ["session_id"],
        [
            "session_id",
            "title",
            "subject",
            "author",
            "keywords",
            "category",
            "comments",
            "company",
            "manager",
            "custom_set",
            "custom_delete"
        ]
    ),
    tool_spec!(
        "word.insert_paragraph",
        ["session_id", "anchor", "text"],
        [
            "session_id",
            "anchor",
            "text",
            "style",
            "heading_level",
            "match_case"
        ]
    ),
    tool_spec!(
        "word.insert_table",
        ["session_id", "anchor", "rows", "cols"],
        [
            "session_id",
            "anchor",
            "rows",
            "cols",
            "data",
            "style",
            "match_case"
        ]
    ),
    tool_spec!(
        "word.insert_image",
        ["session_id", "anchor", "image"],
        [
            "session_id",
            "anchor",
            "image",
            "placement",
            "alt_text",
            "width_pt",
            "height_pt",
            "caption",
            "match_case",
            "validate_only"
        ]
    ),
    tool_spec!("word.list_images", ["session_id"], ["session_id"]),
    tool_spec!(
        "word.list_shapes",
        ["session_id"],
        ["session_id", "scope", "anchor", "limit"]
    ),
    tool_spec!(
        "word.get_image",
        ["session_id", "image"],
        ["session_id", "image"]
    ),
    tool_spec!(
        "word.insert_shape",
        ["session_id", "shape_type"],
        [
            "session_id",
            "shape_type",
            "anchor",
            "text",
            "image",
            "name",
            "width_pt",
            "height_pt",
            "left_pt",
            "top_pt",
            "alt_text_description",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.update_image",
        ["session_id", "image", "action"],
        [
            "session_id",
            "image",
            "action",
            "width_pt",
            "height_pt",
            "preserve_aspect_ratio",
            "alt_text_title",
            "alt_text_description",
            "hyperlink",
            "base64",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.update_shape",
        ["session_id", "shape_id", "action"],
        [
            "session_id",
            "shape_id",
            "action",
            "text",
            "name",
            "width_pt",
            "height_pt",
            "left_pt",
            "top_pt",
            "alt_text_description",
            "fill_color",
            "line_color",
            "wrap_type",
            "visible",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.delete_shape",
        ["session_id", "shape_id"],
        ["session_id", "shape_id", "validate_only"]
    ),
    tool_spec!(
        "word.insert_break",
        ["session_id", "anchor"],
        ["session_id", "anchor", "break_type", "match_case"]
    ),
    tool_spec!(
        "word.update_page_setup",
        ["session_id"],
        [
            "session_id",
            "section_index",
            "orientation",
            "paper_size",
            "margins_pt",
            "page_width_pt",
            "page_height_pt"
        ]
    ),
    tool_spec!(
        "word.insert_list",
        ["session_id", "anchor", "items"],
        [
            "session_id",
            "anchor",
            "items",
            "ordered",
            "style",
            "match_case"
        ]
    ),
    tool_spec!(
        "word.list_lists",
        ["session_id"],
        ["session_id", "offset", "limit"]
    ),
    tool_spec!(
        "word.update_list",
        ["session_id", "action"],
        [
            "session_id",
            "action",
            "list_id",
            "paragraph_index",
            "text",
            "position",
            "level",
            "numbering",
            "bullet_char",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.insert_hyperlink",
        ["session_id", "anchor", "url"],
        [
            "session_id",
            "anchor",
            "url",
            "text",
            "extent",
            "match_case",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.list_hyperlinks",
        ["session_id"],
        ["session_id", "offset", "limit"]
    ),
    tool_spec!(
        "word.remove_hyperlink",
        ["session_id", "anchor"],
        ["session_id", "anchor", "keep_text", "match_case"]
    ),
    tool_spec!(
        "word.update_header_footer",
        ["session_id", "location", "action"],
        [
            "session_id",
            "location",
            "header_footer_type",
            "section_index",
            "action",
            "text",
            "style",
            "formatting",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.insert_field",
        ["session_id", "anchor", "field_type"],
        [
            "session_id",
            "anchor",
            "field_type",
            "code_options",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.insert_bookmark",
        ["session_id", "name", "anchor"],
        ["session_id", "name", "anchor", "extent", "overwrite"]
    ),
    tool_spec!(
        "word.replace_text",
        ["session_id", "find", "replace"],
        [
            "session_id",
            "find",
            "replace",
            "match_case",
            "whole_word",
            "wildcards",
            "scope",
            "limit",
            "dry_run",
            "partial_ok",
            "all",
            "occurrence",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.update_paragraph",
        ["session_id", "index", "text"],
        [
            "session_id",
            "index",
            "text",
            "style",
            "heading_level",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.delete_range",
        ["session_id", "anchor"],
        ["session_id", "anchor", "match_case", "validate_only"]
    ),
    tool_spec!(
        "word.delete_bookmark",
        ["session_id", "name"],
        ["session_id", "name"]
    ),
    tool_spec!(
        "word.insert_note",
        ["session_id", "anchor", "kind", "text"],
        ["session_id", "anchor", "kind", "text", "validate_only"]
    ),
    tool_spec!(
        "word.list_notes",
        ["session_id", "kind"],
        ["session_id", "kind", "offset", "limit"]
    ),
    tool_spec!(
        "word.update_note",
        ["session_id", "kind", "index", "text"],
        ["session_id", "kind", "index", "text", "validate_only"]
    ),
    tool_spec!(
        "word.delete_note",
        ["session_id", "kind", "index"],
        ["session_id", "kind", "index", "validate_only"]
    ),
    tool_spec!(
        "word.update_field",
        ["session_id", "action"],
        [
            "session_id",
            "action",
            "field_index",
            "expected_count",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.delete_field",
        ["session_id", "field_index"],
        ["session_id", "field_index", "validate_only"]
    ),
    tool_spec!(
        "word.create_style",
        ["session_id", "name", "type"],
        [
            "session_id",
            "name",
            "type",
            "base_style",
            "font",
            "paragraph",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.update_style",
        ["session_id", "name"],
        [
            "session_id",
            "name",
            "base_style",
            "font",
            "paragraph",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.apply_formatting",
        ["session_id", "anchor"],
        [
            "session_id",
            "anchor",
            "formatting",
            "paragraph",
            "match_case"
        ]
    ),
    tool_spec!(
        "word.apply_style",
        ["session_id", "anchor"],
        [
            "session_id",
            "anchor",
            "style",
            "heading_level",
            "match_case"
        ]
    ),
    tool_spec!(
        "word.read_table",
        ["session_id", "table_index"],
        ["session_id", "table_index", "include_formatting"]
    ),
    tool_spec!(
        "word.update_table",
        ["session_id", "table_index", "action"],
        [
            "session_id",
            "table_index",
            "action",
            "row",
            "col",
            "text",
            "index",
            "values",
            "row_range",
            "col_range",
            "width_pt",
            "header_row",
            "borders",
            "data",
            "rows",
            "cols"
        ]
    ),
    tool_spec!(
        "word.list_content_controls",
        ["session_id"],
        ["session_id", "type", "tag", "title"]
    ),
    tool_spec!(
        "word.insert_content_control",
        ["session_id"],
        [
            "session_id",
            "anchor",
            "text",
            "tag",
            "title",
            "type",
            "checked",
            "list_items",
            "match_case"
        ]
    ),
    tool_spec!(
        "word.update_content_control",
        ["session_id", "content_control_id"],
        [
            "session_id",
            "content_control_id",
            "text",
            "checked",
            "selected_value",
            "list_items_add",
            "list_items_delete",
            "list_items_clear",
            "tag",
            "title"
        ]
    ),
    tool_spec!(
        "word.delete_content_control",
        ["session_id", "content_control_id"],
        ["session_id", "content_control_id", "delete_contents"]
    ),
    tool_spec!(
        "word.add_comment",
        ["session_id", "anchor", "text"],
        ["session_id", "anchor", "text", "match_case"]
    ),
    tool_spec!(
        "word.update_comment",
        ["session_id", "comment_id", "action"],
        [
            "session_id",
            "comment_id",
            "action",
            "text",
            "reply_id",
            "validate_only"
        ]
    ),
    tool_spec!(
        "word.set_change_tracking",
        ["session_id", "mode"],
        ["session_id", "mode"]
    ),
    tool_spec!(
        "word.update_tracked_change",
        ["session_id", "action"],
        [
            "session_id",
            "action",
            "change_index",
            "expected_fingerprint",
            "expected_count"
        ]
    ),
    tool_spec!("excel.get_workbook_info", ["session_id"], ["session_id"]),
    tool_spec!("excel.save", ["session_id"], ["session_id"]),
    tool_spec!("excel.calculate", ["session_id"], ["session_id", "type"]),
    tool_spec!(
        "excel.list_named_items",
        ["session_id"],
        ["session_id", "scope", "sheet"]
    ),
    tool_spec!(
        "excel.update_named_item",
        ["session_id", "action", "name"],
        [
            "session_id",
            "action",
            "name",
            "scope",
            "sheet",
            "reference",
            "formula",
            "comment"
        ]
    ),
    tool_spec!(
        "excel.get_document_properties",
        ["session_id"],
        ["session_id", "include_custom"]
    ),
    tool_spec!(
        "excel.update_document_properties",
        ["session_id"],
        [
            "session_id",
            "title",
            "subject",
            "author",
            "keywords",
            "category",
            "comments",
            "company",
            "manager",
            "custom_set",
            "custom_delete"
        ]
    ),
    tool_spec!(
        "excel.add_comment",
        ["session_id", "cell", "text"],
        ["session_id", "sheet", "cell", "text", "validate_only"]
    ),
    tool_spec!(
        "excel.list_comments",
        ["session_id"],
        ["session_id", "sheet", "resolved"]
    ),
    tool_spec!(
        "excel.update_comment",
        ["session_id", "comment_id", "action"],
        [
            "session_id",
            "comment_id",
            "action",
            "text",
            "reply_id",
            "validate_only"
        ]
    ),
    tool_spec!("excel.list_sheets", ["session_id"], ["session_id"]),
    tool_spec!(
        "excel.add_sheet",
        ["session_id", "name"],
        ["session_id", "name", "activate"]
    ),
    tool_spec!(
        "excel.update_sheet",
        ["session_id", "sheet"],
        [
            "session_id",
            "sheet",
            "name",
            "position",
            "visibility",
            "tab_color",
            "activate",
            "freeze",
            "show_gridlines",
            "show_headings"
        ]
    ),
    tool_spec!(
        "excel.delete_sheet",
        ["session_id", "sheet"],
        ["session_id", "sheet"]
    ),
    tool_spec!(
        "excel.read_range",
        ["session_id"],
        [
            "session_id",
            "sheet",
            "address",
            "metadata_only",
            "values_only",
            "include_formulas",
            "include_formatting",
            "include_hyperlinks",
            "include_validation"
        ]
    ),
    tool_spec!(
        "excel.write_range",
        ["session_id", "address", "values"],
        ["session_id", "sheet", "address", "values"]
    ),
    tool_spec!(
        "excel.update_range_structure",
        ["session_id", "address", "action", "shift"],
        ["session_id", "sheet", "address", "action", "shift", "count"]
    ),
    tool_spec!(
        "excel.clear_range",
        ["session_id", "address"],
        ["session_id", "sheet", "address", "apply_to"]
    ),
    tool_spec!(
        "excel.set_hyperlink",
        ["session_id", "address", "action"],
        [
            "session_id",
            "sheet",
            "address",
            "action",
            "url",
            "document_reference",
            "text_to_display",
            "screen_tip"
        ]
    ),
    tool_spec!(
        "excel.set_data_validation",
        ["session_id", "address", "action"],
        [
            "session_id",
            "sheet",
            "address",
            "action",
            "rule",
            "ignore_blanks",
            "error_alert",
            "input_prompt",
            "validate_only"
        ]
    ),
    tool_spec!(
        "excel.copy_range",
        ["session_id", "action"],
        [
            "session_id",
            "sheet",
            "source_sheet",
            "source_address",
            "destination_sheet",
            "destination_address",
            "action",
            "copy_type",
            "skip_blanks",
            "transpose",
            "autofill_type",
            "validate_only"
        ]
    ),
    tool_spec!(
        "excel.find_replace_cells",
        ["session_id", "find"],
        [
            "session_id",
            "sheet",
            "address",
            "find",
            "replace",
            "complete_match",
            "match_case",
            "search_direction"
        ]
    ),
    tool_spec!(
        "excel.set_formula",
        ["session_id", "address"],
        ["session_id", "sheet", "address", "formula", "formulas"]
    ),
    tool_spec!(
        "excel.format_range",
        ["session_id", "address"],
        [
            "session_id",
            "sheet",
            "address",
            "bold",
            "italic",
            "underline",
            "font_color",
            "fill_color",
            "number_format",
            "number_formats",
            "horizontal_alignment",
            "vertical_alignment",
            "wrap_text",
            "merge",
            "column_width_pt",
            "row_height_pt",
            "hidden_columns",
            "hidden_rows",
            "style",
            "autofit",
            "borders"
        ]
    ),
    tool_spec!(
        "excel.list_conditional_formats",
        ["session_id"],
        ["session_id", "sheet", "address"]
    ),
    tool_spec!(
        "excel.update_conditional_format",
        ["session_id", "action"],
        [
            "session_id",
            "sheet",
            "address",
            "action",
            "id",
            "rule",
            "priority",
            "stop_if_true",
            "validate_only"
        ]
    ),
    tool_spec!(
        "excel.sort_range",
        ["session_id", "fields"],
        [
            "session_id",
            "sheet",
            "address",
            "table",
            "target_type",
            "fields",
            "match_case",
            "has_headers",
            "orientation",
            "method"
        ]
    ),
    tool_spec!(
        "excel.apply_filter",
        ["session_id", "action"],
        [
            "session_id",
            "sheet",
            "address",
            "table",
            "column",
            "action",
            "criteria"
        ]
    ),
    tool_spec!(
        "excel.create_table",
        ["session_id", "address"],
        [
            "session_id",
            "sheet",
            "address",
            "has_headers",
            "name",
            "style"
        ]
    ),
    tool_spec!(
        "excel.update_table",
        ["session_id", "table", "action"],
        [
            "session_id",
            "table",
            "action",
            "name",
            "style",
            "show_headers",
            "show_totals",
            "highlight_first_column",
            "highlight_last_column",
            "show_banded_columns",
            "show_banded_rows",
            "show_filter_button",
            "rows",
            "columns",
            "values",
            "address"
        ]
    ),
    tool_spec!(
        "excel.create_chart",
        ["session_id", "address"],
        [
            "session_id",
            "sheet",
            "address",
            "type",
            "title",
            "series_by"
        ]
    ),
    tool_spec!(
        "excel.update_chart",
        ["session_id", "chart", "action"],
        [
            "session_id",
            "sheet",
            "chart",
            "action",
            "title",
            "visible",
            "position",
            "overlay",
            "axis",
            "axis_group",
            "title_visible",
            "address",
            "series_by",
            "start_cell",
            "end_cell",
            "width",
            "height",
            "fitting_mode"
        ]
    ),
    tool_spec!(
        "excel.create_pivot_table",
        ["session_id", "destination"],
        [
            "session_id",
            "sheet",
            "address",
            "table",
            "name",
            "destination"
        ]
    ),
    tool_spec!(
        "excel.update_pivot_table",
        ["session_id", "pivot_table", "action"],
        [
            "session_id",
            "pivot_table",
            "action",
            "axis",
            "hierarchy",
            "field",
            "summarize_by",
            "number_format",
            "layout_type",
            "show_column_grand_totals",
            "show_row_grand_totals",
            "selected_items",
            "values"
        ]
    ),
    tool_spec!(
        "excel.insert_image",
        ["session_id", "image"],
        [
            "session_id",
            "sheet",
            "image",
            "left_pt",
            "top_pt",
            "width_pt",
            "height_pt",
            "alt_text",
            "validate_only"
        ]
    ),
    tool_spec!("excel.list_shapes", ["session_id"], ["session_id", "sheet"]),
    tool_spec!(
        "excel.update_shape",
        ["session_id", "shape_id", "action"],
        [
            "session_id",
            "sheet",
            "shape_id",
            "action",
            "left_pt",
            "top_pt",
            "width_pt",
            "height_pt",
            "alt_text",
            "text",
            "z_order",
            "validate_only"
        ]
    ),
    tool_spec!(
        "powerpoint.get_presentation_info",
        ["session_id"],
        ["session_id", "include_selection"]
    ),
    tool_spec!(
        "powerpoint.export_file",
        ["session_id"],
        ["session_id", "format", "slice_size"]
    ),
    tool_spec!(
        "powerpoint.update_tags",
        ["session_id"],
        ["session_id", "action", "key", "value"]
    ),
    tool_spec!(
        "powerpoint.list_slides",
        ["session_id"],
        ["session_id", "include_tags"]
    ),
    tool_spec!(
        "powerpoint.add_slide",
        ["session_id"],
        [
            "session_id",
            "layout",
            "title",
            "content",
            "title_box",
            "content_box"
        ]
    ),
    tool_spec!(
        "powerpoint.update_slide",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "action",
            "tags",
            "hidden",
            "background_color"
        ]
    ),
    tool_spec!(
        "powerpoint.delete_slide",
        ["session_id"],
        ["session_id", "slide_id", "slide_index", "format"]
    ),
    tool_spec!(
        "powerpoint.move_slide",
        ["session_id", "target_index"],
        ["session_id", "slide_id", "slide_index", "target_index"]
    ),
    tool_spec!(
        "powerpoint.export_slide",
        ["session_id"],
        ["session_id", "slide_id", "slide_index", "format"]
    ),
    tool_spec!("powerpoint.list_layouts", ["session_id"], ["session_id"]),
    tool_spec!(
        "powerpoint.apply_layout",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "layout_id",
            "layout_name",
            "layout_type"
        ]
    ),
    tool_spec!("powerpoint.get_selection", ["session_id"], ["session_id"]),
    tool_spec!(
        "powerpoint.set_selection",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_id",
            "text_range"
        ]
    ),
    tool_spec!(
        "powerpoint.list_shapes",
        ["session_id"],
        ["session_id", "slide_id", "slide_index"]
    ),
    tool_spec!(
        "powerpoint.add_shape",
        ["session_id", "shape_type"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_type",
            "text",
            "left",
            "top",
            "width",
            "height",
            "name"
        ]
    ),
    tool_spec!(
        "powerpoint.insert_image",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "image",
            "base64",
            "left",
            "top",
            "width",
            "height",
            "alt_text"
        ]
    ),
    tool_spec!(
        "powerpoint.update_shape",
        ["session_id", "action"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_id",
            "shape_ids",
            "action",
            "name",
            "left",
            "top",
            "width",
            "height",
            "rotation",
            "alt_text_title",
            "alt_text_description",
            "is_decorative",
            "visible",
            "fill_color",
            "fill_transparency",
            "clear_fill",
            "line_color",
            "line_weight",
            "line_transparency",
            "line_visible",
            "z_order"
        ]
    ),
    tool_spec!(
        "powerpoint.read_text",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_id",
            "offset",
            "limit"
        ]
    ),
    tool_spec!(
        "powerpoint.replace_text",
        ["session_id", "find", "replace"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_id",
            "find",
            "replace",
            "match_case",
            "all"
        ]
    ),
    tool_spec!(
        "powerpoint.format_text",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_id",
            "bold",
            "italic",
            "underline",
            "font_color",
            "font_size",
            "font_name"
        ]
    ),
    tool_spec!(
        "powerpoint.add_table",
        ["session_id"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "rows",
            "columns",
            "values",
            "left",
            "top",
            "width",
            "height"
        ]
    ),
    tool_spec!(
        "powerpoint.read_table",
        ["session_id"],
        ["session_id", "slide_id", "slide_index", "shape_id"]
    ),
    tool_spec!(
        "powerpoint.update_table",
        ["session_id", "action"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "shape_id",
            "action",
            "values",
            "row_index",
            "column_index",
            "row_indices",
            "column_indices",
            "row_count",
            "column_count",
            "count",
            "value",
            "all",
            "text",
            "format",
            "style"
        ]
    ),
];

fn tool_input_spec(tool: &str) -> ToolInputSpec {
    TOOL_INPUT_SPECS
        .iter()
        .find(|(name, _)| *name == tool)
        .map_or(ToolInputSpec::EMPTY, |(_, spec)| *spec)
}

fn object_schema(tool: &str, required: &[&str], properties: &[&str]) -> Value {
    let mut map = serde_json::Map::new();
    for property in properties {
        map.insert((*property).to_string(), property_schema(tool, property));
    }
    if supports_validate_only(tool) {
        map.entry("validate_only".to_string())
            .or_insert_with(|| property_schema(tool, "validate_only"));
    }
    let mut schema = json!({
        "type": "object",
        "required": required,
        "properties": map,
        "additionalProperties": false
    });
    if tool == "word.apply_formatting" {
        schema["anyOf"] = json!([
            { "required": ["formatting"] },
            { "required": ["paragraph"] }
        ]);
    }
    if tool == "word.update_style" {
        schema["anyOf"] = json!([
            { "required": ["base_style"] },
            { "required": ["font"] },
            { "required": ["paragraph"] }
        ]);
    }
    if tool == "word.update_document_properties" || tool == "excel.update_document_properties" {
        schema["anyOf"] = json!([
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
        ]);
    }
    if tool == "word.update_tracked_change" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "enum": ["accept", "reject"] } } },
                "then": { "required": ["change_index", "expected_fingerprint"] }
            },
            {
                "if": { "properties": { "action": { "enum": ["accept_all", "reject_all"] } } },
                "then": { "required": ["expected_count"] }
            }
        ]);
    }
    if tool == "word.update_field" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "enum": ["refresh", "lock", "unlock"] } } },
                "then": { "required": ["field_index"] }
            },
            {
                "if": { "properties": { "action": { "const": "refresh_all" } } },
                "then": { "required": ["expected_count"] }
            }
        ]);
    }
    if tool == "excel.update_named_item" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "const": "add" } } },
                "then": {
                    "anyOf": [
                        { "required": ["reference"] },
                        { "required": ["formula"] }
                    ]
                }
            },
            {
                "if": { "properties": { "action": { "const": "edit" } } },
                "then": {
                    "anyOf": [
                        { "required": ["formula"] },
                        { "required": ["comment"] }
                    ]
                }
            }
        ]);
    }
    if tool == "excel.update_comment" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "enum": ["reply", "edit"] } } },
                "then": { "required": ["text"] }
            }
        ]);
    }
    add_excel_data_validation_schema_rules(tool, &mut schema);
    add_excel_set_hyperlink_schema_rules(tool, &mut schema);
    add_excel_conditional_format_schema_rules(tool, &mut schema);
    schema
}

fn add_excel_data_validation_schema_rules(tool: &str, schema: &mut Value) {
    if tool == "excel.set_data_validation" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "const": "set" } } },
                "then": { "required": ["rule"] }
            }
        ]);
    }
}

fn add_excel_conditional_format_schema_rules(tool: &str, schema: &mut Value) {
    if tool == "excel.update_conditional_format" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "const": "add" } } },
                "then": { "required": ["address", "rule"] }
            },
            {
                "if": { "properties": { "action": { "const": "delete" } } },
                "then": { "required": ["id"] }
            },
            {
                "if": { "properties": { "action": { "const": "clear_range" } } },
                "then": { "required": ["address"] }
            }
        ]);
    }
}

fn add_excel_set_hyperlink_schema_rules(tool: &str, schema: &mut Value) {
    if tool == "excel.set_hyperlink" {
        schema["allOf"] = json!([
            {
                "if": { "properties": { "action": { "const": "set" } } },
                "then": {
                    "oneOf": [
                        { "required": ["url"] },
                        { "required": ["document_reference"] }
                    ]
                }
            },
            {
                "if": { "properties": { "action": { "const": "clear" } } },
                "then": {
                    "not": {
                        "anyOf": [
                            { "required": ["url"] },
                            { "required": ["document_reference"] },
                            { "required": ["text_to_display"] },
                            { "required": ["screen_tip"] }
                        ]
                    }
                }
            }
        ]);
    }
}

fn property_schema(tool: &str, name: &str) -> Value {
    if tool == "word.get_text" && name == "limit" {
        return json!({ "type": "integer", "minimum": 1, "maximum": 1000 });
    }
    if let Some(schema) = word_field_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_style_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_document_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_review_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_table_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_image_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_shape_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_header_footer_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_note_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_range_marker_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_list_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_html_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_content_control_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = powerpoint_shape_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = excel_action_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = excel_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = powerpoint_action_property_schema(tool, name) {
        return schema;
    }
    if (tool != "word.list_notes" || name != "limit")
        && !(tool == "word.update_list" && matches!(name, "position" | "level"))
        && let Some(schema) = generic_property_schema(name)
    {
        return schema;
    }
    match name {
        "session_id" => json!({ "type": "string", "description": "Office document session ID." }),
        "anchor" => anchor_schema_for_tool(tool),
        "mode" if tool == "word.set_selection" => {
            json!({ "enum": ["select", "cursor_start", "cursor_end"], "default": "select" })
        }
        "scope" if tool == "word.replace_text" => word_replace_text_scope_schema(),
        "image" => image_schema(),
        "placement" if tool == "word.insert_image" => word_insert_image_placement_schema(),
        "break_type" => {
            json!({ "enum": ["page", "line", "section_next", "section_continuous", "section_even", "section_odd"], "default": "page" })
        }
        "orientation" => json!({ "enum": ["portrait", "landscape"] }),
        "margins_pt" => json!({
            "type": "object",
            "properties": {
                "top": { "type": "number", "minimum": 0 },
                "bottom": { "type": "number", "minimum": 0 },
                "left": { "type": "number", "minimum": 0 },
                "right": { "type": "number", "minimum": 0 }
            },
            "additionalProperties": false
        }),
        "formatting" | "font" => formatting_schema(),
        "paragraph"
            if tool == "word.apply_formatting"
                || tool == "word.create_style"
                || tool == "word.update_style" =>
        {
            paragraph_formatting_schema()
        }
        "title_box" | "content_box" => shape_box_schema(),
        "tools" | "values" | "data" | "formulas" | "number_formats" | "items" | "fields"
        | "borders" | "criteria" | "selected_items" | "shape_ids" | "row_indices"
        | "column_indices" => {
            json!({ "type": "array" })
        }
        _ => json!({ "type": "string" }),
    }
}

fn word_list_property_schema(tool: &str, name: &str) -> Option<Value> {
    if tool != "word.update_list" {
        return None;
    }
    match name {
        "list_id" | "paragraph_index" => Some(json!({ "type": "integer", "minimum": 0 })),
        "level" => Some(json!({ "type": "integer", "minimum": 0, "maximum": 8 })),
        "action" => Some(json!({
            "enum": ["add_item", "set_item_level", "attach_paragraph", "detach_paragraph", "set_level_format"]
        })),
        "position" => {
            Some(json!({ "enum": ["start", "end", "after_paragraph"], "default": "end" }))
        }
        "numbering" => Some(json!({
            "enum": ["bullet", "arabic", "upper_roman", "lower_roman", "upper_letter", "lower_letter", "none"]
        })),
        "text" => Some(json!({ "type": "string", "minLength": 1 })),
        "bullet_char" => Some(json!({ "type": "string", "maxLength": 1 })),
        _ => None,
    }
}

fn excel_property_schema(tool: &str, name: &str) -> Option<Value> {
    document_property_schema(tool, name)
        .or_else(|| excel_workbook_property_schema(tool, name))
        .or_else(|| excel_sheet_property_schema(tool, name))
        .or_else(|| excel_shape_property_schema(tool, name))
}

fn excel_sheet_property_schema(tool: &str, name: &str) -> Option<Value> {
    if tool != "excel.update_sheet" || name != "freeze" {
        return None;
    }
    Some(json!({
        "type": "object",
        "properties": {
            "rows": { "type": "integer", "minimum": 0 },
            "columns": { "type": "integer", "minimum": 0 },
            "at": { "type": "string" },
            "unfreeze": { "type": "boolean" }
        },
        "additionalProperties": false
    }))
}

fn excel_workbook_property_schema(tool: &str, name: &str) -> Option<Value> {
    match (tool, name) {
        ("excel.calculate", "type") => Some(
            json!({ "enum": ["recalculate", "full", "full_rebuild"], "default": "recalculate" }),
        ),
        ("excel.list_named_items", "scope") => {
            Some(json!({ "enum": ["workbook", "sheet", "all"], "default": "all" }))
        }
        ("excel.update_named_item", "scope") => {
            Some(json!({ "enum": ["workbook", "sheet"], "default": "workbook" }))
        }
        ("excel.update_named_item", "name")
        | ("excel.add_comment", "cell")
        | ("excel.add_comment" | "excel.update_comment", "text")
        | ("excel.update_comment", "comment_id" | "reply_id")
        | ("excel.set_hyperlink", "document_reference")
        | ("excel.update_conditional_format", "id")
        | ("excel.copy_range", "source_address" | "destination_address") => {
            Some(json!({ "type": "string", "minLength": 1 }))
        }
        ("excel.copy_range", "copy_type") => Some(
            json!({ "enum": ["all", "values", "formulas", "formats", "link"], "default": "all" }),
        ),
        ("excel.copy_range", "autofill_type") => Some(
            json!({ "enum": ["default", "copy", "series", "formats", "values", "flash_fill"], "default": "default" }),
        ),
        ("excel.update_range_structure", "action") => Some(json!({ "enum": ["insert", "delete"] })),
        ("excel.update_range_structure", "shift") => {
            Some(json!({ "enum": ["down", "right", "up", "left"] }))
        }
        ("excel.update_range_structure", "count") => {
            Some(json!({ "type": "integer", "minimum": 1, "default": 1 }))
        }
        ("excel.clear_range", "apply_to") => {
            Some(json!({ "enum": ["contents", "formats", "all"], "default": "contents" }))
        }
        ("excel.set_hyperlink", "url") => Some(json!({ "type": "string", "format": "uri" })),
        ("excel.set_hyperlink", "text_to_display" | "screen_tip")
        | ("excel.copy_range", "source_sheet" | "destination_sheet") => {
            Some(json!({ "type": "string" }))
        }
        ("excel.format_range", "merge") => {
            Some(json!({ "enum": ["merge", "merge_across", "unmerge"] }))
        }
        ("excel.format_range", "column_width_pt" | "row_height_pt") => {
            Some(json!({ "type": "number", "minimum": 0 }))
        }
        ("excel.format_range", "style") => Some(json!({ "type": "string", "minLength": 1 })),
        ("excel.set_data_validation", "rule") => Some(excel_data_validation_rule_schema()),
        ("excel.set_data_validation", "error_alert") => Some(json!({
            "type": "object",
            "properties": {
                "style": { "enum": ["stop", "warning", "information"], "default": "stop" },
                "title": { "type": "string" },
                "message": { "type": "string" },
                "show_alert": { "type": "boolean", "default": true }
            },
            "additionalProperties": false
        })),
        ("excel.set_data_validation", "input_prompt") => Some(json!({
            "type": "object",
            "properties": {
                "title": { "type": "string" },
                "message": { "type": "string" },
                "show_prompt": { "type": "boolean", "default": true }
            },
            "additionalProperties": false
        })),
        ("excel.update_conditional_format", "rule") => Some(excel_conditional_format_rule_schema()),
        ("excel.update_conditional_format", "priority") => {
            Some(json!({ "type": "integer", "minimum": 0 }))
        }
        ("excel.update_conditional_format", "stop_if_true")
        | ("excel.set_data_validation", "ignore_blanks")
        | ("excel.copy_range", "skip_blanks" | "transpose")
        | ("excel.list_comments", "resolved") => Some(json!({ "type": "boolean" })),
        _ => None,
    }
}

fn excel_shape_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_shape_tool = matches!(
        tool,
        "excel.insert_image" | "excel.list_shapes" | "excel.update_shape"
    );
    if !is_shape_tool {
        return None;
    }
    match name {
        "image" if tool == "excel.insert_image" => Some(json!({
            "oneOf": [
                {
                    "type": "object",
                    "required": ["base64"],
                    "properties": {
                        "base64": { "type": "string", "minLength": 1 },
                        "mime_type": { "type": "string" },
                        "byte_length": { "type": "integer", "minimum": 0 }
                    },
                    "additionalProperties": false
                },
                {
                    "type": "object",
                    "required": ["url"],
                    "properties": {
                        "url": { "type": "string", "format": "uri" }
                    },
                    "additionalProperties": false
                }
            ]
        })),
        "shape_id" if tool == "excel.update_shape" => {
            Some(json!({ "type": "string", "minLength": 1 }))
        }
        "action" if tool == "excel.update_shape" => Some(json!({
            "enum": ["move", "resize", "set_alt_text", "set_text", "set_z_order", "delete"]
        })),
        "left_pt" | "top_pt" => Some(json!({ "type": "number" })),
        "width_pt" | "height_pt" => Some(json!({ "type": "number", "exclusiveMinimum": 0 })),
        "alt_text" | "text" => Some(json!({ "type": "string" })),
        "z_order" if tool == "excel.update_shape" => Some(json!({
            "enum": ["bring_forward", "send_backward", "bring_to_front", "send_to_back"]
        })),
        _ => None,
    }
}

fn excel_data_validation_rule_schema() -> Value {
    json!({
        "type": "object",
        "required": ["type"],
        "properties": {
            "type": { "enum": ["list", "whole_number", "decimal", "date", "time", "text_length", "custom"] },
            "operator": { "enum": ["between", "not_between", "equal_to", "not_equal_to", "greater_than", "less_than", "greater_than_or_equal_to", "less_than_or_equal_to"] },
            "value1": { "oneOf": [{ "type": "string" }, { "type": "number" }] },
            "value2": { "oneOf": [{ "type": "string" }, { "type": "number" }] },
            "list_source": {
                "oneOf": [
                    { "type": "array", "items": { "type": "string" }, "minItems": 1 },
                    { "type": "string", "minLength": 1 }
                ]
            },
            "in_cell_dropdown": { "type": "boolean", "default": true },
            "formula": { "type": "string", "minLength": 1 }
        },
        "additionalProperties": false
    })
}

fn excel_conditional_format_rule_schema() -> Value {
    json!({
        "type": "object",
        "required": ["type"],
        "properties": {
            "type": {
                "enum": [
                    "cell_value",
                    "color_scale",
                    "data_bar",
                    "icon_set",
                    "top_bottom",
                    "preset_criteria",
                    "contains_text",
                    "custom_formula"
                ]
            },
            "operator": { "type": "string", "minLength": 1 },
            "values": { "type": "array" },
            "formula": { "type": "string", "minLength": 1 },
            "text": { "type": "string", "minLength": 1 },
            "preset": { "type": "string", "minLength": 1 },
            "rank": { "type": "integer", "minimum": 1 },
            "percent": { "type": "boolean" },
            "colors": {
                "type": "array",
                "items": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
                "minItems": 2,
                "maxItems": 3
            },
            "icon_set": { "type": "string", "minLength": 1 },
            "format": {
                "type": "object",
                "properties": {
                    "fill_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
                    "font_color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" },
                    "bold": { "type": "boolean" },
                    "italic": { "type": "boolean" }
                },
                "additionalProperties": false
            }
        },
        "additionalProperties": false
    })
}

fn excel_action_property_schema(tool: &str, name: &str) -> Option<Value> {
    if name != "action" {
        return None;
    }
    match tool {
        "excel.update_table" => Some(json!({
            "enum": ["metadata", "read", "add_rows", "add_columns", "resize", "rename", "options", "style", "delete"]
        })),
        "excel.update_named_item" => Some(json!({
            "enum": ["add", "edit", "delete"]
        })),
        "excel.update_comment" => Some(json!({
            "enum": ["reply", "edit", "resolve", "reopen", "delete"]
        })),
        "excel.set_hyperlink" | "excel.set_data_validation" => Some(json!({
            "enum": ["set", "clear"]
        })),
        "excel.copy_range" => Some(json!({
            "enum": ["copy", "autofill"]
        })),
        "excel.update_conditional_format" => Some(json!({
            "enum": ["add", "delete", "clear_range"]
        })),
        "excel.update_shape" => Some(json!({
            "enum": ["move", "resize", "set_alt_text", "set_text", "set_z_order", "delete"]
        })),
        "excel.update_chart" => Some(json!({
            "enum": ["metadata", "read", "title", "legend", "axis", "data", "series_source", "position", "size", "export_image", "delete"]
        })),
        "excel.update_pivot_table" => Some(json!({
            "enum": ["metadata", "read", "refresh", "add_hierarchy", "remove_hierarchy", "layout", "filter", "clear_filters", "delete"]
        })),
        _ => None,
    }
}

fn powerpoint_action_property_schema(tool: &str, name: &str) -> Option<Value> {
    if name != "action" {
        return None;
    }
    match tool {
        "powerpoint.update_tags" => {
            Some(json!({ "enum": ["list", "set", "delete"], "default": "list" }))
        }
        "powerpoint.update_shape" => Some(json!({
            "enum": ["move", "resize", "rotate", "rename", "set_alt_text", "set_fill", "set_line", "set_z_order", "group", "ungroup", "delete"]
        })),
        "powerpoint.update_table" => Some(json!({
            "enum": ["set_values", "set_cell", "add_rows", "delete_rows", "add_columns", "delete_columns", "merge_cells", "split_cell", "clear", "style", "delete"]
        })),
        _ => None,
    }
}

fn word_content_control_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_content_control_tool = matches!(
        tool,
        "word.list_content_controls"
            | "word.insert_content_control"
            | "word.update_content_control"
    );
    if !is_content_control_tool {
        return None;
    }
    match name {
        "type" => Some(
            json!({ "enum": ["rich_text", "plain_text", "checkbox", "dropdown_list", "combo_box"] }),
        ),
        "checked" => Some(json!({ "type": "boolean" })),
        "list_items" if tool == "word.insert_content_control" => {
            Some(content_control_list_items_schema(false))
        }
        "list_items_add" if tool == "word.update_content_control" => {
            Some(content_control_list_items_schema(true))
        }
        "list_items_delete" if tool == "word.update_content_control" => {
            Some(json!({ "type": "array", "items": { "type": "string" } }))
        }
        "list_items_clear" if tool == "word.update_content_control" => {
            Some(json!({ "type": "boolean", "default": false }))
        }
        "selected_value" if tool == "word.update_content_control" => {
            Some(json!({ "type": "string" }))
        }
        _ => None,
    }
}

fn powerpoint_shape_property_schema(tool: &str, name: &str) -> Option<Value> {
    if tool != "powerpoint.add_shape" {
        return None;
    }
    match name {
        "shape_type" => Some(
            json!({ "enum": ["text_box", "rectangle", "ellipse", "rounded_rectangle", "line"] }),
        ),
        "text" => Some(json!({ "type": "string" })),
        _ => None,
    }
}

fn content_control_list_items_schema(with_index: bool) -> Value {
    let mut properties = json!({
        "display_text": { "type": "string", "minLength": 1 },
        "value": { "type": "string" }
    });
    if with_index {
        properties["index"] = json!({ "type": "integer", "minimum": 0 });
    }
    json!({
        "type": "array",
        "minItems": 1,
        "items": {
            "type": "object",
            "required": ["display_text"],
            "properties": properties,
            "additionalProperties": false
        }
    })
}

fn word_html_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_html_tool = matches!(tool, "word.get_html" | "word.insert_html");
    if !is_html_tool {
        return None;
    }
    match name {
        "html" if tool == "word.insert_html" => {
            Some(json!({ "type": "string", "minLength": 1, "maxLength": 1_000_000 }))
        }
        "insert_location" if tool == "word.insert_html" => Some(
            json!({ "enum": ["replace", "before", "after", "start", "end"], "default": "after" }),
        ),
        _ => None,
    }
}

fn word_image_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_image_tool = matches!(
        tool,
        "word.list_images" | "word.get_image" | "word.update_image"
    );
    match (is_image_tool, name) {
        (true, "image") => Some(inline_image_locator_schema()),
        (true, "action") => Some(
            json!({ "enum": ["resize", "set_alt_text", "set_hyperlink", "replace", "delete"] }),
        ),
        (true, "width_pt" | "height_pt") => {
            Some(json!({ "type": "number", "exclusiveMinimum": 0 }))
        }
        (true, "preserve_aspect_ratio") => Some(json!({ "type": "boolean", "default": true })),
        (true, "alt_text_title" | "alt_text_description" | "base64") => {
            Some(json!({ "type": "string" }))
        }
        (true, "hyperlink") => Some(json!({ "type": "string", "format": "uri" })),
        _ => None,
    }
}

fn word_shape_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_shape_tool = matches!(
        tool,
        "word.list_shapes" | "word.insert_shape" | "word.update_shape" | "word.delete_shape"
    );
    if !is_shape_tool {
        return None;
    }
    match name {
        "scope" if tool == "word.list_shapes" => {
            Some(json!({ "enum": ["body", "paragraph", "anchor"], "default": "body" }))
        }
        "shape_type" if tool == "word.insert_shape" => Some(
            json!({ "enum": ["text_box", "rectangle", "ellipse", "rounded_rectangle", "line", "picture"] }),
        ),
        "shape_id" => Some(json!({ "type": "integer", "minimum": 0 })),
        "image" if tool == "word.insert_shape" => Some(json!({
            "oneOf": [
                {
                    "type": "object",
                    "required": ["base64"],
                    "properties": {
                        "base64": { "type": "string", "minLength": 1 },
                        "mime_type": { "type": "string" },
                        "byte_length": { "type": "integer", "minimum": 0 }
                    },
                    "additionalProperties": false
                }
            ]
        })),
        "action" if tool == "word.update_shape" => Some(json!({
            "enum": ["move", "resize", "set_text", "set_alt_text", "set_fill", "set_line", "set_wrap", "set_visibility"]
        })),
        "wrap_type" if tool == "word.update_shape" => {
            Some(json!({ "enum": ["inline", "square", "tight", "behind", "front", "top_bottom"] }))
        }
        "name" | "text" | "alt_text_description" | "fill_color" | "line_color" => {
            Some(json!({ "type": "string" }))
        }
        "visible" => Some(json!({ "type": "boolean" })),
        _ => None,
    }
}

fn word_document_property_schema(tool: &str, name: &str) -> Option<Value> {
    document_property_schema(tool, name)
}

fn document_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_document_property_tool = matches!(
        tool,
        "word.get_document_properties"
            | "word.update_document_properties"
            | "excel.get_document_properties"
            | "excel.update_document_properties"
    );
    if !is_document_property_tool {
        return None;
    }
    match name {
        "include_custom" => Some(json!({ "type": "boolean", "default": true })),
        "title" | "subject" | "author" | "keywords" | "category" | "comments" | "company"
        | "manager" => Some(json!({ "type": "string" })),
        "custom_set" => Some(json!({
            "type": "array",
            "items": {
                "type": "object",
                "required": ["key", "value"],
                "properties": {
                    "key": { "type": "string", "minLength": 1 },
                    "value": {
                        "oneOf": [
                            { "type": "string" },
                            { "type": "number" },
                            { "type": "boolean" }
                        ]
                    }
                },
                "additionalProperties": false
            }
        })),
        "custom_delete" => Some(json!({
            "type": "array",
            "items": { "type": "string", "minLength": 1 }
        })),
        _ => None,
    }
}

fn generic_property_schema(name: &str) -> Option<Value> {
    match name {
        "index" | "offset" | "limit" | "occurrence" | "rows" | "cols" | "row" | "col"
        | "table_index" | "change_index" | "content_control_id" | "position" | "slice_size"
        | "section_index" | "slide_index" | "target_index" | "row_index" | "column_index"
        | "row_count" | "column_count" | "count" | "max_depth" | "expected_count" => {
            Some(json!({ "type": "integer", "minimum": 0 }))
        }
        "heading_level" | "level" | "columns" => Some(json!({ "type": "integer", "minimum": 1 })),
        "width_pt" | "height_pt" | "left_pt" | "top_pt" | "scale_percent" | "width" | "height"
        | "left" | "top" | "page_width_pt" | "page_height_pt" | "rotation"
        | "fill_transparency" | "line_weight" | "line_transparency" | "font_size" => {
            Some(json!({ "type": "number" }))
        }
        "match_case"
        | "whole_word"
        | "wildcards"
        | "dry_run"
        | "partial_ok"
        | "include_metadata"
        | "include_text_preview"
        | "include_formatting"
        | "include_formulas"
        | "include_hyperlinks"
        | "include_validation"
        | "include_selection"
        | "include_page_setup"
        | "include_tags"
        | "ordered"
        | "all"
        | "activate"
        | "has_headers"
        | "metadata_only"
        | "values_only"
        | "complete_match"
        | "autofit"
        | "wrap_text"
        | "hidden_columns"
        | "hidden_rows"
        | "show_headers"
        | "show_gridlines"
        | "show_headings"
        | "show_totals"
        | "highlight_first_column"
        | "highlight_last_column"
        | "show_banded_columns"
        | "show_banded_rows"
        | "show_filter_button"
        | "include_hidden"
        | "visible"
        | "overlay"
        | "title_visible"
        | "hidden"
        | "is_decorative"
        | "clear_fill"
        | "line_visible"
        | "lock_aspect_ratio"
        | "delete_contents"
        | "keep_text"
        | "overwrite"
        | "validate_only" => Some(json!({ "type": "boolean" })),
        _ => None,
    }
}

fn word_table_property_schema(tool: &str, name: &str) -> Option<Value> {
    if tool != "word.update_table" {
        return None;
    }
    match name {
        "action" => Some(json!({
            "enum": [
                "update_cell",
                "add_row",
                "add_column",
                "format_cell",
                "delete",
                "delete_row",
                "delete_column",
                "merge_cells",
                "set_column_width",
                "distribute_columns",
                "set_borders",
                "set_header_row"
            ]
        })),
        "row_range" | "col_range" => Some(json!({
            "type": "array",
            "items": { "type": "integer", "minimum": 0 },
            "minItems": 2,
            "maxItems": 2
        })),
        "width_pt" => Some(json!({ "type": "number", "exclusiveMinimum": 0 })),
        "header_row" => Some(json!({ "type": "boolean" })),
        "borders" => Some(json!({
            "type": "object",
            "properties": {
                "edges": {
                    "type": "array",
                    "items": { "enum": ["top", "bottom", "left", "right", "inside_horizontal", "inside_vertical", "all"] }
                },
                "style": { "enum": ["single", "double", "dotted", "dashed", "none"] },
                "width_pt": { "type": "number", "minimum": 0 },
                "color": { "type": "string", "pattern": "^#[0-9A-Fa-f]{6}$" }
            },
            "additionalProperties": false
        })),
        _ => None,
    }
}

fn word_style_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_style_tool = matches!(
        tool,
        "word.list_styles" | "word.create_style" | "word.update_style"
    );
    match (is_style_tool, name) {
        (true, "type") => Some(json!({ "enum": ["paragraph", "character", "table", "list"] })),
        (true, "name" | "base_style") => Some(json!({ "type": "string", "minLength": 1 })),
        (true, "built_in" | "in_use_only") => Some(json!({ "type": "boolean" })),
        _ => None,
    }
}

fn word_field_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_field_tool = matches!(
        tool,
        "word.list_fields" | "word.insert_field" | "word.update_field" | "word.delete_field"
    );
    match (is_field_tool, tool, name) {
        (true, "word.list_fields", "limit") => {
            Some(json!({ "type": "integer", "minimum": 1, "maximum": 200 }))
        }
        (true, _, "field_index") => Some(json!({ "type": "integer", "minimum": 0 })),
        (true, "word.insert_field", "field_type") => Some(json!({
            "enum": ["toc", "page", "num_pages", "date", "time", "ref", "hyperlink", "seq", "styleref"]
        })),
        (true, "word.update_field", "action") => {
            Some(json!({ "enum": ["refresh", "refresh_all", "lock", "unlock"] }))
        }
        (true, _, "type" | "code_options") => Some(json!({ "type": "string" })),
        _ => None,
    }
}

fn word_header_footer_property_schema(tool: &str, name: &str) -> Option<Value> {
    match (tool, name) {
        ("word.get_header_footer" | "word.update_header_footer", "location") => {
            Some(json!({ "enum": ["header", "footer"] }))
        }
        (_, "header_footer_type") => {
            Some(json!({ "enum": ["primary", "first_page", "even_pages"], "default": "primary" }))
        }
        ("word.update_header_footer", "action") => {
            Some(json!({ "enum": ["set_text", "append_paragraph", "clear"] }))
        }
        _ => None,
    }
}

fn word_review_property_schema(tool: &str, name: &str) -> Option<Value> {
    match (tool, name) {
        ("word.update_tracked_change", "action") => {
            Some(json!({ "enum": ["accept", "reject", "accept_all", "reject_all"] }))
        }
        ("word.update_comment", "action") => {
            Some(json!({ "enum": ["reply", "edit", "resolve", "reopen", "delete"] }))
        }
        ("word.update_comment", "comment_id" | "reply_id") => {
            Some(json!({ "type": "string", "minLength": 1 }))
        }
        ("word.set_change_tracking", "mode") => {
            Some(json!({ "enum": ["off", "track_all", "track_mine_only"] }))
        }
        _ => None,
    }
}
fn word_range_marker_property_schema(tool: &str, name: &str) -> Option<Value> {
    match (tool, name) {
        ("word.list_hyperlinks", "limit") => {
            Some(json!({ "type": "integer", "minimum": 1, "maximum": 200 }))
        }
        ("word.remove_hyperlink", "keep_text") => {
            Some(json!({ "type": "boolean", "default": true }))
        }
        ("word.insert_bookmark", "name") => Some(json!({
            "type": "string",
            "minLength": 1,
            "pattern": "^[A-Za-z_][A-Za-z0-9_]{0,39}$"
        })),
        ("word.delete_bookmark", "name") => Some(json!({ "type": "string", "minLength": 1 })),
        _ => None,
    }
}

fn word_note_property_schema(tool: &str, name: &str) -> Option<Value> {
    let is_note_tool = matches!(
        tool,
        "word.insert_note" | "word.list_notes" | "word.update_note" | "word.delete_note"
    );
    match (is_note_tool, name) {
        (true, "kind") => Some(json!({ "enum": ["footnote", "endnote"] })),
        (true, "limit") if tool == "word.list_notes" => {
            Some(json!({ "type": "integer", "minimum": 1, "maximum": 200 }))
        }
        (true, "text") if tool == "word.insert_note" => {
            Some(json!({ "type": "string", "minLength": 1 }))
        }
        _ => None,
    }
}

fn validate_anchor_argument(
    tool: &str,
    arguments: &serde_json::Map<String, Value>,
) -> Result<(), String> {
    let Some(anchor) = arguments.get("anchor") else {
        return Ok(());
    };
    let Some(kind) = anchor.get("kind").and_then(Value::as_str) else {
        return Err(format!("{tool} anchor requires string kind."));
    };
    if supported_anchor_kinds(tool).contains(&kind) {
        Ok(())
    } else {
        Err(format!("{tool} does not support anchor kind {kind}."))
    }
}

fn supported_anchor_kinds(_tool: &str) -> &'static [&'static str] {
    &[
        "selection",
        "start_of_document",
        "end_of_document",
        "paragraph_index",
        "before_paragraph_index",
        "after_paragraph_index",
        "before_text",
        "after_text",
        "heading",
        "bookmark",
    ]
}

fn anchor_schema_for_tool(tool: &str) -> Value {
    anchor_schema(supported_anchor_kinds(tool))
}

fn supports_anchor_kind(kinds: &[&str], kind: &str) -> bool {
    kinds.contains(&kind)
}

fn anchor_schema(kinds: &[&str]) -> Value {
    let mut variants = Vec::new();
    if supports_anchor_kind(kinds, "selection") {
        variants.push(json!({ "type": "object", "required": ["kind"], "properties": { "kind": { "const": "selection" } }, "additionalProperties": false }));
    }
    let document_kinds = ["start_of_document", "end_of_document"]
        .into_iter()
        .filter(|kind| supports_anchor_kind(kinds, kind))
        .collect::<Vec<_>>();
    if !document_kinds.is_empty() {
        variants.push(json!({ "type": "object", "required": ["kind"], "properties": { "kind": { "enum": document_kinds } }, "additionalProperties": false }));
    }
    let paragraph_kinds = [
        "paragraph_index",
        "before_paragraph_index",
        "after_paragraph_index",
    ]
    .into_iter()
    .filter(|kind| supports_anchor_kind(kinds, kind))
    .collect::<Vec<_>>();
    if !paragraph_kinds.is_empty() {
        variants.push(json!({ "type": "object", "required": ["kind", "index"], "properties": { "kind": { "enum": paragraph_kinds }, "index": { "type": "integer", "minimum": 0 } }, "additionalProperties": false }));
    }
    let text_kinds = ["after_text", "before_text"]
        .into_iter()
        .filter(|kind| supports_anchor_kind(kinds, kind))
        .collect::<Vec<_>>();
    if !text_kinds.is_empty() {
        variants.push(json!({ "type": "object", "required": ["kind", "text"], "properties": { "kind": { "enum": text_kinds }, "text": { "type": "string", "minLength": 1 }, "occurrence": { "type": "integer", "minimum": 1 } }, "additionalProperties": false }));
    }
    if supports_anchor_kind(kinds, "heading") {
        variants.push(json!({ "type": "object", "required": ["kind", "text"], "properties": { "kind": { "const": "heading" }, "text": { "type": "string", "minLength": 1 }, "level": { "type": "integer", "minimum": 1, "maximum": 9 } }, "additionalProperties": false }));
    }
    if supports_anchor_kind(kinds, "bookmark") {
        variants.push(json!({ "type": "object", "required": ["kind", "name"], "properties": { "kind": { "const": "bookmark" }, "name": { "type": "string", "minLength": 1 } }, "additionalProperties": false }));
    }

    json!({
        "oneOf": variants
    })
}

fn image_schema() -> Value {
    json!({
        "oneOf": [
            { "type": "object", "required": ["base64"], "properties": { "base64": { "type": "string" }, "mime_type": { "type": "string" }, "byte_length": { "type": "integer", "minimum": 0 } }, "additionalProperties": false },
            { "type": "object", "required": ["url"], "properties": { "url": { "type": "string", "format": "uri" }, "mime_type": { "type": "string" }, "byte_length": { "type": "integer", "minimum": 0 } }, "additionalProperties": false }
        ]
    })
}

fn inline_image_locator_schema() -> Value {
    json!({
        "type": "object",
        "required": ["kind", "index"],
        "properties": {
            "kind": { "const": "paragraph_index" },
            "index": { "type": "integer", "minimum": 0 },
            "image_index": { "type": "integer", "minimum": 0, "default": 0 }
        },
        "additionalProperties": false
    })
}

fn word_insert_image_placement_schema() -> Value {
    json!({
        "enum": [
            "inline",
            "before_paragraph",
            "after_paragraph",
            "new_paragraph_before",
            "new_paragraph_after",
            "replace_paragraph",
            "selection"
        ],
        "default": "inline"
    })
}

fn word_replace_text_scope_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "paragraph_range": {
                "type": "array",
                "items": { "type": "integer", "minimum": 0 },
                "minItems": 2,
                "maxItems": 2
            },
            "selection_only": { "type": "boolean" }
        },
        "additionalProperties": false
    })
}

fn formatting_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "bold": { "type": "boolean" },
            "italic": { "type": "boolean" },
            "underline": { "type": "boolean" },
            "strikethrough": { "type": "boolean" },
            "font_name": { "type": "string" },
            "font_size_pt": { "type": "number", "exclusiveMinimum": 0 },
            "color": { "type": "string" },
            "highlight": { "type": "string" }
        },
        "additionalProperties": false
    })
}

fn paragraph_formatting_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "alignment": { "enum": ["left", "center", "right", "justified"] },
            "left_indent_pt": { "type": "number", "minimum": 0 },
            "right_indent_pt": { "type": "number", "minimum": 0 },
            "first_line_indent_pt": { "type": "number" },
            "line_spacing_pt": { "type": "number", "exclusiveMinimum": 0 },
            "line_unit_before": { "type": "number", "minimum": 0 },
            "line_unit_after": { "type": "number", "minimum": 0 },
            "space_before_pt": { "type": "number", "minimum": 0 },
            "space_after_pt": { "type": "number", "minimum": 0 },
            "outline_level": { "type": "integer", "minimum": 0, "maximum": 9 }
        },
        "additionalProperties": false
    })
}

fn shape_box_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "left": { "type": "number" },
            "top": { "type": "number" },
            "width": { "type": "number" },
            "height": { "type": "number" }
        },
        "additionalProperties": false
    })
}

#[cfg(test)]
#[path = "catalog_tests.rs"]
mod catalog_tests;
