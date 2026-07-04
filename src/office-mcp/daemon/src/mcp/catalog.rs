use crate::mcp::{ToolAccessPolicy, ToolSideEffect, tool_metadata};
use serde_json::{Value, json};

pub const WORD_V1_TOOLS: &[&str] = &[
    "word.add_comment",
    "word.apply_formatting",
    "word.apply_style",
    "word.delete_bookmark",
    "word.delete_content_control",
    "word.delete_field",
    "word.delete_range",
    "word.find_text",
    "word.get_header_footer",
    "word.get_outline",
    "word.get_paragraph",
    "word.get_selection",
    "word.get_text",
    "word.insert_bookmark",
    "word.insert_content_control",
    "word.insert_break",
    "word.insert_field",
    "word.insert_hyperlink",
    "word.insert_image",
    "word.insert_list",
    "word.insert_note",
    "word.insert_paragraph",
    "word.insert_table",
    "word.list_bookmarks",
    "word.list_content_controls",
    "word.list_fields",
    "word.list_hyperlinks",
    "word.list_notes",
    "word.list_sections",
    "word.read_table",
    "word.remove_hyperlink",
    "word.replace_text",
    "word.resolve_anchor",
    "word.resolve_comment",
    "word.resize_image",
    "word.save",
    "word.set_change_tracking",
    "word.update_header_footer",
    "word.update_content_control",
    "word.update_field",
    "word.update_note",
    "word.update_page_setup",
    "word.update_paragraph",
    "word.update_table",
    "word.update_tracked_change",
    "word.delete_note",
];

const EXCEL_V1_TOOLS: &[OfficeToolDefinition] = &[
    "excel.add_sheet",
    "excel.apply_filter",
    "excel.clear_range",
    "excel.create_chart",
    "excel.create_pivot_table",
    "excel.create_table",
    "excel.delete_sheet",
    "excel.find_replace_cells",
    "excel.format_range",
    "excel.get_used_range",
    "excel.get_workbook_info",
    "excel.list_sheets",
    "excel.read_range",
    "excel.set_formula",
    "excel.sort_range",
    "excel.update_chart",
    "excel.update_pivot_table",
    "excel.update_table",
    "excel.update_sheet",
    "excel.write_range",
];

const POWERPOINT_V1_TOOLS: &[OfficeToolDefinition] = &[
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
        "word.insert_bookmark" | "word.list_bookmarks" | "word.delete_bookmark" => {
            bookmark_examples_for_tool(tool)
        }
        "word.insert_note" | "word.list_notes" | "word.update_note" | "word.delete_note" => {
            note_examples_for_tool(tool)
        }
        "word.list_fields" | "word.insert_field" | "word.update_field" | "word.delete_field" => {
            field_examples_for_tool(tool)
        }
        "word.update_table" => vec![json!({
            "description": "Replace one table cell by row and column index.",
            "arguments": {
                "session_id": "session-1",
                "table_index": 0,
                "action": "set_cell_text",
                "row": 1,
                "col": 2,
                "text": "Approved"
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
        ["session_id", "offset", "limit", "include_metadata"]
    ),
    tool_spec!(
        "word.get_outline",
        ["session_id"],
        ["session_id", "max_depth"]
    ),
    tool_spec!(
        "word.get_paragraph",
        ["session_id", "index"],
        ["session_id", "index", "include_formatting"]
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
    tool_spec!(
        "word.resize_image",
        ["session_id", "image"],
        [
            "session_id",
            "image",
            "width_pt",
            "height_pt",
            "scale_percent",
            "lock_aspect_ratio"
        ]
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
            "data",
            "rows",
            "cols"
        ]
    ),
    tool_spec!(
        "word.list_content_controls",
        ["session_id"],
        ["session_id", "tag", "title"]
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
            "match_case"
        ]
    ),
    tool_spec!(
        "word.update_content_control",
        ["session_id", "content_control_id"],
        ["session_id", "content_control_id", "text", "tag", "title"]
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
        "word.resolve_comment",
        ["session_id", "comment_id"],
        ["session_id", "comment_id"]
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
            "activate"
        ]
    ),
    tool_spec!(
        "excel.delete_sheet",
        ["session_id", "sheet"],
        ["session_id", "sheet"]
    ),
    tool_spec!(
        "excel.get_used_range",
        ["session_id"],
        ["session_id", "sheet", "values_only"]
    ),
    tool_spec!(
        "excel.read_range",
        ["session_id", "address"],
        [
            "session_id",
            "sheet",
            "address",
            "include_formulas",
            "include_formatting"
        ]
    ),
    tool_spec!(
        "excel.write_range",
        ["session_id", "address", "values"],
        ["session_id", "sheet", "address", "values"]
    ),
    tool_spec!(
        "excel.clear_range",
        ["session_id", "address"],
        ["session_id", "sheet", "address", "apply_to", "delete_shift"]
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
            "autofit",
            "borders"
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
        "powerpoint.get_presentation_info",
        ["session_id"],
        ["session_id", "include_selection"]
    ),
    tool_spec!("powerpoint.get_active_view", ["session_id"], ["session_id"]),
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
        "powerpoint.add_text_box",
        ["session_id", "text"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "text",
            "left",
            "top",
            "width",
            "height",
            "name"
        ]
    ),
    tool_spec!(
        "powerpoint.add_shape",
        ["session_id", "type"],
        [
            "session_id",
            "slide_id",
            "slide_index",
            "type",
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
    schema
}

fn property_schema(tool: &str, name: &str) -> Value {
    if let Some(schema) = word_field_property_schema(tool, name) {
        return schema;
    }
    if let Some(schema) = word_review_property_schema(tool, name) {
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
    match name {
        "session_id" => json!({ "type": "string", "description": "Office document session ID." }),
        "limit" if tool == "word.list_notes" => {
            json!({ "type": "integer", "minimum": 1, "maximum": 200 })
        }
        "index" | "offset" | "limit" | "occurrence" | "rows" | "cols" | "row" | "col"
        | "table_index" | "change_index" | "content_control_id" | "position" | "slice_size"
        | "section_index" | "slide_index" | "target_index" | "row_index" | "column_index"
        | "row_count" | "column_count" | "count" | "max_depth" | "expected_count" => {
            json!({ "type": "integer", "minimum": 0 })
        }
        "heading_level" | "level" | "columns" => json!({ "type": "integer", "minimum": 1 }),
        "width_pt" | "height_pt" | "scale_percent" | "width" | "height" | "left" | "top"
        | "page_width_pt" | "page_height_pt" | "rotation" | "fill_transparency" | "line_weight"
        | "line_transparency" | "font_size" => {
            json!({ "type": "number" })
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
        | "include_selection"
        | "include_page_setup"
        | "include_tags"
        | "ordered"
        | "all"
        | "activate"
        | "has_headers"
        | "values_only"
        | "complete_match"
        | "autofit"
        | "wrap_text"
        | "show_headers"
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
        | "validate_only" => json!({ "type": "boolean" }),
        "anchor" => anchor_schema_for_tool(tool),
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
        "formatting" => formatting_schema(),
        "paragraph" if tool == "word.apply_formatting" => paragraph_formatting_schema(),
        "title_box" | "content_box" => shape_box_schema(),
        "tools" | "values" | "data" | "formulas" | "number_formats" | "items" | "fields"
        | "borders" | "criteria" | "selected_items" | "shape_ids" | "row_indices"
        | "column_indices" => {
            json!({ "type": "array" })
        }
        _ => json!({ "type": "string" }),
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
            "font_color": { "type": "string" },
            "highlight_color": { "type": "string" },
            "font_size": { "type": "number" }
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
