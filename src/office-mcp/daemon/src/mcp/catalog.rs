use crate::mcp::{ToolAccessPolicy, ToolSideEffect, tool_metadata};
use serde_json::{Value, json};

pub const WORD_V1_TOOLS: &[&str] = &[
    "word.add_comment",
    "word.apply_formatting",
    "word.apply_style",
    "word.delete_content_control",
    "word.delete_range",
    "word.find_text",
    "word.get_outline",
    "word.get_paragraph",
    "word.get_selection",
    "word.get_text",
    "word.insert_content_control",
    "word.insert_image",
    "word.insert_list",
    "word.insert_page_break",
    "word.insert_paragraph",
    "word.insert_table",
    "word.list_content_controls",
    "word.read_table",
    "word.replace_text",
    "word.resolve_anchor",
    "word.resolve_comment",
    "word.resize_image",
    "word.save",
    "word.update_content_control",
    "word.update_paragraph",
    "word.update_table",
    "word.update_tracked_change",
];

const EXCEL_V1_TOOLS: &[ExcelToolDefinition] = &[
    ExcelToolDefinition {
        name: "excel.add_sheet",
    },
    ExcelToolDefinition {
        name: "excel.apply_filter",
    },
    ExcelToolDefinition {
        name: "excel.clear_range",
    },
    ExcelToolDefinition {
        name: "excel.create_chart",
    },
    ExcelToolDefinition {
        name: "excel.create_pivot_table",
    },
    ExcelToolDefinition {
        name: "excel.create_table",
    },
    ExcelToolDefinition {
        name: "excel.delete_sheet",
    },
    ExcelToolDefinition {
        name: "excel.find_replace_cells",
    },
    ExcelToolDefinition {
        name: "excel.format_range",
    },
    ExcelToolDefinition {
        name: "excel.get_used_range",
    },
    ExcelToolDefinition {
        name: "excel.get_workbook_info",
    },
    ExcelToolDefinition {
        name: "excel.list_sheets",
    },
    ExcelToolDefinition {
        name: "excel.read_range",
    },
    ExcelToolDefinition {
        name: "excel.set_formula",
    },
    ExcelToolDefinition {
        name: "excel.sort_range",
    },
    ExcelToolDefinition {
        name: "excel.update_chart",
    },
    ExcelToolDefinition {
        name: "excel.update_pivot_table",
    },
    ExcelToolDefinition {
        name: "excel.update_table",
    },
    ExcelToolDefinition {
        name: "excel.update_sheet",
    },
    ExcelToolDefinition {
        name: "excel.write_range",
    },
];

const POWERPOINT_V1_TOOLS: &[PowerPointToolDefinition] = &[
    PowerPointToolDefinition {
        name: "powerpoint.get_presentation_info",
    },
    PowerPointToolDefinition {
        name: "powerpoint.get_active_view",
    },
    PowerPointToolDefinition {
        name: "powerpoint.export_file",
    },
    PowerPointToolDefinition {
        name: "powerpoint.update_tags",
    },
    PowerPointToolDefinition {
        name: "powerpoint.list_slides",
    },
    PowerPointToolDefinition {
        name: "powerpoint.add_slide",
    },
    PowerPointToolDefinition {
        name: "powerpoint.update_slide",
    },
    PowerPointToolDefinition {
        name: "powerpoint.delete_slide",
    },
    PowerPointToolDefinition {
        name: "powerpoint.move_slide",
    },
    PowerPointToolDefinition {
        name: "powerpoint.export_slide",
    },
    PowerPointToolDefinition {
        name: "powerpoint.list_layouts",
    },
    PowerPointToolDefinition {
        name: "powerpoint.apply_layout",
    },
    PowerPointToolDefinition {
        name: "powerpoint.get_selection",
    },
    PowerPointToolDefinition {
        name: "powerpoint.set_selection",
    },
    PowerPointToolDefinition {
        name: "powerpoint.list_shapes",
    },
    PowerPointToolDefinition {
        name: "powerpoint.add_text_box",
    },
    PowerPointToolDefinition {
        name: "powerpoint.add_shape",
    },
    PowerPointToolDefinition {
        name: "powerpoint.insert_image",
    },
    PowerPointToolDefinition {
        name: "powerpoint.update_shape",
    },
    PowerPointToolDefinition {
        name: "powerpoint.read_text",
    },
    PowerPointToolDefinition {
        name: "powerpoint.replace_text",
    },
    PowerPointToolDefinition {
        name: "powerpoint.format_text",
    },
    PowerPointToolDefinition {
        name: "powerpoint.add_table",
    },
    PowerPointToolDefinition {
        name: "powerpoint.read_table",
    },
    PowerPointToolDefinition {
        name: "powerpoint.update_table",
    },
];

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExcelToolDefinition {
    pub name: &'static str,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointToolDefinition {
    pub name: &'static str,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExcelToolCatalog;

impl ExcelToolCatalog {
    #[must_use]
    pub const fn tools() -> &'static [ExcelToolDefinition] {
        EXCEL_V1_TOOLS
    }

    #[must_use]
    pub fn contains(name: &str) -> bool {
        Self::tools().iter().any(|tool| tool.name == name)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointToolCatalog;

impl PowerPointToolCatalog {
    #[must_use]
    pub const fn tools() -> &'static [PowerPointToolDefinition] {
        POWERPOINT_V1_TOOLS
    }

    #[must_use]
    pub fn contains(name: &str) -> bool {
        Self::tools().iter().any(|tool| tool.name == name)
    }
}

#[must_use]
pub fn tool_catalog_json() -> Vec<Value> {
    tool_catalog_json_for_policy(&ToolAccessPolicy::default())
}

#[must_use]
pub fn tool_catalog_json_for_policy(policy: &ToolAccessPolicy) -> Vec<Value> {
    let mut tools = vec![
        tool_json(
            "office.list_sessions",
            "List Office Sessions",
            "List connected Office document sessions.",
        ),
        tool_json(
            "office.get_session_info",
            "Get Office Session Info",
            "Return metadata and supported tools for one Office document session.",
        ),
        tool_json(
            "office.describe_tools",
            "Describe Office Tools",
            "Return runtime contracts, examples, and common errors for multiple Office MCP tools.",
        ),
    ];
    tools.extend(
        WORD_V1_TOOLS
            .iter()
            .filter(|name| policy.allows_tool(name))
            .map(|name| {
                tool_json(
                    name,
                    name,
                    "Forward this Word tool call to the selected Office document session.",
                )
            }),
    );
    tools.extend(
        ExcelToolCatalog::tools()
            .iter()
            .filter(|tool| policy.allows_tool(tool.name))
            .map(|tool| {
                tool_json(
                    tool.name,
                    tool.name,
                    "Forward this Excel tool call to the selected Office workbook session.",
                )
            }),
    );
    tools.extend(PowerPointToolCatalog::tools().iter().filter(|tool| policy.allows_tool(tool.name)).map(|tool| {
        tool_json(
            tool.name,
            tool.name,
            "Forward this PowerPoint tool call to the selected Office presentation session.",
        )
    }));
    tools
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
    let side_effect = tool_side_effect(name).unwrap_or("read");
    let mut tool = json!({
        "name": name,
        "title": title,
        "description": description,
        "inputSchema": input_schema_for_tool(name),
        "annotations": tool_annotations(side_effect),
        "_meta": {
            "com.office-mcp/side_effects": side_effect,
            "com.office-mcp/common_errors": common_errors_for_tool(name),
            "com.office-mcp/examples": examples_for_tool(name)
        }
    });
    if let Some(metadata) = tool_metadata(name) {
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
    let spec = tool_input_spec(tool);
    object_schema(tool, spec.required, spec.properties)
}

#[must_use]
pub fn describe_tool_contract(tool: &str) -> Option<Value> {
    let side_effect = tool_side_effect(tool)?;
    let input_schema = input_schema_for_tool(tool);
    let mut contract = json!({
        "name": tool,
        "input_schema": input_schema,
        "parameters": parameters_for_schema(&input_schema),
        "examples": examples_for_tool(tool),
        "side_effect": side_effect,
        "common_errors": common_errors_for_tool(tool)
    });
    if let Some(metadata) = tool_metadata(tool) {
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
        "word.insert_content_control" | "word.update_content_control" | "word.delete_content_control" => {
            errors.push(json!({
                "code": "INVALID_ARGUMENTS",
                "cause": "Content control IDs are runtime identifiers; refresh the list before updating stale IDs."
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
        ["session_id", "index"]
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
    tool_spec!("word.get_selection", ["session_id"], ["session_id"]),
    tool_spec!("word.save", ["session_id"], ["session_id"]),
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
        "word.insert_page_break",
        ["session_id", "anchor"],
        ["session_id", "anchor", "match_case"]
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
        "word.apply_formatting",
        ["session_id", "anchor", "formatting"],
        ["session_id", "anchor", "formatting", "match_case"]
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
        "word.update_tracked_change",
        [
            "session_id",
            "action",
            "change_index",
            "expected_fingerprint"
        ],
        [
            "session_id",
            "action",
            "change_index",
            "expected_fingerprint"
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
    json!({
        "type": "object",
        "required": required,
        "properties": map,
        "additionalProperties": false
    })
}

fn property_schema(tool: &str, name: &str) -> Value {
    match name {
        "session_id" => json!({ "type": "string", "description": "Office document session ID." }),
        "index" | "offset" | "limit" | "occurrence" | "rows" | "cols" | "row" | "col"
        | "table_index" | "change_index" | "content_control_id" | "position" | "slice_size"
        | "slide_index" | "target_index" | "row_index" | "column_index" | "row_count"
        | "column_count" | "count" | "max_depth" => json!({ "type": "integer", "minimum": 0 }),
        "heading_level" | "level" | "columns" => json!({ "type": "integer", "minimum": 1 }),
        "width_pt" | "height_pt" | "scale_percent" | "width" | "height" | "left" | "top"
        | "rotation" | "fill_transparency" | "line_weight" | "line_transparency" | "font_size" => {
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
        | "visible"
        | "overlay"
        | "title_visible"
        | "hidden"
        | "is_decorative"
        | "clear_fill"
        | "line_visible"
        | "lock_aspect_ratio"
        | "delete_contents"
        | "validate_only" => json!({ "type": "boolean" }),
        "anchor" => anchor_schema_for_tool(tool),
        "scope" if tool == "word.replace_text" => word_replace_text_scope_schema(),
        "image" => image_schema(),
        "placement" if tool == "word.insert_image" => word_insert_image_placement_schema(),
        "formatting" => formatting_schema(),
        "title_box" | "content_box" => shape_box_schema(),
        "tools" | "values" | "data" | "formulas" | "number_formats" | "items" | "fields"
        | "borders" | "criteria" | "selected_items" | "shape_ids" | "row_indices"
        | "column_indices" => {
            json!({ "type": "array" })
        }
        _ => json!({ "type": "string" }),
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
