use serde_json::{Value, json};

pub const WORD_V1_TOOLS: &[&str] = &[
    "word.accept_change",
    "word.add_column",
    "word.add_comment",
    "word.add_row",
    "word.apply_formatting",
    "word.apply_style",
    "word.delete_range",
    "word.find_text",
    "word.format_cell",
    "word.get_outline",
    "word.get_paragraph",
    "word.get_selection",
    "word.get_text",
    "word.insert_heading",
    "word.insert_image",
    "word.insert_list",
    "word.insert_page_break",
    "word.insert_paragraph",
    "word.insert_table",
    "word.read_table",
    "word.reject_change",
    "word.replace_text",
    "word.resolve_comment",
    "word.save",
    "word.set_heading_level",
    "word.update_cell",
    "word.update_paragraph",
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
        name: "powerpoint.add_slide",
    },
    PowerPointToolDefinition {
        name: "powerpoint.apply_layout",
    },
    PowerPointToolDefinition {
        name: "powerpoint.export_pdf",
    },
    PowerPointToolDefinition {
        name: "powerpoint.insert_image",
    },
    PowerPointToolDefinition {
        name: "powerpoint.replace_text",
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
    ];
    tools.extend(WORD_V1_TOOLS.iter().map(|name| {
        tool_json(
            name,
            name,
            "Forward this Word tool call to the selected Office document session.",
        )
    }));
    tools.extend(ExcelToolCatalog::tools().iter().map(|tool| {
        tool_json(
            tool.name,
            tool.name,
            "Forward this Excel tool call to the selected Office workbook session.",
        )
    }));
    tools.extend(PowerPointToolCatalog::tools().iter().map(|tool| {
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
    json!({
        "name": name,
        "title": title,
        "description": description,
        "inputSchema": {
            "type": "object",
            "additionalProperties": true
        }
    })
}

#[cfg(test)]
#[path = "catalog_tests.rs"]
mod catalog_tests;
