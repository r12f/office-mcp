#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AccessMode {
    Read,
    Write,
    #[default]
    All,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ToolSideEffect {
    Read,
    Mutating,
    Destructive,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ToolMetadata {
    pub name: &'static str,
    pub app: &'static str,
    pub category: &'static str,
    pub side_effect: ToolSideEffect,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiToolAccessPolicySnapshot {
    pub access_mode: AccessMode,
    pub disabled_apps: Vec<String>,
    pub disabled_categories: Vec<(String, String)>,
    pub disabled_tools: Vec<String>,
}

impl Default for UiToolAccessPolicySnapshot {
    fn default() -> Self {
        Self {
            access_mode: AccessMode::All,
            disabled_apps: Vec::new(),
            disabled_categories: Vec::new(),
            disabled_tools: Vec::new(),
        }
    }
}

impl AccessMode {
    #[must_use]
    pub const fn allows(self, side_effect: ToolSideEffect) -> bool {
        match self {
            Self::Read => matches!(side_effect, ToolSideEffect::Read),
            Self::Write => !matches!(side_effect, ToolSideEffect::Destructive),
            Self::All => true,
        }
    }
}

#[must_use]
pub fn tool_metadata(name: &str) -> Option<ToolMetadata> {
    TOOL_METADATA
        .iter()
        .copied()
        .find(|metadata| metadata.name == name)
}

#[must_use]
pub fn tool_metadata_catalog() -> &'static [ToolMetadata] {
    TOOL_METADATA
}

const TOOL_METADATA: &[ToolMetadata] = &[
    tool(
        "word.get_text",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.get_outline",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.get_header_footer",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.get_paragraph",
        "word",
        "Paragraphs & lists",
        ToolSideEffect::Read,
    ),
    tool(
        "word.find_text",
        "word",
        "Range & selection",
        ToolSideEffect::Read,
    ),
    tool(
        "word.resolve_anchor",
        "word",
        "Range & selection",
        ToolSideEffect::Read,
    ),
    tool(
        "word.get_selection",
        "word",
        "Range & selection",
        ToolSideEffect::Read,
    ),
    tool(
        "word.insert_paragraph",
        "word",
        "Paragraphs & lists",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.insert_image",
        "word",
        "Media",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.resize_image",
        "word",
        "Media",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.insert_table",
        "word",
        "Tables",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.insert_break",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.list_sections",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.update_page_setup",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.insert_list",
        "word",
        "Paragraphs & lists",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.insert_hyperlink",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.list_hyperlinks",
        "word",
        "Range & selection",
        ToolSideEffect::Read,
    ),
    tool(
        "word.remove_hyperlink",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_header_footer",
        "word",
        "Document & structure",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.replace_text",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_paragraph",
        "word",
        "Paragraphs & lists",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.delete_range",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.insert_bookmark",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.list_bookmarks",
        "word",
        "Range & selection",
        ToolSideEffect::Read,
    ),
    tool(
        "word.delete_bookmark",
        "word",
        "Range & selection",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.insert_note",
        "word",
        "Notes",
        ToolSideEffect::Mutating,
    ),
    tool("word.list_notes", "word", "Notes", ToolSideEffect::Read),
    tool(
        "word.update_note",
        "word",
        "Notes",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.delete_note",
        "word",
        "Notes",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.apply_formatting",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool("word.read_table", "word", "Tables", ToolSideEffect::Read),
    tool(
        "word.update_table",
        "word",
        "Tables",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.list_content_controls",
        "word",
        "Content controls",
        ToolSideEffect::Read,
    ),
    tool(
        "word.insert_content_control",
        "word",
        "Content controls",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_content_control",
        "word",
        "Content controls",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.delete_content_control",
        "word",
        "Content controls",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.apply_style",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.add_comment",
        "word",
        "Review",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.resolve_comment",
        "word",
        "Review",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.set_change_tracking",
        "word",
        "Review",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_tracked_change",
        "word",
        "Review",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.save",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.get_workbook_info",
        "excel",
        "Workbook",
        ToolSideEffect::Read,
    ),
    tool(
        "excel.list_sheets",
        "excel",
        "Worksheet",
        ToolSideEffect::Read,
    ),
    tool(
        "excel.add_sheet",
        "excel",
        "Worksheet",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.update_sheet",
        "excel",
        "Worksheet",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.delete_sheet",
        "excel",
        "Worksheet",
        ToolSideEffect::Destructive,
    ),
    tool(
        "excel.get_used_range",
        "excel",
        "Range",
        ToolSideEffect::Read,
    ),
    tool("excel.read_range", "excel", "Range", ToolSideEffect::Read),
    tool(
        "excel.write_range",
        "excel",
        "Range",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.clear_range",
        "excel",
        "Range",
        ToolSideEffect::Destructive,
    ),
    tool(
        "excel.find_replace_cells",
        "excel",
        "Range",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.set_formula",
        "excel",
        "Formula",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.format_range",
        "excel",
        "Format",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.sort_range",
        "excel",
        "Data",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.apply_filter",
        "excel",
        "Data",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.create_table",
        "excel",
        "Table",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.update_table",
        "excel",
        "Table",
        ToolSideEffect::Destructive,
    ),
    tool(
        "excel.create_chart",
        "excel",
        "Chart",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.update_chart",
        "excel",
        "Chart",
        ToolSideEffect::Destructive,
    ),
    tool(
        "excel.create_pivot_table",
        "excel",
        "PivotTable",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.update_pivot_table",
        "excel",
        "PivotTable",
        ToolSideEffect::Destructive,
    ),
    tool(
        "powerpoint.get_presentation_info",
        "powerpoint",
        "Presentation",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.get_active_view",
        "powerpoint",
        "Presentation",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.export_file",
        "powerpoint",
        "Presentation",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.update_tags",
        "powerpoint",
        "Metadata",
        ToolSideEffect::Destructive,
    ),
    tool(
        "powerpoint.list_slides",
        "powerpoint",
        "Slides",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.add_slide",
        "powerpoint",
        "Slides",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.update_slide",
        "powerpoint",
        "Slides",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.delete_slide",
        "powerpoint",
        "Slides",
        ToolSideEffect::Destructive,
    ),
    tool(
        "powerpoint.move_slide",
        "powerpoint",
        "Slides",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.export_slide",
        "powerpoint",
        "Slides",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.list_layouts",
        "powerpoint",
        "Layout",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.apply_layout",
        "powerpoint",
        "Slides",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.get_selection",
        "powerpoint",
        "Selection",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.set_selection",
        "powerpoint",
        "Selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.list_shapes",
        "powerpoint",
        "Shapes",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.add_text_box",
        "powerpoint",
        "Shapes",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.add_shape",
        "powerpoint",
        "Shapes",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.insert_image",
        "powerpoint",
        "Shapes",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.update_shape",
        "powerpoint",
        "Shapes",
        ToolSideEffect::Destructive,
    ),
    tool(
        "powerpoint.read_text",
        "powerpoint",
        "Text",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.replace_text",
        "powerpoint",
        "Text",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.format_text",
        "powerpoint",
        "Text",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.add_table",
        "powerpoint",
        "Tables",
        ToolSideEffect::Mutating,
    ),
    tool(
        "powerpoint.read_table",
        "powerpoint",
        "Tables",
        ToolSideEffect::Read,
    ),
    tool(
        "powerpoint.update_table",
        "powerpoint",
        "Tables",
        ToolSideEffect::Destructive,
    ),
];

const fn tool(
    name: &'static str,
    app: &'static str,
    category: &'static str,
    side_effect: ToolSideEffect,
) -> ToolMetadata {
    ToolMetadata {
        name,
        app,
        category,
        side_effect,
    }
}

#[cfg(test)]
#[path = "tool_metadata_tests.rs"]
mod tool_metadata_tests;
