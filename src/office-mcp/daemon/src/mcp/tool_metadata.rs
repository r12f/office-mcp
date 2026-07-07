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
    pub action_side_effects: Option<&'static [ActionSideEffect]>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ActionSideEffect {
    pub action: &'static str,
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

impl ToolMetadata {
    #[must_use]
    pub fn side_effect_for_action(self, action: &str) -> Option<ToolSideEffect> {
        self.action_side_effects.and_then(|actions| {
            actions
                .iter()
                .find(|entry| entry.action == action)
                .map(|entry| entry.side_effect)
        })
    }

    #[must_use]
    pub fn has_action_allowed_by(self, mode: AccessMode) -> bool {
        self.action_side_effects
            .is_some_and(|actions| actions.iter().any(|entry| mode.allows(entry.side_effect)))
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
        "word.set_selection",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.get_html",
        "word",
        "Range & selection",
        ToolSideEffect::Read,
    ),
    tool(
        "word.insert_html",
        "word",
        "Range & selection",
        ToolSideEffect::Mutating,
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
    tool("word.list_images", "word", "Media", ToolSideEffect::Read),
    tool("word.get_image", "word", "Media", ToolSideEffect::Read),
    tool("word.list_shapes", "word", "Media", ToolSideEffect::Read),
    tool(
        "word.insert_shape",
        "word",
        "Media",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_shape",
        "word",
        "Media",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.delete_shape",
        "word",
        "Media",
        ToolSideEffect::Destructive,
    ),
    tool_with_actions(
        "word.update_image",
        "word",
        "Media",
        ToolSideEffect::Destructive,
        WORD_UPDATE_IMAGE_ACTIONS,
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
        "word.list_fields",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.insert_field",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_field",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.delete_field",
        "word",
        "Document & structure",
        ToolSideEffect::Destructive,
    ),
    tool(
        "word.list_styles",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.get_document_properties",
        "word",
        "Document & structure",
        ToolSideEffect::Read,
    ),
    tool(
        "word.update_document_properties",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.create_style",
        "word",
        "Document & structure",
        ToolSideEffect::Mutating,
    ),
    tool(
        "word.update_style",
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
        "word.list_lists",
        "word",
        "Paragraphs & lists",
        ToolSideEffect::Read,
    ),
    tool_with_actions(
        "word.update_list",
        "word",
        "Paragraphs & lists",
        ToolSideEffect::Destructive,
        WORD_UPDATE_LIST_ACTIONS,
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
    tool_with_actions(
        "word.update_header_footer",
        "word",
        "Document & structure",
        ToolSideEffect::Destructive,
        WORD_UPDATE_HEADER_FOOTER_ACTIONS,
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
    tool_with_actions(
        "word.update_table",
        "word",
        "Tables",
        ToolSideEffect::Destructive,
        WORD_UPDATE_TABLE_ACTIONS,
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
    tool_with_actions(
        "word.update_comment",
        "word",
        "Review",
        ToolSideEffect::Destructive,
        WORD_UPDATE_COMMENT_ACTIONS,
    ),
    tool(
        "word.set_change_tracking",
        "word",
        "Review",
        ToolSideEffect::Mutating,
    ),
    tool_with_actions(
        "word.update_tracked_change",
        "word",
        "Review",
        ToolSideEffect::Destructive,
        WORD_UPDATE_TRACKED_CHANGE_ACTIONS,
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
    tool("excel.save", "excel", "Workbook", ToolSideEffect::Mutating),
    tool(
        "excel.calculate",
        "excel",
        "Workbook",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.list_named_items",
        "excel",
        "Workbook",
        ToolSideEffect::Read,
    ),
    tool_with_actions(
        "excel.update_named_item",
        "excel",
        "Workbook",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_NAMED_ITEM_ACTIONS,
    ),
    tool(
        "excel.get_document_properties",
        "excel",
        "Workbook",
        ToolSideEffect::Read,
    ),
    tool(
        "excel.update_document_properties",
        "excel",
        "Workbook",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.add_comment",
        "excel",
        "Review",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.list_comments",
        "excel",
        "Review",
        ToolSideEffect::Read,
    ),
    tool_with_actions(
        "excel.update_comment",
        "excel",
        "Review",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_COMMENT_ACTIONS,
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
    tool("excel.read_range", "excel", "Range", ToolSideEffect::Read),
    tool(
        "excel.write_range",
        "excel",
        "Range",
        ToolSideEffect::Mutating,
    ),
    tool(
        "excel.insert_range",
        "excel",
        "Range",
        ToolSideEffect::Mutating,
    ),
    tool_with_actions(
        "excel.set_hyperlink",
        "excel",
        "Range",
        ToolSideEffect::Mutating,
        EXCEL_SET_HYPERLINK_ACTIONS,
    ),
    tool_with_actions(
        "excel.set_data_validation",
        "excel",
        "Range",
        ToolSideEffect::Destructive,
        EXCEL_SET_DATA_VALIDATION_ACTIONS,
    ),
    tool_with_actions(
        "excel.copy_range",
        "excel",
        "Range",
        ToolSideEffect::Mutating,
        EXCEL_COPY_RANGE_ACTIONS,
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
        "excel.list_conditional_formats",
        "excel",
        "Format",
        ToolSideEffect::Read,
    ),
    tool_with_actions(
        "excel.update_conditional_format",
        "excel",
        "Format",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_CONDITIONAL_FORMAT_ACTIONS,
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
    tool_with_actions(
        "excel.update_table",
        "excel",
        "Table",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_TABLE_ACTIONS,
    ),
    tool(
        "excel.create_chart",
        "excel",
        "Chart",
        ToolSideEffect::Mutating,
    ),
    tool_with_actions(
        "excel.update_chart",
        "excel",
        "Chart",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_CHART_ACTIONS,
    ),
    tool(
        "excel.create_pivot_table",
        "excel",
        "PivotTable",
        ToolSideEffect::Mutating,
    ),
    tool_with_actions(
        "excel.update_pivot_table",
        "excel",
        "PivotTable",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_PIVOT_TABLE_ACTIONS,
    ),
    tool(
        "excel.insert_image",
        "excel",
        "Shapes",
        ToolSideEffect::Mutating,
    ),
    tool("excel.list_shapes", "excel", "Shapes", ToolSideEffect::Read),
    tool_with_actions(
        "excel.update_shape",
        "excel",
        "Shapes",
        ToolSideEffect::Destructive,
        EXCEL_UPDATE_SHAPE_ACTIONS,
    ),
    tool(
        "powerpoint.get_presentation_info",
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
    tool_with_actions(
        "powerpoint.update_tags",
        "powerpoint",
        "Metadata",
        ToolSideEffect::Destructive,
        POWERPOINT_UPDATE_TAGS_ACTIONS,
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
    tool_with_actions(
        "powerpoint.update_shape",
        "powerpoint",
        "Shapes",
        ToolSideEffect::Destructive,
        POWERPOINT_UPDATE_SHAPE_ACTIONS,
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
    tool_with_actions(
        "powerpoint.update_table",
        "powerpoint",
        "Tables",
        ToolSideEffect::Destructive,
        POWERPOINT_UPDATE_TABLE_ACTIONS,
    ),
];

const WORD_UPDATE_IMAGE_ACTIONS: &[ActionSideEffect] = &[
    action("resize", ToolSideEffect::Mutating),
    action("set_alt_text", ToolSideEffect::Mutating),
    action("set_hyperlink", ToolSideEffect::Mutating),
    action("replace", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const WORD_UPDATE_LIST_ACTIONS: &[ActionSideEffect] = &[
    action("add_item", ToolSideEffect::Mutating),
    action("set_item_level", ToolSideEffect::Mutating),
    action("attach_paragraph", ToolSideEffect::Mutating),
    action("detach_paragraph", ToolSideEffect::Mutating),
    action("set_level_format", ToolSideEffect::Mutating),
];

const WORD_UPDATE_HEADER_FOOTER_ACTIONS: &[ActionSideEffect] = &[
    action("set_text", ToolSideEffect::Mutating),
    action("append_paragraph", ToolSideEffect::Mutating),
    action("clear", ToolSideEffect::Destructive),
];

const WORD_UPDATE_TABLE_ACTIONS: &[ActionSideEffect] = &[
    action("update_cell", ToolSideEffect::Mutating),
    action("add_row", ToolSideEffect::Mutating),
    action("add_column", ToolSideEffect::Mutating),
    action("format_cell", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
    action("delete_row", ToolSideEffect::Destructive),
    action("delete_column", ToolSideEffect::Destructive),
    action("merge_cells", ToolSideEffect::Mutating),
    action("set_column_width", ToolSideEffect::Mutating),
    action("distribute_columns", ToolSideEffect::Mutating),
    action("set_borders", ToolSideEffect::Mutating),
    action("set_header_row", ToolSideEffect::Mutating),
];

const WORD_UPDATE_COMMENT_ACTIONS: &[ActionSideEffect] = &[
    action("reply", ToolSideEffect::Mutating),
    action("edit", ToolSideEffect::Mutating),
    action("resolve", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
    action("reopen", ToolSideEffect::Mutating),
];

const WORD_UPDATE_TRACKED_CHANGE_ACTIONS: &[ActionSideEffect] = &[
    action("accept", ToolSideEffect::Mutating),
    action("reject", ToolSideEffect::Mutating),
    action("accept_all", ToolSideEffect::Destructive),
    action("reject_all", ToolSideEffect::Destructive),
];

const EXCEL_UPDATE_TABLE_ACTIONS: &[ActionSideEffect] = &[
    action("metadata", ToolSideEffect::Read),
    action("read", ToolSideEffect::Read),
    action("add_rows", ToolSideEffect::Mutating),
    action("add_columns", ToolSideEffect::Mutating),
    action("resize", ToolSideEffect::Mutating),
    action("rename", ToolSideEffect::Mutating),
    action("options", ToolSideEffect::Mutating),
    action("style", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const EXCEL_UPDATE_CHART_ACTIONS: &[ActionSideEffect] = &[
    action("metadata", ToolSideEffect::Read),
    action("read", ToolSideEffect::Read),
    action("title", ToolSideEffect::Mutating),
    action("legend", ToolSideEffect::Mutating),
    action("axis", ToolSideEffect::Mutating),
    action("data", ToolSideEffect::Mutating),
    action("series_source", ToolSideEffect::Mutating),
    action("position", ToolSideEffect::Mutating),
    action("size", ToolSideEffect::Mutating),
    action("export_image", ToolSideEffect::Read),
    action("delete", ToolSideEffect::Destructive),
];

const EXCEL_UPDATE_PIVOT_TABLE_ACTIONS: &[ActionSideEffect] = &[
    action("metadata", ToolSideEffect::Read),
    action("read", ToolSideEffect::Read),
    action("refresh", ToolSideEffect::Mutating),
    action("add_hierarchy", ToolSideEffect::Mutating),
    action("remove_hierarchy", ToolSideEffect::Mutating),
    action("layout", ToolSideEffect::Mutating),
    action("filter", ToolSideEffect::Mutating),
    action("clear_filters", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const EXCEL_UPDATE_SHAPE_ACTIONS: &[ActionSideEffect] = &[
    action("move", ToolSideEffect::Mutating),
    action("resize", ToolSideEffect::Mutating),
    action("set_alt_text", ToolSideEffect::Mutating),
    action("set_text", ToolSideEffect::Mutating),
    action("set_z_order", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const EXCEL_UPDATE_NAMED_ITEM_ACTIONS: &[ActionSideEffect] = &[
    action("add", ToolSideEffect::Mutating),
    action("edit", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const EXCEL_UPDATE_COMMENT_ACTIONS: &[ActionSideEffect] = &[
    action("reply", ToolSideEffect::Mutating),
    action("edit", ToolSideEffect::Mutating),
    action("resolve", ToolSideEffect::Mutating),
    action("reopen", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const EXCEL_SET_HYPERLINK_ACTIONS: &[ActionSideEffect] = &[
    action("set", ToolSideEffect::Mutating),
    action("clear", ToolSideEffect::Mutating),
];

const EXCEL_SET_DATA_VALIDATION_ACTIONS: &[ActionSideEffect] = &[
    action("set", ToolSideEffect::Mutating),
    action("clear", ToolSideEffect::Destructive),
];

const EXCEL_COPY_RANGE_ACTIONS: &[ActionSideEffect] = &[
    action("copy", ToolSideEffect::Mutating),
    action("autofill", ToolSideEffect::Mutating),
];

const EXCEL_UPDATE_CONDITIONAL_FORMAT_ACTIONS: &[ActionSideEffect] = &[
    action("add", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
    action("clear_range", ToolSideEffect::Destructive),
];

const POWERPOINT_UPDATE_TAGS_ACTIONS: &[ActionSideEffect] = &[
    action("list", ToolSideEffect::Read),
    action("set", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const POWERPOINT_UPDATE_SHAPE_ACTIONS: &[ActionSideEffect] = &[
    action("move", ToolSideEffect::Mutating),
    action("resize", ToolSideEffect::Mutating),
    action("rotate", ToolSideEffect::Mutating),
    action("rename", ToolSideEffect::Mutating),
    action("set_alt_text", ToolSideEffect::Mutating),
    action("set_fill", ToolSideEffect::Mutating),
    action("set_line", ToolSideEffect::Mutating),
    action("set_z_order", ToolSideEffect::Mutating),
    action("group", ToolSideEffect::Mutating),
    action("ungroup", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const POWERPOINT_UPDATE_TABLE_ACTIONS: &[ActionSideEffect] = &[
    action("set_values", ToolSideEffect::Mutating),
    action("set_cell", ToolSideEffect::Mutating),
    action("add_rows", ToolSideEffect::Mutating),
    action("delete_rows", ToolSideEffect::Destructive),
    action("add_columns", ToolSideEffect::Mutating),
    action("delete_columns", ToolSideEffect::Destructive),
    action("merge_cells", ToolSideEffect::Mutating),
    action("split_cell", ToolSideEffect::Mutating),
    action("clear", ToolSideEffect::Mutating),
    action("style", ToolSideEffect::Mutating),
    action("delete", ToolSideEffect::Destructive),
];

const fn action(action: &'static str, side_effect: ToolSideEffect) -> ActionSideEffect {
    ActionSideEffect {
        action,
        side_effect,
    }
}

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
        action_side_effects: None,
    }
}

const fn tool_with_actions(
    name: &'static str,
    app: &'static str,
    category: &'static str,
    side_effect: ToolSideEffect,
    action_side_effects: &'static [ActionSideEffect],
) -> ToolMetadata {
    ToolMetadata {
        name,
        app,
        category,
        side_effect,
        action_side_effects: Some(action_side_effects),
    }
}

#[cfg(test)]
#[path = "tool_metadata_tests.rs"]
mod tool_metadata_tests;
