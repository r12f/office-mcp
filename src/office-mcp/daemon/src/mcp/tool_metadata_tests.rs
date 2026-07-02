use super::{ToolSideEffect, tool_metadata, tool_metadata_catalog};
use crate::mcp::all_office_tool_names;
use std::collections::BTreeSet;

#[test]
fn tool_metadata_classifies_app_category_and_side_effect() {
    let read = tool_metadata("word.get_text").expect("word metadata");
    assert_eq!(read.app, "word");
    assert_eq!(read.category, "Document & structure");
    assert_eq!(read.side_effect, ToolSideEffect::Read);

    let anchor = tool_metadata("word.resolve_anchor").expect("anchor metadata");
    assert_eq!(anchor.app, "word");
    assert_eq!(anchor.category, "Range & selection");
    assert_eq!(anchor.side_effect, ToolSideEffect::Read);

    let mutating = tool_metadata("excel.write_range").expect("excel metadata");
    assert_eq!(mutating.app, "excel");
    assert_eq!(mutating.category, "Range");
    assert_eq!(mutating.side_effect, ToolSideEffect::Mutating);

    let destructive = tool_metadata("powerpoint.delete_slide").expect("powerpoint metadata");
    assert_eq!(destructive.app, "powerpoint");
    assert_eq!(destructive.category, "Slides");
    assert_eq!(destructive.side_effect, ToolSideEffect::Destructive);
}

#[test]
fn unknown_tool_has_no_metadata() {
    assert_eq!(tool_metadata("office.list_sessions"), None);
    assert_eq!(tool_metadata("word.future_tool"), None);
}

#[test]
fn tool_metadata_catalog_is_daemon_ui_source_of_truth() {
    let catalog = tool_metadata_catalog();

    assert_eq!(catalog.first().expect("first tool").name, "word.get_text");
    assert!(catalog.iter().any(|tool| tool.name == "excel.update_table"));
    assert!(
        catalog
            .iter()
            .any(|tool| tool.name == "powerpoint.list_slides")
    );
    assert!(catalog.iter().all(|tool| !tool.name.starts_with("office.")));
}

#[test]
fn tool_metadata_covers_every_forwarded_office_tool() {
    let forwarded_tools = all_office_tool_names().collect::<BTreeSet<_>>();
    let metadata_tools = tool_metadata_catalog()
        .iter()
        .map(|tool| tool.name)
        .collect::<BTreeSet<_>>();

    assert_eq!(metadata_tools, forwarded_tools);
}
