use super::{ToolSideEffect, tool_metadata};

#[test]
fn tool_metadata_classifies_app_category_and_side_effect() {
    let read = tool_metadata("word.get_text").expect("word metadata");
    assert_eq!(read.app, "word");
    assert_eq!(read.category, "Document & structure");
    assert_eq!(read.side_effect, ToolSideEffect::Read);

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
