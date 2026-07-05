use super::{ToolSideEffect, tool_metadata, tool_metadata_catalog};
use crate::mcp::all_office_tool_names;
use std::collections::BTreeSet;

#[test]
#[allow(clippy::too_many_lines)]
fn tool_metadata_classifies_app_category_and_side_effect() {
    let read = tool_metadata("word.get_text").expect("word metadata");
    assert_eq!(read.app, "word");
    assert_eq!(read.category, "Document & structure");
    assert_eq!(read.side_effect, ToolSideEffect::Read);

    let anchor = tool_metadata("word.resolve_anchor").expect("anchor metadata");
    assert_eq!(anchor.app, "word");
    assert_eq!(anchor.category, "Range & selection");
    assert_eq!(anchor.side_effect, ToolSideEffect::Read);

    let set_selection = tool_metadata("word.set_selection").expect("selection metadata");
    assert_eq!(set_selection.app, "word");
    assert_eq!(set_selection.category, "Range & selection");
    assert_eq!(set_selection.side_effect, ToolSideEffect::Mutating);

    let list_bookmarks = tool_metadata("word.list_bookmarks").expect("bookmark list metadata");
    assert_eq!(list_bookmarks.app, "word");
    assert_eq!(list_bookmarks.category, "Range & selection");
    assert_eq!(list_bookmarks.side_effect, ToolSideEffect::Read);

    let insert_bookmark = tool_metadata("word.insert_bookmark").expect("bookmark insert metadata");
    assert_eq!(insert_bookmark.app, "word");
    assert_eq!(insert_bookmark.category, "Range & selection");
    assert_eq!(insert_bookmark.side_effect, ToolSideEffect::Mutating);

    let delete_bookmark = tool_metadata("word.delete_bookmark").expect("bookmark delete metadata");
    assert_eq!(delete_bookmark.app, "word");
    assert_eq!(delete_bookmark.category, "Range & selection");
    assert_eq!(delete_bookmark.side_effect, ToolSideEffect::Destructive);

    let list_notes = tool_metadata("word.list_notes").expect("note list metadata");
    assert_eq!(list_notes.app, "word");
    assert_eq!(list_notes.category, "Notes");
    assert_eq!(list_notes.side_effect, ToolSideEffect::Read);

    let insert_note = tool_metadata("word.insert_note").expect("note insert metadata");
    assert_eq!(insert_note.app, "word");
    assert_eq!(insert_note.category, "Notes");
    assert_eq!(insert_note.side_effect, ToolSideEffect::Mutating);

    let update_note = tool_metadata("word.update_note").expect("note update metadata");
    assert_eq!(update_note.app, "word");
    assert_eq!(update_note.category, "Notes");
    assert_eq!(update_note.side_effect, ToolSideEffect::Mutating);

    let delete_note = tool_metadata("word.delete_note").expect("note delete metadata");
    assert_eq!(delete_note.app, "word");
    assert_eq!(delete_note.category, "Notes");
    assert_eq!(delete_note.side_effect, ToolSideEffect::Destructive);

    let list_fields = tool_metadata("word.list_fields").expect("field list metadata");
    assert_eq!(list_fields.app, "word");
    assert_eq!(list_fields.category, "Document & structure");
    assert_eq!(list_fields.side_effect, ToolSideEffect::Read);

    let insert_field = tool_metadata("word.insert_field").expect("field insert metadata");
    assert_eq!(insert_field.app, "word");
    assert_eq!(insert_field.category, "Document & structure");
    assert_eq!(insert_field.side_effect, ToolSideEffect::Mutating);

    let update_field = tool_metadata("word.update_field").expect("field update metadata");
    assert_eq!(update_field.app, "word");
    assert_eq!(update_field.category, "Document & structure");
    assert_eq!(update_field.side_effect, ToolSideEffect::Mutating);

    let delete_field = tool_metadata("word.delete_field").expect("field delete metadata");
    assert_eq!(delete_field.app, "word");
    assert_eq!(delete_field.category, "Document & structure");
    assert_eq!(delete_field.side_effect, ToolSideEffect::Destructive);

    let list_styles = tool_metadata("word.list_styles").expect("style list metadata");
    assert_eq!(list_styles.app, "word");
    assert_eq!(list_styles.category, "Document & structure");
    assert_eq!(list_styles.side_effect, ToolSideEffect::Read);

    let create_style = tool_metadata("word.create_style").expect("style create metadata");
    assert_eq!(create_style.app, "word");
    assert_eq!(create_style.category, "Document & structure");
    assert_eq!(create_style.side_effect, ToolSideEffect::Mutating);

    let update_style = tool_metadata("word.update_style").expect("style update metadata");
    assert_eq!(update_style.app, "word");
    assert_eq!(update_style.category, "Document & structure");
    assert_eq!(update_style.side_effect, ToolSideEffect::Mutating);

    let sections = tool_metadata("word.list_sections").expect("sections metadata");
    assert_eq!(sections.app, "word");
    assert_eq!(sections.category, "Document & structure");
    assert_eq!(sections.side_effect, ToolSideEffect::Read);

    let page_setup = tool_metadata("word.update_page_setup").expect("page setup metadata");
    assert_eq!(page_setup.app, "word");
    assert_eq!(page_setup.category, "Document & structure");
    assert_eq!(page_setup.side_effect, ToolSideEffect::Mutating);

    let list_lists = tool_metadata("word.list_lists").expect("list lists metadata");
    assert_eq!(list_lists.app, "word");
    assert_eq!(list_lists.category, "Paragraphs & lists");
    assert_eq!(list_lists.side_effect, ToolSideEffect::Read);

    let update_list = tool_metadata("word.update_list").expect("list update metadata");
    assert_eq!(update_list.app, "word");
    assert_eq!(update_list.category, "Paragraphs & lists");
    assert_eq!(update_list.side_effect, ToolSideEffect::Destructive);

    let set_change_tracking =
        tool_metadata("word.set_change_tracking").expect("change tracking metadata");
    assert_eq!(set_change_tracking.app, "word");
    assert_eq!(set_change_tracking.category, "Review");
    assert_eq!(set_change_tracking.side_effect, ToolSideEffect::Mutating);

    let tracked_change =
        tool_metadata("word.update_tracked_change").expect("tracked change metadata");
    assert_eq!(tracked_change.app, "word");
    assert_eq!(tracked_change.category, "Review");
    assert_eq!(tracked_change.side_effect, ToolSideEffect::Destructive);

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
fn word_html_tool_metadata_classifies_range_side_effects() {
    let get_html = tool_metadata("word.get_html").expect("HTML read metadata");
    assert_eq!(get_html.app, "word");
    assert_eq!(get_html.category, "Range & selection");
    assert_eq!(get_html.side_effect, ToolSideEffect::Read);

    let insert_html = tool_metadata("word.insert_html").expect("HTML insert metadata");
    assert_eq!(insert_html.app, "word");
    assert_eq!(insert_html.category, "Range & selection");
    assert_eq!(insert_html.side_effect, ToolSideEffect::Mutating);
}

#[test]
fn word_document_property_tool_metadata_classifies_document_side_effects() {
    let get_document_properties =
        tool_metadata("word.get_document_properties").expect("document properties metadata");
    assert_eq!(get_document_properties.app, "word");
    assert_eq!(get_document_properties.category, "Document & structure");
    assert_eq!(get_document_properties.side_effect, ToolSideEffect::Read);

    let update_document_properties = tool_metadata("word.update_document_properties")
        .expect("document properties update metadata");
    assert_eq!(update_document_properties.app, "word");
    assert_eq!(update_document_properties.category, "Document & structure");
    assert_eq!(
        update_document_properties.side_effect,
        ToolSideEffect::Mutating
    );
}

#[test]
fn word_image_tool_metadata_classifies_media_side_effects() {
    let list_images = tool_metadata("word.list_images").expect("image list metadata");
    assert_eq!(list_images.app, "word");
    assert_eq!(list_images.category, "Media");
    assert_eq!(list_images.side_effect, ToolSideEffect::Read);

    let get_image = tool_metadata("word.get_image").expect("image get metadata");
    assert_eq!(get_image.app, "word");
    assert_eq!(get_image.category, "Media");
    assert_eq!(get_image.side_effect, ToolSideEffect::Read);

    let update_image = tool_metadata("word.update_image").expect("image update metadata");
    assert_eq!(update_image.app, "word");
    assert_eq!(update_image.category, "Media");
    assert_eq!(update_image.side_effect, ToolSideEffect::Mutating);

    let delete_image = tool_metadata("word.delete_image").expect("image delete metadata");
    assert_eq!(delete_image.app, "word");
    assert_eq!(delete_image.category, "Media");
    assert_eq!(delete_image.side_effect, ToolSideEffect::Destructive);

    let list_shapes = tool_metadata("word.list_shapes").expect("shape list metadata");
    assert_eq!(list_shapes.app, "word");
    assert_eq!(list_shapes.category, "Media");
    assert_eq!(list_shapes.side_effect, ToolSideEffect::Read);

    let insert_shape = tool_metadata("word.insert_shape").expect("shape insert metadata");
    assert_eq!(insert_shape.app, "word");
    assert_eq!(insert_shape.category, "Media");
    assert_eq!(insert_shape.side_effect, ToolSideEffect::Mutating);

    let update_shape = tool_metadata("word.update_shape").expect("shape update metadata");
    assert_eq!(update_shape.app, "word");
    assert_eq!(update_shape.category, "Media");
    assert_eq!(update_shape.side_effect, ToolSideEffect::Mutating);

    let delete_shape = tool_metadata("word.delete_shape").expect("shape delete metadata");
    assert_eq!(delete_shape.app, "word");
    assert_eq!(delete_shape.category, "Media");
    assert_eq!(delete_shape.side_effect, ToolSideEffect::Destructive);
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
