use super::ToolAccessPolicy;
use crate::mcp::AccessMode;

#[test]
fn default_policy_allows_tools_until_disabled() {
    let policy = ToolAccessPolicy::default();

    assert!(policy.allows_tool("word.get_text"));

    let policy = policy.with_disabled_tool("word.get_text");

    assert!(!policy.allows_tool("word.get_text"));
    assert!(policy.allows_tool("word.get_outline"));
}

#[test]
fn policy_applies_access_mode_app_and_category_before_tool_switches() {
    let read_only = ToolAccessPolicy::default().with_access_mode(AccessMode::Read);

    assert!(read_only.allows_tool("word.get_text"));
    assert!(!read_only.allows_tool("word.insert_paragraph"));
    assert!(!read_only.allows_tool("word.update_table"));

    let word_disabled = ToolAccessPolicy::default().with_disabled_app("word");

    assert!(!word_disabled.allows_tool("word.get_text"));
    assert!(word_disabled.allows_tool("excel.read_range"));

    let range_disabled = ToolAccessPolicy::default().with_disabled_category("excel", "Range");

    assert!(!range_disabled.allows_tool("excel.read_range"));
    assert!(!range_disabled.allows_tool("excel.write_range"));
    assert!(range_disabled.allows_tool("excel.list_sheets"));

    let explicitly_disabled =
        ToolAccessPolicy::default().with_disabled_tool("powerpoint.delete_slide");

    assert!(!explicitly_disabled.allows_tool("powerpoint.delete_slide"));
    assert!(explicitly_disabled.allows_tool("powerpoint.list_slides"));
}
