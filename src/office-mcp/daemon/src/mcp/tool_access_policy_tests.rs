use super::ToolAccessPolicy;

#[test]
fn default_policy_allows_tools_until_disabled() {
    let policy = ToolAccessPolicy::default();

    assert!(policy.allows_tool("word.get_text"));

    let policy = policy.with_disabled_tool("word.get_text");

    assert!(!policy.allows_tool("word.get_text"));
    assert!(policy.allows_tool("word.get_outline"));
}
