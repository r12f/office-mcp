use super::{prompt_catalog_json, prompt_description, prompt_messages};
use serde_json::json;

#[test]
fn prompt_catalog_lists_supported_prompts() {
    let prompts = prompt_catalog_json();
    let names = prompts
        .iter()
        .filter_map(|prompt| prompt["name"].as_str())
        .collect::<Vec<_>>();

    assert_eq!(
        names,
        vec![
            "summarize_document",
            "polish_section",
            "extract_action_items"
        ]
    );
    assert_eq!(
        prompts[1]["arguments"]["required"],
        json!(["session_id", "heading"])
    );
}

#[test]
fn prompt_descriptions_match_catalog_names() {
    assert!(prompt_description("summarize_document").contains("summary comment"));
    assert!(prompt_description("polish_section").contains("user approval"));
    assert_eq!(prompt_description("unknown"), "");
}

#[test]
fn prompt_messages_are_session_aware() {
    let messages = prompt_messages(
        "polish_section",
        Some(&json!({ "session_id": "session-1", "heading": "Scope" })),
    )
    .expect("messages");
    let text = messages[0]["content"]["text"].as_str().expect("text");

    assert!(text.contains("session-1"));
    assert!(text.contains("Scope"));
    assert!(prompt_messages("unknown", None).is_none());
}
