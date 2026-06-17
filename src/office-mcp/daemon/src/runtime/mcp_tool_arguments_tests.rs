use super::McpToolArgumentPreprocessor;
use crate::addin_mgr::{ImageFetcher, PartialEffect};
use base64::Engine;
use serde_json::json;

const PNG_1X1_HEADER: &[u8] = &[0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 0x00];

#[test]
fn leaves_non_image_tools_unchanged() {
    let arguments = json!({ "session_id": "session-1", "text": "Hello" });

    let processed = McpToolArgumentPreprocessor::preprocess(
        &ImageFetcher::new(),
        "word.insert_paragraph",
        &arguments,
    )
    .expect("arguments");

    assert_eq!(processed, arguments);
}

#[test]
fn normalizes_insert_image_base64_arguments() {
    let base64 = base64::engine::general_purpose::STANDARD.encode(PNG_1X1_HEADER);
    let arguments = json!({
        "session_id": "session-1",
        "image": { "base64": base64 }
    });

    let processed = McpToolArgumentPreprocessor::preprocess(
        &ImageFetcher::new(),
        "word.insert_image",
        &arguments,
    )
    .expect("arguments");

    assert_eq!(processed["image"]["base64"], base64);
    assert_eq!(processed["image"]["mime_type"], "image/png");
    assert_eq!(processed["image"]["byte_length"], PNG_1X1_HEADER.len());
}

#[test]
fn maps_invalid_insert_image_input_to_command_failure() {
    let failure = McpToolArgumentPreprocessor::preprocess(
        &ImageFetcher::new(),
        "word.insert_image",
        &json!({
            "session_id": "session-1",
            "image": { "base64": "not-valid-base64" }
        }),
    )
    .expect_err("invalid image");

    assert_eq!(failure.office_mcp_code, "IMAGE_FETCH_FAILED");
    assert_eq!(failure.tool.as_deref(), Some("word.insert_image"));
    assert_eq!(failure.partial_effect, Some(PartialEffect::None));
    assert!(!failure.retriable);
}

#[test]
fn rejects_non_https_insert_image_url_without_network_fetch() {
    let failure = McpToolArgumentPreprocessor::preprocess(
        &ImageFetcher::new(),
        "word.insert_image",
        &json!({
            "session_id": "session-1",
            "image": { "url": "http://example.com/image.png" }
        }),
    )
    .expect_err("invalid url");

    assert_eq!(failure.office_mcp_code, "IMAGE_FETCH_FAILED");
    assert!(failure.message.contains("https"));
}
