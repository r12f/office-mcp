use crate::addin_mgr::{ImageFetcher, PartialEffect};
use crate::api::CommandFailure;
use serde_json::{Value, json};

pub(crate) struct McpToolArgumentPreprocessor;

impl McpToolArgumentPreprocessor {
    pub(crate) fn preprocess(
        image_fetcher: &ImageFetcher,
        tool: &str,
        arguments: &Value,
    ) -> Result<Value, CommandFailure> {
        if !is_insert_image_tool(tool) {
            return Ok(arguments.clone());
        }
        let Some(image) = arguments.get("image") else {
            return Ok(arguments.clone());
        };
        let processed = if let Some(base64) = image.get("base64").and_then(Value::as_str) {
            image_fetcher.validate_base64(base64)
        } else if let Some(url) = image.get("url").and_then(Value::as_str) {
            image_fetcher.fetch_url(url)
        } else {
            return Ok(arguments.clone());
        };
        let fetched = processed.map_err(|error| CommandFailure {
            office_mcp_code: "IMAGE_FETCH_FAILED".to_string(),
            message: error.to_string(),
            tool: Some(tool.to_string()),
            retriable: false,
            partial_effect: Some(PartialEffect::None),
        })?;
        let mut updated = arguments.clone();
        if let Some(object) = updated.as_object_mut() {
            object.insert(
                "image".to_string(),
                json!({
                    "base64": fetched.base64,
                    "mime_type": fetched.mime_type.as_str(),
                    "byte_length": fetched.byte_length
                }),
            );
        }
        Ok(updated)
    }
}

fn is_insert_image_tool(tool: &str) -> bool {
    matches!(tool, "word.insert_image" | "powerpoint.insert_image")
}

#[cfg(test)]
#[path = "mcp_tool_arguments_tests.rs"]
mod mcp_tool_arguments_tests;
