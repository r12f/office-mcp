use crate::addin_mgr::SessionRegistry;
use serde_json::{Value, json};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ResourceReadRequest {
    Sessions,
    Forwarded {
        uri: String,
        tool: &'static str,
        arguments: Value,
        check_capability: bool,
    },
}

/// Converts an MCP resource URI into the daemon-internal read request.
///
/// # Errors
///
/// Returns an error if the URI is unsupported, malformed, or references an
/// unknown document session.
pub fn resource_request_from_uri(
    registry: &SessionRegistry,
    uri: &str,
) -> Result<ResourceReadRequest, String> {
    if uri == "office://sessions" {
        return Ok(ResourceReadRequest::Sessions);
    }
    let Some(rest) = uri.strip_prefix("office://word/") else {
        return Err(format!("Unsupported resource URI {uri}."));
    };
    let (path, query) = rest
        .split_once('?')
        .map_or((rest, ""), |(path, query)| (path, query));
    let segments = path.split('/').collect::<Vec<_>>();
    if segments.len() < 2 {
        return Err(format!("Malformed Word resource URI {uri}."));
    }
    let session_id = segments[0];
    if registry.get_session_info(session_id).is_none() {
        return Err(format!("Session {session_id} is not registered."));
    }
    match segments.as_slice() {
        [_, "document"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "word.get_text",
            arguments: json!({
                "session_id": session_id,
                "offset": query_param_usize(query, "offset", 0)?,
                "limit": query_param_usize(query, "limit", 200)?,
            }),
            check_capability: true,
        }),
        [_, "structure"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "word._get_structure",
            arguments: json!({ "session_id": session_id }),
            check_capability: false,
        }),
        [_, "paragraph", index] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "word.get_paragraph",
            arguments: json!({
                "session_id": session_id,
                "index": index.parse::<usize>().map_err(|_| "paragraph index must be a non-negative integer.".to_string())?,
            }),
            check_capability: true,
        }),
        [_, "comments"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "word._get_comments",
            arguments: json!({ "session_id": session_id }),
            check_capability: false,
        }),
        [_, "track_changes"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "word._get_tracked_changes",
            arguments: json!({ "session_id": session_id }),
            check_capability: false,
        }),
        [_, "selection"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "word.get_selection",
            arguments: json!({ "session_id": session_id }),
            check_capability: true,
        }),
        _ => Err(format!("Unsupported Word resource URI {uri}.")),
    }
}

fn query_param_usize(query: &str, name: &str, default: usize) -> Result<usize, String> {
    for part in query.split('&').filter(|part| !part.is_empty()) {
        let (key, value) = part.split_once('=').unwrap_or((part, ""));
        if key == name {
            return value
                .parse::<usize>()
                .map_err(|_| format!("{name} must be a non-negative integer."));
        }
    }
    Ok(default)
}

#[cfg(test)]
#[path = "resource_request_tests.rs"]
mod resource_request_tests;
