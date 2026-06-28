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
    if let Some(rest) = uri.strip_prefix("office://word/") {
        return word_resource_request_from_uri(registry, uri, rest);
    }
    if let Some(rest) = uri.strip_prefix("office://excel/") {
        return excel_resource_request_from_uri(registry, uri, rest);
    }
    if let Some(rest) = uri.strip_prefix("office://powerpoint/") {
        return powerpoint_resource_request_from_uri(registry, uri, rest);
    }
    Err(format!("Unsupported resource URI {uri}."))
}

fn word_resource_request_from_uri(
    registry: &SessionRegistry,
    uri: &str,
    rest: &str,
) -> Result<ResourceReadRequest, String> {
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

fn excel_resource_request_from_uri(
    registry: &SessionRegistry,
    uri: &str,
    rest: &str,
) -> Result<ResourceReadRequest, String> {
    let (path, query) = rest
        .split_once('?')
        .map_or((rest, ""), |(path, query)| (path, query));
    let segments = path.split('/').collect::<Vec<_>>();
    if segments.len() < 2 {
        return Err(format!("Malformed Excel resource URI {uri}."));
    }
    let session_id = segments[0];
    if registry.get_session_info(session_id).is_none() {
        return Err(format!("Session {session_id} is not registered."));
    }
    match segments.as_slice() {
        [_, "workbook"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "excel.get_workbook_info",
            arguments: json!({ "session_id": session_id }),
            check_capability: true,
        }),
        [_, "sheets"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "excel.list_sheets",
            arguments: json!({ "session_id": session_id }),
            check_capability: true,
        }),
        [_, "used-range"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "excel.get_used_range",
            arguments: excel_arguments_with_optional_sheet(session_id, query),
            check_capability: true,
        }),
        [_, "range", address] => {
            if address.is_empty() {
                return Err("Excel range resource requires a non-empty address.".to_string());
            }
            let mut arguments = excel_arguments_with_optional_sheet(session_id, query);
            arguments["address"] = json!(*address);
            Ok(ResourceReadRequest::Forwarded {
                uri: uri.to_string(),
                tool: "excel.read_range",
                arguments,
                check_capability: true,
            })
        }
        _ => Err(format!("Unsupported Excel resource URI {uri}.")),
    }
}

fn powerpoint_resource_request_from_uri(
    registry: &SessionRegistry,
    uri: &str,
    rest: &str,
) -> Result<ResourceReadRequest, String> {
    let (path, query) = rest
        .split_once('?')
        .map_or((rest, ""), |(path, query)| (path, query));
    let segments = path.split('/').collect::<Vec<_>>();
    if segments.len() < 2 {
        return Err(format!("Malformed PowerPoint resource URI {uri}."));
    }
    let session_id = segments[0];
    if registry.get_session_info(session_id).is_none() {
        return Err(format!("Session {session_id} is not registered."));
    }
    match segments.as_slice() {
        [_, "presentation"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "powerpoint.get_presentation_info",
            arguments: json!({ "session_id": session_id }),
            check_capability: true,
        }),
        [_, "slides"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "powerpoint.list_slides",
            arguments: json!({ "session_id": session_id }),
            check_capability: true,
        }),
        [_, "slides", "text"] => {
            let mut arguments = json!({
                "session_id": session_id,
                "start": query_param_usize(query, "start", 0)?,
            });
            if let Some(end) = query_param_optional_usize(query, "end")? {
                arguments["end"] = json!(end);
            }
            Ok(ResourceReadRequest::Forwarded {
                uri: uri.to_string(),
                tool: "powerpoint.read_text",
                arguments,
                check_capability: true,
            })
        }
        [_, "slide", index, "text"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "powerpoint.read_text",
            arguments: json!({
                "session_id": session_id,
                "slide_index": parse_index(index, "slide index")?,
                "offset": query_param_usize(query, "offset", 0)?,
                "limit": query_param_usize(query, "limit", 200)?,
            }),
            check_capability: true,
        }),
        [_, "slide", index, "shapes"] => Ok(ResourceReadRequest::Forwarded {
            uri: uri.to_string(),
            tool: "powerpoint.list_shapes",
            arguments: json!({
                "session_id": session_id,
                "slide_index": parse_index(index, "slide index")?,
            }),
            check_capability: true,
        }),
        _ => Err(format!("Unsupported PowerPoint resource URI {uri}.")),
    }
}

fn excel_arguments_with_optional_sheet(session_id: &str, query: &str) -> Value {
    let mut arguments = json!({ "session_id": session_id });
    if let Some(sheet) = query_param_string(query, "sheet") {
        arguments["sheet"] = json!(sheet);
    }
    arguments
}

fn parse_index(value: &str, name: &str) -> Result<usize, String> {
    value
        .parse::<usize>()
        .map_err(|_| format!("{name} must be a non-negative integer."))
}

fn query_param_usize(query: &str, name: &str, default: usize) -> Result<usize, String> {
    Ok(query_param_optional_usize(query, name)?.unwrap_or(default))
}

fn query_param_optional_usize(query: &str, name: &str) -> Result<Option<usize>, String> {
    for part in query.split('&').filter(|part| !part.is_empty()) {
        let (key, value) = part.split_once('=').unwrap_or((part, ""));
        if key == name {
            return value
                .parse::<usize>()
                .map(Some)
                .map_err(|_| format!("{name} must be a non-negative integer."));
        }
    }
    Ok(None)
}

fn query_param_string(query: &str, name: &str) -> Option<String> {
    query
        .split('&')
        .filter(|part| !part.is_empty())
        .find_map(|part| {
            let (key, value) = part.split_once('=').unwrap_or((part, ""));
            (key == name).then(|| value.to_string())
        })
}

#[cfg(test)]
#[path = "resource_request_tests.rs"]
mod resource_request_tests;
