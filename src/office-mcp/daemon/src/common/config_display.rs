use crate::common::{DaemonConfig, DaemonConfigService, EndpointConfig};

#[must_use]
pub fn render_endpoints(endpoints: &EndpointConfig) -> String {
    format!(
        "{{\n  \"mcp\": \"{}\",\n  \"addin_origin\": \"{}\",\n  \"addin_wss\": \"{}\"\n}}",
        json_escape(&endpoints.mcp),
        json_escape(&endpoints.addin_origin),
        json_escape(&endpoints.addin_wss)
    )
}

#[must_use]
pub fn render_redacted_config(config: &DaemonConfig) -> String {
    let redacted = DaemonConfigService::redacted(config);
    let endpoints = config.endpoints();
    format!(
        concat!(
            "{{\n",
            "  \"addin\": {{\n",
            "    \"host\": \"{}\",\n",
            "    \"port\": {},\n",
            "    \"origin\": \"{}\",\n",
            "    \"pfxPath\": \"{}\",\n",
            "    \"pfxPassphrase\": \"{}\"\n",
            "  }},\n",
            "  \"mcp\": {{\n",
            "    \"host\": \"{}\",\n",
            "    \"port\": {}\n",
            "  }},\n",
            "  \"endpoints\": {{\n",
            "    \"mcp\": \"{}\",\n",
            "    \"addin_origin\": \"{}\",\n",
            "    \"addin_wss\": \"{}\"\n",
            "  }}\n",
            "}}"
        ),
        json_escape(&redacted.addin.host),
        redacted.addin.port,
        json_escape(&redacted.addin.origin),
        json_escape(&redacted.addin.pfx_path),
        json_escape(&redacted.addin.pfx_passphrase),
        json_escape(&redacted.mcp.host),
        redacted.mcp.port,
        json_escape(&endpoints.mcp),
        json_escape(&endpoints.addin_origin),
        json_escape(&endpoints.addin_wss)
    )
}

#[must_use]
pub fn json_escape(value: &str) -> String {
    value
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}

#[cfg(test)]
#[path = "config_display_tests.rs"]
mod config_display_tests;
