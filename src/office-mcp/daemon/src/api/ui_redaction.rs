pub(crate) fn redact_text(value: &str) -> String {
    let mut result = redact_bearer_tokens(value);
    result = redact_key_value_secret(&result);
    result = redact_base64_data(&result);
    result.truncate(500);
    result
}

fn redact_bearer_tokens(value: &str) -> String {
    let parts = value.split_whitespace().collect::<Vec<_>>();
    let mut output = Vec::with_capacity(parts.len());
    let mut redact_next = false;
    for part in parts {
        if redact_next {
            output.push("[redacted]".to_string());
            redact_next = false;
            continue;
        }
        if part.eq_ignore_ascii_case("bearer") {
            output.push("Bearer".to_string());
            redact_next = true;
        } else {
            output.push(part.to_string());
        }
    }
    output.join(" ")
}

fn redact_key_value_secret(value: &str) -> String {
    value
        .split_whitespace()
        .map(|part| {
            let Some((key, _secret)) = part.split_once('=') else {
                return part.to_string();
            };
            let normalized = key.replace(['_', '-'], "").to_ascii_lowercase();
            if matches!(normalized.as_str(), "password" | "passphrase" | "token")
                || normalized.ends_with("passphrase")
            {
                format!("{key}=[redacted]")
            } else {
                part.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn redact_base64_data(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    let mut rest = value;
    while let Some(index) = rest.find("base64,") {
        output.push_str(&rest[..index]);
        output.push_str("base64,[redacted]");
        let after_marker = &rest[index + "base64,".len()..];
        let end = after_marker
            .find(|character: char| character.is_whitespace())
            .unwrap_or(after_marker.len());
        rest = &after_marker[end..];
    }
    output.push_str(rest);
    output
}

#[cfg(test)]
#[path = "ui_redaction_tests.rs"]
mod ui_redaction_tests;
