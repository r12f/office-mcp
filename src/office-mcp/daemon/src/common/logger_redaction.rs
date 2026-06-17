#[must_use]
pub(crate) fn redact_log_text(value: &str) -> String {
    let mut output = redact_bearer_tokens(value);
    output = redact_key_values(&output);
    output = redact_base64(&output);
    output.truncate(500);
    output
}

#[must_use]
pub(crate) fn redact_log_field(name: &str, value: &str) -> String {
    if is_sensitive_key(name) {
        "[redacted]".to_string()
    } else {
        redact_log_text(value)
    }
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

fn redact_key_values(value: &str) -> String {
    value
        .split_whitespace()
        .map(|part| {
            let Some((key, _value)) = part.split_once('=') else {
                return part.to_string();
            };
            if is_sensitive_key(key) {
                format!("{key}=[redacted]")
            } else {
                part.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn is_sensitive_key(key: &str) -> bool {
    let normalized = key.replace(['_', '-'], "").to_ascii_lowercase();
    matches!(normalized.as_str(), "password" | "passphrase" | "token")
        || normalized.ends_with("passphrase")
}

fn redact_base64(value: &str) -> String {
    let mut output = String::new();
    let mut rest = value;
    while let Some(index) = rest.find("base64,") {
        output.push_str(&rest[..index]);
        output.push_str("base64,[redacted]");
        let after = &rest[index + "base64,".len()..];
        let end = after
            .find(|character: char| character.is_whitespace())
            .unwrap_or(after.len());
        rest = &after[end..];
    }
    output.push_str(rest);
    output
}

#[cfg(test)]
#[path = "logger_redaction_tests.rs"]
mod logger_redaction_tests;
