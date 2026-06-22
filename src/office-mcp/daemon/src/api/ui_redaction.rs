pub(crate) fn redact_text(value: &str) -> String {
    redact_text_with_limit(value, 500)
}

pub(crate) fn redact_text_with_limit(value: &str, max_chars: usize) -> String {
    let mut result = redact_bearer_tokens(value);
    result = redact_key_value_secret(&result);
    result = redact_base64_data(&result);
    result.truncate(max_chars);
    result
}

fn redact_bearer_tokens(value: &str) -> String {
    let parts = split_text_tokens(value);
    let mut output = String::with_capacity(value.len());
    let mut redact_next = false;
    for (part, separator) in parts {
        if redact_next {
            output.push_str("[redacted]");
            output.push_str(separator);
            redact_next = false;
            continue;
        }
        if part.eq_ignore_ascii_case("bearer") {
            output.push_str("Bearer");
            redact_next = true;
        } else {
            output.push_str(part);
        }
        output.push_str(separator);
    }
    output
}

fn redact_key_value_secret(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for (part, separator) in split_text_tokens(value) {
        let redacted = || {
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
        };
        output.push_str(&redacted());
        output.push_str(separator);
    }
    output
}

fn split_text_tokens(value: &str) -> Vec<(&str, &str)> {
    let mut tokens = Vec::new();
    let mut rest = value;
    while !rest.is_empty() {
        let token_end = rest.find(char::is_whitespace).unwrap_or(rest.len());
        let token = &rest[..token_end];
        let separator_end = rest[token_end..]
            .find(|character: char| !character.is_whitespace())
            .map_or(rest.len(), |index| token_end + index);
        tokens.push((token, &rest[token_end..separator_end]));
        rest = &rest[separator_end..];
    }
    tokens
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
