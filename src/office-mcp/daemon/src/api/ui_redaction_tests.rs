use super::redact_text;

#[test]
fn redacts_bearer_token_values() {
    assert_eq!(
        redact_text("Authorization: bearer abc123"),
        "Authorization: Bearer [redacted]"
    );
}

#[test]
fn redacts_key_value_secrets() {
    assert_eq!(
        redact_text("password=secret token=abc safe=value cert_passphrase=open"),
        "password=[redacted] token=[redacted] safe=value cert_passphrase=[redacted]"
    );
}

#[test]
fn preserves_diagnostic_line_breaks_while_redacting() {
    assert_eq!(
        redact_text("first line\npassword=secret\nAuthorization: bearer abc123"),
        "first line\npassword=[redacted]\nAuthorization: Bearer [redacted]"
    );
}

#[test]
fn redacts_base64_payloads() {
    assert_eq!(
        redact_text("image=data:image/png;base64,iVBORw0KGgoAAA next=ok"),
        "image=data:image/png;base64,[redacted] next=ok"
    );
}

#[test]
fn truncates_long_redacted_text() {
    let long = "x".repeat(600);

    assert_eq!(redact_text(&long).len(), 500);
}
