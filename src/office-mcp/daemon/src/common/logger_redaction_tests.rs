use super::{redact_log_field, redact_log_text};

#[test]
fn redacts_log_text_boundaries() {
    assert_eq!(
        redact_log_text(
            "Authorization: bearer abc token=secret image=data:image/png;base64,QUJD next=ok"
        ),
        "Authorization: Bearer [redacted] token=[redacted] image=data:image/png;base64,[redacted] next=ok"
    );
}

#[test]
fn redacts_sensitive_field_by_name() {
    assert_eq!(
        redact_log_field("certificate_passphrase", "open"),
        "[redacted]"
    );
    assert_eq!(
        redact_log_field("message", "password=open"),
        "password=[redacted]"
    );
}

#[test]
fn truncates_redacted_log_text() {
    let long = "a".repeat(600);

    assert_eq!(redact_log_text(&long).len(), 500);
}
