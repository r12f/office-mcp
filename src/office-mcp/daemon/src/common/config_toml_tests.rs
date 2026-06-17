use super::parse_toml;

#[test]
fn parses_supported_toml_subset() {
    let parsed = parse_toml(
        r#"
[addin_channel]
bind = "localhost"
port = 8765 # inline comments are ignored
certificate_passphrase = "not#comment"
enabled = true

[mcp_http]
bind = "127.0.0.1"
port = 8800
"#,
    )
    .expect("valid TOML subset");
    let addin = parsed.section("addin_channel");

    assert_eq!(
        addin.string_value("bind", "fallback").expect("bind"),
        "localhost"
    );
    assert_eq!(addin.int_value("port", 1).expect("port"), 8765);
    assert_eq!(
        addin
            .string_value("certificate_passphrase", "fallback")
            .expect("passphrase"),
        "not#comment"
    );
    assert!(addin.bool_value("enabled", false).expect("enabled"));
}

#[test]
fn rejects_unsupported_toml_syntax() {
    let error = parse_toml("bind = \"localhost\"").expect_err("top-level key rejected");

    assert!(error.to_string().contains("Unsupported TOML syntax"));
}

#[test]
fn typed_section_accessors_reject_wrong_value_kinds() {
    let parsed = parse_toml(
        r#"
[logging]
level = "info"
enabled = true
"#,
    )
    .expect("valid TOML subset");
    let logging = parsed.section("logging");

    assert!(logging.int_value("level", 1).is_err());
    assert!(logging.string_value("enabled", "false").is_err());
}
