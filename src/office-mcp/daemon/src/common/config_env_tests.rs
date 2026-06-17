use super::ConfigEnv;
use crate::common::ConfigError;
use std::collections::BTreeMap;

#[test]
fn string_override_uses_first_present_name() {
    let values = env_map(&[("LEGACY", "old"), ("SECTION", "new")]);
    let env = ConfigEnv::new(&values);

    assert_eq!(
        env.string_any(&["SECTION", "LEGACY"], "fallback".to_string()),
        "new"
    );
}

#[test]
fn positive_int_override_rejects_zero_and_non_numeric_values() {
    let zero_values = env_map(&[("PORT", "0")]);
    let zero_env = ConfigEnv::new(&zero_values);
    let zero_error = zero_env
        .positive_int_any(&["PORT"], 8800)
        .expect_err("zero rejected");
    assert!(matches!(
        zero_error,
        ConfigError::Validation(message) if message == "PORT must be a positive integer"
    ));

    let bad_values = env_map(&[("PORT", "abc")]);
    let bad_env = ConfigEnv::new(&bad_values);
    let bad_error = bad_env
        .positive_int_any(&["PORT"], 8800)
        .expect_err("text rejected");
    assert!(matches!(bad_error, ConfigError::Validation(_)));
}

#[test]
fn bool_override_accepts_only_lowercase_true_or_false() {
    let enabled_values = env_map(&[("ENABLED", "true")]);
    let enabled = ConfigEnv::new(&enabled_values);
    assert!(enabled.bool_any(&["ENABLED"], false).expect("true"));

    let invalid_values = env_map(&[("ENABLED", "yes")]);
    let invalid = ConfigEnv::new(&invalid_values);
    let error = invalid
        .bool_any(&["ENABLED"], false)
        .expect_err("invalid bool rejected");
    assert!(matches!(
        error,
        ConfigError::Validation(message) if message == "ENABLED must be true or false"
    ));
}

#[test]
fn optional_path_value_replaces_empty_values_with_default_path() {
    assert_eq!(
        ConfigEnv::optional_path_value(String::new(), "default.log"),
        "default.log"
    );
    assert_eq!(
        ConfigEnv::optional_path_value("explicit.log".to_string(), "default.log"),
        "explicit.log"
    );
}

fn env_map(values: &[(&str, &str)]) -> BTreeMap<String, String> {
    values
        .iter()
        .map(|(key, value)| ((*key).to_string(), (*value).to_string()))
        .collect()
}
