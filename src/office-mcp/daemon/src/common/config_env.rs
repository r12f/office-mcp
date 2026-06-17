use crate::common::ConfigError;
use std::collections::BTreeMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ConfigEnv<'a> {
    values: &'a BTreeMap<String, String>,
}

impl<'a> ConfigEnv<'a> {
    #[must_use]
    pub(crate) const fn new(values: &'a BTreeMap<String, String>) -> Self {
        Self { values }
    }

    #[must_use]
    pub(crate) fn string_any(&self, names: &[&str], fallback: String) -> String {
        names
            .iter()
            .find_map(|name| self.values.get(*name).cloned())
            .unwrap_or(fallback)
    }

    pub(crate) fn positive_int_any(
        &self,
        names: &[&str],
        fallback: u64,
    ) -> Result<u64, ConfigError> {
        let Some((name, raw)) = names
            .iter()
            .find_map(|name| self.values.get(*name).map(|raw| (*name, raw)))
        else {
            return Ok(fallback);
        };
        raw.parse::<u64>()
            .ok()
            .filter(|value| *value > 0)
            .ok_or_else(|| ConfigError::Validation(format!("{name} must be a positive integer")))
    }

    pub(crate) fn bool_any(&self, names: &[&str], fallback: bool) -> Result<bool, ConfigError> {
        let Some((name, raw)) = names
            .iter()
            .find_map(|name| self.values.get(*name).map(|raw| (*name, raw)))
        else {
            return Ok(fallback);
        };
        match raw.as_str() {
            "true" => Ok(true),
            "false" => Ok(false),
            _ => Err(ConfigError::Validation(format!(
                "{name} must be true or false"
            ))),
        }
    }

    #[must_use]
    pub(crate) fn optional_path_value(value: String, fallback: &str) -> String {
        if value.is_empty() {
            fallback.to_string()
        } else {
            value
        }
    }
}

#[cfg(test)]
#[path = "config_env_tests.rs"]
mod config_env_tests;
