use crate::common::ConfigError;
use std::collections::BTreeMap;

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RawToml(BTreeMap<String, BTreeMap<String, RawTomlValue>>);

impl RawToml {
    #[must_use]
    pub(crate) fn section(&self, name: &str) -> RawTomlSection<'_> {
        RawTomlSection(self.0.get(name))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RawTomlValue {
    String(String),
    Integer(u64),
    Boolean(bool),
}

pub(crate) struct RawTomlSection<'a>(Option<&'a BTreeMap<String, RawTomlValue>>);

impl RawTomlSection<'_> {
    pub(crate) fn string_value(&self, key: &str, fallback: &str) -> Result<String, ConfigError> {
        let Some(value) = self.0.and_then(|section| section.get(key)) else {
            return Ok(fallback.to_string());
        };
        match value {
            RawTomlValue::String(value) => Ok(value.clone()),
            other => Err(ConfigError::Parse(format!(
                "Expected string config value for {key}, got {}.",
                other.kind()
            ))),
        }
    }

    pub(crate) fn int_value(&self, key: &str, fallback: u64) -> Result<u64, ConfigError> {
        let Some(value) = self.0.and_then(|section| section.get(key)) else {
            return Ok(fallback);
        };
        match value {
            RawTomlValue::Integer(value) if *value > 0 => Ok(*value),
            _ => Err(ConfigError::Parse(format!(
                "Expected positive integer config value for {key}."
            ))),
        }
    }

    pub(crate) fn bool_value(&self, key: &str, fallback: bool) -> Result<bool, ConfigError> {
        let Some(value) = self.0.and_then(|section| section.get(key)) else {
            return Ok(fallback);
        };
        match value {
            RawTomlValue::Boolean(value) => Ok(*value),
            other => Err(ConfigError::Parse(format!(
                "Expected boolean config value for {key}, got {}.",
                other.kind()
            ))),
        }
    }
}

impl RawTomlValue {
    const fn kind(&self) -> &'static str {
        match self {
            Self::String(_) => "string",
            Self::Integer(_) => "integer",
            Self::Boolean(_) => "boolean",
        }
    }
}

/// Parses the small TOML subset accepted by the daemon config file.
///
/// # Errors
///
/// Returns an error when the input contains unsupported syntax or unsupported
/// scalar values.
pub fn parse_toml(input: &str) -> Result<RawToml, ConfigError> {
    let mut result = BTreeMap::<String, BTreeMap<String, RawTomlValue>>::new();
    let mut current_section = String::new();
    for (index, raw_line) in input.lines().enumerate() {
        let line = strip_toml_comment(raw_line).trim().to_string();
        if line.is_empty() {
            continue;
        }
        if let Some(section) = parse_section(&line) {
            current_section = section;
            result.entry(current_section.clone()).or_default();
            continue;
        }
        let Some((key, value)) = line.split_once('=') else {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML syntax at line {}: {raw_line}",
                index + 1
            )));
        };
        if current_section.is_empty() {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML syntax at line {}: {raw_line}",
                index + 1
            )));
        }
        let key = key.trim();
        if !is_identifier(key) {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML syntax at line {}: {raw_line}",
                index + 1
            )));
        }
        let parsed_value = parse_toml_value(value.trim(), index + 1)?;
        result
            .entry(current_section.clone())
            .or_default()
            .insert(key.to_string(), parsed_value);
    }
    Ok(RawToml(result))
}

fn parse_section(line: &str) -> Option<String> {
    let inner = line.strip_prefix('[')?.strip_suffix(']')?;
    is_identifier(inner).then(|| inner.to_string())
}

fn parse_toml_value(raw: &str, line_number: usize) -> Result<RawTomlValue, ConfigError> {
    if raw.starts_with('"') && raw.ends_with('"') {
        return Ok(RawTomlValue::String(unescape_toml_string(
            raw,
            line_number,
        )?));
    }
    match raw {
        "true" => return Ok(RawTomlValue::Boolean(true)),
        "false" => return Ok(RawTomlValue::Boolean(false)),
        _ => {}
    }
    if raw.chars().all(|char| char.is_ascii_digit()) {
        return raw.parse::<u64>().map(RawTomlValue::Integer).map_err(|_| {
            ConfigError::Parse(format!(
                "Unsupported TOML value at line {line_number}: {raw}"
            ))
        });
    }
    Err(ConfigError::Parse(format!(
        "Unsupported TOML value at line {line_number}: {raw}"
    )))
}

fn unescape_toml_string(raw: &str, line_number: usize) -> Result<String, ConfigError> {
    let mut result = String::new();
    let mut chars = raw[1..raw.len() - 1].chars();
    while let Some(char) = chars.next() {
        if char != '\\' {
            result.push(char);
            continue;
        }
        let Some(escaped) = chars.next() else {
            return Err(ConfigError::Parse(format!(
                "Unsupported TOML value at line {line_number}: {raw}"
            )));
        };
        match escaped {
            '"' => result.push('"'),
            '\\' => result.push('\\'),
            'n' => result.push('\n'),
            'r' => result.push('\r'),
            't' => result.push('\t'),
            other => {
                result.push('\\');
                result.push(other);
            }
        }
    }
    Ok(result)
}

fn strip_toml_comment(line: &str) -> String {
    let mut in_string = false;
    let mut escaped = false;
    for (index, char) in line.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        if char == '\\' && in_string {
            escaped = true;
            continue;
        }
        if char == '"' {
            in_string = !in_string;
        }
        if char == '#' && !in_string {
            return line[..index].to_string();
        }
    }
    line.to_string()
}

fn is_identifier(value: &str) -> bool {
    !value.is_empty()
        && value
            .chars()
            .all(|char| char.is_ascii_alphanumeric() || char == '_')
}

#[cfg(test)]
#[path = "config_toml_tests.rs"]
mod config_toml_tests;
