use std::error::Error;
use std::fmt::{Display, Formatter};

#[derive(Debug)]
pub enum ConfigError {
    Io(std::io::Error),
    Parse(String),
    Validation(String),
}

impl Display for ConfigError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::Parse(message) | Self::Validation(message) => formatter.write_str(message),
        }
    }
}

impl Error for ConfigError {}
