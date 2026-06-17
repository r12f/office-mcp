use std::error::Error;
use std::fmt::{Display, Formatter};
use std::fs;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiAsset {
    pub content: Vec<u8>,
    pub content_type: &'static str,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiAssetStore {
    root: PathBuf,
}

impl UiAssetStore {
    #[must_use]
    pub const fn with_root(root: PathBuf) -> Self {
        Self { root }
    }

    /// Reads a daemon UI asset by file name.
    ///
    /// # Errors
    ///
    /// Returns an error when the asset name is unsafe, the default source tree
    /// path cannot be found, or the asset cannot be read.
    pub fn read(&self, name: &str) -> Result<UiAsset, UiAssetError> {
        if !is_safe_asset_name(name) {
            return Err(UiAssetError::UnsafeName);
        }
        let path = self.root.join(name);
        let content = fs::read(&path).map_err(UiAssetError::Io)?;
        Ok(UiAsset {
            content,
            content_type: content_type(&path),
        })
    }
}

impl Default for UiAssetStore {
    fn default() -> Self {
        Self {
            root: default_ui_asset_dir().unwrap_or_else(|| {
                PathBuf::from("src")
                    .join("office-mcp")
                    .join("daemon")
                    .join("src")
                    .join("ui")
                    .join("assets")
            }),
        }
    }
}

#[derive(Debug)]
pub enum UiAssetError {
    UnsafeName,
    Io(std::io::Error),
}

impl Display for UiAssetError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::UnsafeName => formatter.write_str("unsafe daemon UI asset name"),
            Self::Io(error) => write!(formatter, "daemon UI asset read failed: {error}"),
        }
    }
}

impl Error for UiAssetError {}

fn default_ui_asset_dir() -> Option<PathBuf> {
    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let candidate = ancestor
            .join("src")
            .join("office-mcp")
            .join("daemon")
            .join("src")
            .join("ui")
            .join("assets");
        if candidate.join("index.html").is_file() {
            return Some(candidate);
        }
    }
    None
}

fn is_safe_asset_name(name: &str) -> bool {
    !name.is_empty() && !name.contains("..") && !name.contains('/') && !name.contains('\\')
}

fn content_type(path: &Path) -> &'static str {
    match path.extension().and_then(|extension| extension.to_str()) {
        Some("html") => "text/html; charset=utf-8",
        Some("js") => "text/javascript; charset=utf-8",
        Some("css") => "text/css; charset=utf-8",
        _ => "application/octet-stream",
    }
}

#[cfg(test)]
#[path = "assets_tests.rs"]
mod assets_tests;
