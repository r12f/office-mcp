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
            root: default_ui_asset_dir_from_env(std::env::vars()).unwrap_or_else(|| {
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

fn default_ui_asset_dir_from_env<I>(env: I) -> Option<PathBuf>
where
    I: IntoIterator<Item = (String, String)>,
{
    let mut install_root = None;
    for (name, value) in env {
        match name.as_str() {
            "OFFICE_MCP_UI_ASSET_DIR" => {
                let path = PathBuf::from(value);
                if is_ui_asset_dir(&path) {
                    return Some(path);
                }
            }
            "OFFICE_MCP_INSTALL_ROOT" => install_root = Some(PathBuf::from(value)),
            _ => {}
        }
    }
    if let Some(root) = install_root {
        let path = root.join("office-mcp").join("ui");
        if is_ui_asset_dir(&path) {
            return Some(path);
        }
    }
    default_ui_asset_dir()
}

fn default_ui_asset_dir() -> Option<PathBuf> {
    if let Ok(exe_path) = std::env::current_exe()
        && let Some(exe_dir) = exe_path.parent()
    {
        let candidate = exe_dir.join("office-mcp").join("ui");
        if is_ui_asset_dir(&candidate) {
            return Some(candidate);
        }
    }

    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let candidate = ancestor
            .join("src")
            .join("office-mcp")
            .join("daemon")
            .join("src")
            .join("ui")
            .join("assets");
        if is_ui_asset_dir(&candidate) {
            return Some(candidate);
        }
    }
    None
}

fn is_ui_asset_dir(path: &Path) -> bool {
    path.join("index.html").is_file()
        && path.join("app.css").is_file()
        && path.join("app.js").is_file()
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
