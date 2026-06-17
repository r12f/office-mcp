use std::path::{Path, PathBuf};

#[must_use]
pub fn static_asset_content_type(path: &Path) -> &'static str {
    match path.extension().and_then(|extension| extension.to_str()) {
        Some("html") => "text/html; charset=utf-8",
        Some("js") => "text/javascript; charset=utf-8",
        Some("css") => "text/css; charset=utf-8",
        _ => "application/octet-stream",
    }
}

#[must_use]
pub fn default_addin_public_dir() -> PathBuf {
    find_addin_public_dir_from(&std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
        .unwrap_or_else(|| {
            PathBuf::from("src")
                .join("office-ctl")
                .join("word")
                .join("public")
        })
}

#[must_use]
pub fn default_office_ctl_common_dir() -> Option<PathBuf> {
    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let source_candidate = ancestor.join("src").join("office-ctl").join("common");
        if source_candidate.join("browser-ui.js").is_file() {
            return Some(source_candidate);
        }
        let installed_candidate = ancestor.join("office-ctl").join("common");
        if installed_candidate.join("browser-ui.js").is_file() {
            return Some(installed_candidate);
        }
    }
    None
}

#[must_use]
pub fn default_office_ctl_host_public_dir(host: &str) -> Option<PathBuf> {
    let current = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    for ancestor in current.ancestors() {
        let source_candidate = ancestor
            .join("src")
            .join("office-ctl")
            .join(host)
            .join("public");
        if source_candidate.is_dir() {
            return Some(source_candidate);
        }
        let installed_candidate = ancestor.join("office-ctl").join(host).join("public");
        if installed_candidate.is_dir() {
            return Some(installed_candidate);
        }
    }
    None
}

#[must_use]
pub fn find_addin_public_dir_from(start: &Path) -> Option<PathBuf> {
    for ancestor in start.ancestors() {
        let candidate = ancestor
            .join("src")
            .join("office-ctl")
            .join("word")
            .join("public");
        if candidate.join("taskpane.html").is_file() {
            return Some(candidate);
        }
    }
    None
}

#[cfg(test)]
#[path = "static_assets_tests.rs"]
mod static_assets_tests;
