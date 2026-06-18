use crate::addin_mgr::{
    default_office_ctl_assets_dir, default_office_ctl_common_dir,
    default_office_ctl_host_public_dir, static_asset_content_type,
};
use crate::runtime::http_wire::WireHttpResponse;
use crate::ui::UiAssetStore;
use std::collections::BTreeMap;
use std::fs;
use std::path::PathBuf;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct StaticResponseService {
    addin_public_dir: PathBuf,
}

impl StaticResponseService {
    #[must_use]
    pub(crate) fn new(addin_public_dir: PathBuf) -> Self {
        Self { addin_public_dir }
    }

    #[must_use]
    pub(crate) fn serve_ui_asset(&self, name: &str) -> WireHttpResponse {
        let Ok(asset) = UiAssetStore::default().read(name) else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        WireHttpResponse::binary(200, asset.content_type, asset.content, no_store_headers())
    }

    #[must_use]
    pub(crate) fn serve_addin_asset(&self, path: &str) -> WireHttpResponse {
        let path = path.split_once('?').map_or(path, |(path, _)| path);
        if let Some(relative) = path.strip_prefix("/assets/") {
            return Self::serve_brand_asset(relative);
        }
        if let Some(common_path) = path.strip_prefix("/common/") {
            return Self::serve_common_asset(common_path);
        }
        let (host_root, relative) = if let Some(relative) = path.strip_prefix("/powerpoint/") {
            (default_office_ctl_host_public_dir("powerpoint"), relative)
        } else if let Some(relative) = path.strip_prefix("/excel/") {
            (default_office_ctl_host_public_dir("excel"), relative)
        } else if let Some(relative) = path.strip_prefix("/word/") {
            (default_office_ctl_host_public_dir("word"), relative)
        } else {
            let relative = if path == "/" {
                "taskpane.html"
            } else {
                path.trim_start_matches('/')
            };
            (Some(self.addin_public_dir.clone()), relative)
        };
        if relative.contains("..") || relative.contains('\\') || relative.is_empty() {
            return WireHttpResponse::text(403, "Forbidden".to_string());
        }
        let Some(host_root) = host_root else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        let file_path = host_root.join(relative);
        let Ok(content) = fs::read(&file_path) else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        WireHttpResponse::binary(
            200,
            static_asset_content_type(&file_path),
            content,
            no_store_headers(),
        )
    }

    fn serve_common_asset(relative: &str) -> WireHttpResponse {
        if relative.contains("..") || relative.contains('\\') || relative.is_empty() {
            return WireHttpResponse::text(403, "Forbidden".to_string());
        }
        let Some(common_dir) = default_office_ctl_common_dir() else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        let file_path = common_dir.join(relative);
        let Ok(content) = fs::read(&file_path) else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        WireHttpResponse::binary(
            200,
            static_asset_content_type(&file_path),
            content,
            no_store_headers(),
        )
    }

    fn serve_brand_asset(relative: &str) -> WireHttpResponse {
        if relative.contains("..") || relative.contains('\\') || relative.is_empty() {
            return WireHttpResponse::text(403, "Forbidden".to_string());
        }
        let Some(asset_dir) = default_office_ctl_assets_dir() else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        let file_path = asset_dir.join(relative);
        let Ok(content) = fs::read(&file_path) else {
            return WireHttpResponse::text(404, "Not found".to_string());
        };
        WireHttpResponse::binary(
            200,
            static_asset_content_type(&file_path),
            content,
            no_store_headers(),
        )
    }
}

fn no_store_headers() -> BTreeMap<String, String> {
    BTreeMap::from([("Cache-Control".to_string(), "no-store".to_string())])
}

#[cfg(test)]
#[path = "static_response_tests.rs"]
mod static_response_tests;
