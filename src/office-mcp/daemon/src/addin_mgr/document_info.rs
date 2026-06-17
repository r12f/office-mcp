#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DocumentInfo {
    pub title: Option<String>,
    pub url: Option<String>,
    pub filename: Option<String>,
    pub is_dirty: Option<bool>,
    pub is_read_only: Option<bool>,
    pub is_protected: Option<bool>,
    pub protection: Option<ProtectionInfo>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionInfo {
    pub kind: Option<String>,
    pub rights: Option<Vec<String>>,
    pub rights_source: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentDescriptor {
    pub title: Option<String>,
    pub url: Option<String>,
    pub filename: Option<String>,
    pub is_dirty: Option<bool>,
    pub is_read_only: Option<bool>,
    pub is_protected: Option<bool>,
    pub protection_kind: Option<String>,
    pub rights: Option<Vec<String>>,
    pub rights_source: Option<String>,
}

impl From<&DocumentInfo> for DocumentDescriptor {
    fn from(document: &DocumentInfo) -> Self {
        Self {
            title: document.title.clone().or_else(|| document.filename.clone()),
            url: document.url.clone(),
            filename: document.filename.clone(),
            is_dirty: document.is_dirty,
            is_read_only: document.is_read_only,
            is_protected: document.is_protected,
            protection_kind: document
                .protection
                .as_ref()
                .and_then(|protection| protection.kind.clone()),
            rights: document
                .protection
                .as_ref()
                .and_then(|protection| protection.rights.clone()),
            rights_source: document
                .protection
                .as_ref()
                .and_then(|protection| protection.rights_source.clone()),
        }
    }
}

#[cfg(test)]
#[path = "document_info_tests.rs"]
mod document_info_tests;
