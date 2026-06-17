use super::{DocumentDescriptor, DocumentInfo, ProtectionInfo};

#[test]
fn descriptor_uses_explicit_title_when_present() {
    let descriptor = DocumentDescriptor::from(&DocumentInfo {
        title: Some("Quarterly Review".to_string()),
        filename: Some("Review.docx".to_string()),
        ..DocumentInfo::default()
    });

    assert_eq!(descriptor.title.as_deref(), Some("Quarterly Review"));
    assert_eq!(descriptor.filename.as_deref(), Some("Review.docx"));
}

#[test]
fn descriptor_falls_back_to_filename_for_title() {
    let descriptor = DocumentDescriptor::from(&DocumentInfo {
        filename: Some("Draft.docx".to_string()),
        ..DocumentInfo::default()
    });

    assert_eq!(descriptor.title.as_deref(), Some("Draft.docx"));
}

#[test]
fn descriptor_flattens_protection_metadata() {
    let descriptor = DocumentDescriptor::from(&DocumentInfo {
        is_dirty: Some(true),
        is_read_only: Some(false),
        is_protected: Some(true),
        protection: Some(ProtectionInfo {
            kind: Some("tracked_changes".to_string()),
            rights: Some(vec!["comment".to_string(), "review".to_string()]),
            rights_source: Some("office".to_string()),
        }),
        ..DocumentInfo::default()
    });

    assert_eq!(descriptor.is_dirty, Some(true));
    assert_eq!(descriptor.is_read_only, Some(false));
    assert_eq!(descriptor.is_protected, Some(true));
    assert_eq!(
        descriptor.protection_kind.as_deref(),
        Some("tracked_changes")
    );
    assert_eq!(
        descriptor.rights,
        Some(vec!["comment".to_string(), "review".to_string()])
    );
    assert_eq!(descriptor.rights_source.as_deref(), Some("office"));
}
