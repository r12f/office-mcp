use super::same_major_version;

#[test]
fn matches_same_major_protocol_version() {
    assert!(same_major_version("1.0", "1.3"));
    assert!(same_major_version("2", "2.1"));
    assert!(!same_major_version("1.0", "2.0"));
}
