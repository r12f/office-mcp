use super::AddinChannelError;

#[test]
fn formats_upgrade_and_origin_errors() {
    assert_eq!(
        AddinChannelError::InvalidUpgradePath("/wrong".to_string()).to_string(),
        "Invalid add-in path /wrong."
    );
    assert_eq!(
        AddinChannelError::ForbiddenOrigin("https://example.invalid".to_string()).to_string(),
        "Forbidden add-in origin https://example.invalid."
    );
}

#[test]
fn formats_registration_errors() {
    assert_eq!(
        AddinChannelError::MalformedRegister.to_string(),
        "Malformed register request."
    );
    assert_eq!(
        AddinChannelError::ProtocolVersionMismatch {
            offered: "2.0".to_string(),
            supported: "1.0".to_string(),
        }
        .to_string(),
        "Protocol version mismatch: server supports 1.0, add-in offered 2.0."
    );
}

#[test]
fn formats_connection_and_session_errors() {
    assert_eq!(
        AddinChannelError::UnknownConnection("connection-1".to_string()).to_string(),
        "Unknown add-in connection connection-1."
    );
    assert_eq!(
        AddinChannelError::InstanceMismatch {
            expected: "instance-1".to_string(),
            actual: "instance-2".to_string(),
        }
        .to_string(),
        "Add-in instance mismatch: expected instance-1, got instance-2."
    );
    assert_eq!(
        AddinChannelError::MalformedSessionEvent.to_string(),
        "Malformed session event."
    );
    assert_eq!(
        AddinChannelError::UnknownSession("session-1".to_string()).to_string(),
        "Unknown session session-1."
    );
}
