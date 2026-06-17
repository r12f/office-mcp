use super::websocket_accept_key;

#[test]
fn computes_rfc6455_accept_key() {
    assert_eq!(
        websocket_accept_key("dGhlIHNhbXBsZSBub25jZQ=="),
        "s3pPLMBiTxaQ9kYGzzhZRbK+xOo="
    );
}

#[test]
fn computes_accept_key_for_office_runtime_nonce() {
    assert_eq!(
        websocket_accept_key("office-mcp-runtime"),
        "icF4hSYJz06v8InPEmA+69sJ+lw="
    );
}
