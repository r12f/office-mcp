use super::AddinRegistrationPolicy;
use crate::addin_mgr::addin_channel::{ADDIN_PROTOCOL_VERSION, SERVER_VERSION};
use crate::addin_mgr::{
    AddInInfo, AddinChannelConfig, AddinChannelError, HostInfo, JsonRpcId, RegisterRequest,
};
use std::time::{Duration, SystemTime};

#[test]
fn rejects_malformed_register_requests() {
    let policy = policy();

    let error = policy
        .validate(
            "connection-1",
            &register_request("", ADDIN_PROTOCOL_VERSION),
        )
        .expect_err("malformed register");

    assert_eq!(error, AddinChannelError::MalformedRegister);
}

#[test]
fn rejects_protocol_major_mismatch() {
    let policy = policy();

    let error = policy
        .validate("connection-1", &register_request("instance-1", "2.0"))
        .expect_err("protocol mismatch");

    assert_eq!(
        error,
        AddinChannelError::ProtocolVersionMismatch {
            offered: "2.0".to_string(),
            supported: ADDIN_PROTOCOL_VERSION.to_string(),
        }
    );
}

#[test]
fn builds_runtime_from_register_request() {
    let now = SystemTime::UNIX_EPOCH + Duration::from_secs(42);

    let runtime = AddinRegistrationPolicy::runtime_from(
        register_request("instance-1", ADDIN_PROTOCOL_VERSION),
        now,
    );

    assert_eq!(runtime.instance_id, "instance-1");
    assert_eq!(runtime.host.app, "word");
    assert_eq!(runtime.add_in.protocol_version, ADDIN_PROTOCOL_VERSION);
    assert_eq!(runtime.registered_at, now);
}

#[test]
fn builds_register_result_from_config() {
    let policy = policy();
    let runtime = AddinRegistrationPolicy::runtime_from(
        register_request("instance-1", ADDIN_PROTOCOL_VERSION),
        SystemTime::UNIX_EPOCH,
    );
    let config = AddinChannelConfig {
        heartbeat_interval: Duration::from_secs(7),
        session_grace: Duration::from_secs(31),
        max_pending_per_session: 9,
        ..AddinChannelConfig::default()
    };

    let result = policy.register_result(&runtime, &config);

    assert_eq!(result.server_version, SERVER_VERSION);
    assert_eq!(result.protocol_version, ADDIN_PROTOCOL_VERSION);
    assert_eq!(result.session_grace_sec, 31);
    assert_eq!(result.heartbeat_interval_sec, 7);
    assert_eq!(result.max_pending_per_session, 9);
    assert_eq!(result.assigned_instance_id, "instance-1");
}

fn policy() -> AddinRegistrationPolicy {
    AddinRegistrationPolicy::new(SERVER_VERSION, ADDIN_PROTOCOL_VERSION)
}

fn register_request(instance_id: &str, protocol_version: &str) -> RegisterRequest {
    RegisterRequest {
        id: JsonRpcId::String("register-1".to_string()),
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: "word".to_string(),
            version: Some("16.0".to_string()),
            platform: Some("windows".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.0".to_string(),
            protocol_version: protocol_version.to_string(),
            supported_features: vec!["doc.read".to_string()],
        },
    }
}
