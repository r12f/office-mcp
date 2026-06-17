use crate::addin_mgr::{
    AddInInfo, DocumentInfo, HostInfo, NewSessionInfo, PartialEffect, ProtectionInfo, RuntimeInfo,
    SessionRegistry,
};
use crate::api::{
    CommandFailure, CommandResult, RegisterClientInput, StartCommandInput, UiClientTransport,
    UiHealth, UiStateStore,
};
use crate::runtime::RuntimeSeedState;
use std::time::{Duration, SystemTime};

#[must_use]
pub(crate) fn seeded_state(now: SystemTime) -> RuntimeSeedState {
    let mut registry = SessionRegistry::new();
    registry.register_runtime(runtime("word-instance", "word", now));
    registry.add_session(word_session(), now);
    registry.register_runtime(runtime("excel-instance", "excel", now));
    registry.add_session(excel_session(), now);
    registry.mark_session_stale("33333333-3333-4333-8333-333333333333", now);

    let mut ui_state = UiStateStore::new();
    ui_state.set_health(
        UiHealth::Degraded,
        Some("Certificate reload failed. Check the configured local PFX path."),
    );
    let client_id = ui_state.register_client(RegisterClientInput {
        client_id: Some("client-1".to_string()),
        transport: UiClientTransport::Http,
        name: Some("copilot-cli/1.0 token=secret-value".to_string()),
    });
    ui_state.start_command(StartCommandInput {
        client_id: Some(client_id.clone()),
        client_name: Some("copilot-cli/1.0".to_string()),
        session_id: Some("11111111-1111-4111-8111-111111111111".to_string()),
        host_app: Some("word".to_string()),
        tool: "word.insert_paragraph".to_string(),
        user_intent: Some("running smoke task".to_string()),
        timeout_ms: Some(30_000),
        ..start_command_default()
    });
    for index in 0..12 {
        let command_id = ui_state.start_command(StartCommandInput {
            client_id: Some(client_id.clone()),
            client_name: Some("copilot-cli/1.0".to_string()),
            session_id: Some("11111111-1111-4111-8111-111111111111".to_string()),
            host_app: Some("word".to_string()),
            tool: if index % 2 == 0 {
                "word.get_text"
            } else {
                "word.insert_paragraph"
            }
            .to_string(),
            user_intent: Some(format!(
                "summarize status token=secret-value base64,QUJDREVGRw== {index}"
            )),
            timeout_ms: Some(30_000),
            started_at: Some(now + Duration::from_millis(index * 10)),
            ..start_command_default()
        });
        let result = match index % 4 {
            0 => CommandResult::Success,
            1 => CommandResult::Failure(CommandFailure {
                office_mcp_code: "IRM_DENIED".to_string(),
                message: "certificate_passphrase=secret-value blocked by document policy."
                    .to_string(),
                tool: Some("word.insert_paragraph".to_string()),
                retriable: false,
                partial_effect: Some(PartialEffect::None),
            }),
            2 => CommandResult::Failure(CommandFailure {
                office_mcp_code: "TIMEOUT".to_string(),
                message: "The add-in did not respond before the deadline.".to_string(),
                tool: Some("word.get_text".to_string()),
                retriable: true,
                partial_effect: Some(PartialEffect::Unknown),
            }),
            _ => CommandResult::Failure(CommandFailure {
                office_mcp_code: "CANCELLED".to_string(),
                message: "The client cancelled the command.".to_string(),
                tool: Some("word.insert_paragraph".to_string()),
                retriable: true,
                partial_effect: Some(PartialEffect::Unknown),
            }),
        };
        ui_state.finish_command(
            &command_id,
            result,
            now + Duration::from_millis(index * 10 + 5),
        );
    }
    RuntimeSeedState { ui_state, registry }
}

#[must_use]
pub(crate) fn empty_state() -> RuntimeSeedState {
    let registry = SessionRegistry::new();
    let mut ui_state = UiStateStore::new();
    ui_state.set_health(UiHealth::Up, None);
    RuntimeSeedState { ui_state, registry }
}

fn runtime(instance_id: &str, app: &str, registered_at: SystemTime) -> RuntimeInfo {
    RuntimeInfo {
        instance_id: instance_id.to_string(),
        host: HostInfo {
            app: app.to_string(),
            version: Some("16.0".to_string()),
            platform: Some("pc".to_string()),
            build: Some("Desktop".to_string()),
        },
        add_in: AddInInfo {
            version: "0.1.6".to_string(),
            protocol_version: "1.0".to_string(),
            supported_features: vec!["doc.read".to_string(), "doc.write".to_string()],
        },
        registered_at,
    }
}

fn word_session() -> NewSessionInfo {
    NewSessionInfo {
        session_id: "11111111-1111-4111-8111-111111111111".to_string(),
        instance_id: "word-instance".to_string(),
        document: DocumentInfo {
            title: Some("Runtime Evidence.docx".to_string()),
            filename: Some("Runtime Evidence.docx".to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(true),
            protection: Some(ProtectionInfo {
                kind: Some("irm".to_string()),
                rights: None,
                rights_source: Some("unavailable".to_string()),
            }),
            ..DocumentInfo::default()
        },
        available_tools: vec![
            "word.get_text".to_string(),
            "word.insert_paragraph".to_string(),
        ],
        is_active: Some(true),
    }
}

fn excel_session() -> NewSessionInfo {
    NewSessionInfo {
        session_id: "33333333-3333-4333-8333-333333333333".to_string(),
        instance_id: "excel-instance".to_string(),
        document: DocumentInfo {
            title: Some("Budget.xlsx".to_string()),
            filename: Some("Budget.xlsx".to_string()),
            is_dirty: Some(true),
            is_read_only: Some(false),
            is_protected: Some(false),
            ..DocumentInfo::default()
        },
        available_tools: vec![
            "excel.read_range".to_string(),
            "excel.write_range".to_string(),
            "excel.add_sheet".to_string(),
            "excel.set_formula".to_string(),
            "excel.format_range".to_string(),
            "excel.create_table".to_string(),
            "excel.create_chart".to_string(),
        ],
        is_active: Some(false),
    }
}

fn start_command_default() -> StartCommandInput {
    StartCommandInput {
        command_id: None,
        mcp_request_id: None,
        client_id: None,
        client_name: None,
        session_id: None,
        host_app: None,
        tool: String::new(),
        user_intent: None,
        timeout_ms: None,
        started_at: None,
    }
}

#[cfg(test)]
#[path = "evidence_fixture_seed_tests.rs"]
mod evidence_fixture_seed_tests;
