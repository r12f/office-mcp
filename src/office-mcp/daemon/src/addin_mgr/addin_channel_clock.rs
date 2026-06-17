use std::time::SystemTime;

#[must_use]
pub(crate) fn format_system_time(value: SystemTime) -> String {
    let seconds = value
        .duration_since(SystemTime::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    format!("unix:{seconds}")
}

#[cfg(test)]
#[path = "addin_channel_clock_tests.rs"]
mod addin_channel_clock_tests;
