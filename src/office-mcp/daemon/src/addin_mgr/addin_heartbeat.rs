#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(crate) struct AddinHeartbeatState {
    pending_ping_id: Option<String>,
    missed_pongs: u8,
}

impl AddinHeartbeatState {
    pub(crate) fn start_ping(&mut self, ping_id: String) {
        self.pending_ping_id = Some(ping_id);
    }

    pub(crate) fn handle_pong(&mut self, response_id: &str) -> bool {
        if self.pending_ping_id.as_deref() != Some(response_id) {
            return false;
        }
        self.pending_ping_id = None;
        self.missed_pongs = 0;
        true
    }

    pub(crate) fn record_timeout(&mut self) -> AddinHeartbeatTimeout {
        self.pending_ping_id = None;
        self.missed_pongs = self.missed_pongs.saturating_add(1);
        if self.missed_pongs < 2 {
            AddinHeartbeatTimeout::KeepOpen {
                missed_pongs: self.missed_pongs,
            }
        } else {
            AddinHeartbeatTimeout::Close {
                missed_pongs: self.missed_pongs,
                close_code: 4002,
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum AddinHeartbeatTimeout {
    KeepOpen { missed_pongs: u8 },
    Close { missed_pongs: u8, close_code: u16 },
}

#[cfg(test)]
#[path = "addin_heartbeat_tests.rs"]
mod addin_heartbeat_tests;
