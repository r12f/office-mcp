use std::time::{Duration, Instant};

#[derive(Debug, Clone)]
pub(crate) struct WebSocketHeartbeatState {
    interval: Duration,
    timeout: Duration,
    next_ping_at: Instant,
    deadline: Option<Instant>,
}

impl WebSocketHeartbeatState {
    #[must_use]
    pub(crate) fn new(now: Instant, interval: Duration, timeout: Duration) -> Self {
        Self {
            interval,
            timeout,
            next_ping_at: now + interval,
            deadline: None,
        }
    }

    #[must_use]
    pub(crate) fn deadline_elapsed(&self, now: Instant) -> bool {
        self.deadline.is_some_and(|deadline| now >= deadline)
    }

    #[must_use]
    pub(crate) fn should_start_ping(&self, now: Instant) -> bool {
        self.deadline.is_none() && now >= self.next_ping_at
    }

    pub(crate) fn mark_ping_sent(&mut self, now: Instant) {
        self.deadline = Some(now + self.timeout);
    }

    pub(crate) fn mark_ping_skipped(&mut self, now: Instant) {
        self.deadline = None;
        self.next_ping_at = now + self.interval;
    }

    pub(crate) fn mark_pong_received(&mut self, now: Instant) {
        self.deadline = None;
        self.next_ping_at = now + self.interval;
    }
}

#[cfg(test)]
#[path = "websocket_heartbeat_tests.rs"]
mod websocket_heartbeat_tests;
