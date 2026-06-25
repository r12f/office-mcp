use crate::runtime::mcp_response::RuntimeSharedState;
use std::sync::Arc;
use std::thread;
use std::time::{Duration, SystemTime};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SessionCleanupService {
    interval: Duration,
}

impl SessionCleanupService {
    #[must_use]
    pub(crate) const fn new(interval: Duration) -> Self {
        Self { interval }
    }

    #[must_use]
    pub(crate) fn for_session_grace(session_grace: Duration) -> Self {
        Self::new(cleanup_interval(session_grace))
    }

    pub(crate) fn spawn(self, shared_state: Arc<RuntimeSharedState>) {
        thread::spawn(move || self.run_forever(&shared_state));
    }

    pub(crate) fn run_once(shared_state: &RuntimeSharedState, now: SystemTime) -> usize {
        shared_state.prune_stale_sessions(now)
    }

    fn run_forever(self, shared_state: &RuntimeSharedState) -> ! {
        loop {
            thread::sleep(self.interval);
            Self::run_once(shared_state, SystemTime::now());
        }
    }
}

fn cleanup_interval(session_grace: Duration) -> Duration {
    let half_grace = session_grace / 2;
    half_grace.clamp(Duration::from_secs(5), Duration::from_mins(1))
}

#[cfg(test)]
#[path = "session_cleanup_service_tests.rs"]
mod session_cleanup_service_tests;
