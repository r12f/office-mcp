use crate::addin_mgr::QueuedCommand;
use std::collections::VecDeque;
use std::time::SystemTime;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(crate) struct SessionCommandQueue {
    commands: VecDeque<QueuedCommand>,
    next_sequence: u64,
}

impl SessionCommandQueue {
    pub(crate) fn push(&mut self, command: QueuedCommand) {
        self.commands.push_back(command);
    }

    pub(crate) fn remove(&mut self, request_id: &str) -> Option<QueuedCommand> {
        let index = self
            .commands
            .iter()
            .position(|command| command.request_id == request_id)?;
        self.commands.remove(index)
    }

    pub(crate) fn mark_dispatched(&mut self, request_id: &str) -> bool {
        let Some(command) = self
            .commands
            .iter_mut()
            .find(|command| command.request_id == request_id)
        else {
            return false;
        };
        command.dispatched = true;
        true
    }

    pub(crate) fn expired_request_ids(&self, now: SystemTime) -> Vec<String> {
        self.commands
            .iter()
            .filter(|command| command.deadline_at <= now)
            .map(|command| command.request_id.clone())
            .collect()
    }

    pub(crate) fn next_sequence(&mut self) -> u64 {
        let value = self.next_sequence;
        self.next_sequence += 1;
        value
    }

    pub(crate) fn len(&self) -> usize {
        self.commands.len()
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.commands.is_empty()
    }
}

#[cfg(test)]
#[path = "command_queue_tests.rs"]
mod command_queue_tests;
