(function attachOfficeCtlTaskHistory(global) {
  const DEFAULT_HISTORY_LIMIT = 20;

  class TaskHistoryStore {
    constructor(options = {}) {
      this.historyLimit = options.historyLimit || DEFAULT_HISTORY_LIMIT;
      this.now = options.now || (() => Date.now());
      this.toIso = options.toIso || (() => new Date().toISOString());
      this.redactText = options.redactText || ((value) => String(value || '').slice(0, 300));
      this.currentTask = null;
      this.history = [];
      this.cancelledRequests = new Set();
    }

    start(requestId, tool, args = {}, timeoutMs = null) {
      const startedAt = this.now();
      this.currentTask = {
        requestId,
        tool,
        userIntent: args?.client_meta?.user_intent || '',
        startedAt,
        timeoutMs: timeoutMs || null,
        deadlineAt: timeoutMs ? startedAt + timeoutMs : null,
        cancelRequested: this.cancelledRequests.has(requestId)
      };
      return this.currentTask;
    }

    finish(requestId, status, elapsedMs, error = null) {
      if (!this.currentTask || this.currentTask.requestId !== requestId) return null;
      const task = {
        requestId,
        tool: this.currentTask.tool,
        userIntent: this.currentTask.userIntent,
        status,
        elapsedMs,
        error: error ? this.sanitizeError(error) : null,
        completedAt: this.toIso()
      };
      this.history.unshift(task);
      this.history.splice(this.historyLimit);
      this.currentTask = null;
      return task;
    }

    cancel(requestId) {
      this.cancelledRequests.add(requestId);
      if (this.currentTask?.requestId === requestId) {
        this.currentTask.cancelRequested = true;
      }
      return this.currentTask;
    }

    consumeCancellation(requestId) {
      return this.cancelledRequests.delete(requestId);
    }

    isCancelled(requestId) {
      return this.cancelledRequests.has(requestId);
    }

    snapshot() {
      return {
        currentTask: this.currentTask ? { ...this.currentTask } : null,
        history: this.history.map((task) => ({ ...task, error: task.error ? { ...task.error } : null })),
        historyLimit: this.historyLimit
      };
    }

    sanitizeError(error) {
      return {
        office_mcp_code: error.office_mcp_code || error.officeMcpCode || 'GENERIC_FAILURE',
        message: this.redactText(error.message || 'Command failed.'),
        retriable: Boolean(error.retriable),
        partial_effect: error.partial_effect || error.partialEffect || 'unknown'
      };
    }
  }

  global.OfficeCtlTaskHistory = Object.freeze({
    DEFAULT_HISTORY_LIMIT,
    TaskHistoryStore
  });
})(globalThis);
