import type { AddinConnection, AddinToolResult, RuntimeInfo, SessionDescriptor, SessionInfo, ToolFailure } from './types.js';

export class ToolInvocationError extends Error {
  readonly failure: ToolFailure;

  constructor(failure: ToolFailure) {
    super(failure.message);
    this.failure = failure;
  }
}

export class SessionRegistry {
  private readonly connectionsByInstance = new Map<string, AddinConnection>();
  private readonly sessionsById = new Map<string, SessionInfo>();
  private readonly connectionsBySession = new Map<string, AddinConnection>();

  constructor(private readonly maxPendingPerSession = 4) {}

  registerRuntime(connection: AddinConnection, runtime: RuntimeInfo): void {
    const existing = this.connectionsByInstance.get(runtime.instance_id);
    if (existing && existing !== connection) {
      existing.socket.close(4005, 'Add-in replaced');
      this.markConnectionStale(existing);
    }
    connection.runtime = runtime;
    this.connectionsByInstance.set(runtime.instance_id, connection);
  }

  addSession(connection: AddinConnection, session: Omit<SessionInfo, 'status' | 'registered_at'>): SessionInfo {
    const full: SessionInfo = {
      ...session,
      status: 'active',
      registered_at: new Date().toISOString()
    };
    connection.session = full;
    this.sessionsById.set(full.session_id, full);
    this.connectionsBySession.set(full.session_id, connection);
    return full;
  }

  updateSession(sessionId: string, patch: Partial<SessionInfo>): void {
    const existing = this.sessionsById.get(sessionId);
    if (!existing) return;
    this.sessionsById.set(sessionId, { ...existing, ...patch });
  }

  removeSession(sessionId: string): void {
    const connection = this.connectionsBySession.get(sessionId);
    if (connection?.session?.session_id === sessionId) {
      connection.session = undefined;
    }
    this.sessionsById.delete(sessionId);
    this.connectionsBySession.delete(sessionId);
  }

  removeConnection(connection: AddinConnection): void {
    if (connection.session) {
      this.markSessionStale(connection.session.session_id);
    }
    this.connectionsByInstance.delete(connection.runtime.instance_id);
  }

  markConnectionStale(connection: AddinConnection): void {
    if (connection.session) {
      this.markSessionStale(connection.session.session_id);
    }
  }

  markSessionStale(sessionId: string): void {
    const session = this.sessionsById.get(sessionId);
    if (!session || session.status === 'stale') return;
    const stale = { ...session, status: 'stale' as const, stale_since: new Date().toISOString() };
    this.sessionsById.set(sessionId, stale);
  }

  pruneStaleSessions(graceSec: number): void {
    const now = Date.now();
    for (const session of this.sessionsById.values()) {
      if (session.status !== 'stale' || !session.stale_since) continue;
      if (now - Date.parse(session.stale_since) > graceSec * 1000) {
        this.removeSession(session.session_id);
      }
    }
  }

  listSessions(): SessionDescriptor[] {
    return [...this.sessionsById.values()].map((session) => this.describe(session));
  }

  getSessionInfo(sessionId: string): (SessionDescriptor & { available_tools: string[] }) | undefined {
    const session = this.sessionsById.get(sessionId);
    if (!session) return undefined;
    return { ...this.describe(session), available_tools: session.available_tools };
  }

  async invoke(sessionId: string, tool: string, args: Record<string, unknown>, timeoutMs: number): Promise<AddinToolResult> {
    return this.invokeWithCapabilityCheck(sessionId, tool, args, timeoutMs, true);
  }

  async invokeInternal(sessionId: string, tool: string, args: Record<string, unknown>, timeoutMs: number): Promise<AddinToolResult> {
    return this.invokeWithCapabilityCheck(sessionId, tool, args, timeoutMs, false);
  }

  private async invokeWithCapabilityCheck(sessionId: string, tool: string, args: Record<string, unknown>, timeoutMs: number, checkCapability: boolean): Promise<AddinToolResult> {
    const session = this.sessionsById.get(sessionId);
    if (!session) {
      if (this.sessionsById.size === 0) {
        throw new ToolInvocationError({
          office_mcp_code: 'NO_SESSIONS',
          message: 'No Office document sessions are connected. Activate the office-mcp add-in in Word and try again.',
          session_id: sessionId,
          tool,
          retriable: true
        });
      }
      throw new ToolInvocationError({
        office_mcp_code: 'SESSION_NOT_FOUND',
        message: `Session ${sessionId} is not registered.`,
        session_id: sessionId,
        tool,
        retriable: false
      });
    }
    if (session.status === 'stale') {
      throw new ToolInvocationError({
        office_mcp_code: 'SESSION_STALE',
        message: `Session ${sessionId} is stale while the add-in reconnects.`,
        session_id: sessionId,
        tool,
        retriable: true
      });
    }
    if (checkCapability && !session.available_tools.includes(tool)) {
      throw new ToolInvocationError({
        office_mcp_code: 'HOST_CAPABILITY_UNAVAILABLE',
        message: `The selected Word session does not support ${tool}.`,
        session_id: sessionId,
        tool,
        retriable: false
      });
    }
    const connection = this.connectionsBySession.get(sessionId);
    if (!connection || connection.socket.readyState !== connection.socket.OPEN) {
      this.markSessionStale(sessionId);
      throw new ToolInvocationError({
        office_mcp_code: 'SESSION_LOST',
        message: `Session ${sessionId} lost its add-in connection.`,
        session_id: sessionId,
        tool,
        retriable: false,
        partial_effect: 'unknown'
      });
    }

    const queued = connection.pending.size;
    if (queued >= this.maxPendingPerSession) {
      throw new ToolInvocationError({
        office_mcp_code: 'MAX_PENDING_EXCEEDED',
        message: `Session ${sessionId} has too many pending tool calls.`,
        session_id: sessionId,
        tool,
        retriable: true
      });
    }

    const run = () => connection.invokeTool(sessionId, tool, args, timeoutMs);
    const result = connection.queue.then(run, run);
    connection.queue = result.catch(() => undefined);
    return result;
  }

  private describe(session: SessionInfo): SessionDescriptor {
    const doc = session.document;
    return {
      session_id: session.session_id,
      instance_id: session.instance_id,
      app: 'word',
      document: {
        title: doc.title ?? doc.filename ?? null,
        url: doc.url ?? null,
        filename: doc.filename ?? null,
        is_dirty: doc.is_dirty ?? null,
        is_protected: doc.is_protected ?? null,
        protection_kind: doc.protection?.kind ?? null,
        rights: doc.protection?.rights ?? null,
        rights_source: doc.protection?.rights_source ?? null
      },
      is_active: session.is_active ?? null,
      capability_tiers: inferCapabilityTiers(session.available_tools),
      available_tool_count: session.available_tools.length,
      registered_at: session.registered_at,
      status: session.status
    };
  }
}

function inferCapabilityTiers(tools: string[]): string[] {
  const tiers = ['core'];
  if (tools.includes('word.add_comment')) tiers.push('review');
  if (tools.includes('word.accept_change')) tiers.push('tracked_changes');
  return tiers;
}


