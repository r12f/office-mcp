import { randomUUID } from 'node:crypto';
import type { IncomingMessage } from 'node:http';
import type { Duplex } from 'node:stream';
import { WebSocketServer, type WebSocket } from 'ws';
import type { DaemonConfig } from './config.js';
import type { AddinConnection, AddinToolResult, JsonRpcRequest, JsonRpcResponse, RuntimeInfo, SessionInfo } from './types.js';
import { ADDIN_PROTOCOL_VERSION, SERVER_VERSION } from './types.js';
import { SessionRegistry, ToolInvocationError } from './session-registry.js';
import { constantTimeEquals } from './security.js';
import { createLogger } from './logger.js';

export function createAddinChannel(config: DaemonConfig, registry: SessionRegistry): WebSocketServer {
  const wss = new WebSocketServer({ noServer: true, maxPayload: config.limits.maxWsFrameBytes });
  const logger = createLogger(config);

  wss.on('connection', (socket) => {
    const connection: AddinConnection = {
      socket,
      runtime: {
        instance_id: 'unregistered',
        host: { app: 'unknown' },
        add_in: { version: 'unknown', protocol_version: 'unknown' },
        registered_at: new Date().toISOString()
      },
      pending: new Map(),
      queue: Promise.resolve(),
      invokeTool: (sessionId, tool, args, timeoutMs) => invokeTool(connection, sessionId, tool, args, timeoutMs)
    };

    startHeartbeat(config, registry, connection);
    socket.on('message', (data) => handleMessage(config, registry, connection, data.toString()));
    socket.on('close', () => {
      logger.info('add-in websocket closed', { component: 'addin_channel', event: 'addin_closed', instance_id: connection.runtime.instance_id, session_id: connection.session?.session_id });
      stopHeartbeat(connection);
      for (const pending of connection.pending.values()) {
        clearTimeout(pending.timeout);
        pending.reject(new Error('Add-in connection closed'));
      }
      connection.pending.clear();
      registry.removeConnection(connection);
    });
  });

  return wss;
}

export function handleUpgrade(config: DaemonConfig, wss: WebSocketServer, req: IncomingMessage, socket: Duplex, head: Buffer): void {
  if (req.url !== '/addin') {
    socket.destroy();
    return;
  }
  const origin = req.headers.origin;
  if (origin !== config.addin.origin) {
    socket.write('HTTP/1.1 403 Forbidden\r\n\r\n');
    socket.destroy();
    return;
  }
  wss.handleUpgrade(req, socket, head, (ws) => wss.emit('connection', ws, req));
}

function handleMessage(config: DaemonConfig, registry: SessionRegistry, connection: AddinConnection, raw: string): void {
  let message: JsonRpcRequest | JsonRpcResponse;
  try {
    message = JSON.parse(raw) as JsonRpcRequest | JsonRpcResponse;
  } catch {
    sendError(connection.socket, null, -32700, 'Parse error');
    return;
  }

  if ('id' in message && ('result' in message || 'error' in message)) {
    handleResponse(connection, message);
    return;
  }

  if (!('method' in message)) {
    sendError(connection.socket, null, -32600, 'Invalid JSON-RPC message');
    return;
  }

  switch (message.method) {
    case 'register':
      handleRegister(config, registry, connection, message);
      break;
    case 'session.added':
      handleSessionAdded(config, registry, connection, message);
      break;
    case 'session.updated':
      handleSessionUpdated(registry, message);
      break;
    case 'session.removed':
      handleSessionRemoved(registry, message);
      break;
    default:
      if (message.id !== undefined) sendError(connection.socket, message.id, -32601, `Unknown method ${message.method}`);
  }
}

function handleRegister(config: DaemonConfig, registry: SessionRegistry, connection: AddinConnection, request: JsonRpcRequest): void {
  const params = request.params as Record<string, unknown> | undefined;
  const instanceId = typeof params?.instance_id === 'string' ? params.instance_id : '';
  const addIn = params?.add_in as RuntimeInfo['add_in'] | undefined;
  const host = params?.host as RuntimeInfo['host'] | undefined;
  const auth = params?.auth as { shared_secret?: string } | undefined;

  if (!instanceId || !addIn?.protocol_version || !host?.app) {
    sendError(connection.socket, request.id ?? null, -32602, 'Malformed register request');
    connection.socket.close(4003, 'Malformed register request');
    return;
  }
  if (!sameMajor(addIn.protocol_version, ADDIN_PROTOCOL_VERSION)) {
    sendError(connection.socket, request.id ?? null, -32000, `Protocol version mismatch: server supports ${ADDIN_PROTOCOL_VERSION}.`, {
      office_mcp_code: 'PROTOCOL_VERSION_MISMATCH',
      server_protocol: ADDIN_PROTOCOL_VERSION
    });
    connection.socket.close(4003, 'Protocol mismatch');
    return;
  }
  if (config.addin.sharedSecret && !constantTimeEquals(config.addin.sharedSecret, auth?.shared_secret ?? '')) {
    sendError(connection.socket, request.id ?? null, -32000, 'Add-in authentication failed.', { office_mcp_code: 'AUTH_FAILED' });
    connection.socket.close(4003, 'Authentication failed');
    return;
  }

  const runtime: RuntimeInfo = {
    instance_id: instanceId,
    host,
    add_in: addIn,
    registered_at: new Date().toISOString()
  };
  registry.registerRuntime(connection, runtime);
  createLogger(config).info('add-in runtime registered', { component: 'addin_channel', event: 'addin_registered', instance_id: instanceId, host_app: host.app, addin_version: addIn.version });
  sendResult(connection.socket, request.id ?? null, {
    server_version: SERVER_VERSION,
    protocol_version: ADDIN_PROTOCOL_VERSION,
    session_grace_sec: config.addin.sessionGraceSec,
    heartbeat_interval_sec: config.addin.heartbeatIntervalSec,
    max_pending_per_session: config.addin.maxPendingPerSession,
    assigned_instance_id: instanceId
  });
}

function handleSessionAdded(config: DaemonConfig, registry: SessionRegistry, connection: AddinConnection, request: JsonRpcRequest): void {
  const params = request.params as Record<string, unknown> | undefined;
  if (!params || typeof params.session_id !== 'string' || typeof params.instance_id !== 'string') return;
  if (params.instance_id !== connection.runtime.instance_id) return;

  registry.addSession(connection, {
    session_id: params.session_id,
    instance_id: params.instance_id,
    document: (params.document as SessionInfo['document']) ?? {},
    available_tools: Array.isArray(params.available_tools) ? params.available_tools.filter((tool): tool is string => typeof tool === 'string') : [],
    is_active: typeof params.is_active === 'boolean' ? params.is_active : null
  });
  const document = params.document as { title?: string; filename?: string; url?: string } | undefined;
  createLogger(config).info('add-in session registered', { component: 'addin_channel', event: 'session_added', session_id: params.session_id, instance_id: params.instance_id, document_title: document?.title, document_url: document?.url });
}

function handleSessionUpdated(registry: SessionRegistry, request: JsonRpcRequest): void {
  const params = request.params as Record<string, unknown> | undefined;
  if (!params || typeof params.session_id !== 'string') return;
  registry.updateSession(params.session_id, params as Partial<SessionInfo>);
}

function handleSessionRemoved(registry: SessionRegistry, request: JsonRpcRequest): void {
  const params = request.params as Record<string, unknown> | undefined;
  if (!params || typeof params.session_id !== 'string') return;
  registry.removeSession(params.session_id);
}

function handleResponse(connection: AddinConnection, response: JsonRpcResponse): void {
  if (typeof response.id !== 'string') return;
  if (response.id === connection.heartbeat?.pendingPingId) {
    if (connection.heartbeat.timeout) clearTimeout(connection.heartbeat.timeout);
    connection.heartbeat.pendingPingId = undefined;
    connection.heartbeat.timeout = undefined;
    connection.heartbeat.missedPongs = 0;
    return;
  }
  const pending = connection.pending.get(response.id);
  if (!pending) return;
  clearTimeout(pending.timeout);
  connection.pending.delete(response.id);
  if (response.error) {
    pending.reject(new Error(response.error.message));
  } else {
    pending.resolve(response.result as AddinToolResult);
  }
}

function startHeartbeat(config: DaemonConfig, registry: SessionRegistry, connection: AddinConnection): void {
  const sendPing = () => {
    if (connection.socket.readyState !== connection.socket.OPEN) return;
    if (connection.heartbeat?.pendingPingId) {
      return;
    }

    const id = randomUUID();
    connection.heartbeat!.pendingPingId = id;
    connection.socket.send(JSON.stringify({
      jsonrpc: '2.0',
      id,
      method: 'ping',
      params: { ts: new Date().toISOString() }
    } satisfies JsonRpcRequest));
    connection.heartbeat!.timeout = setTimeout(() => {
      if (connection.heartbeat?.pendingPingId === id) {
        connection.heartbeat.missedPongs += 1;
        connection.heartbeat.pendingPingId = undefined;
        connection.heartbeat.timeout = undefined;
        if (connection.heartbeat.missedPongs >= 2) {
          registry.markConnectionStale(connection);
          connection.socket.close(4002, 'Heartbeat missed');
        }
      }
    }, config.addin.heartbeatTimeoutSec * 1000);
  };

  connection.heartbeat = {
    interval: setInterval(sendPing, config.addin.heartbeatIntervalSec * 1000),
    missedPongs: 0
  };
}

function stopHeartbeat(connection: AddinConnection): void {
  if (!connection.heartbeat) return;
  clearInterval(connection.heartbeat.interval);
  if (connection.heartbeat.timeout) clearTimeout(connection.heartbeat.timeout);
  connection.heartbeat = undefined;
}

function invokeTool(connection: AddinConnection, sessionId: string, tool: string, args: Record<string, unknown>, timeoutMs: number): Promise<AddinToolResult> {
  const id = randomUUID();
  const payload: JsonRpcRequest = {
    jsonrpc: '2.0',
    id,
    method: 'tool.invoke',
    params: {
      session_id: sessionId,
      tool,
      args,
      timeout_ms: timeoutMs
    }
  };
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      connection.pending.delete(id);
      sendCancel(connection.socket, id, 'deadline_expired');
      reject(new ToolInvocationError({
        office_mcp_code: 'TIMEOUT',
        message: `Tool ${tool} timed out after ${timeoutMs}ms.`,
        session_id: sessionId,
        tool,
        retriable: true,
        partial_effect: 'unknown'
      }));
    }, timeoutMs);
    connection.pending.set(id, { resolve, reject, timeout, sessionId, tool });
    connection.socket.send(JSON.stringify(payload), (error) => {
      if (!error) return;
      clearTimeout(timeout);
      connection.pending.delete(id);
      reject(error);
    });
  }) as Promise<AddinToolResult>;
}

function sendCancel(socket: WebSocket, requestId: string, reason: string): void {
  if (socket.readyState !== socket.OPEN) return;
  socket.send(JSON.stringify({
    jsonrpc: '2.0',
    method: 'tool.cancel',
    params: { request_id: requestId, reason }
  } satisfies JsonRpcRequest));
}

function sendResult(socket: WebSocket, id: string | number | null, result: unknown): void {
  socket.send(JSON.stringify({ jsonrpc: '2.0', id, result } satisfies JsonRpcResponse));
}

function sendError(socket: WebSocket, id: string | number | null, code: number, message: string, data?: unknown): void {
  socket.send(JSON.stringify({ jsonrpc: '2.0', id, error: { code, message, data } } satisfies JsonRpcResponse));
}

function sameMajor(left: string, right: string): boolean {
  return left.split('.')[0] === right.split('.')[0];
}
