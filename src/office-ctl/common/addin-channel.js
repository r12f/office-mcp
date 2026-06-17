(function attachOfficeCtlAddinChannel(global) {
  const ENDPOINT_STORAGE_KEY = 'office-mcp.addin-endpoint';
  const INSTANCE_STORAGE_KEY = 'office-mcp.instance-id';
  const REGISTER_REQUEST_STORAGE_KEY = 'office-mcp.register-request-id';
  const SESSION_STORAGE_KEY = 'office-mcp.session-id';
  const WEBSOCKET_OPEN = 1;

  function defaultEndpoint(locationLike = global.location) {
    const origin = String(locationLike?.origin || 'https://localhost:8765');
    return `${origin.replace(/^https:/i, 'wss:').replace(/^http:/i, 'ws:')}/addin`;
  }

  function configuredEndpoint(options = {}) {
    const storage = options.storage || global.localStorage;
    const locationLike = options.location || global.location;
    return storage?.getItem(ENDPOINT_STORAGE_KEY) || defaultEndpoint(locationLike);
  }

  function currentOriginEndpoint(options = {}) {
    return defaultEndpoint(options.location || global.location);
  }

  function clearEndpointOverride(options = {}) {
    const storage = options.storage || global.localStorage;
    storage?.removeItem(ENDPOINT_STORAGE_KEY);
    return currentOriginEndpoint(options);
  }

  function saveEndpointOverride(value, options = {}) {
    const endpoint = String(value || '').trim();
    validateEndpoint(endpoint);
    const storage = options.storage || global.localStorage;
    storage?.setItem(ENDPOINT_STORAGE_KEY, endpoint);
    return endpoint;
  }

  function validateEndpoint(value) {
    let parsed;
    try {
      parsed = new URL(String(value || '').trim());
    } catch {
      throw new Error('Enter a valid wss://localhost endpoint.');
    }
    if (parsed.protocol !== 'wss:' || parsed.hostname !== 'localhost') {
      throw new Error('Use a wss://localhost endpoint, for example wss://localhost:8765/addin.');
    }
    return parsed;
  }

  function runtimeIds(options = {}) {
    const storage = options.storage || global.sessionStorage;
    const randomUUID = options.randomUUID || global.crypto?.randomUUID?.bind(global.crypto);
    if (!randomUUID) throw new Error('crypto.randomUUID is required.');
    const instanceId = storage?.getItem(INSTANCE_STORAGE_KEY) || randomUUID();
    const sessionId = storage?.getItem(SESSION_STORAGE_KEY) || randomUUID();
    storage?.setItem(INSTANCE_STORAGE_KEY, instanceId);
    storage?.setItem(SESSION_STORAGE_KEY, sessionId);
    return { instanceId, sessionId };
  }

  function rememberRegisterRequest(id, options = {}) {
    const storage = options.storage || global.sessionStorage;
    storage?.setItem(REGISTER_REQUEST_STORAGE_KEY, String(id));
  }

  function clearRegisterRequest(options = {}) {
    const storage = options.storage || global.sessionStorage;
    storage?.removeItem(REGISTER_REQUEST_STORAGE_KEY);
  }

  function isRegisterResponse(message, options = {}) {
    const storage = options.storage || global.sessionStorage;
    return Boolean(message?.id && storage?.getItem(REGISTER_REQUEST_STORAGE_KEY) === String(message.id));
  }

  function parseJsonRpc(raw) {
    try {
      const message = JSON.parse(raw);
      return message && typeof message === 'object' ? message : null;
    } catch {
      return null;
    }
  }

  function sendJsonRpc(socket, message) {
    if (socket && socket.readyState === WEBSOCKET_OPEN) {
      socket.send(JSON.stringify(message));
      return true;
    }
    return false;
  }

  function reply(socket, id, result) {
    return sendJsonRpc(socket, { jsonrpc: '2.0', id, result });
  }

  function registerRequest(id, params) {
    return {
      jsonrpc: '2.0',
      id,
      method: 'register',
      params
    };
  }

  function sessionAddedNotification(params) {
    return {
      jsonrpc: '2.0',
      method: 'session.added',
      params
    };
  }

  function sessionUpdatedNotification(params) {
    return {
      jsonrpc: '2.0',
      method: 'session.updated',
      params
    };
  }

  function registerResult(message, fallbackProtocolVersion = '1.0') {
    return {
      serverVersion: message?.result?.server_version || 'Unknown',
      protocolVersion: message?.result?.protocol_version || fallbackProtocolVersion
    };
  }

  function isMethod(message, method) {
    return message?.method === method;
  }

  function isToolInvoke(message) {
    return isMethod(message, 'tool.invoke');
  }

  function isPing(message) {
    return isMethod(message, 'ping');
  }

  function isToolCancel(message) {
    return isMethod(message, 'tool.cancel') && Boolean(message?.params?.request_id);
  }

  function reconnectDelay(attempt, random = Math.random) {
    return Math.min(10000, 1000 + attempt * 500 + Math.floor(random() * 500));
  }

  global.OfficeCtlAddinChannel = Object.freeze({
    ENDPOINT_STORAGE_KEY,
    INSTANCE_STORAGE_KEY,
    REGISTER_REQUEST_STORAGE_KEY,
    SESSION_STORAGE_KEY,
    clearEndpointOverride,
    clearRegisterRequest,
    configuredEndpoint,
    currentOriginEndpoint,
    defaultEndpoint,
    isPing,
    isRegisterResponse,
    isToolCancel,
    isToolInvoke,
    parseJsonRpc,
    registerRequest,
    registerResult,
    reconnectDelay,
    rememberRegisterRequest,
    reply,
    runtimeIds,
    saveEndpointOverride,
    sessionAddedNotification,
    sessionUpdatedNotification,
    sendJsonRpc,
    validateEndpoint
  });
})(globalThis);
