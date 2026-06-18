(() => {
  const ADDIN_VERSION = '0.1.0';
  const PROTOCOL_VERSION = '1.0';
  const { escapeHtml, fileName, formatDuration, formatTime, titleCase, redactText } = window.OfficeCtlCommon;
  const {
    clearEndpointOverride,
    clearRegisterRequest,
    configuredEndpoint,
    currentOriginEndpoint,
    isPing,
    isRegisterResponse,
    isToolCancel,
    isToolInvoke,
    parseJsonRpc,
    registerRequest,
    registerResult,
    reconnectDelay,
    rememberRegisterRequest,
    reply: replyJsonRpc,
    runtimeIds,
    saveEndpointOverride: storeEndpointOverride,
    sessionAddedNotification,
    sendJsonRpc,
    validateEndpoint
  } = window.OfficeCtlAddinChannel;
  const { AddinLogger } = window.OfficeCtlLogger;
  const { TaskHistoryStore } = window.OfficeCtlTaskHistory;

  const PLANNED_TOOLS = [
    'powerpoint.add_slide',
    'powerpoint.replace_text',
    'powerpoint.insert_image',
    'powerpoint.apply_layout',
    'powerpoint.export_pdf'
  ];
  const TOOL_GROUPS = [
    { label: 'Slides', tools: ['powerpoint.add_slide', 'powerpoint.apply_layout'] },
    { label: 'Content', tools: ['powerpoint.replace_text', 'powerpoint.insert_image'] },
    { label: 'Export', tools: ['powerpoint.export_pdf'] }
  ];
  const TOOL_METADATA = new Map([
    ['powerpoint.add_slide', { category: 'Slides', sideEffect: 'mutating', description: 'Add a slide to the current presentation.' }],
    ['powerpoint.replace_text', { category: 'Content', sideEffect: 'mutating', description: 'Replace matching text in slide content.' }],
    ['powerpoint.insert_image', { category: 'Content', sideEffect: 'mutating', description: 'Insert an image on a slide.' }],
    ['powerpoint.apply_layout', { category: 'Slides', sideEffect: 'mutating', description: 'Apply a layout to a slide.' }],
    ['powerpoint.export_pdf', { category: 'Export', sideEffect: 'read', description: 'Export the presentation to PDF when host support is available.' }]
  ]);

  const { instanceId, sessionId } = runtimeIds();
  const logger = new AddinLogger({ redactText });
  const taskStore = new TaskHistoryStore({ redactText });
  let socket;
  let reconnectTimer;
  let reconnectAttempt = 0;
  let endpointDirty = false;
  let suppressNextSettingsClick = false;
  let serverInfo = { serverVersion: 'Unknown', protocolVersion: PROTOCOL_VERSION };
  let documentInfo = null;
  let sessionAnnounced = false;

  const connectionBadgeEl = document.getElementById('connectionBadge');
  const sessionEl = document.getElementById('session');
  const daemonEl = document.getElementById('daemon');
  const serverVersionEl = document.getElementById('serverVersion');
  const protocolVersionEl = document.getElementById('protocolVersion');
  const hostPlatformEl = document.getElementById('hostPlatform');
  const documentTitleEl = document.getElementById('documentTitle');
  const protectionEl = document.getElementById('protection');
  const documentStateEl = document.getElementById('documentState');
  const connectionDetailEl = document.getElementById('connectionDetail');
  const toolCountEl = document.getElementById('toolCount');
  const toolListEl = document.getElementById('toolList');
  const currentTaskEl = document.getElementById('currentTask');
  const currentTaskStateEl = document.getElementById('currentTaskState');
  const historyListEl = document.getElementById('historyList');
  const historyCountEl = document.getElementById('historyCount');
  const settingsToggleEl = document.getElementById('settingsToggle');
  const settingsPanelEl = document.getElementById('settingsPanel');
  const settingsFormEl = document.getElementById('settingsForm');
  const endpointInputEl = document.getElementById('endpointInput');
  const endpointErrorEl = document.getElementById('endpointError');
  const saveEndpointEl = document.getElementById('saveEndpoint');
  const announcerEl = document.getElementById('announcer');

  settingsToggleEl.addEventListener('click', handleSettingsClick);
  settingsToggleEl.addEventListener('keydown', activateSettingsWithKeyboard);
  settingsFormEl.addEventListener('submit', saveEndpointOverride);
  endpointInputEl.addEventListener('input', () => {
    endpointDirty = endpointInputEl.value.trim() !== configuredEndpoint();
  });
  window.addEventListener('beforeunload', (event) => {
    if (!endpointDirty) return;
    event.preventDefault();
    event.returnValue = '';
  });
  endpointInputEl.value = configuredEndpoint();
  renderStaticState();
  window.__OFFICE_MCP_TASKPANE_READY__ = true;

  whenOfficeReady((info) => {
    if (!isPowerPointHost(info)) {
      setStatus('Unsupported host');
      return;
    }
    setConnectionState('connecting', 'Connecting...');
    connect();
  });

  function whenOfficeReady(callback) {
    if (!window.Office || typeof Office.onReady !== 'function') {
      setStatus('Unsupported host');
      return;
    }
    Office.onReady(callback);
  }

  function isPowerPointHost(info) {
    const host = String(info?.host || '').toLowerCase();
    const expected = String(Office.HostType?.PowerPoint || 'PowerPoint').toLowerCase();
    const diagnosticsHost = String(Office.context?.diagnostics?.host || '').toLowerCase();
    const hasPowerPointRuntime = typeof window.PowerPoint?.run === 'function';
    const hasPowerPointRequirementSet = Office.context?.requirements?.isSetSupported?.('PowerPointApi', '1.1') === true;
    return host === expected || host === 'powerpoint' || diagnosticsHost === 'powerpoint' || hasPowerPointRuntime || hasPowerPointRequirementSet;
  }

  function connect() {
    clearTimeout(reconnectTimer);
    const endpoint = configuredEndpoint();
    daemonEl.textContent = endpoint;
    endpointInputEl.value = endpoint;
    endpointDirty = false;
    setConnectionState('connecting', 'Connecting...');
    socket = new WebSocket(endpoint);
    socket.addEventListener('open', () => register());
    socket.addEventListener('message', (event) => handleMessage(event.data));
    socket.addEventListener('close', () => {
      logger.warn('websocket.closed');
      if (tryCurrentOriginEndpointFallback(endpoint)) return;
      scheduleReconnect();
    });
    socket.addEventListener('error', () => {
      if (tryCurrentOriginEndpointFallback(endpoint)) return;
      connectionDetailEl.textContent = 'Connection failed. Check that the local daemon is running and the endpoint uses wss://localhost.';
      setConnectionState('failed', 'Failed');
    });
  }

  function tryCurrentOriginEndpointFallback(failedEndpoint) {
    const originEndpoint = currentOriginEndpoint();
    if (failedEndpoint === originEndpoint) return false;
    const nextEndpoint = clearEndpointOverride();
    logger.warn('websocket.fallback_to_manifest_origin', { failedEndpoint, nextEndpoint });
    connectionDetailEl.textContent = `Endpoint ${failedEndpoint} failed. Retrying ${nextEndpoint}.`;
    setConnectionState('reconnecting', 'Reconnecting...');
    reconnectAttempt = 0;
    clearTimeout(reconnectTimer);
    reconnectTimer = setTimeout(connect, 250);
    return true;
  }

  function register() {
    reconnectAttempt = 0;
    setConnectionState('connecting', 'Registering...');
    const requestId = crypto.randomUUID();
    send(registerRequest(requestId, {
      instance_id: instanceId,
      host: {
        app: 'powerpoint',
        version: Office.context.diagnostics?.version || null,
        platform: String(Office.context.platform || 'unknown').toLowerCase(),
        build: Office.context.diagnostics?.host || 'Desktop'
      },
      add_in: {
        version: ADDIN_VERSION,
        protocol_version: PROTOCOL_VERSION,
        requirement_sets: probeRequirementSets(),
        supported_features: ['presentation.session']
      }
    }));
    rememberRegisterRequest(requestId);
  }

  async function announceSession() {
    const presentation = await getPresentationInfo();
    documentInfo = presentation;
    logger.info('session.added', { sessionId, presentation });
    send(sessionAddedNotification({
      session_id: sessionId,
      instance_id: instanceId,
      document: presentation,
      available_tools: [],
      is_active: null
    }));
    sessionEl.textContent = sessionId;
    sessionAnnounced = true;
    renderDocumentState();
    setConnectionState('connected', 'Connected');
  }

  function handleMessage(raw) {
    const message = parseJsonRpc(raw);
    if (!message) return;
    if (isRegisterResponse(message)) {
      handleRegisterResponse(message);
      return;
    }
    if (isToolInvoke(message)) {
      invokeTool(message).catch((error) => reply(message.id, {
        ok: false,
        error: mapError(error, message.params?.tool),
        elapsed_ms: 0
      }));
    } else if (isPing(message)) {
      reply(message.id, { ts: new Date().toISOString() });
    } else if (isToolCancel(message)) {
      taskStore.cancel(message.params.request_id);
      renderCurrentTask();
    }
  }

  function handleRegisterResponse(message) {
    clearRegisterRequest();
    if (message.error) {
      const error = message.error.data || message.error;
      connectionDetailEl.textContent = error.message || message.error.message || 'Registration failed. Check the daemon endpoint and protocol version.';
      setConnectionState('failed', 'Failed');
      return;
    }
    serverInfo = registerResult(message, PROTOCOL_VERSION);
    serverVersionEl.textContent = `Server ${serverInfo.serverVersion}`;
    protocolVersionEl.textContent = `Protocol ${serverInfo.protocolVersion}`;
    connectionDetailEl.textContent = 'None';
    enableAutoOpen().then(() => announceSession()).catch((error) => {
      logger.error('session.announce.failed', error);
      connectionDetailEl.textContent = error.message || 'Failed to announce presentation session.';
      setConnectionState('failed', 'Failed');
    });
  }

  function enableAutoOpen() {
    return new Promise((resolve) => {
      try {
        if (!Office.context.document?.settings?.set || !Office.context.document?.settings?.saveAsync) {
          resolve();
          return;
        }
        Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
        Office.context.document.settings.saveAsync(() => resolve());
      } catch (error) {
        logger.warn('autoopen.failed', error);
        resolve();
      }
    });
  }

  async function invokeTool(message) {
    const started = performance.now();
    const requestId = message.params?.request_id || String(message.id);
    const tool = message.params?.tool || 'powerpoint.unknown';
    taskStore.start(requestId, tool, message.params || {}, message.params?.timeout_ms);
    renderCurrentTask();
    const error = Object.assign(new Error(`${tool} is declared by the daemon contract but is not implemented in the PowerPoint add-in yet.`), {
      officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE',
      partialEffect: 'none'
    });
    taskStore.finish(requestId, 'failure', Math.round(performance.now() - started), mapError(error, tool));
    renderCurrentTask();
    renderHistory();
    throw error;
  }

  async function getPresentationInfo() {
    const title = Office.context.document?.url ? fileName(Office.context.document.url) : 'PowerPoint Presentation';
    return {
      title,
      filename: title,
      is_dirty: null,
      is_read_only: false,
      protection: { kind: 'none', label: 'Not protected' },
      host: {
        app: 'powerpoint',
        platform: String(Office.context.platform || 'unknown').toLowerCase(),
        version: Office.context.diagnostics?.version || null
      }
    };
  }

  function probeRequirementSets() {
    const sets = [];
    for (const version of ['1.1', '1.2', '1.3']) {
      try {
        if (Office.context?.requirements?.isSetSupported?.('PowerPointApi', version)) sets.push(`PowerPointApi ${version}`);
      } catch (error) {
        logger.warn('requirements.probe.failed', { version, error });
      }
    }
    return sets;
  }

  function renderStaticState() {
    sessionEl.textContent = sessionId;
    daemonEl.textContent = configuredEndpoint();
    serverVersionEl.textContent = `Server ${serverInfo.serverVersion}`;
    protocolVersionEl.textContent = `Protocol ${serverInfo.protocolVersion}`;
    renderToolSummary();
    renderCurrentTask();
    renderHistory();
  }

  function renderToolSummary() {
    toolCountEl.textContent = `Available 0 of ${PLANNED_TOOLS.length}`;
    toolListEl.innerHTML = TOOL_GROUPS.map((group) => {
      const rows = group.tools.map((tool) => toolRowMarkup(tool)).join('');
      return `<details class="tool-group" open><summary class="tool-group-title"><span>${escapeHtml(group.label)}</span><span>0 of ${group.tools.length}</span></summary>${rows}</details>`;
    }).join('');
  }

  function toolRowMarkup(tool) {
    const meta = TOOL_METADATA.get(tool) || { sideEffect: 'unknown', description: 'No metadata.' };
    return `<div class="tool-permission-row is-disabled"><div><strong>${escapeHtml(tool)}</strong><span>${escapeHtml(meta.description)}</span></div><span class="tool-side-effect">${escapeHtml(titleCase(meta.sideEffect))}</span></div>`;
  }

  function renderDocumentState() {
    if (!documentInfo) return;
    documentTitleEl.textContent = documentInfo.title || documentInfo.filename || 'PowerPoint Presentation';
    hostPlatformEl.textContent = `PowerPoint / ${titleCase(documentInfo.host?.platform || 'unknown')}`;
    protectionEl.textContent = protectionLabel(documentInfo);
    documentStateEl.textContent = documentStateLabel(documentInfo);
  }

  function protectionLabel(info) {
    const label = info?.protection?.label || info?.protection?.kind;
    if (!label || String(label).toLowerCase() === 'none') return 'Not protected';
    return titleCase(String(label));
  }

  function documentStateLabel(info) {
    if (info?.is_read_only === true) return 'Read-only';
    return 'Editable';
  }

  function renderCurrentTask() {
    const { currentTask } = taskStore.snapshot();
    if (!currentTask) {
      currentTaskStateEl.textContent = 'Idle';
      currentTaskEl.className = 'empty-state';
      currentTaskEl.textContent = 'No command is running.';
      return;
    }
    currentTaskStateEl.textContent = 'Running';
    currentTaskEl.className = 'task-card';
    currentTaskEl.innerHTML = taskMarkup(currentTask);
  }

  function renderHistory() {
    const { history, historyLimit } = taskStore.snapshot();
    historyCountEl.textContent = `${history.length} / ${historyLimit}`;
    historyListEl.innerHTML = history.map((task) => `<li>${taskMarkup(task)}</li>`).join('');
  }

  function taskMarkup(task) {
    const status = task.status ? titleCase(task.status) : 'Running';
    const elapsed = typeof task.elapsedMs === 'number' ? formatDuration(task.elapsedMs) : 'in progress';
    const error = task.error ? `<div class="task-error">${escapeHtml(task.error.message || 'Command failed.')}</div>` : '';
    return `<article class="task-card"><div class="task-title"><span>${escapeHtml(task.tool)}</span><span>${escapeHtml(status)}</span></div><div class="task-meta">${escapeHtml(formatTime(task.startedAt))} / ${escapeHtml(elapsed)}</div>${error}</article>`;
  }

  function scheduleReconnect() {
    reconnectAttempt += 1;
    const delay = reconnectDelay(reconnectAttempt);
    const seconds = Math.round(delay / 1000);
    setConnectionState('reconnecting', `Reconnecting in ${seconds}s`);
    reconnectTimer = setTimeout(connect, delay);
  }

  function setStatus(label) {
    connectionDetailEl.textContent = label;
    setConnectionState('failed', label);
  }

  function setConnectionState(state, label) {
    connectionBadgeEl.textContent = label;
    connectionBadgeEl.className = 'status-badge';
    connectionBadgeEl.classList.add(state === 'connected' ? 'status-success' : state === 'failed' ? 'status-danger' : state === 'reconnecting' ? 'status-warning' : 'status-neutral');
    announcerEl.textContent = label;
  }

  function reply(id, result) {
    send(replyJsonRpc(id, result));
  }

  function send(message) {
    if (!sendJsonRpc(socket, message)) {
      connectionDetailEl.textContent = 'Daemon connection is not open.';
      setConnectionState('failed', 'Failed');
    }
  }

  function mapError(error, tool) {
    return {
      office_mcp_code: error.officeMcpCode || 'HOST_ERROR',
      message: error.message || String(error),
      tool,
      retriable: error.officeMcpCode === 'HOST_CAPABILITY_UNAVAILABLE',
      partial_effect: error.partialEffect || 'unknown'
    };
  }

  function handleSettingsClick(event) {
    if (suppressNextSettingsClick) {
      suppressNextSettingsClick = false;
      return;
    }
    toggleSettings(event);
  }

  function activateSettingsWithKeyboard(event) {
    if (event.key !== 'Enter' && event.key !== ' ') return;
    suppressNextSettingsClick = true;
    toggleSettings(event);
  }

  function toggleSettings(event) {
    event.preventDefault();
    const opening = settingsPanelEl.hidden;
    if (!opening && endpointDirty && !confirm('Discard unsaved endpoint changes?')) return;
    settingsPanelEl.hidden = !opening;
    settingsToggleEl.setAttribute('aria-expanded', String(opening));
    settingsToggleEl.setAttribute('aria-label', opening ? 'Close Settings' : 'Open Settings');
    if (opening) endpointInputEl.focus();
    if (!opening) {
      endpointInputEl.value = configuredEndpoint();
      endpointDirty = false;
      endpointErrorEl.textContent = '';
    }
  }

  function saveEndpointOverride(event) {
    event.preventDefault();
    const value = endpointInputEl.value.trim();
    const validation = validateEndpoint(value);
    if (!validation.ok) {
      endpointErrorEl.textContent = validation.message;
      return;
    }
    saveEndpointEl.disabled = true;
    saveEndpointEl.textContent = 'Saving...';
    try {
      storeEndpointOverride(value);
      endpointDirty = false;
      endpointErrorEl.textContent = '';
      settingsPanelEl.hidden = true;
      settingsToggleEl.setAttribute('aria-expanded', 'false');
      settingsToggleEl.setAttribute('aria-label', 'Open Settings');
      connect();
    } finally {
      saveEndpointEl.disabled = false;
      saveEndpointEl.textContent = 'Save Endpoint';
    }
  }
})();
