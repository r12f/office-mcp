(() => {
  const ADDIN_VERSION = '0.1.18';
  const PROTOCOL_VERSION = '1.0';
  const {
    boolLabel,
    escapeHtml,
    fileName,
    formatDuration,
    formatTime,
    redactText,
    titleCase
  } = window.OfficeCtlCommon;
  const {
    clearEndpointOverride,
    clearRegisterRequest,
    configuredEndpoint,
    createRequestId,
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
    sessionUpdatedNotification,
    sendJsonRpc,
    validateEndpoint
  } = window.OfficeCtlAddinChannel;
  const { AddinLogger } = window.OfficeCtlLogger;
  const { TaskHistoryStore } = window.OfficeCtlTaskHistory;
  const {
    bindDetailsControl,
    commandIdMarkup,
    copyMetadataValue,
    documentStateLabel,
    isToolAllowedByCapabilityMode,
    middleTruncate,
    protectionLabel,
    renderStaticMetadata,
    renderToolModeControl: renderSharedToolModeControl,
    setConnectionState: setSharedConnectionState,
    setCopyableMetadata,
    statusClass,
    taskMetadataMarkup,
    taskStatusClass,
    taskStatusLabel
  } = window.OfficeCtlMainUi;
  const CONNECT_TIMEOUT_MS = 8000;
  const INSERT_IMAGE_PLACEMENTS = new Set([
    'inline',
    'before_paragraph',
    'after_paragraph',
    'new_paragraph_before',
    'new_paragraph_after',
    'replace_paragraph',
    'selection'
  ]);
  const WORD_MUTATING_TOOLS = new Set([
    'word.insert_paragraph',
    'word.insert_table',
    'word.insert_image',
    'word.resize_image',
    'word.insert_break',
    'word.insert_page_break',
    'word.update_page_setup',
    'word.update_header_footer',
    'word.insert_list',
    'word.replace_text',
    'word.update_paragraph',
    'word.delete_range',
    'word.insert_bookmark',
    'word.delete_bookmark',
    'word.apply_formatting',
    'word.update_table',
    'word.insert_content_control',
    'word.update_content_control',
    'word.delete_content_control',
    'word.apply_style',
    'word.add_comment',
    'word.resolve_comment',
    'word.update_tracked_change',
    'word.save'
  ]);
  const AVAILABLE_TOOLS = [
    'word.get_text',
    'word.get_outline',
    'word.get_paragraph',
    'word.find_text',
    'word.resolve_anchor',
    'word.insert_bookmark',
    'word.list_bookmarks',
    'word.delete_bookmark',
    'word.get_selection',
    'word.get_header_footer',
    'word.insert_paragraph',
    'word.insert_image',
    'word.resize_image',
    'word.insert_table',
    'word.update_header_footer',
    'word.insert_break',
    'word.list_sections',
    'word.update_page_setup',
    'word.insert_list',
    'word.replace_text',
    'word.update_paragraph',
    'word.delete_range',
    'word.apply_formatting',
    'word.read_table',
    'word.update_table',
    'word.list_content_controls',
    'word.insert_content_control',
    'word.update_content_control',
    'word.delete_content_control',
    'word.apply_style',
    'word.add_comment',
    'word.resolve_comment',
    'word.update_tracked_change',
    'word.save'
  ];
  const TOOL_GROUPS = [
    { label: 'Document & structure', tools: ['word.get_text', 'word.get_outline', 'word.get_header_footer', 'word.update_header_footer', 'word.insert_break', 'word.list_sections', 'word.update_page_setup', 'word.save'] },
    { label: 'Range & selection', tools: ['word.get_selection', 'word.find_text', 'word.resolve_anchor', 'word.insert_bookmark', 'word.list_bookmarks', 'word.delete_bookmark', 'word.replace_text', 'word.delete_range', 'word.apply_formatting', 'word.apply_style'] },
    { label: 'Paragraphs & lists', tools: ['word.get_paragraph', 'word.insert_paragraph', 'word.update_paragraph', 'word.insert_list'] },
    { label: 'Tables', tools: ['word.read_table', 'word.update_table'] },
    { label: 'Media', tools: ['word.insert_image', 'word.resize_image'] },
    { label: 'Content controls', tools: ['word.list_content_controls', 'word.insert_content_control', 'word.update_content_control', 'word.delete_content_control'] },
    { label: 'Review', tools: ['word.add_comment', 'word.resolve_comment', 'word.update_tracked_change'] }
  ];
  const TOOL_METADATA = new Map([
    ['word.get_text', { category: 'Document & structure', sideEffect: 'read', description: 'Read document text by paragraph range.' }],
    ['word.get_outline', { category: 'Document & structure', sideEffect: 'read', description: 'Read heading outline and structure.' }],
    ['word.get_header_footer', { category: 'Document & structure', sideEffect: 'read', description: 'Read section header or footer text.' }],
    ['word.get_paragraph', { category: 'Paragraphs & lists', sideEffect: 'read', description: 'Read a single paragraph by index.' }],
    ['word.find_text', { category: 'Range & selection', sideEffect: 'read', description: 'Find text matches in the document body.' }],
    ['word.resolve_anchor', { category: 'Range & selection', sideEffect: 'read', description: 'Resolve an anchor to safe diagnostic metadata.' }],
    ['word.insert_bookmark', { category: 'Range & selection', sideEffect: 'mutating', description: 'Create a named bookmark around an anchored range.' }],
    ['word.list_bookmarks', { category: 'Range & selection', sideEffect: 'read', description: 'List bookmark names and locations.' }],
    ['word.delete_bookmark', { category: 'Range & selection', sideEffect: 'destructive', description: 'Delete a bookmark marker without deleting text.' }],
    ['word.get_selection', { category: 'Range & selection', sideEffect: 'read', description: 'Read the current selection.' }],
    ['word.insert_paragraph', { category: 'Paragraphs & lists', sideEffect: 'mutating', description: 'Insert a paragraph near an anchor.' }],
    ['word.insert_image', { category: 'Media', sideEffect: 'mutating', description: 'Insert an image into the document.' }],
    ['word.resize_image', { category: 'Media', sideEffect: 'mutating', description: 'Resize an existing inline image.' }],
    ['word.insert_table', { category: 'Tables', sideEffect: 'mutating', description: 'Insert a table with provided values.' }],
    ['word.insert_break', { category: 'Document & structure', sideEffect: 'mutating', description: 'Insert a page, line, or section break.' }],
    ['word.insert_page_break', { category: 'Document & structure', sideEffect: 'mutating', description: 'Insert a page break.' }],
    ['word.list_sections', { category: 'Document & structure', sideEffect: 'read', description: 'List document sections.' }],
    ['word.update_page_setup', { category: 'Document & structure', sideEffect: 'mutating', description: 'Update document or section page setup.' }],
    ['word.update_header_footer', { category: 'Document & structure', sideEffect: 'destructive', description: 'Replace, append, or clear a section header or footer.' }],
    ['word.insert_list', { category: 'Paragraphs & lists', sideEffect: 'mutating', description: 'Insert a list.' }],
    ['word.replace_text', { category: 'Range & selection', sideEffect: 'mutating', description: 'Replace matching document text.' }],
    ['word.update_paragraph', { category: 'Paragraphs & lists', sideEffect: 'mutating', description: 'Update paragraph text and style.' }],
    ['word.delete_range', { category: 'Range & selection', sideEffect: 'mutating', description: 'Delete text resolved from an anchor.' }],
    ['word.apply_formatting', { category: 'Range & selection', sideEffect: 'mutating', description: 'Apply formatting to an anchored range.' }],
    ['word.read_table', { category: 'Tables', sideEffect: 'read', description: 'Read table dimensions and cell values.' }],
    ['word.update_table', { category: 'Tables', sideEffect: 'destructive', description: 'Update table cells, rows, columns, formatting, or lifecycle.' }],
    ['word.list_content_controls', { category: 'Content controls', sideEffect: 'read', description: 'List content-control metadata.' }],
    ['word.insert_content_control', { category: 'Content controls', sideEffect: 'mutating', description: 'Create a content control around an anchored range.' }],
    ['word.update_content_control', { category: 'Content controls', sideEffect: 'mutating', description: 'Update content-control metadata, locks, or text.' }],
    ['word.delete_content_control', { category: 'Content controls', sideEffect: 'destructive', description: 'Delete a content control with explicit content handling.' }],
    ['word.apply_style', { category: 'Range & selection', sideEffect: 'mutating', description: 'Apply an Office style to an anchored range.' }],
    ['word.add_comment', { category: 'Review', sideEffect: 'mutating', description: 'Add a comment to an anchored range.' }],
    ['word.resolve_comment', { category: 'Review', sideEffect: 'mutating', description: 'Resolve an existing comment.' }],
    ['word.update_tracked_change', { category: 'Review', sideEffect: 'destructive', description: 'Accept or reject a tracked change by fingerprint.' }],
    ['word.save', { category: 'Document & structure', sideEffect: 'mutating', description: 'Save the current document.' }]
  ]);
  let socket;
  let connectGeneration = 0;
  let connectTimeoutTimer;
  const { instanceId, sessionId } = runtimeIds();
  let runtimeInstanceId = instanceId;
  const TOOL_PERMISSION_STORAGE_KEY = `office-mcp.word.tool-permissions.${sessionId}`;
  const TOOL_PERMISSION_MODE_STORAGE_KEY = `office-mcp.word.tool-permission-mode.${sessionId}`;
  let documentInfo = null;
  let serverInfo = { serverVersion: 'Unknown', protocolVersion: PROTOCOL_VERSION };
  let reconnectTimer;
  let reconnectAttempt = 0;
  let endpointDirty = false;
  let toolPermissions = loadToolPermissions();
  let toolPermissionMode = loadToolPermissionMode();
  let sessionAnnounced = false;
  const logger = new AddinLogger({ redactText });
  const taskStore = new TaskHistoryStore({ redactText });

  window.addEventListener('error', (event) => {
    recordDiagnostic('runtime.error', {
      message: event.message,
      source: event.filename,
      line: event.lineno,
      column: event.colno
    });
  });
  window.addEventListener('unhandledrejection', (event) => {
    recordDiagnostic('runtime.unhandled_rejection', {
      message: event.reason?.message || String(event.reason || '')
    });
  });

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
  const toolModeControlEl = document.getElementById('toolModeControl');
  const currentTaskEl = document.getElementById('currentTask');
  const currentTaskStateEl = document.getElementById('currentTaskState');
  const historyListEl = document.getElementById('historyList');
  const historyCountEl = document.getElementById('historyCount');
  const settingsFormEl = document.getElementById('settingsForm');
  const endpointInputEl = document.getElementById('endpointInput');
  const saveEndpointEl = document.getElementById('saveEndpoint');
  const announcerEl = document.getElementById('announcer');

  settingsFormEl.addEventListener('submit', saveEndpointOverride);
  toolModeControlEl?.querySelectorAll('[data-tool-mode]').forEach((button) => button.addEventListener('click', handleToolModeChange));
  document.addEventListener('click', handleMetadataCopy);
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
  recordDiagnostic('taskpane.ready', { endpoint: configuredEndpoint() });

  whenOfficeReady(async (info) => {
    recordDiagnostic('office.ready', {
      host: info?.host,
      platform: Office.context?.platform,
      diagnosticsHost: Office.context?.diagnostics?.host,
      diagnosticsVersion: Office.context?.diagnostics?.version
    });
    if (!isWordHost(info)) {
      setStatus('Unsupported host');
      return;
    }
    if (Office.context.platform !== Office.PlatformType.PC) {
      setStatus('Unsupported platform');
      return;
    }
    setConnectionState('connecting', 'Connecting…');
    connect();
  });

  function whenOfficeReady(callback) {
    if (!window.Office || typeof Office.onReady !== 'function') {
      setStatus('Unsupported host');
      return;
    }
    Office.onReady(callback);
  }

  function isWordHost(info) {
    const host = String(info?.host || '').toLowerCase();
    const expected = String(Office.HostType?.Word || 'Word').toLowerCase();
    const diagnosticsHost = String(Office.context?.diagnostics?.host || '').toLowerCase();
    const hasWordRuntime = typeof window.Word?.run === 'function';
    const hasWordRequirementSet = Office.context?.requirements?.isSetSupported?.('WordApi', '1.3') === true;
    return host === expected || host === 'word' || diagnosticsHost === 'word' || hasWordRuntime || hasWordRequirementSet;
  }

  function connect() {
    clearTimeout(reconnectTimer);
    clearTimeout(connectTimeoutTimer);
    const generation = ++connectGeneration;
    const endpoint = configuredEndpoint();
    recordDiagnostic('websocket.connecting', { endpoint, generation });
    setCopyableMetadata(daemonEl, endpoint);
    endpointInputEl.value = endpoint;
    endpointDirty = false;
    setConnectionState('connecting', 'Connecting…');
    let failureHandled = false;
    const handleConnectionFailure = (message) => {
      if (failureHandled || generation !== connectGeneration) return;
      failureHandled = true;
      clearTimeout(connectTimeoutTimer);
      connectionDetailEl.textContent = message;
      setConnectionState('failed', 'Failed');
      scheduleReconnect();
    };
    try {
      socket = new WebSocket(endpoint);
    } catch (error) {
      logger.error('websocket.create.failed', error);
      recordDiagnostic('websocket.create_failed', { endpoint, message: error.message });
      if (tryCurrentOriginEndpointFallback(endpoint)) return;
      handleConnectionFailure(error.message || 'Connection failed before the daemon socket could open.');
      return;
    }
    connectTimeoutTimer = setTimeout(() => {
      logger.warn('websocket.open.timeout', { endpoint, timeoutMs: CONNECT_TIMEOUT_MS });
      recordDiagnostic('websocket.open_timeout', { endpoint, timeoutMs: CONNECT_TIMEOUT_MS });
      try {
        socket?.close(1000, 'Open timeout');
      } catch {
        // Nothing else can be recovered from a socket that never opened.
      }
      handleConnectionFailure('Connection timed out before the daemon socket opened. Check the daemon log and Office WebView runtime.');
    }, CONNECT_TIMEOUT_MS);
    socket.addEventListener('open', () => {
      if (generation !== connectGeneration) return;
      clearTimeout(connectTimeoutTimer);
      recordDiagnostic('websocket.open', { endpoint, generation });
      register();
    });
    socket.addEventListener('message', (event) => handleMessage(event.data));
    socket.addEventListener('close', (event) => {
      if (generation !== connectGeneration) return;
      clearTimeout(connectTimeoutTimer);
      logger.warn('websocket.closed', { code: event.code, reason: event.reason, wasClean: event.wasClean });
      recordDiagnostic('websocket.closed', { endpoint, code: event.code, reason: event.reason, wasClean: event.wasClean });
      if (tryCurrentOriginEndpointFallback(endpoint)) return;
      if (!sessionAnnounced) {
        handleConnectionFailure('Connection closed before the document registered with the daemon.');
        return;
      }
      scheduleReconnect();
    });
    socket.addEventListener('error', () => {
      if (generation !== connectGeneration) return;
      recordDiagnostic('websocket.error', { endpoint });
      if (tryCurrentOriginEndpointFallback(endpoint)) return;
      handleConnectionFailure('Connection failed. Check that the local daemon is running and the endpoint uses wss://localhost.');
    });
  }

  function tryCurrentOriginEndpointFallback(failedEndpoint) {
    const originEndpoint = currentOriginEndpoint();
    if (failedEndpoint === originEndpoint) return false;
    const nextEndpoint = clearEndpointOverride();
    logger.warn('websocket.fallback_to_manifest_origin', { failedEndpoint, nextEndpoint });
    connectionDetailEl.textContent = `Endpoint ${failedEndpoint} failed. Retrying ${nextEndpoint}.`;
    setConnectionState('reconnecting', 'Reconnecting…');
    reconnectAttempt = 0;
    clearTimeout(reconnectTimer);
    reconnectTimer = setTimeout(connect, 250);
    return true;
  }

  function register() {
    reconnectAttempt = 0;
    setConnectionState('connecting', 'Registering…');
    const requestId = createRequestId();
    send(registerRequest(requestId, {
      instance_id: instanceId,
      host: {
        app: 'word',
        version: Office.context.diagnostics?.version || null,
        platform: String(Office.context.platform || 'unknown').toLowerCase(),
        build: Office.context.diagnostics?.host || 'Desktop'
      },
      add_in: {
        version: ADDIN_VERSION,
        protocol_version: PROTOCOL_VERSION,
        requirement_sets: probeRequirementSets(),
        supported_features: ['doc.read', 'doc.write']
      }
    }));
    rememberRegisterRequest(requestId);
  }

  async function announceSession() {
    const document = await getDocumentInfo();
    documentInfo = document;
    await enableDocumentAutoOpen();
    logger.info('session.added', { sessionId, document });
    send(sessionAddedNotification({
      session_id: sessionId,
      instance_id: runtimeInstanceId,
      document,
      available_tools: effectiveTools(),
      is_active: null
    }));
    setCopyableMetadata(sessionEl, sessionId);
    sessionAnnounced = true;
    renderDocumentState();
    setConnectionState('connected', 'Connected');
  }

  function enableDocumentAutoOpen() {
    return new Promise((resolve) => {
      try {
        Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
        Office.context.document.settings.saveAsync(() => resolve());
      } catch {
        resolve();
      }
    });
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
        error: mapError(error, message.params?.tool, message.params?.args),
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
    runtimeInstanceId = serverInfo.assignedInstanceId || instanceId;
    renderStaticState();
    announceSession().catch((error) => {
      logger.error('session.announce.failed', error);
      connectionDetailEl.textContent = error.message || 'Failed to announce document session.';
      setConnectionState('failed', 'Failed');
    });
  }

  async function invokeTool(message) {
    const started = performance.now();
    const requestId = message.params?.request_id || String(message.id);
    const { tool, args } = message.params;
    startTask(requestId, tool, message.params || {}, message.params.timeout_ms);
    let data;
    try {
      if (!isToolEnabled(tool)) throw toolDisabledError(tool);
      if (taskStore.isCancelled(requestId)) throw cancelledError(tool);
      preflightWordMutatingTool(tool, args || {});
      if (args?.validate_only) data = await validateWordMutationOnly(tool, args || {});
      else
      switch (tool) {
        case 'word.get_text':
          data = await getText(args);
          break;
        case 'word.get_paragraph':
          data = await getParagraph(args);
          break;
        case 'word.get_outline':
          data = await getOutline(args);
          break;
        case 'word.find_text':
          data = await findText(args);
          break;
        case 'word.resolve_anchor':
          data = await resolveAnchorTool(args);
          break;
        case 'word.insert_bookmark':
          data = await insertBookmark(args);
          break;
        case 'word.list_bookmarks':
          data = await listBookmarks(args || {});
          break;
        case 'word.delete_bookmark':
          data = await deleteBookmark(args);
          break;
        case 'word.get_selection':
          data = await getSelection(args);
          break;
        case 'word.get_header_footer':
          data = await getHeaderFooter(args);
          break;
        case 'word.insert_paragraph':
          data = await insertParagraph(args);
          break;
        case 'word.insert_table':
          data = await insertTable(args);
          break;
        case 'word.insert_image':
          data = await insertImage(args);
          break;
        case 'word.resize_image':
          data = await resizeImage(args);
          break;
        case 'word.insert_break':
          data = await insertBreak(args);
          break;
        case 'word.insert_page_break':
          data = await insertPageBreak(args);
          break;
        case 'word.list_sections':
          data = await listSections(args || {});
          break;
        case 'word.update_page_setup':
          data = await updatePageSetup(args || {});
          break;
        case 'word.update_header_footer':
          data = args?.validate_only ? await validateWordMutationOnly(tool, args) : await updateHeaderFooter(args);
          break;
        case 'word.insert_list':
          data = await insertList(args);
          break;
        case 'word.replace_text':
          data = await replaceText(args);
          break;
        case 'word.update_paragraph':
          data = await updateParagraph(args);
          break;
        case 'word.delete_range':
          data = await deleteRange(args);
          break;
        case 'word.apply_formatting':
          data = await applyFormatting(args);
          break;
        case 'word.read_table':
          data = await readTable(args);
          break;
        case 'word.update_table':
          data = await updateTable(args);
          break;
        case 'word.list_content_controls':
          data = await listContentControls(args);
          break;
        case 'word.insert_content_control':
          data = await insertContentControl(args);
          break;
        case 'word.update_content_control':
          data = await updateContentControl(args);
          break;
        case 'word.delete_content_control':
          data = await deleteContentControl(args);
          break;
        case 'word.apply_style':
          data = await applyStyle(args);
          break;
        case 'word.add_comment':
          data = await addComment(args);
          break;
        case 'word.resolve_comment':
          data = await resolveComment(args);
          break;
        case 'word.update_tracked_change':
          data = await updateTrackedChange(args);
          break;
        case 'word._get_comments':
          data = await getComments(args);
          break;
        case 'word._get_tracked_changes':
          data = await getTrackedChanges(args);
          break;
        case 'word._get_structure':
          data = await getStructure(args);
          break;
        case 'word.save':
          data = await saveDocument(args);
          break;
        default:
          throw Object.assign(new Error(`Unsupported tool ${tool}`), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE' });
      }
      if (taskStore.consumeCancellation(requestId)) throw cancelledError(tool);
      const elapsedMs = Math.round(performance.now() - started);
      finishTask(requestId, 'success', elapsedMs);
      reply(message.id, { ok: true, data, elapsed_ms: elapsedMs });
    } catch (error) {
      const mapped = mapError(error, tool, args);
      finishTask(requestId, mapped.office_mcp_code === 'CANCELLED' ? 'cancelled' : 'failure', Math.round(performance.now() - started), mapped);
      throw error;
    }
  }

  function cancelledError(tool) {
    return Object.assign(new Error(`Tool ${tool} was cancelled.`), { officeMcpCode: 'CANCELLED', partialEffect: 'unknown' });
  }

  function toolDisabledError(tool) {
    return Object.assign(new Error(`Tool ${tool} is disabled by task pane settings.`), {
      officeMcpCode: 'TOOL_DISABLED_BY_USER',
      partialEffect: 'none'
    });
  }

  function preflightWordMutatingTool(tool, args) {
    if (!WORD_MUTATING_TOOLS.has(tool)) return;
    switch (tool) {
      case 'word.insert_paragraph':
        requireAnchor(tool, args.anchor);
        validateHeadingLevelArg(tool, args.heading_level, true);
        validateFormattingArg(tool, args.formatting);
        break;
      case 'word.insert_table':
        requireAnchor(tool, args.anchor);
        requirePositiveInteger(tool, 'rows', args.rows);
        requirePositiveInteger(tool, 'cols', args.cols);
        validateTableData(args.rows, args.cols, args.data);
        break;
      case 'word.insert_image':
        requireAnchor(tool, args.anchor);
        validateInsertImagePreflight(args);
        break;
      case 'word.resize_image':
        validateResizeImageArgs(args);
        break;
      case 'word.insert_break':
        requireAnchor(tool, args.anchor);
        validateBreakType(args.break_type);
        break;
      case 'word.insert_page_break':
        requireAnchor(tool, args.anchor);
        break;
      case 'word.update_page_setup':
        validateUpdatePageSetupArgs(args);
        break;
      case 'word.update_header_footer':
        validateHeaderFooterArgs(tool, args, true);
        break;
      case 'word.insert_list':
        requireAnchor(tool, args.anchor);
        validateInsertListArgs(args);
        break;
      case 'word.replace_text':
        validateReplaceTextArgs(args);
        break;
      case 'word.update_paragraph':
        requireNonNegativeInteger(tool, 'index', args.index);
        break;
      case 'word.delete_range':
        validateExtentToolArgs(tool, args);
        break;
      case 'word.insert_bookmark':
        requireAnchor(tool, args.anchor);
        validateBookmarkName(tool, args.name);
        break;
      case 'word.delete_bookmark':
        validateBookmarkName(tool, args.name, { strictPattern: false });
        break;
      case 'word.apply_formatting':
        validateExtentToolArgs(tool, args);
        validateFormattingArg(tool, args.formatting, true);
        break;
      case 'word.update_table':
        validateUpdateTableArgs(args);
        break;
      case 'word.insert_content_control':
        if (args.anchor !== undefined) requireAnchor(tool, args.anchor);
        validateContentControlArgs(tool, args);
        break;
      case 'word.update_content_control':
        validateContentControlTargetArgs(tool, args);
        validateContentControlArgs(tool, args);
        break;
      case 'word.delete_content_control':
        validateContentControlTargetArgs(tool, args);
        validateDeleteContentControlMode(args.mode);
        break;
      case 'word.apply_style':
        requireAnchor(tool, args.anchor);
        if (!args.style && args.heading_level === undefined) {
          throw invalidArgument('word.apply_style requires style or heading_level.');
        }
        validateHeadingLevelArg(tool, args.heading_level, true);
        break;
      case 'word.add_comment':
        requireAnchor(tool, args.anchor);
        break;
      case 'word.resolve_comment':
        if (!args.comment_id) throw invalidArgument('word.resolve_comment requires comment_id.');
        break;
      case 'word.update_tracked_change':
        requireNonNegativeInteger(tool, 'change_index', args.change_index);
        validateTrackedChangeAction(args.action);
        if (!args.expected_fingerprint) throw invalidArgument('word.update_tracked_change requires expected_fingerprint.');
        break;
      default:
        break;
    }
  }

  function invalidArgument(message) {
    return Object.assign(new Error(message), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  function invalidArgumentWithSuggestion(message, suggestion) {
    return Object.assign(invalidArgument(message), { suggestion });
  }

  async function validateWordMutationOnly(tool, args) {
    switch (tool) {
      case 'word.insert_image':
        return validateInsertImageOnly(args);
      case 'word.replace_text':
        return validateReplaceTextOnly(args);
      case 'word.update_paragraph':
        return validateUpdateParagraphOnly(args);
      case 'word.delete_range':
        return validateDeleteRangeOnly(args);
      case 'word.update_header_footer':
        return validateUpdateHeaderFooterOnly(args);
      default:
        throw invalidArgument(`${tool} does not support validate_only.`);
    }
    return { valid: true, partial_effect: 'none' };
  }

  async function validateInsertImageOnly(args) {
    return Word.run(async (context) => {
      validateInsertImagePlacement(args.anchor, args.placement);
      const resolved = await resolveValidationAnchor(context, args.anchor);
      return validationSuccess('word.insert_image', {
        resolved_target: {
          ...resolved,
          placement: args.placement || 'inline',
          image_mime_type: args.image?.mime_type ?? null,
          image_byte_length: args.image?.byte_length ?? null
        }
      });
    });
  }

  async function validateReplaceTextOnly(args) {
    return Word.run(async (context) => {
      const scope = args.scope || {};
      const searchRoot = scope.selection_only ? context.document.getSelection() : context.document.body;
      const ranges = searchRoot.search(args.find, {
        matchCase: args.match_case ?? false,
        matchWholeWord: args.whole_word ?? false,
        matchWildcards: args.wildcards ?? false
      });
      ranges.load('items/text');
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();

      const limit = args.limit ?? 500;
      const limitedRanges = ranges.items.slice(0, limit);
      const paragraphMatches = mapMatchesToParagraphs(paragraphs.items, args.find, args.match_case ?? false, limit, limitedRanges);
      const filtered = filterReplaceRanges(limitedRanges, paragraphMatches, scope.paragraph_range);
      filtered.skipped_count += Math.max(0, ranges.items.length - limitedRanges.length);
      if (!args.partial_ok && scope.paragraph_range && filtered.skipped_count > 0) {
        throw invalidArgument('replace_text scope excluded one or more matches. Pass partial_ok to replace only scoped matches.');
      }
      const matches = filtered.matches.map((match) => ({
        paragraph_index: match.paragraph_index,
        occurrence_in_paragraph: match.occurrence_in_paragraph,
        text: match.range.text,
        snippet: match.snippet
      }));
      return validationSuccess('word.replace_text', {
        replaced_count: 0,
        matches,
        skipped_count: filtered.skipped_count,
        dry_run: true
      });
    });
  }

  async function validateUpdateParagraphOnly(args) {
    return Word.run(async (context) => {
      const paragraph = await getParagraphByIndex(context, args.index);
      paragraph.load('text,style');
      await context.sync();
      return validationSuccess('word.update_paragraph', {
        resolved_target: {
          type: 'Paragraph',
          paragraph_index: args.index,
          current_text_length: (paragraph.text || '').length,
          style: paragraph.style || null
        }
      });
    });
  }

  async function validateDeleteRangeOnly(args) {
    return Word.run(async (context) => {
      const target = args.extent === 'selection' ? context.document.getSelection() : await resolveAnchor(context, args.anchor);
      target.load('text');
      await context.sync();
      return validationSuccess('word.delete_range', {
        resolved_target: {
          ...validationTargetForAnchor(args.anchor || { kind: 'selection' }),
          extent: args.extent ?? 'paragraph',
          current_text_length: (target.text || '').length
        }
      });
    });
  }

  async function resolveValidationAnchor(context, anchor) {
    if (anchor.kind === 'start_of_document' || anchor.kind === 'end_of_document') {
      return validationTargetForAnchor(anchor);
    }
    const resolved = await resolveAnchor(context, anchor);
    resolved.load('text');
    await context.sync();
    return {
      ...validationTargetForAnchor(anchor),
      current_text_length: (resolved.text || '').length
    };
  }

  function validationTargetForAnchor(anchor) {
    const target = {
      type: resolvedAnchorObjectType(anchor),
      anchor_kind: anchor.kind
    };
    if (Number.isInteger(anchor.index)) target.paragraph_index = anchor.index;
    return target;
  }

  function validationSuccess(operation, data = {}) {
    return {
      valid: true,
      operation,
      partial_effect: 'none',
      ...data
    };
  }

  async function getText(args) {
    return Word.run(async (context) => {
      const offset = args.offset ?? 0;
      const limit = args.limit ?? 200;
      const includeMetadata = args.include_metadata ?? false;
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text,items/style');
      await context.sync();
      const items = paragraphs.items;
      const selected = items.slice(offset, offset + limit);
      const text = selected.map((paragraph) => paragraph.text).join('\n');
      const data = {
        text,
        paragraph_count: items.length,
        returned_paragraphs: { offset, limit },
        truncated: offset + limit < items.length,
        untrusted_source: true
      };
      if (includeMetadata) {
        data.paragraphs = selected.map((paragraph, index) => ({
          index: offset + index,
          text: paragraph.text,
          style: paragraph.style || null,
          level: headingLevelFromStyle(paragraph.style)
        }));
      }
      return data;
    });
  }

  async function getParagraph(args) {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text,items/style');
      await context.sync();
      const paragraph = paragraphs.items[args.index];
      if (!paragraph) throw Object.assign(new Error(`Paragraph index ${args.index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
      return { index: args.index, text: paragraph.text, style: paragraph.style || null, untrusted_source: true };
    });
  }

  async function getOutline(args) {
    return Word.run(async (context) => {
      const maxLevel = args.max_level ?? 6;
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text,items/style');
      await context.sync();
      const flat = [];
      paragraphs.items.forEach((paragraph, index) => {
        const level = headingLevelFromStyle(paragraph.style);
        if (level > 0 && level <= maxLevel) {
          flat.push({ text: paragraph.text, level, paragraph_index: index, children: [] });
        }
      });
      return { outline: nestOutline(flat), headings: flat.map(({ children, ...heading }) => heading), untrusted_source: true };
    });
  }

  async function findText(args) {
    return Word.run(async (context) => {
      const ranges = context.document.body.search(args.query, {
        matchCase: args.match_case ?? false,
        matchWholeWord: args.whole_word ?? false,
        matchWildcards: args.wildcards ?? false
      });
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      ranges.load('items/text');
      await context.sync();
      const limit = args.limit ?? 50;
      const limitedRanges = ranges.items.slice(0, limit);
      const metadata = mapMatchesToParagraphs(paragraphs.items, args.query, args.match_case ?? false, limit, limitedRanges);
      const matches = metadata.map((match, index) => ({
        paragraph_index: match.paragraph_index,
        occurrence_in_paragraph: match.occurrence_in_paragraph,
        text: limitedRanges[index]?.text ?? match.text,
        snippet: match.snippet
      }));
      return { matches, count: matches.length, truncated: ranges.items.length > limit, untrusted_source: true };
    });
  }

  async function resolveAnchorTool(args) {
    requireAnchor('word.resolve_anchor', args.anchor);
    return Word.run(async (context) => {
      const resolved = await resolveAnchor(context, args.anchor);
      return describeResolvedAnchor(context, args.anchor, resolved, args.include_text_preview !== false);
    });
  }

  async function getSelection() {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      selection.load('text');
      paragraphs.load('items');
      await context.sync();
      return {
        text: selection.text,
        paragraph_count: paragraphs.items.length,
        is_empty: selection.text.length === 0,
        untrusted_source: true
      };
    });
  }

  async function getHeaderFooter(args) {
    validateHeaderFooterArgs('word.get_header_footer', args, false);
    return Word.run(async (context) => {
      const target = await headerFooterTarget(context, args);
      const body = target.body;
      const paragraphs = body.paragraphs;
      body.load('text');
      paragraphs.load('items/text,items/style');
      await context.sync();
      const data = {
        text: body.text || '',
        is_empty: !(body.text || '').trim(),
        section_count: target.sectionCount,
        section_index: target.sectionIndex,
        location: target.location,
        header_footer_type: target.headerFooterType,
        untrusted_source: true
      };
      if (args.include_metadata) {
        data.paragraphs = paragraphs.items.map((paragraph, index) => ({
          index,
          text: paragraph.text,
          style: paragraph.style || null
        }));
      }
      return data;
    });
  }

  async function validateUpdateHeaderFooterOnly(args) {
    return Word.run(async (context) => {
      const target = await headerFooterTarget(context, args);
      return validationSuccess('word.update_header_footer', {
        resolved_target: headerFooterResolvedTarget(target)
      });
    });
  }

  async function updateHeaderFooter(args) {
    validateHeaderFooterArgs('word.update_header_footer', args, true);
    return Word.run(async (context) => {
      const target = await headerFooterTarget(context, args);
      const action = normalizedHeaderFooterAction(args.action);
      if (action === 'set_text') {
        target.body.insertText(String(args.text), Word.InsertLocation.replace);
      } else if (action === 'append_paragraph') {
        const paragraph = target.body.insertParagraph(String(args.text), Word.InsertLocation.end);
        if (args.style) paragraph.style = String(args.style);
        if (args.formatting) applyRunFormatting(paragraph.font, args.formatting);
      } else if (action === 'clear') {
        target.body.clear();
      }
      await context.sync();
      return {
        updated: true,
        action,
        resolved_target: headerFooterResolvedTarget(target)
      };
    });
  }

  async function insertParagraph(args) {
    return Word.run(async (context) => {
      const style = paragraphStyleFromArgs(args);
      if (args.anchor.kind === 'start_of_document') {
        const inserted = context.document.body.insertParagraph(args.text, Word.InsertLocation.start);
        if (style) inserted.style = style;
        if (args.formatting) applyRunFormatting(inserted.font, args.formatting);
        await context.sync();
        return { inserted: true };
      }
      if (args.anchor.kind === 'end_of_document') {
        const inserted = context.document.body.insertParagraph(args.text, Word.InsertLocation.end);
        if (style) inserted.style = style;
        if (args.formatting) applyRunFormatting(inserted.font, args.formatting);
        await context.sync();
        return { inserted: true };
      }
      const target = await resolveAnchor(context, args.anchor);
      let inserted;
      switch (args.anchor.kind) {
        case 'start_of_document':
        case 'before_paragraph_index':
        case 'before_text':
          inserted = target.insertParagraph(args.text, Word.InsertLocation.before);
          break;
        case 'selection':
        case 'end_of_document':
        case 'paragraph_index':
        case 'after_paragraph_index':
        case 'after_text':
        default:
          inserted = target.insertParagraph(args.text, Word.InsertLocation.after);
          break;
      }
      if (style) inserted.style = style;
      if (args.formatting) applyRunFormatting(inserted.font, args.formatting);
      await context.sync();
      return { inserted: true };
    });
  }

  async function insertHeading(args) {
    const inserted = await insertParagraph({ text: args.text, anchor: args.anchor, style: `Heading ${args.level}` });
    return { ...inserted, level: args.level };
  }

  async function insertTable(args) {
    validateTableData(args.rows, args.cols, args.data);
    return Word.run(async (context) => {
      const values = args.data ?? blankTable(args.rows, args.cols);
      let table;
      if (args.anchor.kind === 'start_of_document') {
        table = context.document.body.insertTable(args.rows, args.cols, Word.InsertLocation.start, values);
      } else if (args.anchor.kind === 'end_of_document') {
        table = context.document.body.insertTable(args.rows, args.cols, Word.InsertLocation.end, values);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        const location = isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after;
        table = target.insertTable(args.rows, args.cols, location, values);
      }
      if (args.style) table.style = args.style;
      const tables = context.document.body.tables;
      tables.load('items');
      await context.sync();
      const matchedIndex = tables.items.findIndex((item) => item === table);
      const tableIndex = matchedIndex >= 0 ? matchedIndex : Math.max(0, tables.items.length - 1);
      return { inserted: true, table_index: tableIndex, rows: args.rows, cols: args.cols, header_row: args.header_row ?? false };
    });
  }

  async function insertImage(args) {
    return Word.run(async (context) => {
      const base64 = args.image?.base64;
      if (!base64) throw Object.assign(new Error('word.insert_image requires base64 image data after daemon preprocessing.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      validateInsertImagePlacement(args.anchor, args.placement);
      let picture;
      if (args.placement === 'selection') {
        picture = context.document.getSelection().insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
      } else if (args.anchor.kind === 'start_of_document') {
        picture = context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        picture = context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        picture = insertInlinePictureWithPlacement(target, args.anchor, base64, args.placement);
      }
      if (args.alt_text) picture.altTextDescription = args.alt_text;
      if (typeof args.width_pt === 'number') picture.width = args.width_pt;
      if (typeof args.height_pt === 'number') picture.height = args.height_pt;
      await context.sync();
      return { inserted: true, byte_length: args.image.byte_length ?? null, mime_type: args.image.mime_type ?? null };
    });
  }

  async function resizeImage(args) {
    validateResizeImageArgs(args);
    return Word.run(async (context) => {
      const selector = args.image;
      const paragraph = await getParagraphByIndex(context, selector.index);
      const pictures = paragraph.inlinePictures;
      pictures.load('items/width,items/height');
      await context.sync();

      const imageIndex = selector.image_index ?? 0;
      const picture = pictures.items[imageIndex];
      if (!picture) {
        throw Object.assign(new Error(`Inline image index ${imageIndex} is out of range for paragraph ${selector.index}.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE', partialEffect: 'none' });
      }

      const oldWidth = picture.width;
      const oldHeight = picture.height;
      const size = resizedImageSize(args, oldWidth, oldHeight);
      picture.width = size.width;
      picture.height = size.height;
      await context.sync();

      return {
        resized: true,
        image: {
          paragraph_index: selector.index,
          image_index: imageIndex,
          old_width_pt: oldWidth,
          old_height_pt: oldHeight,
          new_width_pt: size.width,
          new_height_pt: size.height,
          preserve_aspect_ratio: args.preserve_aspect_ratio !== false
        }
      };
    });
  }


  async function insertBreak(args) {
    return Word.run(async (context) => {
      const breakType = normalizeBreakType(args.break_type);
      const wordBreakType = wordBreakTypeFrom(breakType);
      if (args.anchor.kind === 'start_of_document') {
        context.document.body.insertBreak(wordBreakType, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        context.document.body.insertBreak(wordBreakType, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        target.insertBreak(wordBreakType, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      await context.sync();
      return { inserted: true, break_type: breakType };
    });
  }

  async function insertPageBreak(args) {
    const result = await insertBreak({ ...args, break_type: 'page' });
    return { ...result, superseded_by: 'word.insert_break' };
  }

  async function listSections(args = {}) {
    return Word.run(async (context) => {
      const sections = context.document.sections;
      const bodyParagraphs = context.document.body.paragraphs;
      sections.load('items');
      bodyParagraphs.load('items/text');
      await context.sync();

      const includePageSetup = args.include_page_setup === true && supportsWordApiDesktop13();
      const sectionViews = [];
      for (let index = 0; index < sections.items.length; index += 1) {
        const section = sections.items[index];
        const body = section.body;
        const paragraphs = body.paragraphs;
        const header = section.getHeader(Word.HeaderFooterType.primary).body;
        const footer = section.getFooter(Word.HeaderFooterType.primary).body;
        body.load('text');
        paragraphs.load('items/text');
        header.load('text');
        footer.load('text');
        if (includePageSetup) section.pageSetup.load('orientation,paperSize,topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight');
        sectionViews.push({ section, body, paragraphs, header, footer });
      }
      await context.sync();

      const entries = [];
      let nextParagraphIndex = 0;
      for (let index = 0; index < sectionViews.length; index += 1) {
        const { section, paragraphs, header, footer } = sectionViews[index];
        const paragraphItems = paragraphs.items || [];
        const entry = {
          index,
          first_paragraph_index: nextParagraphIndex,
          paragraph_count: paragraphItems.length,
          has_header: Boolean((header.text || '').trim()),
          has_footer: Boolean((footer.text || '').trim())
        };
        if (includePageSetup) entry.page_setup = pageSetupMetadata(section.pageSetup);
        entries.push(entry);
        nextParagraphIndex += paragraphItems.length;
      }

      return { sections: entries, count: entries.length };
    });
  }

  async function updatePageSetup(args) {
    requireWordApiDesktop13('word.update_page_setup');
    return Word.run(async (context) => {
      const target = await pageSetupTarget(context, args.section_index);
      applyPageSetup(target.pageSetup, args);
      await context.sync();
      target.pageSetup.load('orientation,paperSize,topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight');
      await context.sync();
      return {
        updated: true,
        section_index: target.sectionIndex,
        page_setup: pageSetupMetadata(target.pageSetup)
      };
    });
  }

  async function insertList(args) {
    return Word.run(async (context) => {
      const text = args.items.join('\n');
      let range;
      if (args.anchor.kind === 'start_of_document') {
        range = context.document.body.insertText(text, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        range = context.document.body.insertText(text, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        range = target.insertText(text, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      range.paragraphs.load('items');
      await context.sync();
      const paragraphs = range.paragraphs.items;
      const level = args.level ?? 0;
      let list = null;
      for (let index = 0; index < paragraphs.length; index++) {
        const paragraph = paragraphs[index];
        paragraph.style = 'List Paragraph';
        if (index === 0) {
          list = paragraph.startNewList();
          if ((args.kind ?? 'bulleted') === 'numbered') {
            list.setLevelNumbering(level, Word.ListNumbering.arabic);
          } else {
            list.setLevelBullet(level, Word.ListBullet.solid);
          }
        } else if (list) {
          list.load('id');
          await context.sync();
          paragraph.attachToList(list.id, level);
        }
        if (level > 0) paragraph.leftIndent = 18 * level;
      }
      await context.sync();
      return { inserted_items: args.items.length, kind: args.kind ?? 'bulleted', level: args.level ?? 0 };
    });
  }

  async function replaceText(args) {
    return Word.run(async (context) => {
      const scope = args.scope || {};
      const searchRoot = scope.selection_only ? context.document.getSelection() : context.document.body;
      const ranges = searchRoot.search(args.find, {
        matchCase: args.match_case ?? false,
        matchWholeWord: args.whole_word ?? false,
        matchWildcards: args.wildcards ?? false
      });
      ranges.load('items/text');
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();

      const limit = args.limit ?? 500;
      const limitedRanges = ranges.items.slice(0, limit);
      const paragraphMatches = mapMatchesToParagraphs(paragraphs.items, args.find, args.match_case ?? false, limit, limitedRanges);
      const filtered = filterReplaceRanges(limitedRanges, paragraphMatches, scope.paragraph_range);
      filtered.skipped_count += Math.max(0, ranges.items.length - limitedRanges.length);
      if (!args.partial_ok && scope.paragraph_range && filtered.skipped_count > 0) {
        throw Object.assign(new Error('replace_text scope excluded one or more matches. Pass partial_ok to replace only scoped matches.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      }
      const matches = filtered.matches.map((match) => ({
        paragraph_index: match.paragraph_index,
        occurrence_in_paragraph: match.occurrence_in_paragraph,
        text: match.range.text,
        snippet: match.snippet
      }));
      if (args.dry_run || args.validate_only) {
        return {
          valid: args.validate_only ? true : undefined,
          operation: args.validate_only ? 'word.replace_text' : undefined,
          partial_effect: args.validate_only ? 'none' : undefined,
          replaced_count: 0,
          matches,
          dry_run: true,
          skipped_count: filtered.skipped_count
        };
      }
      for (const match of filtered.matches) {
        match.range.insertText(args.replace, Word.InsertLocation.replace);
      }
      await context.sync();
      return { replaced_count: filtered.matches.length, matches, skipped_count: filtered.skipped_count };
    });
  }

  function mapMatchesToParagraphs(paragraphs, find, matchCase, limit, ranges) {
    const rangeTexts = ranges ? Array.from(ranges, (range) => range.text || '') : null;
    const query = matchCase ? find : find.toLowerCase();
    const matches = [];
    for (let i = 0; i < paragraphs.length && matches.length < limit; i++) {
      const text = paragraphs[i].text || '';
      const haystack = matchCase ? text : text.toLowerCase();
      let from = 0;
      let occurrence = 0;
      while (matches.length < limit) {
        const currentText = rangeTexts?.[matches.length] || find;
        const currentQuery = matchCase ? currentText : currentText.toLowerCase();
        const index = haystack.indexOf(rangeTexts ? currentQuery : query, from);
        if (index < 0) break;
        occurrence += 1;
        matches.push({
          paragraph_index: i,
          occurrence_in_paragraph: occurrence,
          text: text.slice(index, index + currentText.length),
          snippet: text.slice(Math.max(0, index - 40), Math.min(text.length, index + currentText.length + 40))
        });
        from = index + Math.max(1, currentQuery.length);
        if (rangeTexts && matches.length >= rangeTexts.length) return matches;
      }
    }
    return matches;
  }

  function filterReplaceRanges(ranges, paragraphMatches, paragraphRange) {
    const selected = [];
    let skipped = 0;
    for (let i = 0; i < ranges.length; i++) {
      const metadata = paragraphMatches[i] || { paragraph_index: null, occurrence_in_paragraph: null, snippet: ranges[i].text };
      const inRange = !paragraphRange || (metadata.paragraph_index >= paragraphRange[0] && metadata.paragraph_index <= paragraphRange[1]);
      if (inRange) selected.push({ range: ranges[i], ...metadata });
      else skipped += 1;
    }
    return { matches: selected, skipped_count: skipped };
  }

  async function updateParagraph(args) {
    return Word.run(async (context) => {
      const paragraph = await getParagraphByIndex(context, args.index);
      paragraph.insertText(args.text, Word.InsertLocation.replace);
      await context.sync();
      return { updated_paragraph_index: args.index };
    });
  }

  async function deleteRange(args) {
    return Word.run(async (context) => {
      const target = args.extent === 'selection' ? context.document.getSelection() : await resolveAnchor(context, args.anchor);
      if (args.extent === 'sentence') {
        const sentence = await rangeSentenceOrSelf(context, target);
        sentence.delete();
        await context.sync();
        return { deleted: true, extent: 'sentence' };
      }
      if (target.paragraphs && (args.extent ?? 'paragraph') === 'paragraph') {
        const paragraphs = target.paragraphs;
        paragraphs.load('items');
        await context.sync();
        if (paragraphs.items[0]) paragraphs.items[0].delete();
      } else {
        target.delete();
      }
      await context.sync();
      return { deleted: true, extent: args.extent ?? 'paragraph' };
    });
  }

  async function applyFormatting(args) {
    return Word.run(async (context) => {
      const target = args.extent === 'selection' ? context.document.getSelection() : await resolveAnchor(context, args.anchor);
      const range = args.extent === 'sentence' ? await rangeSentenceOrSelf(context, target) : target;
      applyRunFormatting(range.font, args.formatting);
      await context.sync();
      return { formatted: true };
    });
  }

  async function rangeSentenceOrSelf(context, target) {
    target.load('text');
    await context.sync();
    const sourceText = target.text || '';
    const paragraphs = target.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const paragraph = paragraphs.items[0];
    if (!paragraph) return target;
    const sentences = paragraph.split(['.', '!', '?'], false, true);
    sentences.load('items/text');
    await context.sync();
    if (!sourceText) return sentences.items[0] || target;
    return sentences.items.find((sentence) => sentence.text.includes(sourceText)) || sentences.items[0] || target;
  }

  async function readTable(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.load('rowCount,columnCount,values');
      await context.sync();
      const data = table.values ?? [];
      return { rows: table.rowCount ?? data.length, cols: table.columnCount ?? (data[0]?.length ?? 0), data, header_row: false, untrusted_source: true };
    });
  }

  async function updateTable(args) {
    const action = String(args.action || '').trim().toLowerCase();
    switch (action) {
      case 'update_cell':
      case 'cell':
        return updateCell(args);
      case 'add_row':
        return addRow(args);
      case 'add_column':
        return addColumn(args);
      case 'format_cell':
        return formatCell(args);
      case 'delete':
        return deleteTable(args);
      default:
        throw Object.assign(new Error(`Unsupported table action ${args.action}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
  }

  async function updateCell(args) {
    return Word.run(async (context) => {
      const cell = await getTableCell(context, args.table_index, args.row, args.col);
      cell.value = args.text;
      if (args.formatting) applyRunFormatting(cell.body.getRange().font, args.formatting);
      await context.sync();
      return { updated: true, table_index: args.table_index, row: args.row, col: args.col };
    });
  }

  async function addRow(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.load('rowCount,columnCount,values');
      await context.sync();
      const rowCount = table.rowCount ?? table.values.length;
      const columnCount = table.columnCount ?? (table.values[0]?.length ?? 0);
      const values = args.values ?? blankRow(columnCount);
      if (values.length !== columnCount) throw invalidArgument('Row values length must match table column count.');
      const rowIndex = args.index ?? rowCount;
      if (rowIndex >= rowCount) {
        table.getCell(rowCount - 1, 0).insertRows(Word.InsertLocation.after, 1, [values]);
      } else {
        table.getCell(rowIndex, 0).insertRows(Word.InsertLocation.before, 1, [values]);
      }
      await context.sync();
      return { added_row_index: rowIndex };
    });
  }

  async function addColumn(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.load('rowCount,columnCount,values');
      await context.sync();
      const rowCount = table.rowCount ?? table.values.length;
      const columnCount = table.columnCount ?? (table.values[0]?.length ?? 0);
      const values = args.values ?? blankColumn(rowCount);
      if (values.length !== rowCount) throw invalidArgument('Column values length must match table row count.');
      const colIndex = args.index ?? columnCount;
      const matrix = values.map((value) => [value]);
      if (colIndex >= columnCount) {
        table.getCell(0, columnCount - 1).insertColumns(Word.InsertLocation.after, 1, matrix);
      } else {
        table.getCell(0, colIndex).insertColumns(Word.InsertLocation.before, 1, matrix);
      }
      await context.sync();
      return { added_column_index: colIndex };
    });
  }

  async function formatCell(args) {
    return Word.run(async (context) => {
      const cell = await getTableCell(context, args.table_index, args.row, args.col);
      if (args.background_color) cell.shadingColor = args.background_color;
      if (args.horizontal_alignment) cell.horizontalAlignment = mapHorizontalAlignment(args.horizontal_alignment);
      if (args.vertical_alignment) cell.verticalAlignment = mapVerticalAlignment(args.vertical_alignment);
      if (typeof args.padding_pt === 'number') {
        cell.setCellPadding(Word.CellPaddingLocation.top, args.padding_pt);
        cell.setCellPadding(Word.CellPaddingLocation.bottom, args.padding_pt);
        cell.setCellPadding(Word.CellPaddingLocation.left, args.padding_pt);
        cell.setCellPadding(Word.CellPaddingLocation.right, args.padding_pt);
      }
      if (args.formatting) applyRunFormatting(cell.body.getRange().font, args.formatting);
      await context.sync();
      return { formatted: true, table_index: args.table_index, row: args.row, col: args.col };
    });
  }

  async function deleteTable(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.delete();
      await context.sync();
      return { deleted: true, table_index: args.table_index };
    });
  }

  async function listContentControls(args) {
    return Word.run(async (context) => {
      const controls = context.document.body.getContentControls(contentControlFilterOptions(args));
      controls.load('items/id,items/tag,items/title,items/type,items/subtype,items/cannotDelete,items/cannotEdit');
      await context.sync();
      const filtered = controls.items.filter((control) => {
        if (args.tag !== undefined && control.tag !== String(args.tag)) return false;
        if (args.title !== undefined && control.title !== String(args.title)) return false;
        return true;
      });
      return {
        content_controls: filtered.map(contentControlMetadata),
        count: filtered.length,
        untrusted_source: true
      };
    });
  }

  async function insertContentControl(args) {
    return Word.run(async (context) => {
      const target = args.anchor ? await resolveAnchor(context, args.anchor) : context.document.getSelection();
      const range = args.text !== undefined ? target.insertText(String(args.text), Word.InsertLocation.replace) : target;
      const control = range.insertContentControl(contentControlTypeFrom(args.type));
      applyContentControlProperties(control, args);
      control.load('id,tag,title,type,subtype,cannotDelete,cannotEdit');
      await context.sync();
      return { content_control: contentControlMetadata(control, 0), created: true };
    });
  }

  async function updateContentControl(args) {
    return Word.run(async (context) => {
      const control = await targetContentControl(context, args);
      applyContentControlProperties(control, args);
      if (args.text !== undefined) control.insertText(String(args.text), Word.InsertLocation.replace);
      control.load('id,tag,title,type,subtype,cannotDelete,cannotEdit');
      await context.sync();
      return { content_control: contentControlMetadata(control, 0), updated: true };
    });
  }

  async function deleteContentControl(args) {
    return Word.run(async (context) => {
      const control = await targetContentControl(context, args);
      control.load('id');
      await context.sync();
      const id = control.id;
      const keepContent = args.mode !== 'delete_content';
      control.delete(keepContent);
      await context.sync();
      return { content_control_id: id, deleted: true, mode: keepContent ? 'keep_content' : 'delete_content' };
    });
  }

  async function insertBookmark(args) {
    return Word.run(async (context) => {
      const name = String(args.name);
      await ensureBookmarkAvailable(context, name, Boolean(args.overwrite));
      const target = await resolveAnchor(context, args.anchor);
      const range = target.getRange ? target.getRange() : target;
      range.insertBookmark(name);
      await context.sync();
      const bookmark = await bookmarkMetadata(context, name);
      return { bookmark, overwritten: Boolean(args.overwrite) };
    });
  }

  async function listBookmarks(args) {
    return Word.run(async (context) => {
      const bookmarks = await bookmarkNames(context, Boolean(args.include_hidden));
      const items = [];
      for (const name of bookmarks) {
        const metadata = await bookmarkMetadata(context, name);
        items.push(metadata);
      }
      return { bookmarks: items, count: items.length, untrusted_source: true };
    });
  }

  async function deleteBookmark(args) {
    return Word.run(async (context) => {
      const name = await existingBookmarkName(context, String(args.name));
      context.document.deleteBookmark(name);
      await context.sync();
      const remaining = await bookmarkNames(context, false);
      return { deleted: true, name, count: remaining.length };
    });
  }

  async function setHeadingLevel(args) {
    return Word.run(async (context) => {
      const paragraph = await getParagraphByIndex(context, args.index);
      paragraph.style = args.level === 0 ? 'Normal' : `Heading ${args.level}`;
      await context.sync();
      return { paragraph_index: args.index, level: args.level };
    });
  }

  async function applyStyle(args) {
    return Word.run(async (context) => {
      const target = await resolveAnchor(context, args.anchor);
      target.style = args.style || headingStyleFromLevel(args.heading_level);
      await context.sync();
      return { styled: true, style: target.style };
    });
  }

  async function getComments() {
    return Word.run(async (context) => {
      const comments = context.document.comments;
      comments.load('items/id,items/content,items/resolved,items/authorName,items/creationDate');
      await context.sync();
      return {
        comments: comments.items.map((comment, index) => ({
          index,
          comment_id: comment.id,
          content: comment.content,
          resolved: comment.resolved,
          author: comment.authorName || null,
          created_at: dateToIso(comment.creationDate),
          untrusted_source: true
        })),
        count: comments.items.length,
        untrusted_source: true
      };
    });
  }

  async function getTrackedChanges() {
    return Word.run(async (context) => {
      const changes = context.document.body.getTrackedChanges();
      changes.load('items/author,items/date,items/type,items/text');
      await context.sync();
      return {
        changes: changes.items.map((change, index) => ({
          index,
          author: change.author || null,
          date: dateToIso(change.date),
          type: change.type || null,
          text: change.text || '',
          fingerprint: trackedChangeFingerprint(change, index),
          untrusted_source: true
        })),
        count: changes.items.length,
        untrusted_source: true
      };
    });
  }

  async function getStructure() {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      const tables = context.document.body.tables;
      paragraphs.load('items/text,items/style,items/isListItem');
      tables.load('items/rowCount,items/columnCount');
      await context.sync();
      const listMetadata = new Map();
      const listObjects = [];
      paragraphs.items.forEach((paragraph, index) => {
        if (!paragraph.isListItem) return;
        const listItem = paragraph.listItemOrNullObject;
        const list = paragraph.listOrNullObject;
        listItem.load('isNullObject,level');
        list.load('isNullObject,levelTypes');
        listObjects.push({ index, listItem, list });
      });
      if (listObjects.length > 0) await context.sync();
      for (const item of listObjects) {
        const level = item.listItem.isNullObject ? null : item.listItem.level;
        const levelTypes = item.list.isNullObject ? [] : item.list.levelTypes ?? [];
        listMetadata.set(item.index, { level, kind: listKindFromLevelType(levelTypes[level ?? 0]) });
      }
      const headings = [];
      const lists = [];
      paragraphs.items.forEach((paragraph, index) => {
        const level = headingLevelFromStyle(paragraph.style);
        if (level > 0) headings.push({ text: paragraph.text, level, paragraph_index: index });
        if (paragraph.isListItem) {
          const metadata = listMetadata.get(index) || { level: null, kind: null };
          lists.push({ paragraph_index: index, text: paragraph.text, level: metadata.level, kind: metadata.kind });
        }
      });
      return {
        outline: nestOutline(headings.map((heading) => ({ ...heading, children: [] }))),
        headings,
        lists,
        tables: tables.items.map((table, index) => ({ index, rows: table.rowCount ?? null, cols: table.columnCount ?? null })),
        untrusted_source: true
      };
    });
  }

  async function addComment(args) {
    return Word.run(async (context) => {
      const target = await resolveAnchor(context, args.anchor);
      const comment = target.insertComment(args.text);
      comment.load('id,content,resolved,authorName,creationDate');
      await context.sync();
      return {
        comment_id: comment.id,
        content: comment.content,
        resolved: comment.resolved,
        author: comment.authorName || null,
        created_at: dateToIso(comment.creationDate)
      };
    });
  }

  async function resolveComment(args) {
    return Word.run(async (context) => {
      const comments = context.document.comments;
      comments.load('items/id,items/resolved');
      await context.sync();
      const comment = comments.items.find((item) => item.id === args.comment_id);
      if (!comment) throw Object.assign(new Error(`Comment ${args.comment_id} was not found.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
      comment.resolved = true;
      await context.sync();
      return { comment_id: args.comment_id, resolved: true };
    });
  }

  async function updateTrackedChange(args) {
    const action = String(args.action || '').trim().toLowerCase();
    if (action !== 'accept' && action !== 'reject') {
      throw Object.assign(new Error(`Unsupported tracked-change action ${args.action}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return mutateTrackedChange(args, action);
  }

  async function mutateTrackedChange(args, action) {
    return Word.run(async (context) => {
      const changes = context.document.body.getTrackedChanges();
      changes.load('items/author,items/date,items/type,items/text');
      await context.sync();
      const change = changes.items[args.change_index];
      if (!change) throw Object.assign(new Error(`Tracked change index ${args.change_index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
      const fingerprint = trackedChangeFingerprint(change, args.change_index);
      if (fingerprint !== args.expected_fingerprint) {
        throw Object.assign(new Error('Tracked change fingerprint mismatch; re-read track_changes before mutating.'), { officeMcpCode: 'STALE_INDEX' });
      }
      if (action === 'accept') change.accept();
      else change.reject();
      await context.sync();
      return { change_index: args.change_index, fingerprint, action };
    });
  }

  async function saveDocument() {
    return Word.run(async (context) => {
      const document = context.document;
      document.load('saved');
      await context.sync();
      const wasDirty = typeof document.saved === 'boolean' ? !document.saved : null;
      document.save(Word.SaveBehavior.save);
      document.load('saved');
      await context.sync();
      return { saved: true, was_dirty: wasDirty, is_dirty: typeof document.saved === 'boolean' ? !document.saved : null };
    });
  }

  async function resolveAnchor(context, anchor) {
    switch (anchor.kind) {
      case 'selection':
        return context.document.getSelection();
      case 'start_of_document':
      case 'end_of_document':
        return context.document.body;
      case 'paragraph_index':
      case 'before_paragraph_index':
      case 'after_paragraph_index': {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        const paragraph = paragraphs.items[anchor.index];
        if (!paragraph) throw Object.assign(new Error(`Paragraph index ${anchor.index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
        return paragraph;
      }
      case 'after_text':
      case 'before_text': {
        const ranges = context.document.body.search(anchor.text, { matchCase: false });
        ranges.load('items');
        await context.sync();
        const range = ranges.items[(anchor.occurrence ?? 1) - 1];
        if (!range) throw Object.assign(new Error(`Anchor text not found: ${anchor.text}`), { officeMcpCode: 'ANCHOR_NOT_FOUND' });
        return range;
      }
      case 'heading': {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items/text,items/style');
        await context.sync();
        const expectedLevel = anchor.level || null;
        const paragraph = paragraphs.items.find((item) => item.text === anchor.text && headingLevelFromStyle(item.style) > 0 && (!expectedLevel || headingLevelFromStyle(item.style) === expectedLevel));
        if (!paragraph) throw Object.assign(new Error('Heading anchor not found: ' + anchor.text), { officeMcpCode: 'ANCHOR_NOT_FOUND' });
        return paragraph;
      }
      case 'bookmark': {
        const range = context.document.getBookmarkRangeOrNullObject(anchor.name);
        range.load('isNullObject,text');
        await context.sync();
        if (range.isNullObject) throw Object.assign(new Error('Bookmark anchor not found: ' + anchor.name), { officeMcpCode: 'ANCHOR_NOT_FOUND' });
        return range;
      }
      default:
        throw Object.assign(new Error(`Unsupported anchor ${anchor.kind}`), { officeMcpCode: 'INVALID_ARGUMENT' });
    }
  }

  async function describeResolvedAnchor(context, anchor, resolved, includeTextPreview) {
    const descriptor = {
      resolved: true,
      anchor_kind: anchor.kind,
      object_type: resolvedAnchorObjectType(anchor),
      supported_operations: supportedOperationsForAnchor(anchor),
      unsupported_operations: unsupportedOperationsForAnchor(anchor),
      tool_suitability: toolSuitabilityForAnchor(anchor),
      untrusted_source: true
    };
    if (Number.isInteger(anchor.index)) descriptor.paragraph_index = anchor.index;
    if (includeTextPreview) {
      const preview = await resolvedAnchorTextPreview(context, anchor, resolved);
      if (preview !== null) descriptor.text_preview = preview;
    }
    return descriptor;
  }

  async function resolvedAnchorTextPreview(context, anchor, resolved) {
    switch (anchor.kind) {
      case 'start_of_document':
      case 'end_of_document':
        return null;
      case 'paragraph_index':
      case 'before_paragraph_index':
      case 'after_paragraph_index':
      case 'heading':
        resolved.load('text');
        await context.sync();
        return safeTextPreview(resolved.text);
      case 'selection':
      case 'before_text':
      case 'after_text':
      case 'bookmark':
        resolved.load('text');
        await context.sync();
        return safeTextPreview(resolved.text);
      default:
        return null;
    }
  }

  function safeTextPreview(value) {
    const text = String(value || '').replace(/\s+/g, ' ').trim();
    if (!text) return '';
    return text.length > 80 ? `${text.slice(0, 77)}...` : text;
  }

  async function bookmarkNames(context, includeHidden) {
    const range = context.document.body.getRange();
    const result = range.getBookmarks(Boolean(includeHidden), true);
    await context.sync();
    return Array.isArray(result.value) ? result.value : [];
  }

  async function existingBookmarkName(context, name) {
    const requested = String(name || '');
    const names = await bookmarkNames(context, true);
    const match = names.find((item) => item.toLowerCase() === requested.toLowerCase());
    if (!match) throw invalidArgument(`Bookmark not found: ${requested}.`);
    return match;
  }

  async function ensureBookmarkAvailable(context, name, overwrite) {
    const names = await bookmarkNames(context, true);
    const exists = names.some((item) => item.toLowerCase() === String(name).toLowerCase());
    if (exists && !overwrite) throw invalidArgument(`Bookmark ${name} already exists; pass overwrite=true to move it.`);
  }

  async function bookmarkMetadata(context, name) {
    const range = context.document.getBookmarkRangeOrNullObject(name);
    range.load('isNullObject,text');
    await context.sync();
    if (range.isNullObject) throw invalidArgument(`Bookmark not found: ${name}.`);
    const metadata = { name, text_preview: safeTextPreview(range.text), untrusted_source: true };
    const paragraphIndex = await paragraphIndexForRange(context, range);
    if (paragraphIndex !== null) metadata.paragraph_index = paragraphIndex;
    return metadata;
  }

  async function paragraphIndexForRange(context, range) {
    range.load('text');
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items/text');
    await context.sync();
    const preview = safeTextPreview(range.text);
    if (!preview) return null;
    const index = paragraphs.items.findIndex((paragraph) => safeTextPreview(paragraph.text).includes(preview));
    return index >= 0 ? index : null;
  }

  function resolvedAnchorObjectType(anchor) {
    if (anchor.kind === 'start_of_document' || anchor.kind === 'end_of_document') return 'Body';
    if (isParagraphAnchor(anchor)) return 'Paragraph';
    return 'Range';
  }

  function supportedOperationsForAnchor(anchor) {
    switch (resolvedAnchorObjectType(anchor)) {
      case 'Body':
        return ['insertParagraph', 'insertTable', 'insertInlinePictureFromBase64', 'insertBreak'];
      case 'Paragraph':
        return ['insertParagraph', 'insertText', 'insertTable', 'insertBreak'];
      default:
        return ['insertText', 'delete', 'font', 'insertContentControl'];
    }
  }

  function unsupportedOperationsForAnchor(anchor) {
    switch (resolvedAnchorObjectType(anchor)) {
      case 'Paragraph':
        return ['insertInlinePictureFromBase64'];
      case 'Range':
        return ['insertTable'];
      default:
        return [];
    }
  }

  function toolSuitabilityForAnchor(anchor) {
    const objectType = resolvedAnchorObjectType(anchor);
    return {
      image_insertion: objectType !== 'Paragraph' || 'requires_explicit_paragraph_placement',
      text_replacement: objectType === 'Range',
      deletion: objectType !== 'Body',
      formatting: objectType !== 'Body'
    };
  }

  async function getParagraphByIndex(context, index) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const paragraph = paragraphs.items[index];
    if (!paragraph) throw Object.assign(new Error(`Paragraph index ${index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
    return paragraph;
  }

  async function getTableByIndex(context, index) {
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();
    const table = tables.items[index];
    if (!table) throw Object.assign(new Error(`Table index ${index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
    return table;
  }

  async function getTableCell(context, tableIndex, row, col) {
    const table = await getTableByIndex(context, tableIndex);
    table.load('rowCount,columnCount');
    await context.sync();
    if (row >= table.rowCount || col >= table.columnCount) {
      throw Object.assign(new Error(`Cell ${row},${col} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
    }
    return table.getCell(row, col);
  }

  async function headerFooterTarget(context, args) {
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();
    const sectionIndex = args.section_index ?? 0;
    const section = sections.items[sectionIndex];
    if (!section) throw Object.assign(new Error(`Section index ${sectionIndex} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE', partialEffect: 'none' });
    const headerFooterType = normalizedHeaderFooterType(args.header_footer_type);
    const location = normalizedHeaderFooterLocation(args.location);
    await validateHeaderFooterLayout(section, headerFooterType);
    const body = location === 'header'
      ? section.getHeader(wordHeaderFooterType(headerFooterType))
      : section.getFooter(wordHeaderFooterType(headerFooterType));
    return { body, section, sectionIndex, sectionCount: sections.items.length, location, headerFooterType };
  }

  async function validateHeaderFooterLayout(section, headerFooterType) {
    if (headerFooterType === 'primary') return;
    const pageSetup = section.pageSetup;
    pageSetup.load('differentFirstPageHeaderFooter,oddAndEvenPagesHeaderFooter');
    await section.context.sync();
    if (headerFooterType === 'first_page' && !pageSetup.differentFirstPageHeaderFooter) {
      throw invalidArgument('first_page header/footer requires different-first-page layout to be enabled for the section.');
    }
    if (headerFooterType === 'even_pages' && !pageSetup.oddAndEvenPagesHeaderFooter) {
      throw invalidArgument('even_pages header/footer requires odd/even-pages layout to be enabled for the section.');
    }
  }

  function headerFooterResolvedTarget(target) {
    return {
      section_index: target.sectionIndex,
      section_count: target.sectionCount,
      location: target.location,
      header_footer_type: target.headerFooterType
    };
  }

  function wordHeaderFooterType(headerFooterType) {
    switch (headerFooterType) {
      case 'primary':
        return Word.HeaderFooterType.primary;
      case 'first_page':
        return Word.HeaderFooterType.firstPage;
      case 'even_pages':
        return Word.HeaderFooterType.evenPages;
      default:
        throw invalidArgument(`Unsupported header_footer_type ${headerFooterType}.`);
    }
  }

  function normalizeBreakType(value) {
    const breakType = value || 'page';
    validateBreakType(breakType);
    return breakType;
  }

  function validateBreakType(value) {
    if (value === undefined) return;
    if (!['page', 'line', 'section_next', 'section_continuous', 'section_even', 'section_odd'].includes(value)) {
      throw invalidArgument('word.insert_break break_type must be page, line, section_next, section_continuous, section_even, or section_odd.');
    }
  }

  function wordBreakTypeFrom(value) {
    switch (value) {
      case 'page':
        return Word.BreakType.page;
      case 'line':
        return Word.BreakType.line;
      case 'section_next':
        return Word.BreakType.sectionNext;
      case 'section_continuous':
        return Word.BreakType.sectionContinuous;
      case 'section_even':
        return Word.BreakType.sectionEven;
      case 'section_odd':
        return Word.BreakType.sectionOdd;
      default:
        throw invalidArgument(`Unsupported break_type ${value}.`);
    }
  }

  async function pageSetupTarget(context, sectionIndex) {
    if (sectionIndex === undefined || sectionIndex === null) {
      return { pageSetup: context.document.pageSetup, sectionIndex: null };
    }
    const sections = context.document.sections;
    sections.load('items');
    await context.sync();
    const section = sections.items[sectionIndex];
    if (!section) throw Object.assign(new Error(`Section index ${sectionIndex} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE', partialEffect: 'none' });
    return { pageSetup: section.pageSetup, sectionIndex };
  }

  function validateUpdatePageSetupArgs(args) {
    if (args.section_index !== undefined) requireNonNegativeInteger('word.update_page_setup', 'section_index', args.section_index);
    if (args.orientation !== undefined && !['portrait', 'landscape'].includes(args.orientation)) {
      throw invalidArgument('word.update_page_setup orientation must be portrait or landscape.');
    }
    validateOptionalNonNegativeNumber('word.update_page_setup', 'page_width_pt', args.page_width_pt, true);
    validateOptionalNonNegativeNumber('word.update_page_setup', 'page_height_pt', args.page_height_pt, true);
    if (args.margins_pt !== undefined) {
      if (!args.margins_pt || typeof args.margins_pt !== 'object' || Array.isArray(args.margins_pt)) {
        throw invalidArgument('word.update_page_setup margins_pt must be an object.');
      }
      for (const side of ['top', 'bottom', 'left', 'right']) {
        validateOptionalNonNegativeNumber('word.update_page_setup', `margins_pt.${side}`, args.margins_pt[side], false);
      }
    }
    if (
      args.orientation === undefined
      && args.paper_size === undefined
      && args.margins_pt === undefined
      && args.page_width_pt === undefined
      && args.page_height_pt === undefined
    ) {
      throw invalidArgument('word.update_page_setup requires at least one page setup field.');
    }
  }

  function validateOptionalNonNegativeNumber(tool, name, value, exclusive) {
    if (value === undefined) return;
    if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || (exclusive && value <= 0)) {
      throw invalidArgument(`${tool} ${name} must be a ${exclusive ? 'positive' : 'non-negative'} number.`);
    }
  }

  function applyPageSetup(pageSetup, args) {
    if (args.orientation !== undefined) pageSetup.orientation = args.orientation === 'landscape' ? Word.Orientation.landscape : Word.Orientation.portrait;
    if (args.paper_size !== undefined) pageSetup.paperSize = String(args.paper_size);
    if (args.margins_pt) {
      if (args.margins_pt.top !== undefined) pageSetup.topMargin = args.margins_pt.top;
      if (args.margins_pt.bottom !== undefined) pageSetup.bottomMargin = args.margins_pt.bottom;
      if (args.margins_pt.left !== undefined) pageSetup.leftMargin = args.margins_pt.left;
      if (args.margins_pt.right !== undefined) pageSetup.rightMargin = args.margins_pt.right;
    }
    if (args.page_width_pt !== undefined) pageSetup.pageWidth = args.page_width_pt;
    if (args.page_height_pt !== undefined) pageSetup.pageHeight = args.page_height_pt;
  }

  function pageSetupMetadata(pageSetup) {
    return {
      orientation: normalizeOrientation(pageSetup.orientation),
      paper_size: pageSetup.paperSize || null,
      margins_pt: {
        top: pageSetup.topMargin,
        bottom: pageSetup.bottomMargin,
        left: pageSetup.leftMargin,
        right: pageSetup.rightMargin
      },
      page_width_pt: pageSetup.pageWidth,
      page_height_pt: pageSetup.pageHeight
    };
  }

  function normalizeOrientation(value) {
    const text = String(value || '').toLowerCase();
    return text.includes('landscape') ? 'landscape' : 'portrait';
  }

  function supportsWordApiDesktop13() {
    return Office.context?.requirements?.isSetSupported?.('WordApiDesktop', '1.3') === true;
  }

  function requireWordApiDesktop13(tool) {
    if (supportsWordApiDesktop13()) return;
    throw Object.assign(new Error(`${tool} requires WordApiDesktop 1.3.`), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE', partialEffect: 'none' });
  }

  async function targetContentControl(context, args) {
    const id = Number(args.content_control_id ?? args.id);
    if (!Number.isInteger(id)) {
      throw Object.assign(new Error('Content control id is required.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return context.document.body.getContentControls().getById(id);
  }

  function contentControlFilterOptions(args) {
    if (!args.type) return undefined;
    return { types: [contentControlTypeFrom(args.type)] };
  }

  function contentControlMetadata(control, index) {
    return {
      index,
      content_control_id: control.id,
      tag: control.tag || null,
      title: control.title || null,
      type: control.type || null,
      subtype: control.subtype || null,
      cannot_delete: Boolean(control.cannotDelete),
      cannot_edit: Boolean(control.cannotEdit),
      untrusted_source: true
    };
  }

  function applyContentControlProperties(control, args) {
    if (args.tag !== undefined) control.tag = String(args.tag);
    if (args.title !== undefined) control.title = String(args.title);
    if (args.cannot_delete !== undefined) control.cannotDelete = Boolean(args.cannot_delete);
    if (args.cannot_edit !== undefined) control.cannotEdit = Boolean(args.cannot_edit);
    if (args.lock_content_control !== undefined) control.cannotDelete = Boolean(args.lock_content_control);
    if (args.lock_contents !== undefined) control.cannotEdit = Boolean(args.lock_contents);
    if (args.appearance !== undefined) control.appearance = contentControlAppearanceFrom(args.appearance);
    if (args.color !== undefined) control.color = String(args.color);
    if (args.placeholder_text !== undefined) control.placeholderText = String(args.placeholder_text);
  }

  function contentControlTypeFrom(value) {
    const normalized = String(value || 'rich_text').trim().toLowerCase();
    const values = {
      rich_text: 'RichText',
      richtext: 'RichText',
      rich: 'RichText',
      plain_text: 'PlainText',
      plaintext: 'PlainText',
      plain: 'PlainText'
    };
    if (values[normalized]) return values[normalized];
    throw Object.assign(new Error(`Unsupported content control type ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  function contentControlAppearanceFrom(value) {
    const normalized = String(value || '').trim().toLowerCase();
    if (normalized === 'bounding_box' || normalized === 'boundingbox') return 'BoundingBox';
    if (normalized === 'tags') return 'Tags';
    if (normalized === 'hidden') return 'Hidden';
    throw Object.assign(new Error(`Unsupported content control appearance ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  function requireAnchor(tool, anchor) {
    if (!anchor || typeof anchor !== 'object' || typeof anchor.kind !== 'string') {
      throw invalidArgument(`${tool} requires anchor.kind.`);
    }
  }

  function requireNonNegativeInteger(tool, name, value) {
    if (!Number.isInteger(value) || value < 0) {
      throw invalidArgument(`${tool} ${name} must be a non-negative integer.`);
    }
  }

  function requirePositiveInteger(tool, name, value) {
    if (!Number.isInteger(value) || value < 1) {
      throw invalidArgument(`${tool} ${name} must be a positive integer.`);
    }
  }

  function validateBookmarkName(tool, name, { strictPattern = true } = {}) {
    if (typeof name !== 'string' || name.length < 1) {
      throw invalidArgument(`${tool} requires a non-empty bookmark name.`);
    }
    if (strictPattern && !/^[A-Za-z_][A-Za-z0-9_]{0,39}$/.test(name)) {
      throw invalidArgument(`${tool} bookmark name must start with a letter or underscore and contain only letters, digits, and underscores.`);
    }
  }

  function validateOptionalPositiveNumber(tool, name, value) {
    if (value !== undefined && !isPositiveNumber(value)) {
      throw invalidArgument(`${tool} ${name} must be a positive number.`);
    }
  }

  function validateHeadingLevelArg(tool, level, allowZero = false) {
    if (level === undefined) return;
    const minimum = allowZero ? 0 : 1;
    if (!Number.isInteger(level) || level < minimum || level > 9) {
      throw invalidArgument(`${tool} heading_level must be an integer from ${minimum} to 9.`);
    }
  }

  function validateHeaderFooterArgs(tool, args, mutating) {
    normalizedHeaderFooterLocation(args.location);
    normalizedHeaderFooterType(args.header_footer_type);
    if (args.section_index !== undefined) requireNonNegativeInteger(tool, 'section_index', args.section_index);
    if (!mutating) return;
    const action = normalizedHeaderFooterAction(args.action);
    if ((action === 'set_text' || action === 'append_paragraph') && args.text === undefined) {
      throw invalidArgument(`${tool} ${action} requires text.`);
    }
    validateFormattingArg(tool, args.formatting);
  }

  function normalizedHeaderFooterLocation(value) {
    const normalized = String(value || '').trim().toLowerCase();
    if (normalized === 'header' || normalized === 'footer') return normalized;
    throw invalidArgument('header/footer location must be header or footer.');
  }

  function normalizedHeaderFooterType(value) {
    const normalized = String(value || 'primary').trim().toLowerCase();
    if (normalized === 'primary' || normalized === 'first_page' || normalized === 'even_pages') return normalized;
    throw invalidArgument('header_footer_type must be primary, first_page, or even_pages.');
  }

  function normalizedHeaderFooterAction(value) {
    const normalized = String(value || '').trim().toLowerCase();
    if (normalized === 'set_text' || normalized === 'append_paragraph' || normalized === 'clear') return normalized;
    throw invalidArgument('word.update_header_footer action must be set_text, append_paragraph, or clear.');
  }

  function validateFormattingArg(tool, formatting, required = false) {
    if (formatting === undefined) {
      if (required) throw invalidArgument(`${tool} requires formatting.`);
      return;
    }
    if (!formatting || typeof formatting !== 'object' || Array.isArray(formatting)) {
      throw invalidArgument(`${tool} formatting must be an object.`);
    }
    validateOptionalPositiveNumber(tool, 'formatting.font_size_pt', formatting.font_size_pt);
  }

  function validateInsertImagePreflight(args) {
    const base64 = args.image?.base64;
    if (!base64) throw invalidArgument('word.insert_image requires image.base64 after daemon preprocessing.');
    validateInsertImagePlacement(args.anchor, args.placement);
    validateOptionalPositiveNumber('word.insert_image', 'width_pt', args.width_pt);
    validateOptionalPositiveNumber('word.insert_image', 'height_pt', args.height_pt);
  }

  function validateInsertListArgs(args) {
    if (!Array.isArray(args.items) || args.items.length === 0) {
      throw invalidArgument('word.insert_list requires a non-empty items array.');
    }
    if (args.level !== undefined && (!Number.isInteger(args.level) || args.level < 0 || args.level > 8)) {
      throw invalidArgument('word.insert_list level must be an integer from 0 to 8.');
    }
    const kind = args.kind ?? 'bulleted';
    if (kind !== 'bulleted' && kind !== 'numbered') {
      throw invalidArgument('word.insert_list kind must be bulleted or numbered.');
    }
  }

  function validateReplaceTextArgs(args) {
    if (!args.find) throw invalidArgument('word.replace_text requires non-empty find text.');
    if (args.limit !== undefined && (!Number.isInteger(args.limit) || args.limit < 1)) {
      throw invalidArgument('word.replace_text limit must be a positive integer.');
    }
    const range = args.scope?.paragraph_range;
    if (range !== undefined && (!Array.isArray(range) || range.length !== 2 || !range.every((value) => Number.isInteger(value) && value >= 0) || range[0] > range[1])) {
      throw invalidArgument('word.replace_text scope.paragraph_range must be [start, end] with non-negative integer bounds.');
    }
  }

  function validateExtentToolArgs(tool, args) {
    const extent = args.extent ?? 'paragraph';
    if (extent !== 'selection') requireAnchor(tool, args.anchor);
    if (extent !== 'paragraph' && extent !== 'sentence' && extent !== 'selection') {
      throw invalidArgument(`${tool} extent must be paragraph, sentence, or selection.`);
    }
  }

  function validateUpdateTableArgs(args) {
    requireNonNegativeInteger('word.update_table', 'table_index', args.table_index);
    const action = String(args.action || '').trim().toLowerCase();
    if (!['update_cell', 'cell', 'add_row', 'add_column', 'format_cell', 'delete'].includes(action)) {
      throw invalidArgument(`Unsupported table action ${args.action}.`);
    }
    if (action === 'update_cell' || action === 'cell' || action === 'format_cell') {
      requireNonNegativeInteger('word.update_table', 'row', args.row);
      requireNonNegativeInteger('word.update_table', 'col', args.col);
    }
    if ((action === 'add_row' || action === 'add_column') && args.index !== undefined) {
      requireNonNegativeInteger('word.update_table', 'index', args.index);
    }
    if ((action === 'add_row' || action === 'add_column') && args.values !== undefined && !Array.isArray(args.values)) {
      throw invalidArgument('word.update_table values must be an array.');
    }
    validateOptionalPositiveNumber('word.update_table', 'padding_pt', args.padding_pt);
    validateFormattingArg('word.update_table', args.formatting);
  }

  function validateContentControlTargetArgs(tool, args) {
    if (args.content_control_id === undefined && !args.tag && !args.title) {
      throw invalidArgument(`${tool} requires content_control_id, tag, or title.`);
    }
    if (args.content_control_id !== undefined) requireNonNegativeInteger(tool, 'content_control_id', args.content_control_id);
  }

  function validateContentControlArgs(tool, args) {
    if (args.type !== undefined) contentControlTypeFrom(args.type);
    if (args.appearance !== undefined) contentControlAppearanceFrom(args.appearance);
    if (args.color !== undefined && !/^#[0-9a-f]{6}$/i.test(String(args.color))) {
      throw invalidArgument(`${tool} color must be a #RRGGBB value.`);
    }
  }

  function validateDeleteContentControlMode(mode) {
    if (mode !== undefined && mode !== 'keep_content' && mode !== 'delete_content') {
      throw invalidArgument('word.delete_content_control mode must be keep_content or delete_content.');
    }
  }

  function validateTrackedChangeAction(action) {
    const normalized = String(action || '').trim().toLowerCase();
    if (normalized !== 'accept' && normalized !== 'reject') {
      throw invalidArgument(`Unsupported tracked-change action ${action}.`);
    }
  }

  function validateResizeImageArgs(args) {
    if (!args.image || args.image.kind !== 'paragraph_index') {
      throw Object.assign(new Error('word.resize_image requires image.kind="paragraph_index".'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    if (!Number.isInteger(args.image.index) || args.image.index < 0) {
      throw Object.assign(new Error('word.resize_image image.index must be a non-negative integer.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    if (args.image.image_index !== undefined && (!Number.isInteger(args.image.image_index) || args.image.image_index < 0)) {
      throw Object.assign(new Error('word.resize_image image.image_index must be a non-negative integer.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    if (args.width_pt === undefined && args.height_pt === undefined) {
      throw Object.assign(new Error('word.resize_image requires width_pt or height_pt.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    if (args.width_pt !== undefined && !isPositiveNumber(args.width_pt)) {
      throw Object.assign(new Error('word.resize_image width_pt must be a positive number.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    if (args.height_pt !== undefined && !isPositiveNumber(args.height_pt)) {
      throw Object.assign(new Error('word.resize_image height_pt must be a positive number.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    if (args.preserve_aspect_ratio === false && (args.width_pt === undefined || args.height_pt === undefined)) {
      throw Object.assign(new Error('word.resize_image requires both width_pt and height_pt when preserve_aspect_ratio is false.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
  }

  function resizedImageSize(args, oldWidth, oldHeight) {
    const preserve = args.preserve_aspect_ratio !== false;
    let width = args.width_pt;
    let height = args.height_pt;
    if (preserve && width !== undefined && height === undefined) {
      height = oldHeight * (width / oldWidth);
    } else if (preserve && height !== undefined && width === undefined) {
      width = oldWidth * (height / oldHeight);
    }
    return { width, height };
  }

  function isPositiveNumber(value) {
    return typeof value === 'number' && Number.isFinite(value) && value > 0;
  }

  function isBeforeAnchor(anchor) {
    return anchor.kind === 'before_paragraph_index' || anchor.kind === 'before_text' || anchor.kind === 'start_of_document';
  }

  function insertInlinePictureWithPlacement(target, anchor, base64, placement) {
    switch (placement || 'inline') {
      case 'before_paragraph':
        return target.getRange().insertInlinePictureFromBase64(base64, Word.InsertLocation.before);
      case 'after_paragraph':
        return target.getRange().insertInlinePictureFromBase64(base64, Word.InsertLocation.after);
      case 'new_paragraph_before': {
        const paragraph = target.insertParagraph('', Word.InsertLocation.before);
        return paragraph.getRange().insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
      }
      case 'new_paragraph_after': {
        const paragraph = target.insertParagraph('', Word.InsertLocation.after);
        return paragraph.getRange().insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
      }
      case 'replace_paragraph':
        return target.getRange().insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
      default:
        if (isParagraphAnchor(anchor)) {
          const paragraph = target.insertParagraph('', isBeforeAnchor(anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
          return paragraph.getRange().insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
        }
        return target.insertInlinePictureFromBase64(base64, isBeforeAnchor(anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
    }
  }

  function validateInsertImagePlacement(anchor, placement) {
    if (!placement) return;
    if (!INSERT_IMAGE_PLACEMENTS.has(placement)) {
      throw invalidArgument(`Unsupported word.insert_image placement: ${placement}.`);
    }
    if (placement === 'selection' && anchor.kind !== 'selection') {
      throw invalidArgumentWithSuggestion('word.insert_image placement selection requires anchor.kind selection.', { placement: 'inline' });
    }
    if (isParagraphPlacement(placement) && !isParagraphAnchor(anchor)) {
      throw invalidArgumentWithSuggestion(`word.insert_image placement ${placement} requires a paragraph-resolving anchor.`, { placement: 'inline' });
    }
  }

  function isParagraphPlacement(placement) {
    return placement === 'before_paragraph' || placement === 'after_paragraph' || placement === 'new_paragraph_before' || placement === 'new_paragraph_after' || placement === 'replace_paragraph';
  }

  function isParagraphAnchor(anchor) {
    return anchor.kind === 'paragraph_index' || anchor.kind === 'before_paragraph_index' || anchor.kind === 'after_paragraph_index' || anchor.kind === 'heading';
  }

  function applyRunFormatting(font, formatting) {
    if (typeof formatting.bold === 'boolean') font.bold = formatting.bold;
    if (typeof formatting.italic === 'boolean') font.italic = formatting.italic;
    if (typeof formatting.underline === 'boolean') font.underline = formatting.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
    if (typeof formatting.strikethrough === 'boolean') font.strikeThrough = formatting.strikethrough;
    if (formatting.font_name) font.name = formatting.font_name;
    if (typeof formatting.font_size_pt === 'number') font.size = formatting.font_size_pt;
    if (formatting.color) font.color = formatting.color;
    if (formatting.highlight) font.highlightColor = formatting.highlight;
  }

  function paragraphStyleFromArgs(args) {
    if (args.style) return args.style;
    if (args.heading_level !== undefined) return headingStyleFromLevel(args.heading_level);
    return null;
  }

  function headingStyleFromLevel(level) {
    const number = Number(level);
    if (!Number.isInteger(number) || number < 0 || number > 9) {
      throw Object.assign(new Error(`Heading level ${level} is out of range.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return number === 0 ? 'Normal' : `Heading ${number}`;
  }

  function headingLevelFromStyle(style) {
    const match = String(style || '').match(/^Heading\s+([1-9])$/i);
    return match ? Number(match[1]) : 0;
  }

  function listKindFromLevelType(levelType) {
    const value = String(levelType || '').toLowerCase();
    if (value === 'number') return 'numbered';
    if (value === 'bullet') return 'bulleted';
    if (value === 'picture') return 'picture';
    return null;
  }

  function nestOutline(flat) {
    const roots = [];
    const stack = [];
    for (const node of flat) {
      const copy = { ...node, children: [] };
      while (stack.length && stack[stack.length - 1].level >= copy.level) stack.pop();
      if (stack.length) stack[stack.length - 1].children.push(copy);
      else roots.push(copy);
      stack.push(copy);
    }
    return roots;
  }

  function validateTableData(rows, cols, data) {
    if (!data) return;
    if (data.length !== rows || data.some((row) => row.length !== cols)) {
      throw invalidArgument('Table data dimensions must match rows and cols.');
    }
  }

  function blankTable(rows, cols) {
    return Array.from({ length: rows }, () => blankRow(cols));
  }

  function blankRow(cols) {
    return Array.from({ length: cols }, () => '');
  }

  function blankColumn(rows) {
    return Array.from({ length: rows }, () => '');
  }

  function mapHorizontalAlignment(value) {
    return value === 'center' ? Word.Alignment.centered : value === 'right' ? Word.Alignment.right : Word.Alignment.left;
  }

  function mapVerticalAlignment(value) {
    return value === 'center' ? Word.VerticalAlignment.center : value === 'bottom' ? Word.VerticalAlignment.bottom : Word.VerticalAlignment.top;
  }

  function trackedChangeFingerprint(change, index) {
    const date = change.date instanceof Date ? change.date.toISOString() : String(change.date || '');
    return stableHash([index, change.author || '', date, change.type || '', change.text || ''].join('\u001f'));
  }

  function stableHash(input) {
    let hash = 2166136261;
    for (let i = 0; i < input.length; i += 1) {
      hash ^= input.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, '0');
  }

  function dateToIso(value) {
    if (!value) return null;
    if (value instanceof Date) return value.toISOString();
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value) : parsed.toISOString();
  }

  async function getDocumentInfo() {
    return Word.run(async (context) => {
      const document = context.document;
      document.load('saved');
      await context.sync();
      return {
        title: Office.context.document.url ? fileName(Office.context.document.url) : 'Word document',
        url: Office.context.document.url || null,
        filename: Office.context.document.url ? fileName(Office.context.document.url) : null,
        is_dirty: typeof document.saved === 'boolean' ? !document.saved : null,
        is_read_only: null,
        is_protected: null,
        protection: { kind: null, rights: null, rights_source: 'unavailable' },
        opened_at: new Date().toISOString()
      };
    });
  }

  function probeRequirementSets() {
    const requirements = Office.context.requirements;
    return {
      WordApi: requirements.isSetSupported('WordApi', '1.3') ? '1.3' : null,
      WordApi_1_4: requirements.isSetSupported('WordApi', '1.4') ? '1.4' : null,
      WordApi_1_5: requirements.isSetSupported('WordApi', '1.5') ? '1.5' : null,
      WordApi_1_6: requirements.isSetSupported('WordApi', '1.6') ? '1.6' : null,
      WordApiDesktop_1_3: requirements.isSetSupported('WordApiDesktop', '1.3') ? '1.3' : null
    };
  }

  function mapError(error, tool, args) {
    const code = error.officeMcpCode || classifyOfficeError(error);
    const mapped = {
      office_mcp_code: code,
      message: errorMessage(error, tool),
      session_id: sessionId,
      tool,
      retriable: Boolean(error.retriable) || code === 'HOST_BUSY' || code === 'TIMEOUT',
      partial_effect: error.partialEffect || 'unknown'
    };
    const debug = officeErrorDebug(error, tool, args);
    if (debug) mapped.debug = debug;
    if (error.suggestion && typeof error.suggestion === 'object') mapped.suggestion = error.suggestion;
    return mapped;
  }

  function classifyOfficeError(error) {
    const code = String(error.code || error.name || '');
    const message = String(error.message || '');
    if (/InvalidArgument|InvalidObjectPath|InvalidSelection|ItemNotFound/i.test(code + message)) return 'INVALID_ARGUMENT';
    if (/permission|denied|IRM|rights/i.test(code + message)) return 'IRM_DENIED';
    if (/read.?only/i.test(code + message)) return 'DOCUMENT_READ_ONLY';
    if (/STALE_INDEX/i.test(code + message)) return 'STALE_INDEX';
    return 'GENERIC_FAILURE';
  }

  function errorMessage(error, tool) {
    const officeCode = officeErrorCode(error);
    if (officeCode && !error.officeMcpCode) {
      return `Word.js ${officeCode} while running ${tool || 'tool'}.`;
    }
    return error.message || String(error);
  }

  function officeErrorDebug(error, tool, args) {
    const officeCode = officeErrorCode(error);
    if (!officeCode && !error.debugInfo) return null;
    return compactObject({
      office_error_code: officeCode,
      office_error_message: safeOfficeMessage(error.message),
      error_location: safeDebugString(error.debugInfo?.errorLocation || error.traceMessages?.[0]),
      statement: safeDebugString(error.debugInfo?.statement),
      tool,
      ...safeArgumentContext(args),
      hint: officeErrorHint(officeCode, tool, args)
    });
  }

  function officeErrorCode(error) {
    const code = error.code || error.name;
    return code && code !== 'Error' ? String(code).slice(0, 80) : null;
  }

  function safeArgumentContext(args = {}) {
    const context = {};
    if (args.anchor?.kind) context.anchor_kind = String(args.anchor.kind).slice(0, 80);
    if (args.placement) context.placement = String(args.placement).slice(0, 80);
    if (args.extent) context.extent = String(args.extent).slice(0, 80);
    if (args.action) context.action = String(args.action).slice(0, 80);
    if (args.image?.mime_type) context.image_mime_type = String(args.image.mime_type).slice(0, 120);
    if (Number.isFinite(args.image?.byte_length)) context.image_byte_length = args.image.byte_length;
    if (Number.isFinite(args.width_pt)) context.width_pt = args.width_pt;
    if (Number.isFinite(args.height_pt)) context.height_pt = args.height_pt;
    if (Number.isInteger(args.index)) context.index = args.index;
    if (Number.isInteger(args.table_index)) context.table_index = args.table_index;
    if (Number.isInteger(args.row)) context.row = args.row;
    if (Number.isInteger(args.col)) context.col = args.col;
    if (Number.isInteger(args.content_control_id)) context.content_control_id = args.content_control_id;
    return context;
  }

  function safeOfficeMessage(message) {
    if (!message) return null;
    const value = String(message);
    if (looksSensitive(value)) return null;
    return value.slice(0, 240);
  }

  function safeDebugString(value) {
    if (!value) return null;
    const text = String(value);
    if (looksSensitive(text)) return null;
    return text.slice(0, 180);
  }

  function looksSensitive(value) {
    return /base64|data:image|[A-Za-z0-9+/]{80,}={0,2}/.test(value);
  }

  function officeErrorHint(officeCode, tool, args = {}) {
    if (/InvalidArgument/i.test(officeCode || '') && tool === 'word.insert_image') {
      if (isParagraphAnchor(args.anchor) && !args.placement) return 'Use an explicit paragraph placement such as new_paragraph_after for paragraph anchors.';
      return 'Check image payload, placement, anchor kind, and dimensions.';
    }
    if (/InvalidObjectPath/i.test(officeCode || '')) return 'Retry after re-reading the target object; the Office.js object path became invalid.';
    return null;
  }

  function compactObject(value) {
    return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== null && item !== undefined));
  }

  function reply(id, result) {
    return replyJsonRpc(socket, id, result);
  }

  function send(message) {
    sendJsonRpc(socket, message);
  }

  function recordDiagnostic(event, fields = {}) {
    try {
      const payload = {
        host_app: 'word',
        addin_version: ADDIN_VERSION,
        event: String(event || '').slice(0, 120),
        fields: logger.redactFields ? logger.redactFields(fields) : fields,
        at: new Date().toISOString()
      };
      fetch('/addin/diagnostics', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        keepalive: true
      }).catch(() => {});
    } catch {
      // Diagnostics must never break the task pane runtime.
    }
  }

  function scheduleReconnect() {
    reconnectAttempt += 1;
    const delay = reconnectDelay(reconnectAttempt);
    const seconds = Math.ceil(delay / 1000);
    connectionDetailEl.textContent = `Retrying in ${seconds}s.`;
    setConnectionState('reconnecting', `Reconnecting in ${seconds}s`);
    reconnectTimer = setTimeout(connect, delay);
  }

  function setConnectionState(state, label) {
    setSharedConnectionState({ badge: connectionBadgeEl, detail: connectionDetailEl, announcer: announcerEl }, state, label);
  }

  function announce(message) {
    if (announcerEl) announcerEl.textContent = message;
  }

  function setStatus(label) {
    connectionDetailEl.textContent = label;
    setConnectionState('failed', label);
  }

  function renderStaticState() {
    renderStaticMetadata({ session: sessionEl, daemon: daemonEl, serverVersion: serverVersionEl, protocolVersion: protocolVersionEl, hostPlatform: hostPlatformEl }, { sessionId, endpoint: configuredEndpoint(), serverInfo, protocolVersion: PROTOCOL_VERSION, defaultHost: 'Word' });
    renderToolModeControl();
    renderToolSummary();
    renderHistory();
  }

  function renderToolModeControl() {
    renderSharedToolModeControl(toolModeControlEl, toolPermissionMode);
  }

  function renderToolSummary() {
    const effective = effectiveTools();
    const openGroups = new Set([...toolListEl.querySelectorAll('[data-tool-group]')]
      .filter((input) => input.closest('details')?.open)
      .map((input) => input.dataset.toolGroup));
    renderToolModeControl();
    toolCountEl.textContent = `${effective.length}/${AVAILABLE_TOOLS.length}`;
    toolListEl.textContent = '';
    for (const group of TOOL_GROUPS) {
      const tools = group.tools.filter((tool) => AVAILABLE_TOOLS.includes(tool));
      if (tools.length === 0) continue;
      const allowedInGroup = tools.filter((tool) => isToolAllowedByMode(tool));
      const enabledInGroup = tools.filter((tool) => effective.includes(tool));
      const groupEl = document.createElement('details');
      groupEl.className = 'tool-group';
      groupEl.open = openGroups.has(group.label);
      groupEl.innerHTML = [
        '<summary class="tool-group-title">',
        `<span>${escapeHtml(group.label)}</span>`,
        `<span class="tool-group-count">${enabledInGroup.length}/${tools.length}</span>`,
        `<input class="group-toggle" type="checkbox" role="switch" data-tool-group="${escapeHtml(group.label)}" aria-label="Toggle ${escapeHtml(group.label)} tools" ${allowedInGroup.length > 0 && enabledInGroup.length === allowedInGroup.length ? 'checked' : ''} ${allowedInGroup.length === 0 ? 'disabled' : ''} />`,
        '</summary>',
        `<div class="tool-permission-list">${tools.map(toolControlMarkup).join('')}</div>`
      ].join('');
      toolListEl.appendChild(groupEl);
    }
    toolListEl.querySelectorAll('[data-tool]').forEach((input) => {
      bindDetailsControl(input, handleToolPermissionChange);
    });
    toolListEl.querySelectorAll('[data-tool-group]').forEach((input) => {
      bindDetailsControl(input, handleToolGroupPermissionChange);
    });
  }

  function toolControlMarkup(tool) {
    const metadata = TOOL_METADATA.get(tool) || { category: 'Tools', sideEffect: 'read', description: 'Office tool.' };
    const id = `toolPermission-${tool.replace(/[^a-z0-9_-]/gi, '-')}`;
    const modeAllowed = isToolAllowedByMode(tool);
    const checked = isToolEnabled(tool) && modeAllowed;
    return [
      `<label class="tool-permission-row${metadata.sideEffect === 'mutating' || metadata.sideEffect === 'destructive' ? ' is-mutating' : ''}${modeAllowed ? '' : ' is-disabled'}" for="${id}">`,
      '<span class="tool-permission-main">',
      '<span class="tool-permission-title">',
      `<span class="tool-permission-name">${escapeHtml(tool)}</span>`,
      `<span class="side-effect-pill ${metadata.sideEffect === 'mutating' || metadata.sideEffect === 'destructive' ? 'mutating' : 'read'}">${escapeHtml(metadata.sideEffect)}</span>`,
      '</span>',
      `<span class="tool-permission-meta">${escapeHtml(metadata.description)}</span>`,
      '</span>',
      `<input id="${id}" class="tool-toggle" type="checkbox" role="switch" data-tool="${escapeHtml(tool)}" aria-label="Toggle ${escapeHtml(tool)}" ${checked ? 'checked' : ''} ${modeAllowed ? '' : 'disabled aria-disabled="true"'} />`,
      '</label>'
    ].join('');
  }

  function handleToolPermissionChange(event) {
    const tool = event.currentTarget.dataset.tool;
    if (!tool) return;
    toolPermissions[tool] = event.currentTarget.checked;
    saveToolPermissions();
    renderToolSummary();
    sendSessionToolUpdate();
  }

  function handleToolGroupPermissionChange(event) {
    const group = TOOL_GROUPS.find((candidate) => candidate.label === event.currentTarget.dataset.toolGroup);
    if (!group) return;
    const enabled = event.currentTarget.checked;
    for (const tool of group.tools) {
      if (AVAILABLE_TOOLS.includes(tool) && isToolAllowedByMode(tool)) toolPermissions[tool] = enabled;
    }
    saveToolPermissions();
    renderToolSummary();
    sendSessionToolUpdate();
  }

  function effectiveTools() {
    return AVAILABLE_TOOLS.filter((tool) => isToolSupported(tool) && isToolEnabled(tool) && isToolAllowedByMode(tool));
  }

  function isToolSupported(tool) {
    if (tool === 'word.update_page_setup') return supportsWordApiDesktop13();
    return true;
  }

  function isToolAllowedByMode(tool) {
    const sideEffect = TOOL_METADATA.get(tool)?.sideEffect || 'read';
    return isToolAllowedByCapabilityMode(toolPermissionMode, sideEffect);
  }

  function handleToolModeChange(event) {
    const mode = event.currentTarget.dataset.toolMode;
    if (!['read', 'write', 'all'].includes(mode) || mode === toolPermissionMode) return;
    toolPermissionMode = mode;
    saveToolPermissionMode();
    renderToolSummary();
    sendSessionToolUpdate();
  }

  function isToolEnabled(tool) {
    return toolPermissions[tool] !== false;
  }

  function loadToolPermissions() {
    try {
      const parsed = JSON.parse(window.localStorage?.getItem(TOOL_PERMISSION_STORAGE_KEY) || '{}');
      if (!parsed || typeof parsed !== 'object') return {};
      return Object.fromEntries(AVAILABLE_TOOLS.map((tool) => [tool, parsed[tool] !== false]));
    } catch {
      return {};
    }
  }

  function loadToolPermissionMode() {
    try {
      const mode = window.localStorage?.getItem(TOOL_PERMISSION_MODE_STORAGE_KEY) || 'all';
      return ['read', 'write', 'all'].includes(mode) ? mode : 'all';
    } catch {
      return 'all';
    }
  }

  function saveToolPermissions() {
    window.localStorage?.setItem(TOOL_PERMISSION_STORAGE_KEY, JSON.stringify(toolPermissions));
  }

  function saveToolPermissionMode() {
    window.localStorage?.setItem(TOOL_PERMISSION_MODE_STORAGE_KEY, toolPermissionMode);
  }

  function sendSessionToolUpdate() {
    if (!sessionAnnounced) return;
    send(sessionUpdatedNotification({
      session_id: sessionId,
      patch: { available_tools: effectiveTools() }
    }));
  }

  function renderDocumentState() {
    if (!documentInfo) return;
    documentTitleEl.textContent = documentInfo.title || documentInfo.filename || 'Word Document';
    protectionEl.textContent = protectionLabel(documentInfo);
    documentStateEl.textContent = documentStateLabel(documentInfo);
  }

  function startTask(requestId, tool, args, timeoutMs) {
    taskStore.start(requestId, tool, args, timeoutMs);
    renderCurrentTask();
  }

  function finishTask(requestId, status, elapsedMs, error) {
    const task = taskStore.finish(requestId, status, elapsedMs, error);
    if (!task) return;
    renderCurrentTask();
    renderHistory();
    announce(`${task.tool} ${status}`);
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
    const elapsed = Math.max(0, Date.now() - currentTask.startedAt);
    currentTaskEl.innerHTML = taskMarkup({
      tool: currentTask.tool,
      requestId: currentTask.requestId,
      status: 'running',
      elapsedMs: elapsed,
      userIntent: currentTask.userIntent,
      deadlineAt: currentTask.deadlineAt,
      cancelRequested: currentTask.cancelRequested,
      error: null
    });
  }

  function renderHistory() {
    const { history, historyLimit } = taskStore.snapshot();
    historyCountEl.textContent = `${history.length} / ${historyLimit}`;
    historyListEl.textContent = '';
    for (const task of history) {
      const item = document.createElement('li');
      item.innerHTML = `<article class="task-card">${taskMarkup(task)}</article>`;
      historyListEl.appendChild(item);
    }
  }

  function taskMarkup(task) {
    const tone = taskStatusClass(task.status || 'running');
    const metadata = taskMetadataMarkup(task, { escapeHtml, formatTime, redactText, valueLabel: boolLabel });
    const commandId = commandIdMarkup(task.requestId, { escapeHtml });
    return [
      '<div class="task-title">',
      `<span>${escapeHtml(task.tool)}</span>`,
      `<span class="status-badge ${tone}">${escapeHtml(taskStatusLabel(task.status || 'running'))}</span>`,
      '</div>',
      commandId,
      `<div class="task-meta">${formatDuration(task.elapsedMs)}</div>`,
      metadata
    ].join('');
  }


  function saveEndpointOverride(event) {
    event.preventDefault();
    connectionDetailEl.textContent = 'None';
    const value = endpointInputEl.value.trim();
    try {
      validateEndpoint(value);
      storeEndpointOverride(value);
      endpointDirty = false;
      saveEndpointEl.disabled = true;
      saveEndpointEl.setAttribute('aria-busy', 'true');
      setTimeout(() => {
        saveEndpointEl.disabled = false;
        saveEndpointEl.removeAttribute('aria-busy');
        if (socket) socket.close(1000, 'Endpoint changed');
        connect();
      }, 0);
    } catch (error) {
      setConnectionState('failed', error.message || 'Enter a valid wss:// endpoint.');
      endpointInputEl.focus();
    }
  }


  async function handleMetadataCopy(event) {
    await copyMetadataValue(event, { document, navigator, announcer: announcerEl, logger });
  }
})();




