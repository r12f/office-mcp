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
    middleTruncate,
    officeHostSummary,
    renderRuntimeVersions,
    setCopyableMetadata,
    statusClass,
    taskMetadataMarkup
  } = window.OfficeCtlMainUi;
  const CONNECT_TIMEOUT_MS = 8000;
  const AVAILABLE_TOOLS = [
    'word.get_text',
    'word.get_outline',
    'word.get_paragraph',
    'word.find_text',
    'word.get_selection',
    'word.insert_paragraph',
    'word.insert_image',
    'word.insert_table',
    'word.insert_page_break',
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
    { label: 'Document & structure', tools: ['word.get_text', 'word.get_outline', 'word.insert_page_break', 'word.save'] },
    { label: 'Range & selection', tools: ['word.get_selection', 'word.find_text', 'word.replace_text', 'word.delete_range', 'word.apply_formatting', 'word.apply_style'] },
    { label: 'Paragraphs & lists', tools: ['word.get_paragraph', 'word.insert_paragraph', 'word.update_paragraph', 'word.insert_list'] },
    { label: 'Tables', tools: ['word.read_table', 'word.update_table'] },
    { label: 'Media', tools: ['word.insert_image'] },
    { label: 'Content controls', tools: ['word.list_content_controls', 'word.insert_content_control', 'word.update_content_control', 'word.delete_content_control'] },
    { label: 'Review', tools: ['word.add_comment', 'word.resolve_comment', 'word.update_tracked_change'] }
  ];
  const TOOL_METADATA = new Map([
    ['word.get_text', { category: 'Document & structure', sideEffect: 'read', description: 'Read document text by paragraph range.' }],
    ['word.get_outline', { category: 'Document & structure', sideEffect: 'read', description: 'Read heading outline and structure.' }],
    ['word.get_paragraph', { category: 'Paragraphs & lists', sideEffect: 'read', description: 'Read a single paragraph by index.' }],
    ['word.find_text', { category: 'Range & selection', sideEffect: 'read', description: 'Find text matches in the document body.' }],
    ['word.get_selection', { category: 'Range & selection', sideEffect: 'read', description: 'Read the current selection.' }],
    ['word.insert_paragraph', { category: 'Paragraphs & lists', sideEffect: 'mutating', description: 'Insert a paragraph near an anchor.' }],
    ['word.insert_image', { category: 'Media', sideEffect: 'mutating', description: 'Insert an image into the document.' }],
    ['word.insert_table', { category: 'Tables', sideEffect: 'mutating', description: 'Insert a table with provided values.' }],
    ['word.insert_page_break', { category: 'Document & structure', sideEffect: 'mutating', description: 'Insert a page break.' }],
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
  const endpointErrorEl = document.getElementById('endpointError');
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
      instance_id: instanceId,
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
        case 'word.get_selection':
          data = await getSelection(args);
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
        case 'word.insert_page_break':
          data = await insertPageBreak(args);
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
      const mapped = mapError(error, tool);
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
      let picture;
      if (args.anchor.kind === 'start_of_document') {
        picture = context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        picture = context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        picture = target.insertInlinePictureFromBase64(base64, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      if (args.alt_text) picture.altTextDescription = args.alt_text;
      if (typeof args.width_pt === 'number') picture.width = args.width_pt;
      if (typeof args.height_pt === 'number') picture.height = args.height_pt;
      await context.sync();
      return { inserted: true, byte_length: args.image.byte_length ?? null, mime_type: args.image.mime_type ?? null };
    });
  }


  async function insertPageBreak(args) {
    return Word.run(async (context) => {
      if (args.anchor.kind === 'start_of_document') {
        context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        target.insertBreak(Word.BreakType.page, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      await context.sync();
      return { inserted: true, break_type: 'page' };
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
      if (args.dry_run) return { replaced_count: 0, matches, dry_run: true, skipped_count: filtered.skipped_count };
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
      if (values.length !== columnCount) throw Object.assign(new Error('Row values length must match table column count.'), { officeMcpCode: 'INVALID_ARGUMENT' });
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
      if (values.length !== rowCount) throw Object.assign(new Error('Column values length must match table row count.'), { officeMcpCode: 'INVALID_ARGUMENT' });
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

  function isBeforeAnchor(anchor) {
    return anchor.kind === 'before_paragraph_index' || anchor.kind === 'before_text' || anchor.kind === 'start_of_document';
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
      throw Object.assign(new Error('Table data dimensions must match rows and cols.'), { officeMcpCode: 'INVALID_ARGUMENT' });
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
      WordApi_1_6: requirements.isSetSupported('WordApi', '1.6') ? '1.6' : null
    };
  }

  function mapError(error, tool) {
    const code = error.officeMcpCode || classifyOfficeError(error);
    return {
      office_mcp_code: code,
      message: error.message || String(error),
      session_id: sessionId,
      tool,
      retriable: Boolean(error.retriable) || code === 'HOST_BUSY' || code === 'TIMEOUT',
      partial_effect: error.partialEffect || 'unknown'
    };
  }

  function classifyOfficeError(error) {
    const code = String(error.code || error.name || '');
    const message = String(error.message || '');
    if (/permission|denied|IRM|rights/i.test(code + message)) return 'IRM_DENIED';
    if (/read.?only/i.test(code + message)) return 'DOCUMENT_READ_ONLY';
    if (/STALE_INDEX/i.test(code + message)) return 'STALE_INDEX';
    return 'GENERIC_FAILURE';
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
    connectionBadgeEl.textContent = label;
    connectionBadgeEl.className = `status-badge ${statusClass(state)}`;
    if (state === 'connected') connectionDetailEl.textContent = 'None';
    announce(label);
  }

  function announce(message) {
    if (announcerEl) announcerEl.textContent = message;
  }

  function setStatus(label) {
    connectionDetailEl.textContent = label;
    setConnectionState('failed', label);
  }

  function renderStaticState() {
    setCopyableMetadata(sessionEl, sessionId);
    setCopyableMetadata(daemonEl, configuredEndpoint());
    renderRuntimeVersions(serverVersionEl, protocolVersionEl, serverInfo, PROTOCOL_VERSION);
    hostPlatformEl.textContent = officeHostSummary('Word');
    renderToolModeControl();
    renderToolSummary();
    renderHistory();
  }

  function renderToolModeControl() {
    toolModeControlEl?.querySelectorAll('[data-tool-mode]').forEach((button) => {
      const selected = button.dataset.toolMode === toolPermissionMode;
      button.setAttribute('aria-checked', selected ? 'true' : 'false');
    });
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
    return AVAILABLE_TOOLS.filter((tool) => isToolEnabled(tool) && isToolAllowedByMode(tool));
  }

  function isToolAllowedByMode(tool) {
    const sideEffect = TOOL_METADATA.get(tool)?.sideEffect || 'read';
    if (toolPermissionMode === 'all') return true;
    if (toolPermissionMode === 'write') return sideEffect !== 'destructive';
    return sideEffect === 'read';
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

  function protectionLabel(info) {
    if (info.is_protected === true || info.protection?.kind) return info.protection?.kind || 'Protected';
    return 'Not protected';
  }

  function documentStateLabel(info) {
    if (info.is_read_only === true) return 'Read-only';
    if (info.is_protected === true || info.protection?.kind) return `Protected${info.protection?.kind ? `: ${info.protection.kind}` : ''}`;
    if (info.is_dirty === true) return 'Editable, unsaved changes';
    return 'Editable';
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
    const tone = task.status === 'success' ? 'status-success' : task.status === 'running' ? 'status-warning' : task.status === 'cancelled' ? 'status-neutral' : 'status-danger';
    const metadata = taskMetadataMarkup(task, { escapeHtml, formatTime, redactText, valueLabel: boolLabel });
    const commandId = commandIdMarkup(task.requestId, { escapeHtml });
    return [
      '<div class="task-title">',
      `<span>${escapeHtml(task.tool)}</span>`,
      `<span class="status-badge ${tone}">${escapeHtml(titleCase(task.status))}</span>`,
      '</div>',
      commandId,
      `<div class="task-meta">${formatDuration(task.elapsedMs)}</div>`,
      metadata
    ].join('');
  }


  function saveEndpointOverride(event) {
    event.preventDefault();
    endpointErrorEl.textContent = '';
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
      endpointErrorEl.textContent = error.message || 'Enter a valid wss:// endpoint.';
      endpointInputEl.focus();
    }
  }


  async function handleMetadataCopy(event) {
    await copyMetadataValue(event, { document, navigator, announcer: announcerEl, logger, fallbackCopy });
  }

  function fallbackCopy(value) {
    const area = document.createElement('textarea');
    area.value = value;
    area.setAttribute('readonly', '');
    area.style.position = 'fixed';
    area.style.opacity = '0';
    document.body.appendChild(area);
    area.select();
    document.execCommand('copy');
    area.remove();
  }
})();




