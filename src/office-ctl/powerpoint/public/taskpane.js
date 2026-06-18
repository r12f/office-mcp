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
    sessionUpdatedNotification,
    sendJsonRpc,
    validateEndpoint
  } = window.OfficeCtlAddinChannel;
  const { AddinLogger } = window.OfficeCtlLogger;
  const { TaskHistoryStore } = window.OfficeCtlTaskHistory;

  const AVAILABLE_TOOLS = [
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
  const TOOL_PERMISSION_STORAGE_KEY = `office-mcp.powerpoint.tool-permissions.${sessionId}`;
  const logger = new AddinLogger({ redactText });
  const taskStore = new TaskHistoryStore({ redactText });
  let socket;
  let reconnectTimer;
  let reconnectAttempt = 0;
  let endpointDirty = false;
  let suppressNextSettingsClick = false;
  let serverInfo = { serverVersion: 'Unknown', protocolVersion: PROTOCOL_VERSION };
  let documentInfo = null;
  let toolPermissions = loadToolPermissions();
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
    setConnectionState('connecting', 'Connecting\u2026');
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
    setConnectionState('connecting', 'Connecting\u2026');
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
    setConnectionState('reconnecting', 'Reconnecting\u2026');
    reconnectAttempt = 0;
    clearTimeout(reconnectTimer);
    reconnectTimer = setTimeout(connect, 250);
    return true;
  }

  function register() {
    reconnectAttempt = 0;
    setConnectionState('connecting', 'Registering\u2026');
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
      available_tools: effectiveTools(),
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
    const args = message.params?.args || {};
    taskStore.start(requestId, tool, message.params || {}, message.params?.timeout_ms);
    renderCurrentTask();
    try {
      if (!isToolEnabled(tool)) throw toolDisabledError(tool);
      if (taskStore.isCancelled(requestId)) throw cancelledError(tool);
      let data;
      switch (tool) {
        case 'powerpoint.add_slide':
          data = await addSlide(args);
          break;
        case 'powerpoint.replace_text':
          data = await replaceText(args);
          break;
        case 'powerpoint.insert_image':
          data = await insertImage(args);
          break;
        case 'powerpoint.apply_layout':
          data = await applyLayout(args);
          break;
        case 'powerpoint.export_pdf':
          data = await exportPdf(args);
          break;
        default:
          throw Object.assign(new Error(`Unsupported tool ${tool}`), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE', partialEffect: 'none' });
      }
      if (taskStore.consumeCancellation(requestId)) throw cancelledError(tool);
      const elapsedMs = Math.round(performance.now() - started);
      taskStore.finish(requestId, 'success', elapsedMs);
      renderCurrentTask();
      renderHistory();
      reply(message.id, { ok: true, data, elapsed_ms: elapsedMs });
    } catch (error) {
      const mapped = mapError(error, tool);
      taskStore.finish(requestId, mapped.office_mcp_code === 'CANCELLED' ? 'cancelled' : 'failure', Math.round(performance.now() - started), mapped);
      renderCurrentTask();
      renderHistory();
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

  async function addSlide(args) {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const beforeCount = slides.getCount();
      await context.sync();
      slides.add(slideOptions(args));
      await context.sync();
      slides.load('items/id,index');
      await context.sync();
      const added = slides.items.find((slide) => slide.index === beforeCount.value) || slides.items[slides.items.length - 1];
      if (!added) throw Object.assign(new Error('PowerPoint did not report the added slide.'), { officeMcpCode: 'HOST_ERROR', partialEffect: 'unknown' });
      const title = stringArg(args, 'title');
      const content = stringArg(args, 'content');
      if (title || content) {
        if (title) added.shapes.addTextBox(title, shapeOptions(args.title_box, { left: 48, top: 48, width: 600, height: 60 }));
        if (content) added.shapes.addTextBox(content, shapeOptions(args.content_box, { left: 72, top: 132, width: 576, height: 260 }));
        await context.sync();
      }
      return { slide_id: added.id, slide_index: added.index, added: true };
    });
  }

  async function replaceText(args) {
    const search = requiredString(args, 'search', 'powerpoint.replace_text requires search text.');
    const replacement = requiredString(args, 'replacement', 'powerpoint.replace_text requires replacement text.');
    const matchCase = Boolean(args.match_case);
    return PowerPoint.run(async (context) => {
      const slides = targetSlides(context, args);
      if (slides.load) slides.load('items/id,index');
      for (const slide of slides.items || []) slide.load('id,index');
      await context.sync();
      for (const slide of slides.items) {
        slide.shapes.load('items/id,textFrame/hasText,textFrame/textRange/text');
      }
      await context.sync();
      const touchedShapes = [];
      let replacements = 0;
      for (const slide of slides.items) {
        for (const shape of slide.shapes.items || []) {
          const range = shape.textFrame?.textRange;
          if (!shape.textFrame?.hasText || !range || typeof range.text !== 'string') continue;
          const nextText = replaceAllText(range.text, search, replacement, matchCase);
          if (nextText === range.text) continue;
          replacements += countMatches(range.text, search, matchCase);
          range.text = nextText;
          touchedShapes.push({ slide_id: slide.id, shape_id: shape.id });
        }
      }
      await context.sync();
      return { replacements, touched_shapes: touchedShapes };
    });
  }

  async function insertImage(args) {
    const base64 = imageBase64(args);
    await officeAsync((callback) => Office.context.document.setSelectedDataAsync(base64, imageInsertOptions(args), callback));
    return { inserted_image: true, placement: 'selection', width: numberOrNull(args.width), height: numberOrNull(args.height) };
  }

  async function applyLayout(args) {
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const layout = await resolveLayout(context, args);
      slide.applyLayout(layout);
      slide.load('id,index');
      layout.load('id,name,type');
      await context.sync();
      return { slide_id: slide.id, slide_index: slide.index, layout_id: layout.id, layout_name: layout.name, layout_type: layout.type };
    });
  }

  async function exportPdf(args) {
    const file = await officeAsync((callback) => Office.context.document.getFileAsync(Office.FileType.Pdf, { sliceSize: positiveInteger(args.slice_size, 4 * 1024 * 1024) }, callback));
    try {
      const chunks = [];
      for (let index = 0; index < file.sliceCount; index += 1) {
        const slice = await officeAsync((callback) => file.getSliceAsync(index, callback));
        chunks.push(sliceDataToBytes(slice.data));
      }
      return { mime_type: 'application/pdf', base64: bytesToBase64(concatBytes(chunks)), size: file.size, slice_count: file.sliceCount };
    } finally {
      await officeAsync((callback) => file.closeAsync(callback)).catch((error) => logger.warn('pdf.close.failed', error));
    }
  }

  function slideOptions(args) {
    const options = {};
    const layoutId = stringArg(args, 'layout_id') || stringArg(args, 'layoutId');
    const slideMasterId = stringArg(args, 'slide_master_id') || stringArg(args, 'slideMasterId');
    if (layoutId) options.layoutId = layoutId;
    if (slideMasterId) options.slideMasterId = slideMasterId;
    return Object.keys(options).length ? options : undefined;
  }

  function shapeOptions(input, defaults) {
    const source = input && typeof input === 'object' ? input : {};
    const options = { ...defaults };
    for (const key of ['left', 'top', 'width', 'height']) {
      const value = numberOrNull(source[key]);
      if (value !== null) options[key] = value;
    }
    return options;
  }

  function targetSlides(context, args) {
    if (Array.isArray(args.slide_ids) && args.slide_ids.length > 0) {
      return { items: args.slide_ids.map((id) => context.presentation.slides.getItem(String(id))) };
    }
    if (Array.isArray(args.slide_indexes) && args.slide_indexes.length > 0) {
      return { items: args.slide_indexes.map((index) => context.presentation.slides.getItemAt(Number(index))) };
    }
    const slideId = stringArg(args, 'slide_id') || stringArg(args, 'slideId');
    if (slideId) return { items: [context.presentation.slides.getItem(slideId)] };
    const index = numberOrNull(args.slide_index ?? args.slideIndex);
    if (index !== null) return { items: [context.presentation.slides.getItemAt(index)] };
    return context.presentation.slides;
  }

  function targetSlide(context, args) {
    const slideId = stringArg(args, 'slide_id') || stringArg(args, 'slideId');
    if (slideId) return context.presentation.slides.getItem(slideId);
    const index = numberOrNull(args.slide_index ?? args.slideIndex);
    if (index !== null) return context.presentation.slides.getItemAt(index);
    return context.presentation.getSelectedSlides().getItemAt(0);
  }

  async function resolveLayout(context, args) {
    const layoutId = stringArg(args, 'layout_id') || stringArg(args, 'layoutId');
    const layoutName = stringArg(args, 'layout_name') || stringArg(args, 'layoutName');
    const layoutType = stringArg(args, 'layout_type') || stringArg(args, 'layoutType') || stringArg(args, 'layout');
    if (!layoutId && !layoutName && !layoutType) {
      throw Object.assign(new Error('powerpoint.apply_layout requires layout_id, layout_name, or layout_type.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    const masters = context.presentation.slideMasters;
    masters.load('items/id,name,layouts/items/id,name,type');
    await context.sync();
    const normalizedName = layoutName ? layoutName.toLowerCase() : '';
    const normalizedType = layoutType ? layoutType.toLowerCase() : '';
    for (const master of masters.items || []) {
      for (const layout of master.layouts.items || []) {
        if (layoutId && layout.id === layoutId) return master.layouts.getItem(layout.id);
        if (normalizedName && String(layout.name || '').toLowerCase() === normalizedName) return master.layouts.getItem(layout.id);
        if (normalizedType && String(layout.type || '').toLowerCase() === normalizedType) return master.layouts.getItem(layout.id);
      }
    }
    throw Object.assign(new Error('Requested PowerPoint slide layout was not found.'), { officeMcpCode: 'NOT_FOUND', partialEffect: 'none' });
  }

  function imageBase64(args) {
    const raw = requiredString(args, 'base64', 'powerpoint.insert_image requires base64 image data.');
    const value = raw.includes(',') ? raw.split(',').pop() : raw;
    if (!/^[A-Za-z0-9+/=\s]+$/.test(value)) {
      throw Object.assign(new Error('powerpoint.insert_image base64 data is invalid.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return value.replace(/\s+/g, '');
  }

  function imageInsertOptions(args) {
    const options = { coercionType: Office.CoercionType.Image };
    const mapping = { left: 'imageLeft', top: 'imageTop', width: 'imageWidth', height: 'imageHeight' };
    for (const [source, target] of Object.entries(mapping)) {
      const value = numberOrNull(args[source]);
      if (value !== null) options[target] = value;
    }
    return options;
  }

  function officeAsync(start) {
    return new Promise((resolve, reject) => {
      start((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
          return;
        }
        reject(Object.assign(new Error(result.error?.message || 'Office async operation failed.'), {
          officeMcpCode: 'HOST_ERROR',
          partialEffect: 'unknown'
        }));
      });
    });
  }

  function sliceDataToBytes(data) {
    if (typeof data === 'string') return Uint8Array.from(atob(data), (char) => char.charCodeAt(0));
    return Uint8Array.from(Array.isArray(data) ? data : Array.from(data || []));
  }

  function concatBytes(chunks) {
    const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const bytes = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.length;
    }
    return bytes;
  }

  function bytesToBase64(bytes) {
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return btoa(binary);
  }

  function requiredString(args, key, message) {
    const value = stringArg(args, key);
    if (!value) throw Object.assign(new Error(message), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    return value;
  }

  function stringArg(args, key) {
    const value = args?.[key];
    return typeof value === 'string' ? value.trim() : '';
  }

  function numberOrNull(value) {
    if (value === undefined || value === null || value === '') return null;
    const number = Number(value);
    if (!Number.isFinite(number)) return null;
    return number;
  }

  function positiveInteger(value, fallback) {
    const number = Number(value || fallback);
    return Number.isInteger(number) && number > 0 ? number : fallback;
  }

  function replaceAllText(source, search, replacement, matchCase) {
    if (matchCase) return source.split(search).join(replacement);
    return source.replace(new RegExp(escapeRegExp(search), 'gi'), replacement);
  }

  function countMatches(source, search, matchCase) {
    const flags = matchCase ? 'g' : 'gi';
    return Array.from(source.matchAll(new RegExp(escapeRegExp(search), flags))).length;
  }

  function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
    for (const version of ['1.1', '1.2', '1.3', '1.4', '1.8', '1.10']) {
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
    const effective = effectiveTools();
    toolCountEl.textContent = `Enabled ${effective.length} of ${AVAILABLE_TOOLS.length}`;
    toolListEl.innerHTML = TOOL_GROUPS.map((group) => {
      const enabledInGroup = group.tools.filter((tool) => isToolEnabled(tool));
      const rows = group.tools.map((tool) => toolControlMarkup(tool)).join('');
      return `<details class="tool-group" open><summary class="tool-group-title"><span>${escapeHtml(group.label)}</span><span>Enabled ${enabledInGroup.length} of ${group.tools.length}</span></summary>${rows}</details>`;
    }).join('');
    for (const checkbox of toolListEl.querySelectorAll('input[data-tool]')) {
      checkbox.addEventListener('change', () => updateToolPermission(checkbox.dataset.tool, checkbox.checked));
    }
  }

  function toolControlMarkup(tool) {
    const meta = TOOL_METADATA.get(tool) || { sideEffect: 'unknown', description: 'No metadata.' };
    const enabled = isToolEnabled(tool);
    const rowStateClass = `${enabled ? '' : ' is-disabled'}${meta.sideEffect === 'mutating' ? ' is-mutating' : ''}`;
    const sideEffectClass = meta.sideEffect === 'mutating' ? ' mutating' : '';
    const id = `toolPermission-${tool.replace(/[^a-z0-9_-]/gi, '-')}`;
    return `<label class="tool-permission-row${rowStateClass}" for="${id}"><input id="${id}" class="tool-toggle" type="checkbox" data-tool="${escapeHtml(tool)}" ${enabled ? 'checked' : ''} /><span class="tool-permission-main"><span class="tool-permission-title"><span class="tool-permission-name">${escapeHtml(tool)}</span><span class="side-effect-pill${sideEffectClass}">${escapeHtml(titleCase(meta.sideEffect))}</span></span><span class="tool-permission-meta">${escapeHtml(meta.description)}</span></span></label>`;
  }

  function effectiveTools() {
    return AVAILABLE_TOOLS.filter((tool) => isToolEnabled(tool));
  }

  function isToolEnabled(tool) {
    return toolPermissions[tool] !== false;
  }

  function updateToolPermission(tool, enabled) {
    if (!AVAILABLE_TOOLS.includes(tool)) return;
    toolPermissions = { ...toolPermissions, [tool]: Boolean(enabled) };
    saveToolPermissions();
    renderToolSummary();
    if (sessionAnnounced) {
      send(sessionUpdatedNotification({
        session_id: sessionId,
        patch: { available_tools: effectiveTools() }
      }));
    }
  }

  function loadToolPermissions() {
    try {
      const raw = window.localStorage?.getItem(TOOL_PERMISSION_STORAGE_KEY);
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (_error) {
      return {};
    }
  }

  function saveToolPermissions() {
    try {
      window.localStorage?.setItem(TOOL_PERMISSION_STORAGE_KEY, JSON.stringify(toolPermissions));
    } catch (error) {
      logger.warn('tool_permissions.save.failed', error);
    }
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
    toolListEl.classList.toggle('is-editing-tools', opening);
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
    saveEndpointEl.textContent = 'Saving\u2026';
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
