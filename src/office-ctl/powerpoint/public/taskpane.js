(() => {
  const ADDIN_VERSION = '0.1.4';
  const PROTOCOL_VERSION = '1.0';
  const POWERPOINT_FILE_EXPORT_TIMEOUT_MS = 10000;
  const { boolLabel, escapeHtml, fileName, formatDuration, formatTime, titleCase, redactText } = window.OfficeCtlCommon;
  const { bindDetailsControl, commandIdMarkup, copyMetadataValue, middleTruncate, officeHostSummary, renderRuntimeVersions, setCopyableMetadata, statusClass, taskMetadataMarkup } = window.OfficeCtlMainUi;
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

  const AVAILABLE_TOOLS = [
    'powerpoint.get_presentation_info',
    'powerpoint.get_active_view',
    'powerpoint.export_file',
    'powerpoint.update_tags',
    'powerpoint.list_slides',
    'powerpoint.add_slide',
    'powerpoint.update_slide',
    'powerpoint.delete_slide',
    'powerpoint.move_slide',
    'powerpoint.export_slide',
    'powerpoint.list_layouts',
    'powerpoint.apply_layout',
    'powerpoint.get_selection',
    'powerpoint.set_selection',
    'powerpoint.list_shapes',
    'powerpoint.add_text_box',
    'powerpoint.add_shape',
    'powerpoint.insert_image',
    'powerpoint.update_shape',
    'powerpoint.read_text',
    'powerpoint.replace_text',
    'powerpoint.format_text',
    'powerpoint.add_table',
    'powerpoint.read_table',
    'powerpoint.update_table'
  ];
  const TOOL_GROUPS = [
    { label: 'Presentation', tools: ['powerpoint.get_presentation_info', 'powerpoint.get_active_view', 'powerpoint.export_file'] },
    { label: 'Metadata', tools: ['powerpoint.update_tags'] },
    { label: 'Slides', tools: ['powerpoint.list_slides', 'powerpoint.add_slide', 'powerpoint.update_slide', 'powerpoint.delete_slide', 'powerpoint.move_slide', 'powerpoint.export_slide'] },
    { label: 'Layout', tools: ['powerpoint.list_layouts', 'powerpoint.apply_layout'] },
    { label: 'Selection', tools: ['powerpoint.get_selection', 'powerpoint.set_selection'] },
    { label: 'Shapes', tools: ['powerpoint.list_shapes', 'powerpoint.add_text_box', 'powerpoint.add_shape', 'powerpoint.insert_image', 'powerpoint.update_shape'] },
    { label: 'Text', tools: ['powerpoint.read_text', 'powerpoint.replace_text', 'powerpoint.format_text'] },
    { label: 'Tables', tools: ['powerpoint.add_table', 'powerpoint.read_table', 'powerpoint.update_table'] }
  ];
  const TOOL_METADATA = new Map([
    ['powerpoint.get_presentation_info', { category: 'Presentation', sideEffect: 'read', description: 'Return presentation metadata, counts, selection summary, and capability gates.' }],
    ['powerpoint.get_active_view', { category: 'Presentation', sideEffect: 'read', description: 'Return the current PowerPoint view mode.' }],
    ['powerpoint.export_file', { category: 'Presentation', sideEffect: 'read', description: 'Export the presentation as PDF or PPTX when supported by the host.' }],
    ['powerpoint.update_tags', { category: 'Metadata', sideEffect: 'destructive', description: 'Read, set, or delete presentation tags.' }],
    ['powerpoint.list_slides', { category: 'Slides', sideEffect: 'read', description: 'List slides with ids, indices, layout, tags, and shape counts.' }],
    ['powerpoint.add_slide', { category: 'Slides', sideEffect: 'mutating', description: 'Add a slide to the current presentation.' }],
    ['powerpoint.update_slide', { category: 'Slides', sideEffect: 'mutating', description: 'Update slide tags, hidden state, or background where supported.' }],
    ['powerpoint.delete_slide', { category: 'Slides', sideEffect: 'destructive', description: 'Delete a slide without leaving the presentation empty.' }],
    ['powerpoint.move_slide', { category: 'Slides', sideEffect: 'mutating', description: 'Move a slide to a target index.' }],
    ['powerpoint.export_slide', { category: 'Slides', sideEffect: 'read', description: 'Export one slide image data when supported by the host.' }],
    ['powerpoint.list_layouts', { category: 'Layout', sideEffect: 'read', description: 'List slide masters and layouts.' }],
    ['powerpoint.apply_layout', { category: 'Slides', sideEffect: 'mutating', description: 'Apply a layout to a slide.' }],
    ['powerpoint.get_selection', { category: 'Selection', sideEffect: 'read', description: 'Return selected slides, shapes, or text metadata.' }],
    ['powerpoint.set_selection', { category: 'Selection', sideEffect: 'mutating', description: 'Select slides or text where supported.' }],
    ['powerpoint.list_shapes', { category: 'Shapes', sideEffect: 'read', description: 'List shapes on a slide with geometry and content metadata.' }],
    ['powerpoint.add_text_box', { category: 'Shapes', sideEffect: 'mutating', description: 'Add a text box to a slide.' }],
    ['powerpoint.add_shape', { category: 'Shapes', sideEffect: 'mutating', description: 'Add a geometric shape or line to a slide.' }],
    ['powerpoint.insert_image', { category: 'Shapes', sideEffect: 'mutating', description: 'Insert an image on a slide or current selection.' }],
    ['powerpoint.update_shape', { category: 'Shapes', sideEffect: 'destructive', description: 'Update shape geometry, visual settings, metadata, z-order, or deletion.' }],
    ['powerpoint.read_text', { category: 'Text', sideEffect: 'read', description: 'Read text from selected text, a shape, one slide, or all slides.' }],
    ['powerpoint.replace_text', { category: 'Text', sideEffect: 'mutating', description: 'Replace matching text in slide content.' }],
    ['powerpoint.format_text', { category: 'Text', sideEffect: 'mutating', description: 'Apply font and paragraph formatting to text.' }],
    ['powerpoint.add_table', { category: 'Tables', sideEffect: 'mutating', description: 'Add a table to a slide.' }],
    ['powerpoint.read_table', { category: 'Tables', sideEffect: 'read', description: 'Read table dimensions, values, and cell metadata.' }],
    ['powerpoint.update_table', { category: 'Tables', sideEffect: 'destructive', description: 'Update table values, layout, style, or delete the table shape.' }]
  ]);

  const { instanceId, sessionId } = runtimeIds();
  const TOOL_PERMISSION_STORAGE_KEY = `office-mcp.powerpoint.tool-permissions.${sessionId}`;
  const TOOL_PERMISSION_MODE_STORAGE_KEY = `office-mcp.powerpoint.tool-permission-mode.${sessionId}`;
  const logger = new AddinLogger({ redactText });
  const taskStore = new TaskHistoryStore({ redactText });
  let socket;
  let reconnectTimer;
  let reconnectAttempt = 0;
  let endpointDirty = false;
  let serverInfo = { serverVersion: 'Unknown', protocolVersion: PROTOCOL_VERSION };
  let documentInfo = null;
  let toolPermissions = loadToolPermissions();
  let toolPermissionMode = loadToolPermissionMode();
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
    setCopyableMetadata(daemonEl, endpoint);
    endpointInputEl.value = endpoint;
    endpointDirty = false;
    setConnectionState('connecting', 'Connecting\u2026');
    try {
      socket = new WebSocket(endpoint);
    } catch (error) {
      logger.error('websocket.create.failed', error);
      if (tryCurrentOriginEndpointFallback(endpoint)) return;
      connectionDetailEl.textContent = error.message || 'Connection failed before the daemon socket could open.';
      setConnectionState('failed', 'Failed');
      return;
    }
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
    const requestId = createRequestId();
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
    setCopyableMetadata(sessionEl, sessionId);
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
    renderRuntimeVersions(serverVersionEl, protocolVersionEl, serverInfo, PROTOCOL_VERSION);
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
        case 'powerpoint.get_presentation_info':
          data = await getPresentationInfoTool(args);
          break;
        case 'powerpoint.get_active_view':
          data = await getActiveView(args);
          break;
        case 'powerpoint.export_file':
          data = await exportFile(args);
          break;
        case 'powerpoint.update_tags':
          data = await updateTags(args);
          break;
        case 'powerpoint.list_slides':
          data = await listSlides(args);
          break;
        case 'powerpoint.add_slide':
          data = await addSlide(args);
          break;
        case 'powerpoint.update_slide':
          data = await updateSlide(args);
          break;
        case 'powerpoint.delete_slide':
          data = await deleteSlide(args);
          break;
        case 'powerpoint.move_slide':
          data = await moveSlide(args);
          break;
        case 'powerpoint.export_slide':
          data = await exportSlide(args);
          break;
        case 'powerpoint.list_layouts':
          data = await listLayouts(args);
          break;
        case 'powerpoint.apply_layout':
          data = await applyLayout(args);
          break;
        case 'powerpoint.get_selection':
          data = await getSelection(args);
          break;
        case 'powerpoint.set_selection':
          data = await setSelection(args);
          break;
        case 'powerpoint.list_shapes':
          data = await listShapes(args);
          break;
        case 'powerpoint.add_text_box':
          data = await addTextBox(args);
          break;
        case 'powerpoint.add_shape':
          data = await addShape(args);
          break;
        case 'powerpoint.insert_image':
          data = await insertImage(args);
          break;
        case 'powerpoint.update_shape':
          data = await updateShape(args);
          break;
        case 'powerpoint.read_text':
          data = await readText(args);
          break;
        case 'powerpoint.replace_text':
          data = await replaceText(args);
          break;
        case 'powerpoint.format_text':
          data = await formatText(args);
          break;
        case 'powerpoint.add_table':
          data = await addTable(args);
          break;
        case 'powerpoint.read_table':
          data = await readTable(args);
          break;
        case 'powerpoint.update_table':
          data = await updateTable(args);
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

  async function getPresentationInfoTool(args) {
    const info = await getPresentationInfo();
    return {
      ...info,
      slide_count: null,
      requirement_sets: probeRequirementSets(),
      capabilities: Object.fromEntries(AVAILABLE_TOOLS.map((tool) => [tool, isToolEnabled(tool)])),
      include_selection: Boolean(args.include_selection)
    };
  }

  async function getActiveView(_args) {
    const view = await officeAsync((callback) => Office.context.document.getActiveViewAsync(callback));
    return { active_view: String(view || 'unknown'), editable: String(view || '').toLowerCase() !== 'read' };
  }

  async function exportFile(args) {
    const format = String(args.format || 'pdf').toLowerCase();
    if (isDesktopPowerPointHost()) {
      throw hostCapabilityUnavailable('PowerPoint desktop file export is not available through Office.context.document.getFileAsync in this host.');
    }
    const fileType = officeFileTypeFrom(format);
    const asyncOptions = { timeout_ms: POWERPOINT_FILE_EXPORT_TIMEOUT_MS, timeout_code: 'HOST_ERROR', timeout_partial_effect: 'none' };
    const file = await officeAsync((callback) => Office.context.document.getFileAsync(fileType, { sliceSize: positiveInteger(args.slice_size, 4 * 1024 * 1024) }, callback), asyncOptions);
    try {
      const chunks = [];
      for (let index = 0; index < file.sliceCount; index += 1) {
        const slice = await officeAsync((callback) => file.getSliceAsync(index, callback), asyncOptions);
        chunks.push(sliceDataToBytes(slice.data));
      }
      return { format, mime_type: mimeTypeForFormat(format), base64: bytesToBase64(concatBytes(chunks)), size: file.size, slice_count: file.sliceCount };
    } finally {
      await officeAsync((callback) => file.closeAsync(callback)).catch((error) => logger.warn('file.close.failed', error));
    }
  }

  async function updateTags(args) {
    requireRequirementSet('PowerPointApi', '1.3', 'presentation tags');
    return PowerPoint.run(async (context) => {
      const tags = context.presentation.tags;
      const action = String(args.action || 'list').toLowerCase();
      const key = args.key === undefined ? undefined : requiredString(args, 'key', `powerpoint.update_tags ${action} requires key.`);
      if (action === 'set') {
        tags.add(key, requiredString(args, 'value', 'powerpoint.update_tags set requires value.'));
      } else if (action === 'delete') {
        tags.delete(key);
      } else if (action !== 'list') {
        throw invalidArgument(`Unsupported powerpoint.update_tags action ${action}.`);
      }
      await context.sync();
      tags.load('items/key,value');
      const selected = key ? tags.getItemOrNullObject(key) : null;
      if (selected) selected.load('key,value,isNullObject');
      await context.sync();
      return {
        action,
        tag: selected && !selected.isNullObject ? tagMetadata(selected, key) : null,
        tags: (tags.items || []).map((tag) => tagMetadata(tag))
      };
    });
  }

  async function listSlides(args) {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items/id,index,tags/items/key,value,layout/id,layout/name,shapes/items/id');
      await context.sync();
      return { slides: (slides.items || []).map((slide) => slideMetadata(slide, Boolean(args.include_tags))) };
    });
  }

  async function updateSlide(args) {
    requireRequirementSet('PowerPointApi', '1.3', 'slide updates');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const action = String(args.action || 'set_tags').toLowerCase();
      let key = null;
      if (action === 'set_tag') {
        key = requiredString(args, 'key', 'powerpoint.update_slide set_tag requires key.');
        slide.tags.add(key, requiredString(args, 'value', 'powerpoint.update_slide set_tag requires value.'));
      } else if (action === 'delete_tag') {
        key = requiredString(args, 'key', 'powerpoint.update_slide delete_tag requires key.');
        slide.tags.delete(key);
      } else if (action === 'set_background') {
        requireRequirementSet('PowerPointApi', '1.10', 'slide background updates');
        if (slide.background?.fill?.setSolidColor) slide.background.fill.setSolidColor(requiredString(args, 'color', 'powerpoint.update_slide set_background requires color.'));
        else throw hostCapabilityUnavailable('Slide background updates are not available in this PowerPoint host.');
      } else if (action !== 'set_tags') {
        throw invalidArgument(`Unsupported powerpoint.update_slide action ${action}.`);
      }
      slide.load('id,index,tags/items/key,value,layout/id,layout/name,shapes/items/id');
      const selectedTag = key ? slide.tags.getItemOrNullObject(key) : null;
      if (selectedTag) selectedTag.load('key,value,isNullObject');
      await context.sync();
      if (action === 'set_tag' && (!selectedTag || selectedTag.isNullObject)) {
        throw hostCapabilityUnavailable('Slide tag updates are not persisted by this PowerPoint host.');
      }
      return { action, slide: slideMetadata(slide, true), tag: selectedTag && !selectedTag.isNullObject ? tagMetadata(selectedTag, key) : null };
    });
  }

  async function deleteSlide(args) {
    requireRequirementSet('PowerPointApi', '1.3', 'slide deletion');
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const count = slides.getCount();
      await context.sync();
      if (count.value <= 1) throw invalidArgument('powerpoint.delete_slide cannot delete the only slide.');
      const slide = targetSlide(context, args);
      slide.load('id,index');
      await context.sync();
      const deleted = { slide_id: slide.id, slide_index: slide.index };
      slide.delete();
      await context.sync();
      return { ...deleted, deleted: true };
    });
  }

  async function moveSlide(args) {
    requireRequirementSet('PowerPointApi', '1.8', 'slide move');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const targetIndex = requiredInteger(args, 'target_index', 'powerpoint.move_slide requires target_index.');
      if (typeof slide.moveTo !== 'function') throw hostCapabilityUnavailable('Slide move is not available in this PowerPoint host.');
      slide.moveTo(targetIndex);
      slide.load('id,index');
      await context.sync();
      return { slide_id: slide.id, slide_index: slide.index, target_index: targetIndex };
    });
  }

  async function exportSlide(args) {
    requireRequirementSet('PowerPointApi', '1.8', 'slide export');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      if (typeof slide.getImageAsBase64 !== 'function') throw hostCapabilityUnavailable('Slide image export is not available in this PowerPoint host.');
      const image = slide.getImageAsBase64();
      slide.load('id,index');
      await context.sync();
      return { slide_id: slide.id, slide_index: slide.index, mime_type: 'image/png', base64: image.value || image };
    });
  }

  async function listLayouts(_args) {
    return PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      masters.load('items/id,name');
      await context.sync();
      for (const master of masters.items || []) master.layouts.load('items/id,name,type');
      await context.sync();
      return {
        masters: (masters.items || []).map((master) => ({
          id: master.id,
          name: master.name || '',
          layouts: (master.layouts.items || []).map((layout) => ({ id: layout.id, name: layout.name || '', type: layout.type || null }))
        }))
      };
    });
  }

  async function replaceText(args) {
    const search = requiredString(args, 'search', 'powerpoint.replace_text requires search text.');
    const replacement = requiredString(args, 'replacement', 'powerpoint.replace_text requires replacement text.');
    const matchCase = Boolean(args.match_case);
    try {
      return await PowerPoint.run(async (context) => {
        const slides = await loadSlidesWithShapes(context, args);
        for (const slide of slides) {
          for (const shape of slide.shapes.items || []) shape.load('id,textFrame/hasText,textFrame/textRange/text');
        }
        await context.sync();
        const touchedShapes = [];
        let replacements = 0;
        for (const slide of slides) {
          for (const shape of slide.shapes.items || []) {
            const range = shape.textFrame?.textRange;
            if (!shape.textFrame?.hasText || !range || typeof range.text !== 'string') continue;
            const nextText = replaceAllText(range.text, search, replacement, matchCase);
            if (nextText === range.text) continue;
            try {
              range.text = nextText;
            } catch (error) {
              if (isOfficeInvalidArgument(error)) {
                throw hostCapabilityUnavailable('PowerPoint text replacement is not available in this host.');
              }
              throw error;
            }
            replacements += countMatches(range.text, search, matchCase);
            touchedShapes.push({ slide_id: slide.id, shape_id: shape.id });
          }
        }
        try {
          await context.sync();
        } catch (error) {
          if (isOfficeInvalidArgument(error)) {
            throw hostCapabilityUnavailable('PowerPoint text replacement is not available in this host.');
          }
          throw error;
        }
        return { replacements, touched_shapes: touchedShapes };
      });
    } catch (error) {
      if (isOfficeInvalidArgument(error)) {
        throw hostCapabilityUnavailable('PowerPoint text replacement is not available in this host.');
      }
      throw error;
    }
  }

  async function insertImage(args) {
    const base64 = imageBase64(args);
    await officeAsync((callback) => Office.context.document.setSelectedDataAsync(base64, imageInsertOptions(args), callback));
    return {
      inserted_image: true,
      placement: 'selection',
      width: numberOrNull(args.width),
      height: numberOrNull(args.height),
      mime_type: args.image?.mime_type || null,
      byte_length: Number.isInteger(args.image?.byte_length) ? args.image.byte_length : null
    };
  }

  async function getSelection(_args) {
    requireRequirementSet('PowerPointApi', '1.5', 'selection reads');
    return PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items/id,index');
      await context.sync();
      return { slides: (selectedSlides.items || []).map((slide) => ({ slide_id: slide.id, slide_index: slide.index })) };
    });
  }

  async function setSelection(args) {
    requireRequirementSet('PowerPointApi', '1.5', 'selection updates');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      if (typeof slide.select !== 'function') throw hostCapabilityUnavailable('Slide selection is not available in this PowerPoint host.');
      slide.select();
      slide.load('id,index');
      await context.sync();
      return { selected: true, slide_id: slide.id, slide_index: slide.index };
    });
  }

  async function listShapes(args) {
    requireRequirementSet('PowerPointApi', '1.3', 'shape listing');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      slide.load('id,index');
      slide.shapes.load('items');
      await context.sync();
      for (const shape of slide.shapes.items || []) shape.load('id,name,type,left,top,width,height,rotation,textFrame/hasText,textFrame/textRange/text');
      await context.sync();
      return { slide_id: slide.id, slide_index: slide.index, shapes: (slide.shapes.items || []).map(shapeMetadata) };
    });
  }

  async function addTextBox(args) {
    requireRequirementSet('PowerPointApi', '1.4', 'text box creation');
    const text = requiredString(args, 'text', 'powerpoint.add_text_box requires text.');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const shape = slide.shapes.addTextBox(text, shapeOptions(args, { left: 72, top: 72, width: 420, height: 80 }));
      shape.load('id,name,type,left,top,width,height,rotation,textFrame/hasText,textFrame/textRange/text');
      slide.load('id,index');
      await context.sync();
      return { slide_id: slide.id, slide_index: slide.index, shape: shapeMetadata(shape) };
    });
  }

  async function addShape(args) {
    requireRequirementSet('PowerPointApi', '1.4', 'shape creation');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const type = shapeTypeFrom(args.type);
      const shape = slide.shapes.addGeometricShape(type, shapeOptions(args, { left: 96, top: 96, width: 160, height: 96 }));
      applyShapeProperties(shape, args);
      shape.load('id,name,type,left,top,width,height,rotation,textFrame/hasText,textFrame/textRange/text');
      slide.load('id,index');
      await context.sync();
      return { slide_id: slide.id, slide_index: slide.index, shape: shapeMetadata(shape) };
    });
  }

  async function updateShape(args) {
    requireRequirementSet('PowerPointApi', '1.4', 'shape updates');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const shape = targetShape(slide, args);
      const action = String(args.action || 'set_properties').toLowerCase();
      if (action === 'delete') {
        shape.load('id');
        await context.sync();
        const shapeId = shape.id;
        shape.delete();
        await context.sync();
        return { action, shape_id: shapeId, deleted: true };
      }
      if (action === 'group') {
        requireRequirementSet('PowerPointApi', '1.8', 'shape grouping');
        if (!Array.isArray(args.shape_ids) || args.shape_ids.length < 2) throw invalidArgument('powerpoint.update_shape group requires at least two shape_ids.');
        if (typeof slide.shapes.addGroup !== 'function') throw hostCapabilityUnavailable('Shape grouping is not available in this PowerPoint host.');
        const grouped = slide.shapes.addGroup(args.shape_ids);
        grouped.load('id,name,type,left,top,width,height,rotation,zOrderPosition');
        slide.load('id,index');
        await context.sync();
        return { action, slide_id: slide.id, slide_index: slide.index, shape: shapeMetadata(grouped), grouped_shape_ids: args.shape_ids.map(String) };
      }
      if (action === 'ungroup') {
        requireRequirementSet('PowerPointApi', '1.8', 'shape grouping');
        if (!shape.group || typeof shape.group.ungroup !== 'function') throw hostCapabilityUnavailable('Shape ungrouping is not available for this shape or host.');
        shape.load('id');
        await context.sync();
        const shapeId = shape.id;
        shape.group.ungroup();
        await context.sync();
        return { action, shape_id: shapeId, ungrouped: true };
      }
      if (['bring_forward', 'bring_to_front', 'send_backward', 'send_to_back'].includes(action)) {
        requireRequirementSet('PowerPointApi', '1.8', 'shape z-order updates');
        if (typeof shape.setZOrder !== 'function') throw hostCapabilityUnavailable('Shape z-order updates are not available in this PowerPoint host.');
        shape.setZOrder(shapeZOrderAction(action));
      } else if (action !== 'set_properties') throw invalidArgument(`Unsupported powerpoint.update_shape action ${action}.`);
      applyShapeProperties(shape, args);
      shape.load('id,name,type,left,top,width,height,rotation,altTextTitle,altTextDescription,isDecorative,visible,zOrderPosition,textFrame/hasText,textFrame/textRange/text');
      slide.load('id,index');
      await context.sync();
      return { action, slide_id: slide.id, slide_index: slide.index, shape: shapeMetadata(shape) };
    });
  }

  async function readText(args) {
    requireRequirementSet('PowerPointApi', '1.4', 'text reads');
    return PowerPoint.run(async (context) => {
      const slides = await loadSlidesWithShapes(context, args);
      for (const slide of slides) {
        for (const shape of slide.shapes.items || []) shape.load('id,textFrame/hasText,textFrame/textRange/text');
      }
      await context.sync();
      const items = [];
      for (const slide of slides) {
        for (const shape of slide.shapes.items || []) {
          if (!shape.textFrame?.hasText) continue;
          items.push({ slide_id: slide.id, slide_index: slide.index, shape_id: shape.id, text: shape.textFrame.textRange?.text || '' });
        }
      }
      return { items, count: items.length };
    });
  }

  async function formatText(args) {
    requireRequirementSet('PowerPointApi', '1.4', 'text formatting');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const shape = targetShape(slide, args);
      shape.load('id,textFrame/hasText,textFrame/textRange/font');
      await context.sync();
      if (!shape.textFrame?.hasText) throw invalidArgument('Target shape does not contain text.');
      applyTextFormat(shape.textFrame.textRange, args);
      await context.sync();
      return { shape_id: shape.id, formatted: true };
    });
  }

  async function addTable(args) {
    requireRequirementSet('PowerPointApi', '1.8', 'table creation');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const rows = positiveInteger(args.rows, Array.isArray(args.values) ? args.values.length : 2);
      const columns = positiveInteger(args.columns, Array.isArray(args.values?.[0]) ? args.values[0].length : 2);
      if (typeof slide.shapes.addTable !== 'function') throw hostCapabilityUnavailable('Table creation is not available in this PowerPoint host.');
      const shape = slide.shapes.addTable(rows, columns, shapeOptions(args, { left: 72, top: 120, width: 480, height: 220 }));
      shape.load('id,type,table/rowCount,table/columnCount');
      await context.sync();
      if (!shape.table) throw hostCapabilityUnavailable('Table creation is not available in this PowerPoint host.');
      if (Array.isArray(args.values)) applyTableValues(shape.table, args.values);
      await context.sync();
      return { shape_id: shape.id, rows, columns, added: true };
    });
  }

  async function readTable(args) {
    requireRequirementSet('PowerPointApi', '1.8', 'table reads');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const shape = targetShape(slide, args);
      shape.load('id,table/rowCount,table/columnCount,table/values');
      await context.sync();
      if (!shape.table) throw invalidArgument('Target shape is not a table.');
      return tableMetadata(shape);
    });
  }

  async function updateTable(args) {
    requireRequirementSet('PowerPointApi', '1.8', 'table updates');
    return PowerPoint.run(async (context) => {
      const slide = targetSlide(context, args);
      const shape = targetShape(slide, args);
      shape.load('id,table/rowCount,table/columnCount,table/values');
      await context.sync();
      if (!shape.table) throw invalidArgument('Target shape is not a table.');
      const table = shape.table;
      const action = String(args.action || 'set_values').toLowerCase();
      if (action === 'delete') {
        const shapeId = shape.id;
        shape.delete();
        await context.sync();
        return { action, shape_id: shapeId, deleted: true };
      }
      if (action === 'set_values') applyTableValues(table, args.values || []);
      else if (action === 'set_cell') tableCell(table, requiredInteger(args, 'row_index', 'powerpoint.update_table set_cell requires row_index.'), requiredInteger(args, 'column_index', 'powerpoint.update_table set_cell requires column_index.')).text = String(args.value ?? '');
      else if (action === 'add_rows') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        table.rows.add(optionalInteger(args.row_index, null), positiveInteger(args.count, 1));
      } else if (action === 'delete_rows') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        table.rows.deleteRows(indexedItems(table.rows, args.row_indices, 'row_indices'));
      } else if (action === 'add_columns') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        table.columns.add(optionalInteger(args.column_index, null), positiveInteger(args.count, 1));
      } else if (action === 'delete_columns') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        table.columns.deleteColumns(indexedItems(table.columns, args.column_indices, 'column_indices'));
      } else if (action === 'merge_cells') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        table.mergeCells(
          requiredInteger(args, 'row_index', 'powerpoint.update_table merge_cells requires row_index.'),
          requiredInteger(args, 'column_index', 'powerpoint.update_table merge_cells requires column_index.'),
          positiveInteger(args.row_count, 1),
          positiveInteger(args.column_count, 1)
        );
      } else if (action === 'split_cell') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        const rowIndex = requiredInteger(args, 'row_index', 'powerpoint.update_table split_cell requires row_index.');
        const columnIndex = requiredInteger(args, 'column_index', 'powerpoint.update_table split_cell requires column_index.');
        tableCell(table, rowIndex, columnIndex).split(positiveInteger(args.row_count, 1), positiveInteger(args.column_count, 1));
      } else if (action === 'clear') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        table.clear(tableClearOptions(args));
      } else if (action === 'style') {
        requireRequirementSet('PowerPointApi', '1.9', 'table structural updates');
        applyTableStyle(table, args);
      } else throw invalidArgument(`Unsupported powerpoint.update_table action ${action}.`);
      await context.sync();
      table.load('rowCount,columnCount,values');
      await context.sync();
      return { action, table: tableMetadata(shape) };
    });
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

  async function loadSlidesWithShapes(context, args) {
    const slides = targetSlides(context, args);
    if (slides.load) slides.load('items');
    await context.sync();
    for (const slide of slides.items || []) {
      slide.load('id,index');
      slide.shapes.load('items');
    }
    await context.sync();
    return slides.items || [];
  }

  function targetSlide(context, args) {
    const slideId = stringArg(args, 'slide_id') || stringArg(args, 'slideId');
    if (slideId) return context.presentation.slides.getItem(slideId);
    const index = numberOrNull(args.slide_index ?? args.slideIndex);
    if (index !== null) return context.presentation.slides.getItemAt(index);
    return context.presentation.getSelectedSlides().getItemAt(0);
  }

  function targetShape(slide, args) {
    const shapeId = stringArg(args, 'shape_id') || stringArg(args, 'shapeId');
    if (!shapeId) throw invalidArgument('PowerPoint shape tools require shape_id.');
    return slide.shapes.getItem(shapeId);
  }

  async function resolveLayout(context, args) {
    const layoutId = stringArg(args, 'layout_id') || stringArg(args, 'layoutId');
    const layoutName = stringArg(args, 'layout_name') || stringArg(args, 'layoutName');
    const layoutType = stringArg(args, 'layout_type') || stringArg(args, 'layoutType') || stringArg(args, 'layout');
    if (!layoutId && !layoutName && !layoutType) {
      throw Object.assign(new Error('powerpoint.apply_layout requires layout_id, layout_name, or layout_type.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    const masters = context.presentation.slideMasters;
    masters.load('items/id,name');
    await context.sync();
    for (const master of masters.items || []) master.layouts.load('items/id,name,type');
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
    const raw = args.image?.base64 || requiredString(args, 'base64', 'powerpoint.insert_image requires base64 image data.');
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

  function slideMetadata(slide, includeTags) {
    return {
      slide_id: slide.id,
      slide_index: slide.index,
      layout_id: slide.layout?.id || null,
      layout_name: slide.layout?.name || null,
      shape_count: slide.shapes?.items?.length ?? null,
      tags: includeTags ? (slide.tags?.items || []).map((tag) => tagMetadata(tag)) : undefined
    };
  }

  function tagMetadata(tag, requestedKey) {
    const hostKey = tag.key || '';
    const normalizedKey = String(requestedKey || hostKey).toLowerCase();
    return { key: normalizedKey, normalized_key: normalizedKey, host_key: hostKey, value: tag.value || '' };
  }

  function shapeMetadata(shape) {
    const altTextTitle = safeLoaded(shape, 'altTextTitle');
    const altTextDescription = safeLoaded(shape, 'altTextDescription');
    const isDecorative = safeLoaded(shape, 'isDecorative');
    const visible = safeLoaded(shape, 'visible');
    const zOrderPosition = safeLoaded(shape, 'zOrderPosition');
    return {
      shape_id: shape.id,
      name: shape.name || '',
      type: shape.type || null,
      left: numberOrNull(shape.left),
      top: numberOrNull(shape.top),
      width: numberOrNull(shape.width),
      height: numberOrNull(shape.height),
      rotation: numberOrNull(shape.rotation),
      alt_text_title: altTextTitle || null,
      alt_text_description: altTextDescription || null,
      is_decorative: typeof isDecorative === 'boolean' ? isDecorative : null,
      visible: typeof visible === 'boolean' ? visible : null,
      z_order_position: numberOrNull(zOrderPosition),
      has_text: Boolean(shape.textFrame?.hasText),
      text_preview: shape.textFrame?.hasText ? String(shape.textFrame.textRange?.text || '').slice(0, 200) : null,
      has_table: Boolean(shape.table)
    };
  }

  function safeLoaded(object, property) {
    try {
      return object?.[property];
    } catch (_error) {
      return null;
    }
  }

  function tableMetadata(shape) {
    return {
      shape_id: shape.id,
      rows: shape.table?.rowCount ?? null,
      columns: shape.table?.columnCount ?? null,
      values: shape.table?.values || []
    };
  }

  function applyShapeProperties(shape, args) {
    for (const key of ['left', 'top', 'width', 'height', 'rotation']) {
      const value = numberOrNull(args[key]);
      if (value !== null) shape[key] = value;
    }
    const name = stringArg(args, 'name');
    if (name) shape.name = name;
    if (args.alt_text_title !== undefined || args.alt_text_description !== undefined || args.is_decorative !== undefined || args.visible !== undefined) {
      requireRequirementSet('PowerPointApi', '1.10', 'shape accessibility and visibility updates');
      if (args.alt_text_title !== undefined) shape.altTextTitle = String(args.alt_text_title ?? '');
      if (args.alt_text_description !== undefined) shape.altTextDescription = String(args.alt_text_description ?? '');
      if (args.is_decorative !== undefined) shape.isDecorative = Boolean(args.is_decorative);
      if (args.visible !== undefined) shape.visible = Boolean(args.visible);
    }
    const fillColor = stringArg(args, 'fill_color');
    if (args.clear_fill === true && shape.fill?.clear) shape.fill.clear();
    if (fillColor && shape.fill?.setSolidColor) shape.fill.setSolidColor(fillColor);
    const fillTransparency = numberOrNull(args.fill_transparency);
    if (fillTransparency !== null && shape.fill) shape.fill.transparency = fillTransparency;
    const lineColor = stringArg(args, 'line_color');
    if (lineColor && shape.lineFormat?.color !== undefined) shape.lineFormat.color = lineColor;
    const lineWeight = numberOrNull(args.line_weight);
    if (lineWeight !== null && shape.lineFormat) shape.lineFormat.weight = lineWeight;
    const lineDashStyle = stringArg(args, 'line_dash_style');
    if (lineDashStyle && shape.lineFormat) shape.lineFormat.dashStyle = lineDashStyle;
    const lineTransparency = numberOrNull(args.line_transparency);
    if (lineTransparency !== null && shape.lineFormat) shape.lineFormat.transparency = lineTransparency;
    if (args.line_visible !== undefined && shape.lineFormat) shape.lineFormat.visible = Boolean(args.line_visible);
  }

  function applyTextFormat(textRange, args) {
    const font = textRange.font;
    if (!font) return;
    if (args.bold !== undefined) font.bold = Boolean(args.bold);
    if (args.italic !== undefined) font.italic = Boolean(args.italic);
    if (args.underline !== undefined) font.underline = Boolean(args.underline);
    const color = stringArg(args, 'color');
    if (color) font.color = color;
    const name = stringArg(args, 'font_name');
    if (name) font.name = name;
    const size = numberOrNull(args.font_size);
    if (size !== null) font.size = size;
  }

  function applyTableValues(table, values) {
    if (!Array.isArray(values)) throw invalidArgument('Table values must be a two-dimensional array.');
    for (let row = 0; row < values.length; row += 1) {
      if (!Array.isArray(values[row])) throw invalidArgument('Table values must be a two-dimensional array.');
      for (let column = 0; column < values[row].length; column += 1) {
        const cell = tableCell(table, row, column);
        cell.text = values[row][column] == null ? '' : String(values[row][column]);
      }
    }
  }

  function tableCell(table, row, column) {
    if (typeof table.getCellOrNullObject === 'function') return table.getCellOrNullObject(row, column);
    if (typeof table.getCell === 'function') return table.getCell(row, column);
    throw hostCapabilityUnavailable('Table cell access is not available in this PowerPoint host.');
  }

  function indexedItems(collection, indices, key) {
    if (!Array.isArray(indices) || indices.length === 0) throw invalidArgument(`powerpoint.update_table requires ${key}.`);
    return indices.map((value) => collection.getItemAt(Number(value)));
  }

  function tableClearOptions(args) {
    if (args.all !== undefined || args.text !== undefined || args.format !== undefined) {
      return { all: Boolean(args.all), text: Boolean(args.text), format: Boolean(args.format) };
    }
    return { text: true };
  }

  function applyTableStyle(table, args) {
    if (!table.styleSettings) throw hostCapabilityUnavailable('Table style settings are not available in this PowerPoint host.');
    const style = stringArg(args, 'style');
    if (style) table.styleSettings.style = style;
    for (const [argKey, property] of [
      ['banded_rows', 'areRowsBanded'],
      ['banded_columns', 'areColumnsBanded'],
      ['first_row', 'isFirstRowHighlighted'],
      ['last_row', 'isLastRowHighlighted'],
      ['first_column', 'isFirstColumnHighlighted'],
      ['last_column', 'isLastColumnHighlighted']
    ]) {
      if (args[argKey] !== undefined) table.styleSettings[property] = Boolean(args[argKey]);
    }
  }

  function shapeZOrderAction(action) {
    return ({
      bring_forward: 'BringForward',
      bring_to_front: 'BringToFront',
      send_backward: 'SendBackward',
      send_to_back: 'SendToBack'
    })[action];
  }

  function shapeTypeFrom(value) {
    const raw = String(value || 'rectangle');
    const shapes = PowerPoint.ShapeType || {};
    return shapes[raw] || shapes[titleCase(raw).replace(/\s+/g, '')] || raw;
  }

  function officeFileTypeFrom(format) {
    const normalized = String(format || 'pdf').toLowerCase();
    if (normalized === 'pdf') return Office.FileType.Pdf;
    if (normalized === 'pptx' || normalized === 'compressed') return Office.FileType.Compressed;
    throw invalidArgument('powerpoint.export_file format must be pdf or pptx.');
  }

  function isDesktopPowerPointHost() {
    return String(Office.context?.platform || '').toLowerCase() === 'pc';
  }

  function mimeTypeForFormat(format) {
    return format === 'pptx' || format === 'compressed'
      ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : 'application/pdf';
  }

  function supportsRequirementSet(name, version) {
    return Office.context?.requirements?.isSetSupported?.(name, version) === true;
  }

  function requireRequirementSet(name, version, feature) {
    if (!supportsRequirementSet(name, version)) throw hostCapabilityUnavailable(`${feature} requires ${name} ${version}.`);
  }

  function hostCapabilityUnavailable(message) {
    return Object.assign(new Error(message), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE', partialEffect: 'none' });
  }

  function invalidArgument(message) {
    return Object.assign(new Error(message), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  function isOfficeInvalidArgument(error) {
    const detail = [
      error?.code,
      error?.name,
      error?.message,
      error?.debugInfo?.code,
      error?.debugInfo?.message,
      error?.debugInfo?.errorLocation
    ].filter(Boolean).join(' ');
    if (/InvalidArgument/i.test(detail)) return true;
    try { return /InvalidArgument/i.test(JSON.stringify(error)); } catch { return false; }
  }

  function officeAsync(start, options = {}) {
    return new Promise((resolve, reject) => {
      let settled = false;
      const timeoutMs = numberOrNull(options.timeout_ms);
      const timeout = timeoutMs === null ? null : setTimeout(() => {
        if (settled) return;
        settled = true;
        reject(Object.assign(new Error(options.timeout_message || `Office async operation timed out after ${timeoutMs}ms.`), {
          officeMcpCode: options.timeout_code || 'HOST_ERROR',
          partialEffect: options.timeout_partial_effect || 'unknown'
        }));
      }, timeoutMs);
      start((result) => {
        if (settled) return;
        settled = true;
        if (timeout) clearTimeout(timeout);
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
    if (!value) throw invalidArgument(message);
    return value;
  }

  function requiredInteger(args, key, message) {
    const value = Number(args?.[key]);
    if (!Number.isInteger(value)) throw invalidArgument(message);
    return value;
  }

  function optionalInteger(value, fallback) {
    if (value === undefined || value === null || value === '') return fallback;
    const number = Number(value);
    if (!Number.isInteger(number)) throw invalidArgument('Optional integer argument must be an integer.');
    return number;
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
    setCopyableMetadata(sessionEl, sessionId);
    setCopyableMetadata(daemonEl, configuredEndpoint());
    renderRuntimeVersions(serverVersionEl, protocolVersionEl, serverInfo, PROTOCOL_VERSION);
    hostPlatformEl.textContent = officeHostSummary('PowerPoint');
    renderToolModeControl();
    renderToolSummary();
    renderCurrentTask();
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
      const rows = tools.map(toolControlMarkup).join('');
      groupEl.innerHTML = [
        '<summary class="tool-group-title">',
        `<span>${escapeHtml(group.label)}</span>`,
        `<span class="tool-group-count">${enabledInGroup.length}/${tools.length}</span>`,
        `<input class="group-toggle" type="checkbox" role="switch" data-tool-group="${escapeHtml(group.label)}" aria-label="Toggle ${escapeHtml(group.label)} tools" ${allowedInGroup.length > 0 && enabledInGroup.length === allowedInGroup.length ? 'checked' : ''} ${allowedInGroup.length === 0 ? 'disabled' : ''} />`,
        '</summary>',
        `<div class="tool-permission-list">${rows}</div>`
      ].join('');
      toolListEl.appendChild(groupEl);
    }
    toolListEl.querySelectorAll('[data-tool]').forEach((input) => bindDetailsControl(input, handleToolPermissionChange));
    toolListEl.querySelectorAll('[data-tool-group]').forEach((input) => bindDetailsControl(input, handleToolGroupPermissionChange));
  }

  function toolControlMarkup(tool) {
    const meta = TOOL_METADATA.get(tool) || { sideEffect: 'unknown', description: 'No metadata.' };
    const modeAllowed = isToolAllowedByMode(tool);
    const enabled = isToolEnabled(tool) && modeAllowed;
    const rowStateClass = `${modeAllowed ? '' : ' is-disabled'}${meta.sideEffect === 'mutating' || meta.sideEffect === 'destructive' ? ' is-mutating' : ''}`;
    const sideEffectClass = meta.sideEffect === 'mutating' || meta.sideEffect === 'destructive' ? ' mutating' : '';
    const id = `toolPermission-${tool.replace(/[^a-z0-9_-]/gi, '-')}`;
    return `<label class="tool-permission-row${rowStateClass}" for="${id}"><span class="tool-permission-main"><span class="tool-permission-title"><span class="tool-permission-name">${escapeHtml(tool)}</span><span class="side-effect-pill${sideEffectClass}">${escapeHtml(titleCase(meta.sideEffect))}</span></span><span class="tool-permission-meta">${escapeHtml(meta.description)}</span></span><input id="${id}" class="tool-toggle" type="checkbox" role="switch" data-tool="${escapeHtml(tool)}" aria-label="Toggle ${escapeHtml(tool)}" ${enabled ? 'checked' : ''} ${modeAllowed ? '' : 'disabled aria-disabled="true"'} /></label>`;
  }

  function handleToolPermissionChange(event) {
    updateToolPermission(event.currentTarget.dataset.tool, event.currentTarget.checked);
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

  function updateToolPermission(tool, enabled) {
    if (!AVAILABLE_TOOLS.includes(tool)) return;
    toolPermissions = { ...toolPermissions, [tool]: Boolean(enabled) };
    saveToolPermissions();
    renderToolSummary();
    sendSessionToolUpdate();
  }

  function sendSessionToolUpdate() {
    if (!sessionAnnounced) return;
    send(sessionUpdatedNotification({
      session_id: sessionId,
      patch: { available_tools: effectiveTools() }
    }));
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

  function loadToolPermissionMode() {
    try {
      const mode = window.localStorage?.getItem(TOOL_PERMISSION_MODE_STORAGE_KEY) || 'all';
      return ['read', 'write', 'all'].includes(mode) ? mode : 'all';
    } catch (_error) {
      return 'all';
    }
  }

  function saveToolPermissions() {
    try {
      window.localStorage?.setItem(TOOL_PERMISSION_STORAGE_KEY, JSON.stringify(toolPermissions));
    } catch (error) {
      logger.warn('tool_permissions.save.failed', error);
    }
  }

  function saveToolPermissionMode() {
    try {
      window.localStorage?.setItem(TOOL_PERMISSION_MODE_STORAGE_KEY, toolPermissionMode);
    } catch (error) {
      logger.warn('tool_permission_mode.save.failed', error);
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
    const tone = task.status === 'success' ? 'status-success' : task.status === 'running' ? 'status-warning' : task.status === 'cancelled' ? 'status-neutral' : 'status-danger';
    const elapsed = typeof task.elapsedMs === 'number' ? formatDuration(task.elapsedMs) : 'in progress';
    const metadata = taskMetadataMarkup(task, { escapeHtml, formatTime, redactText, valueLabel: boolLabel });
    const commandId = commandIdMarkup(task.requestId, { escapeHtml });
    const startedAt = task.startedAt ? `${escapeHtml(formatTime(task.startedAt))} / ` : '';
    return `<article class="task-card"><div class="task-title"><span>${escapeHtml(task.tool)}</span><span class="status-badge ${tone}">${escapeHtml(status)}</span></div>${commandId}<div class="task-meta">${startedAt}${escapeHtml(elapsed)}</div>${metadata}</article>`;
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
    connectionBadgeEl.className = `status-badge ${statusClass(state)}`;
    announcerEl.textContent = label;
  }

  function reply(id, result) {
    return replyJsonRpc(socket, id, result);
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
      connect();
    } catch (error) {
      endpointErrorEl.textContent = error.message || 'Enter a valid wss:// endpoint.';
      endpointInputEl.focus();
    } finally {
      saveEndpointEl.disabled = false;
      saveEndpointEl.removeAttribute('aria-busy');
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
