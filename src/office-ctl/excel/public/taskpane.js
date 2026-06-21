(() => {
  const ADDIN_VERSION = '0.1.11';
  const PROTOCOL_VERSION = '1.0';
  const { boolLabel, escapeHtml, fileName, formatDuration, formatTime, titleCase, redactText } = window.OfficeCtlCommon;
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
    middleTruncate,
    protectionLabel,
    renderRuntimeVersions,
    renderStaticMetadata,
    renderToolModeControl: renderSharedToolModeControl,
    setConnectionState: setSharedConnectionState,
    setCopyableMetadata,
    statusClass,
    taskMetadataMarkup
  } = window.OfficeCtlMainUi;

  const AVAILABLE_TOOLS = [
    'excel.get_workbook_info',
    'excel.list_sheets',
    'excel.add_sheet',
    'excel.update_sheet',
    'excel.delete_sheet',
    'excel.get_used_range',
    'excel.read_range',
    'excel.write_range',
    'excel.clear_range',
    'excel.find_replace_cells',
    'excel.set_formula',
    'excel.format_range',
    'excel.sort_range',
    'excel.apply_filter',
    'excel.create_table',
    'excel.update_table',
    'excel.create_chart',
    'excel.update_chart',
    'excel.create_pivot_table',
    'excel.update_pivot_table'
  ];
  const TOOL_GROUPS = [
    { label: 'Workbook', tools: ['excel.get_workbook_info'] },
    { label: 'Worksheet', tools: ['excel.list_sheets', 'excel.add_sheet', 'excel.update_sheet', 'excel.delete_sheet'] },
    { label: 'Range', tools: ['excel.get_used_range', 'excel.read_range', 'excel.write_range', 'excel.clear_range', 'excel.find_replace_cells'] },
    { label: 'Formula', tools: ['excel.set_formula'] },
    { label: 'Format', tools: ['excel.format_range'] },
    { label: 'Data', tools: ['excel.sort_range', 'excel.apply_filter'] },
    { label: 'Table', tools: ['excel.create_table', 'excel.update_table'] },
    { label: 'Chart', tools: ['excel.create_chart', 'excel.update_chart'] },
    { label: 'PivotTable', tools: ['excel.create_pivot_table', 'excel.update_pivot_table'] }
  ];
  const TOOL_METADATA = new Map([
    ['excel.get_workbook_info', { category: 'Workbook', sideEffect: 'read', description: 'Read workbook state and aggregate object counts.' }],
    ['excel.list_sheets', { category: 'Worksheet', sideEffect: 'read', description: 'List workbook worksheets.' }],
    ['excel.add_sheet', { category: 'Worksheet', sideEffect: 'mutating', description: 'Add a worksheet to the workbook.' }],
    ['excel.update_sheet', { category: 'Worksheet', sideEffect: 'mutating', description: 'Rename, activate, move, or restyle a worksheet tab.' }],
    ['excel.delete_sheet', { category: 'Worksheet', sideEffect: 'destructive', description: 'Delete a worksheet.' }],
    ['excel.get_used_range', { category: 'Range', sideEffect: 'read', description: 'Read the used range address and dimensions.' }],
    ['excel.read_range', { category: 'Range', sideEffect: 'read', description: 'Read values, text, formulas, and number formats from a range.' }],
    ['excel.write_range', { category: 'Range', sideEffect: 'mutating', description: 'Write a value matrix into a range.' }],
    ['excel.clear_range', { category: 'Range', sideEffect: 'destructive', description: 'Clear contents, formats, or delete cells in a range.' }],
    ['excel.find_replace_cells', { category: 'Range', sideEffect: 'mutating', description: 'Find cells in a range and optionally replace matches.' }],
    ['excel.set_formula', { category: 'Formula', sideEffect: 'mutating', description: 'Set formulas in a range.' }],
    ['excel.format_range', { category: 'Format', sideEffect: 'mutating', description: 'Apply formatting to a range.' }],
    ['excel.sort_range', { category: 'Data', sideEffect: 'mutating', description: 'Sort a range or table by one or more keys.' }],
    ['excel.apply_filter', { category: 'Data', sideEffect: 'mutating', description: 'Apply, clear, remove, or reapply range and table filters.' }],
    ['excel.create_table', { category: 'Table', sideEffect: 'mutating', description: 'Create a table from a range.' }],
    ['excel.update_table', { category: 'Table', sideEffect: 'destructive', description: 'Read or update table structure, style, options, and lifecycle.' }],
    ['excel.create_chart', { category: 'Chart', sideEffect: 'mutating', description: 'Create a chart from a range.' }],
    ['excel.update_chart', { category: 'Chart', sideEffect: 'destructive', description: 'Read or update chart title, legend, axes, source, position, size, export, and lifecycle.' }],
    ['excel.create_pivot_table', { category: 'PivotTable', sideEffect: 'mutating', description: 'Create a PivotTable from a range or table source.' }],
    ['excel.update_pivot_table', { category: 'PivotTable', sideEffect: 'destructive', description: 'Read or update PivotTable fields, layout, filters, refresh, and lifecycle.' }]
  ]);
  const { instanceId, sessionId } = runtimeIds();
  const TOOL_PERMISSION_STORAGE_KEY = `office-mcp.excel.tool-permissions.${sessionId}`;
  const TOOL_PERMISSION_MODE_STORAGE_KEY = `office-mcp.excel.tool-permission-mode.${sessionId}`;
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

  whenOfficeReady(async (info) => {
    if (!isExcelHost(info)) {
      setStatus('Unsupported host');
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

  function isExcelHost(info) {
    const host = String(info?.host || '').toLowerCase();
    const expected = String(Office.HostType?.Excel || 'Excel').toLowerCase();
    const diagnosticsHost = String(Office.context?.diagnostics?.host || '').toLowerCase();
    const hasExcelRuntime = typeof window.Excel?.run === 'function';
    const hasExcelRequirementSet = Office.context?.requirements?.isSetSupported?.('ExcelApi', '1.1') === true;
    return host === expected || host === 'excel' || diagnosticsHost === 'excel' || hasExcelRuntime || hasExcelRequirementSet;
  }

  function connect() {
    clearTimeout(reconnectTimer);
    const endpoint = configuredEndpoint();
    setCopyableMetadata(daemonEl, endpoint);
    endpointInputEl.value = endpoint;
    endpointDirty = false;
    setConnectionState('connecting', 'Connecting…');
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
        app: 'excel',
        version: Office.context.diagnostics?.version || null,
        platform: String(Office.context.platform || 'unknown').toLowerCase(),
        build: Office.context.diagnostics?.host || 'Desktop'
      },
      add_in: {
        version: ADDIN_VERSION,
        protocol_version: PROTOCOL_VERSION,
        requirement_sets: probeRequirementSets(),
        supported_features: ['workbook.session']
      }
    }));
    rememberRegisterRequest(requestId);
  }

  async function announceSession() {
    const workbook = await getWorkbookInfo();
    documentInfo = workbook;
    logger.info('session.added', { sessionId, workbook });
    send(sessionAddedNotification({
      session_id: sessionId,
      instance_id: instanceId,
      document: workbook,
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
      connectionDetailEl.textContent = error.message || 'Failed to announce workbook session.';
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
    const tool = message.params?.tool || 'excel.unknown';
    const args = message.params?.args || {};
    taskStore.start(requestId, tool, message.params || {}, message.params?.timeout_ms);
    renderCurrentTask();
    try {
      if (!isToolEnabled(tool)) throw toolDisabledError(tool);
      if (taskStore.isCancelled(requestId)) throw cancelledError(tool);
      let data;
      switch (tool) {
        case 'excel.get_workbook_info':
          data = await getWorkbookInfoTool(args);
          break;
        case 'excel.list_sheets':
          data = await listSheets(args);
          break;
        case 'excel.add_sheet':
          data = await addSheet(args);
          break;
        case 'excel.update_sheet':
          data = await updateSheet(args);
          break;
        case 'excel.delete_sheet':
          data = await deleteSheet(args);
          break;
        case 'excel.get_used_range':
          data = await getUsedRange(args);
          break;
        case 'excel.read_range':
          data = await readRange(args);
          break;
        case 'excel.write_range':
          data = await writeRange(args);
          break;
        case 'excel.clear_range':
          data = await clearRange(args);
          break;
        case 'excel.find_replace_cells':
          data = await findReplaceCells(args);
          break;
        case 'excel.set_formula':
          data = await setFormula(args);
          break;
        case 'excel.format_range':
          data = await formatRange(args);
          break;
        case 'excel.sort_range':
          data = await sortRange(args);
          break;
        case 'excel.apply_filter':
          data = await applyFilter(args);
          break;
        case 'excel.create_table':
          data = await createTable(args);
          break;
        case 'excel.update_table':
          data = await updateTable(args);
          break;
        case 'excel.create_chart':
          data = await createChart(args);
          break;
        case 'excel.update_chart':
          data = await updateChart(args);
          break;
        case 'excel.create_pivot_table':
          data = await createPivotTable(args);
          break;
        case 'excel.update_pivot_table':
          data = await updatePivotTable(args);
          break;
        default:
          throw Object.assign(new Error(`Unsupported tool ${tool}`), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE' });
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

  async function getWorkbookInfoTool(args) {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheets = workbook.worksheets;
      const tables = workbook.tables;
      worksheets.load('items/id,items/name,items/position,items/visibility,items/tabColor');
      tables.load('items/name');
      await context.sync();
      let activeSheet = null;
      try {
        const sheet = worksheets.getActiveWorksheet();
        sheet.load('id,name,visibility');
        await context.sync();
        activeSheet = sheetInfo(sheet, true);
      } catch (error) {
        logger.warn('excel.active_sheet_probe.failed', error);
        if (worksheets.items[0]) activeSheet = sheetInfo(worksheets.items[0], false);
      }
      return {
        title: documentInfo?.title || fileName(Office.context.document?.url || '') || 'Excel Workbook',
        url: Office.context.document?.url || null,
        filename: fileName(Office.context.document?.url || '') || null,
        active_sheet: activeSheet,
        sheet_count: worksheets.items.length,
        table_count: tables.items.length,
        is_dirty: documentInfo?.is_dirty ?? null,
        is_read_only: documentInfo?.is_read_only ?? false,
        is_protected: documentInfo?.is_protected ?? false
      };
    });
  }

  async function listSheets(args) {
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      const activeSheet = worksheets.getActiveWorksheet();
      worksheets.load('items/id,items/name,items/position,items/visibility,items/tabColor');
      activeSheet.load('id');
      await context.sync();
      return {
        sheets: worksheets.items.map((sheet) => sheetInfo(sheet, sheet.id === activeSheet.id))
      };
    });
  }
  async function readRange(args) {
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      range.load('address,values,text,rowCount,columnCount,numberFormat');
      await context.sync();
      return {
        address: range.address,
        values: range.values,
        text: range.text,
        row_count: range.rowCount,
        column_count: range.columnCount,
        number_format: range.numberFormat,
        untrusted_source: true
      };
    });
  }

  async function writeRange(args) {
    if (!Array.isArray(args.values) || !Array.isArray(args.values[0])) {
      throw Object.assign(new Error('excel.write_range requires a two-dimensional values array.'), { officeMcpCode: 'INVALID_ARGUMENT' });
    }
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      range.values = args.values;
      await context.sync();
      return {
        address: args.address,
        row_count: args.values.length,
        column_count: args.values[0].length,
        wrote_values: true
      };
    });
  }

  async function clearRange(args) {
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      const deleteShift = optionalTrimmedString(args.delete_shift);
      if (deleteShift) {
        range.delete(deleteShiftDirectionFrom(deleteShift));
        await context.sync();
        return { address: args.address, deleted: true, delete_shift: deleteShift };
      }
      const applyTo = clearApplyToFrom(args.apply_to || 'contents');
      range.clear(applyTo);
      await context.sync();
      return { address: args.address, cleared: true, apply_to: applyTo };
    });
  }

  async function findReplaceCells(args) {
    if (!supportsRequirementSet('ExcelApi', '1.9')) {
      throw Object.assign(new Error('excel.find_replace_cells requires ExcelApi 1.9.'), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE', partialEffect: 'none' });
    }
    const query = requiredString(args, 'query', 'excel.find_replace_cells requires query.');
    const hasReplacement = args.replacement !== undefined;
    const replacement = hasReplacement ? String(args.replacement) : null;
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      const criteria = searchCriteriaFrom(args);
      if (hasReplacement) {
        const result = range.replaceAll(query, replacement, replaceCriteriaFrom(args));
        await context.sync();
        return { query, replaced_count: result.value, replaced: true };
      }
      const match = range.findOrNullObject(query, criteria);
      match.load('address,text,rowCount,columnCount,isNullObject');
      await context.sync();
      if (match.isNullObject) return { query, matches: [], match_count: 0, untrusted_source: true };
      return {
        query,
        matches: [{ address: match.address, text: match.text, row_count: match.rowCount, column_count: match.columnCount }],
        match_count: 1,
        untrusted_source: true
      };
    });
  }

  async function updateSheet(args) {
    const sheet = requiredString(args, 'sheet', 'excel.update_sheet requires sheet.');
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      const worksheet = worksheets.getItem(sheet);
      const renamedTo = optionalTrimmedString(args.name);
      if (renamedTo) worksheet.name = renamedTo;
      if (args.visibility) worksheet.visibility = String(args.visibility);
      if (args.tab_color) worksheet.tabColor = String(args.tab_color);
      if (Number.isInteger(args.position)) worksheet.position = args.position;
      if (args.activate === true) worksheet.activate();
      worksheet.load('id,name,position,visibility,tabColor');
      await context.sync();
      return { sheet: sheetInfo(worksheet, args.activate === true), updated: true };
    });
  }

  async function deleteSheet(args) {
    const sheet = requiredString(args, 'sheet', 'excel.delete_sheet requires sheet.');
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items/name');
      await context.sync();
      if (worksheets.items.length <= 1) {
        throw Object.assign(new Error('excel.delete_sheet cannot delete the only worksheet.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      }
      const worksheet = worksheets.getItem(sheet);
      worksheet.delete();
      await context.sync();
      return { sheet, deleted: true };
    });
  }

  async function getUsedRange(args) {
    return Excel.run(async (context) => {
      const worksheet = targetWorksheet(context, args);
      const usedRange = worksheet.getUsedRangeOrNullObject(true);
      usedRange.load('address,rowCount,columnCount,isNullObject');
      await context.sync();
      if (usedRange.isNullObject) {
        return { sheet: args.sheet || null, address: null, row_count: 0, column_count: 0, is_empty: true };
      }
      return {
        address: usedRange.address,
        row_count: usedRange.rowCount,
        column_count: usedRange.columnCount,
        is_empty: false
      };
    });
  }
  async function addSheet(args) {
    const name = String(args.name || '').trim();
    return Excel.run(async (context) => {
      const sheet = name ? context.workbook.worksheets.add(name) : context.workbook.worksheets.add();
      sheet.load('name');
      if (args.activate !== false) sheet.activate();
      await context.sync();
      return { sheet: sheet.name, activated: args.activate !== false };
    });
  }

  async function setFormula(args) {
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      range.load('rowCount,columnCount');
      await context.sync();
      range.formulas = formulaMatrixFrom(args, range.rowCount, range.columnCount);
      await context.sync();
      return {
        address: args.address,
        formula: args.formula !== undefined ? String(args.formula) : null,
        formula_matrix: Array.isArray(args.formulas),
        wrote_formula: true
      };
    });
  }

  async function formatRange(args) {
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      if (args.number_format !== undefined || args.number_formats !== undefined) {
        range.load('rowCount,columnCount');
        await context.sync();
      }
      if (args.bold !== undefined) range.format.font.bold = Boolean(args.bold);
      if (args.italic !== undefined) range.format.font.italic = Boolean(args.italic);
      if (args.font_color) range.format.font.color = String(args.font_color);
      if (args.fill_color) range.format.fill.color = String(args.fill_color);
      if (args.number_format !== undefined) range.numberFormat = matrixFromScalar(String(args.number_format), range.rowCount, range.columnCount);
      if (args.number_formats !== undefined) range.numberFormat = numberFormatMatrixFrom(args, range.rowCount, range.columnCount);
      if (args.horizontal_alignment !== undefined) range.format.horizontalAlignment = alignmentFrom(args.horizontal_alignment, 'horizontal');
      if (args.vertical_alignment !== undefined) range.format.verticalAlignment = alignmentFrom(args.vertical_alignment, 'vertical');
      if (args.wrap_text !== undefined) range.format.wrapText = Boolean(args.wrap_text);
      if (args.borders !== undefined) applyBorders(range, args.borders);
      if (args.autofit_columns === true || args.autofit_rows === true) requireRequirementSet('ExcelApi', '1.2', 'autofit formatting');
      if (args.autofit_columns === true) range.format.autofitColumns();
      if (args.autofit_rows === true) range.format.autofitRows();
      await context.sync();
      return { address: args.address, formatted: true };
    });
  }

  async function sortRange(args) {
    requireRequirementSet('ExcelApi', '1.2', 'range sorting');
    return Excel.run(async (context) => {
      const target = targetSortObject(context, args);
      const action = String(args.action || 'apply').trim().toLowerCase();
      if (action === 'clear') {
        if (!target.table) throw Object.assign(new Error('excel.sort_range clear is only supported for table targets.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
        target.table.sort.clear();
        await context.sync();
        return { target_type: 'table', table: args.table, sorted: false, cleared: true };
      }
      if (action === 'reapply') {
        if (!target.table) throw Object.assign(new Error('excel.sort_range reapply is only supported for table targets.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
        target.table.sort.reapply();
        await context.sync();
        return { target_type: 'table', table: args.table, reapplied: true };
      }
      if (target.table) {
        const table = target.table;
        table.sort.apply(sortFieldsFrom(args.fields), Boolean(args.match_case), sortMethodFrom(args.method));
        await context.sync();
        return { target_type: 'table', table: args.table, sorted: true };
      }
      const sort = target.range.sort;
      sort.apply(sortFieldsFrom(args.fields), Boolean(args.match_case), args.has_headers !== false, sortOrientationFrom(args.orientation), sortMethodFrom(args.method));
      await context.sync();
      return { target_type: 'range', address: args.address, sorted: true };
    });
  }

  async function applyFilter(args) {
    const targetType = String(args.target_type || (args.table ? 'table' : 'range')).trim().toLowerCase();
    const action = String(args.action || 'apply').trim().toLowerCase();
    if (targetType === 'range') requireRequirementSet('ExcelApi', '1.9', 'range filtering');
    else if (targetType === 'table') requireRequirementSet('ExcelApi', '1.2', 'table filtering');
    else throw Object.assign(new Error(`Unsupported filter target ${args.target_type}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    return Excel.run(async (context) => {
      if (targetType === 'table') {
        const table = targetTable(context, args);
        if (action === 'clear' || action === 'remove') {
          table.autoFilter.clearCriteria();
          await context.sync();
          return { target_type: 'table', table: args.table, filtered: false, cleared: true };
        }
        if (action === 'reapply') {
          table.autoFilter.reapply();
          await context.sync();
          return { target_type: 'table', table: args.table, reapplied: true };
        }
        table.columns.getItem(requiredString(args, 'column', 'excel.apply_filter requires column for table filters.')).filter.apply(filterCriteriaFrom(args.criteria));
        await context.sync();
        return { target_type: 'table', table: args.table, column: String(args.column), filtered: true };
      }
      const range = targetRange(context, args);
      const autoFilter = targetWorksheet(context, args).autoFilter;
      if (action === 'clear') {
        autoFilter.clearCriteria();
        await context.sync();
        return { target_type: 'range', address: args.address, filtered: false, cleared: true };
      }
      if (action === 'remove') {
        autoFilter.remove();
        await context.sync();
        return { target_type: 'range', address: args.address, filtered: false, removed: true };
      }
      if (action === 'reapply') {
        autoFilter.reapply();
        await context.sync();
        return { target_type: 'range', address: args.address, reapplied: true };
      }
      if (!Number.isInteger(Number(args.column_index))) {
        throw Object.assign(new Error('excel.apply_filter requires column_index for range filters.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
      }
      autoFilter.apply(range, Number(args.column_index), filterCriteriaFrom(args.criteria));
      await context.sync();
      return { target_type: 'range', address: args.address, column_index: Number(args.column_index), filtered: true };
    });
  }

  async function createTable(args) {
    const hasHeaders = args.has_headers !== false;
    const name = String(args.name || '').trim();
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      const table = context.workbook.tables.add(range, hasHeaders);
      if (name) table.name = name;
      table.load('name');
      await context.sync();
      return { table: table.name, address: args.address, has_headers: hasHeaders };
    });
  }

  async function updateTable(args) {
    const action = String(args.action || 'metadata').trim().toLowerCase();
    return Excel.run(async (context) => {
      const table = targetTable(context, args);
      if (action === 'metadata' || action === 'read') {
        return readTableMetadata(context, table);
      }
      if (action === 'add_rows') {
        table.rows.add(optionalIndex(args.index), tableValuesFrom(args.values, 'rows'), args.always_insert !== false);
        await context.sync();
        return { table: args.table, action, added_rows: Array.isArray(args.values) ? args.values.length : 1 };
      }
      if (action === 'add_columns') {
        table.columns.add(optionalIndex(args.index), tableValuesFrom(args.values, 'columns'), optionalName(args.name));
        await context.sync();
        return { table: args.table, action, added_columns: 1 };
      }
      if (action === 'resize') {
        requireRequirementSet('ExcelApi', '1.13', 'table resize');
        table.resize(requiredString(args, 'address', 'excel.update_table resize requires address.'));
        await context.sync();
        return { table: args.table, action, resized: true, address: args.address };
      }
      if (action === 'rename') {
        table.name = requiredString(args, 'name', 'excel.update_table rename requires name.');
        table.load('id,name');
        await context.sync();
        return { table: table.name, action, renamed: true, id: table.id || null };
      }
      if (action === 'options' || action === 'style') {
        applyTableOptions(table, args);
        await context.sync();
        return { table: args.table, action, updated: true };
      }
      if (action === 'delete') {
        table.delete();
        await context.sync();
        return { table: args.table, action, deleted: true };
      }
      throw Object.assign(new Error(`Unsupported table action ${args.action}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    });
  }

  async function createChart(args) {
    const typeName = String(args.type || 'columnClustered');
    const type = chartTypeFrom(typeName);
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      const worksheet = args.sheet
        ? context.workbook.worksheets.getItem(String(args.sheet))
        : context.workbook.worksheets.getActiveWorksheet();
      const chart = worksheet.charts.add(type, range, Excel.ChartSeriesBy.auto);
      if (args.title) chart.title.text = String(args.title);
      chart.load('name');
      await context.sync();
      return { chart: chart.name, chart_type: typeName, source: args.address };
    });
  }

  async function updateChart(args) {
    const action = String(args.action || 'metadata').trim().toLowerCase();
    return Excel.run(async (context) => {
      const chart = targetChart(context, args);
      if (action === 'metadata' || action === 'read') {
        const chartProperties = supportsRequirementSet('ExcelApi', '1.7')
          ? 'id,name,chartType,left,top,width,height'
          : 'name,left,top,width,height';
        chart.load(chartProperties);
        chart.title.load('text,visible');
        chart.legend.load('visible,position,overlay');
        chart.series.load('count');
        await context.sync();
        return chartMetadata(chart);
      }
      if (action === 'title') {
        chart.title.text = requiredString(args, 'title', 'excel.update_chart title requires title.');
        if (args.visible !== undefined) chart.title.visible = Boolean(args.visible);
        await context.sync();
        return { chart: args.chart, action, updated: true };
      }
      if (action === 'legend') {
        if (args.visible !== undefined) chart.legend.visible = Boolean(args.visible);
        if (args.position !== undefined) chart.legend.position = chartLegendPositionFrom(args.position);
        if (args.overlay !== undefined) chart.legend.overlay = Boolean(args.overlay);
        await context.sync();
        return { chart: args.chart, action, updated: true };
      }
      if (action === 'axis') {
        requireRequirementSet('ExcelApi', '1.7', 'chart axis selection');
        const axis = chart.axes.getItem(chartAxisTypeFrom(args.axis), chartAxisGroupFrom(args.axis_group));
        if (args.title !== undefined) axis.title.text = String(args.title);
        if (args.title_visible !== undefined) axis.title.visible = Boolean(args.title_visible);
        if (args.visible !== undefined) axis.visible = Boolean(args.visible);
        await context.sync();
        return { chart: args.chart, action, axis: String(args.axis || 'value'), updated: true };
      }
      if (action === 'data' || action === 'series_source') {
        chart.setData(targetRange(context, args), chartSeriesByFrom(args.series_by));
        await context.sync();
        return { chart: args.chart, action, source: args.address, updated: true };
      }
      if (action === 'position') {
        chart.setPosition(requiredString(args, 'start_cell', 'excel.update_chart position requires start_cell.'), optionalName(args.end_cell));
        await context.sync();
        return { chart: args.chart, action, positioned: true };
      }
      if (action === 'size') {
        if (args.width !== undefined) chart.width = Number(args.width);
        if (args.height !== undefined) chart.height = Number(args.height);
        await context.sync();
        return { chart: args.chart, action, width: args.width ?? null, height: args.height ?? null, resized: true };
      }
      if (action === 'export_image') {
        requireRequirementSet('ExcelApi', '1.2', 'chart image export');
        const image = chart.getImage(optionalNumber(args.width), optionalNumber(args.height), chartImageFittingModeFrom(args.fitting_mode));
        await context.sync();
        return { chart: args.chart, action, image_base64: image.value, mime_type: 'image/png', untrusted_source: true };
      }
      if (action === 'delete') {
        chart.delete();
        await context.sync();
        return { chart: args.chart, action, deleted: true };
      }
      throw Object.assign(new Error(`Unsupported chart action ${args.action}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    });
  }

  async function createPivotTable(args) {
    requireRequirementSet('ExcelApi', '1.8', 'pivot table creation');
    return Excel.run(async (context) => {
      const pivot = context.workbook.pivotTables.add(requiredString(args, 'name', 'excel.create_pivot_table requires name.'), pivotSource(context, args), requiredString(args, 'destination', 'excel.create_pivot_table requires destination.'));
      pivot.load('name');
      await context.sync();
      return { pivot_table: pivot.name, source: args.table || args.address, destination: args.destination, created: true };
    });
  }

  async function updatePivotTable(args) {
    const action = String(args.action || 'metadata').trim().toLowerCase();
    return Excel.run(async (context) => {
      const pivot = targetPivotTable(context, args);
      if (action === 'metadata' || action === 'read') {
        return readPivotTableMetadata(context, pivot);
      }
      if (action === 'refresh') {
        pivot.refresh();
        await context.sync();
        return { pivot_table: args.pivot_table, action, refreshed: true };
      }
      if (action === 'add_hierarchy' || action === 'remove_hierarchy') {
        requireRequirementSet('ExcelApi', '1.8', 'pivot table hierarchy updates');
        updatePivotHierarchy(pivot, args);
        await context.sync();
        return { pivot_table: args.pivot_table, action, axis: args.axis || null, hierarchy: args.hierarchy || null, updated: true };
      }
      if (action === 'layout') {
        requireRequirementSet('ExcelApi', '1.8', 'pivot table layout');
        applyPivotLayoutOptions(pivot, args);
        await context.sync();
        return { pivot_table: args.pivot_table, action, updated: true };
      }
      if (action === 'filter') {
        requireRequirementSet('ExcelApi', '1.12', 'pivot table filters');
        pivotField(pivot, args).applyFilter(pivotFiltersFrom(args));
        await context.sync();
        return { pivot_table: args.pivot_table, action, field: args.field, filtered: true };
      }
      if (action === 'clear_filters') {
        requireRequirementSet('ExcelApi', '1.12', 'pivot table filters');
        pivotField(pivot, args).clearAllFilters();
        await context.sync();
        return { pivot_table: args.pivot_table, action, field: args.field, cleared: true };
      }
      if (action === 'delete') {
        requireRequirementSet('ExcelApi', '1.8', 'pivot table deletion');
        pivot.delete();
        await context.sync();
        return { pivot_table: args.pivot_table, action, deleted: true };
      }
      throw Object.assign(new Error(`Unsupported PivotTable action ${args.action}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    });
  }

  function targetWorksheet(context, args) {
    return args.sheet
      ? context.workbook.worksheets.getItem(String(args.sheet))
      : context.workbook.worksheets.getActiveWorksheet();
  }

  function targetTable(context, args) {
    const table = requiredString(args, 'table', 'Table name is required.');
    return context.workbook.tables.getItem(table);
  }

  function targetChart(context, args) {
    const chart = requiredString(args, 'chart', 'Chart name is required.');
    const worksheet = args.sheet
      ? context.workbook.worksheets.getItem(String(args.sheet))
      : context.workbook.worksheets.getActiveWorksheet();
    return worksheet.charts.getItem(chart);
  }

  function targetSortObject(context, args) {
    const targetType = String(args.target_type || (args.table ? 'table' : 'range')).trim().toLowerCase();
    if (targetType === 'table') return { table: targetTable(context, args) };
    if (targetType === 'range') return { range: targetRange(context, args) };
    throw Object.assign(new Error(`Unsupported sort target ${args.target_type}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  async function readTableMetadata(context, table) {
    table.load('id,name,showHeaders,showTotals,style');
    table.rows.load('count');
    table.columns.load('count');
    await context.sync();
    const range = table.getRange();
    const bodyRange = table.getDataBodyRange();
    const headerRange = table.showHeaders ? table.getHeaderRowRange() : null;
    const totalRange = table.showTotals ? table.getTotalRowRange() : null;
    range.load('address,rowCount,columnCount');
    if (bodyRange) bodyRange.load('address,rowCount,columnCount');
    if (headerRange) headerRange.load('address,rowCount,columnCount');
    if (totalRange) totalRange.load('address,rowCount,columnCount');
    await context.sync();
    return tableMetadata(table, range, bodyRange, headerRange, totalRange);
  }

  function tableMetadata(table, range, bodyRange, headerRange, totalRange) {
    return {
      table: table.name,
      id: table.id || null,
      show_headers: Boolean(table.showHeaders),
      show_totals: Boolean(table.showTotals),
      style: table.style || null,
      row_count: table.rows.count,
      column_count: table.columns.count,
      range: rangeInfo(range),
      data_body_range: rangeInfo(bodyRange),
      header_row_range: rangeInfo(headerRange),
      total_row_range: rangeInfo(totalRange)
    };
  }

  function rangeInfo(range) {
    if (!range) return { address: null, row_count: null, column_count: null };
    return {
      address: range.address || null,
      row_count: Number.isInteger(range.rowCount) ? range.rowCount : null,
      column_count: Number.isInteger(range.columnCount) ? range.columnCount : null
    };
  }

  function applyTableOptions(table, args) {
    if (args.style !== undefined) table.style = String(args.style);
    if (args.show_headers !== undefined) table.showHeaders = Boolean(args.show_headers);
    if (args.show_totals !== undefined) table.showTotals = Boolean(args.show_totals);
    const visualOptionKeys = [
      'highlight_first_column',
      'highlight_last_column',
      'show_banded_columns',
      'show_banded_rows',
      'show_filter_button'
    ];
    if (visualOptionKeys.some((key) => args[key] !== undefined)) {
      requireRequirementSet('ExcelApi', '1.3', 'table visual options');
    }
    if (args.highlight_first_column !== undefined) table.highlightFirstColumn = Boolean(args.highlight_first_column);
    if (args.highlight_last_column !== undefined) table.highlightLastColumn = Boolean(args.highlight_last_column);
    if (args.show_banded_columns !== undefined) table.showBandedColumns = Boolean(args.show_banded_columns);
    if (args.show_banded_rows !== undefined) table.showBandedRows = Boolean(args.show_banded_rows);
    if (args.show_filter_button !== undefined) table.showFilterButton = Boolean(args.show_filter_button);
  }

  function tableValuesFrom(values, label) {
    if (values === undefined || values === null) return undefined;
    if (!Array.isArray(values) || !Array.isArray(values[0])) {
      throw Object.assign(new Error(`Table ${label} values must be a two-dimensional array.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return values;
  }

  function optionalIndex(value) {
    if (value === undefined || value === null || value === '') return undefined;
    if (!Number.isInteger(Number(value))) {
      throw Object.assign(new Error('Table index must be an integer.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return Number(value);
  }

  function optionalName(value) {
    const name = optionalTrimmedString(value);
    return name || undefined;
  }

  function chartMetadata(chart) {
    return {
      chart: chart.name,
      id: chart.id || null,
      chart_type: chart.chartType || null,
      title: { text: chart.title.text || '', visible: Boolean(chart.title.visible) },
      legend: {
        visible: Boolean(chart.legend.visible),
        position: chart.legend.position || null,
        overlay: Boolean(chart.legend.overlay)
      },
      series_count: chart.series.count,
      position: {
        left: chart.left,
        top: chart.top,
        width: chart.width,
        height: chart.height
      }
    };
  }

  function chartLegendPositionFrom(value) {
    const positions = {
      top: Excel.ChartLegendPosition.top,
      bottom: Excel.ChartLegendPosition.bottom,
      left: Excel.ChartLegendPosition.left,
      right: Excel.ChartLegendPosition.right,
      corner: Excel.ChartLegendPosition.corner,
      custom: Excel.ChartLegendPosition.custom
    };
    return enumValueFrom(positions, value, 'chart legend position');
  }

  function chartAxisTypeFrom(value) {
    const types = {
      category: Excel.ChartAxisType.category,
      value: Excel.ChartAxisType.value,
      series: Excel.ChartAxisType.series
    };
    return enumValueFrom(types, value || 'value', 'chart axis type');
  }

  function chartAxisGroupFrom(value) {
    if (value === undefined || value === null || value === '') return undefined;
    const groups = {
      primary: Excel.ChartAxisGroup.primary,
      secondary: Excel.ChartAxisGroup.secondary
    };
    return enumValueFrom(groups, value, 'chart axis group');
  }

  function chartSeriesByFrom(value) {
    if (value === undefined || value === null || value === '') return Excel.ChartSeriesBy.auto;
    const values = {
      auto: Excel.ChartSeriesBy.auto,
      columns: Excel.ChartSeriesBy.columns,
      rows: Excel.ChartSeriesBy.rows
    };
    return enumValueFrom(values, value, 'chart series orientation');
  }

  function chartImageFittingModeFrom(value) {
    if (value === undefined || value === null || value === '') return undefined;
    const modes = {
      fit: Excel.ImageFittingMode.fit,
      fitAndCenter: Excel.ImageFittingMode.fitAndCenter,
      fill: Excel.ImageFittingMode.fill
    };
    return enumValueFrom(modes, value, 'chart image fitting mode');
  }

  function optionalNumber(value) {
    if (value === undefined || value === null || value === '') return undefined;
    const number = Number(value);
    if (!Number.isFinite(number)) {
      throw Object.assign(new Error('Expected a finite number.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return number;
  }

  function targetPivotTable(context, args) {
    const name = requiredString(args, 'pivot_table', 'PivotTable name is required.');
    return context.workbook.pivotTables.getItem(name);
  }

  function pivotSource(context, args) {
    if (args.table) return context.workbook.tables.getItem(String(args.table));
    return targetRange(context, args);
  }

  async function readPivotTableMetadata(context, pivot) {
    pivot.load('id,name');
    pivot.layout.load('layoutType,showColumnGrandTotals,showRowGrandTotals');
    pivot.rowHierarchies.load('items/name,items/position');
    pivot.columnHierarchies.load('items/name,items/position');
    pivot.filterHierarchies.load('items/name,items/position');
    pivot.dataHierarchies.load('items/name,items/position,items/summarizeBy,items/numberFormat');
    const range = pivot.layout.getRange();
    range.load('address,rowCount,columnCount');
    await context.sync();
    return pivotTableMetadata(pivot, range);
  }

  function pivotTableMetadata(pivot, range) {
    return {
      pivot_table: pivot.name,
      id: pivot.id || null,
      range: rangeInfo(range),
      layout_type: pivot.layout.layoutType || null,
      show_column_grand_totals: Boolean(pivot.layout.showColumnGrandTotals),
      show_row_grand_totals: Boolean(pivot.layout.showRowGrandTotals),
      row_hierarchies: pivot.rowHierarchies.items.map(pivotHierarchyInfo),
      column_hierarchies: pivot.columnHierarchies.items.map(pivotHierarchyInfo),
      filter_hierarchies: pivot.filterHierarchies.items.map(pivotHierarchyInfo),
      data_hierarchies: pivot.dataHierarchies.items.map((item) => ({
        name: item.name,
        position: item.position,
        summarize_by: item.summarizeBy || null,
        number_format: item.numberFormat || null
      }))
    };
  }

  function pivotHierarchyInfo(item) {
    return { name: item.name, position: item.position };
  }

  function updatePivotHierarchy(pivot, args) {
    const axis = String(args.axis || '').trim().toLowerCase();
    const hierarchy = requiredString(args, 'hierarchy', 'excel.update_pivot_table hierarchy actions require hierarchy.');
    const action = String(args.action || '').trim().toLowerCase();
    if (action === 'remove_hierarchy') {
      removePivotHierarchy(pivot, axis, hierarchy);
      return;
    }
    if (axis === 'row') {
      pivot.rowHierarchies.add(pivot.hierarchies.getItem(hierarchy));
      return;
    }
    if (axis === 'column') {
      pivot.columnHierarchies.add(pivot.hierarchies.getItem(hierarchy));
      return;
    }
    if (axis === 'filter') {
      pivot.filterHierarchies.add(pivot.hierarchies.getItem(hierarchy));
      return;
    }
    if (axis === 'data') {
      const dataHierarchy = pivot.dataHierarchies.add(pivot.hierarchies.getItem(hierarchy));
      if (args.summarize_by !== undefined) dataHierarchy.summarizeBy = aggregationFunctionFrom(args.summarize_by);
      if (args.number_format !== undefined) dataHierarchy.numberFormat = String(args.number_format);
      return;
    }
    throw Object.assign(new Error(`Unsupported PivotTable axis ${args.axis}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  function removePivotHierarchy(pivot, axis, hierarchy) {
    if (axis === 'row') {
      pivot.rowHierarchies.remove(pivot.rowHierarchies.getItem(hierarchy));
      return;
    }
    if (axis === 'column') {
      pivot.columnHierarchies.remove(pivot.columnHierarchies.getItem(hierarchy));
      return;
    }
    if (axis === 'filter') {
      pivot.filterHierarchies.remove(pivot.filterHierarchies.getItem(hierarchy));
      return;
    }
    if (axis === 'data') {
      pivot.dataHierarchies.remove(pivot.dataHierarchies.getItem(hierarchy));
      return;
    }
    throw Object.assign(new Error(`Unsupported PivotTable axis ${axis}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
  }

  function applyPivotLayoutOptions(pivot, args) {
    if (args.layout_type !== undefined) pivot.layout.layoutType = pivotLayoutTypeFrom(args.layout_type);
    if (args.show_column_grand_totals !== undefined) pivot.layout.showColumnGrandTotals = Boolean(args.show_column_grand_totals);
    if (args.show_row_grand_totals !== undefined) pivot.layout.showRowGrandTotals = Boolean(args.show_row_grand_totals);
  }

  function aggregationFunctionFrom(value) {
    const values = {
      automatic: Excel.AggregationFunction.automatic,
      sum: Excel.AggregationFunction.sum,
      count: Excel.AggregationFunction.count,
      average: Excel.AggregationFunction.average,
      max: Excel.AggregationFunction.max,
      min: Excel.AggregationFunction.min,
      product: Excel.AggregationFunction.product,
      countNumbers: Excel.AggregationFunction.countNumbers,
      standardDeviation: Excel.AggregationFunction.standardDeviation,
      standardDeviationP: Excel.AggregationFunction.standardDeviationP,
      variance: Excel.AggregationFunction.variance,
      varianceP: Excel.AggregationFunction.varianceP
    };
    return enumValueFrom(values, value, 'PivotTable aggregation function');
  }

  function pivotLayoutTypeFrom(value) {
    const values = {
      compact: Excel.PivotLayoutType.compact,
      tabular: Excel.PivotLayoutType.tabular,
      outline: Excel.PivotLayoutType.outline
    };
    return enumValueFrom(values, value, 'PivotTable layout type');
  }

  function pivotFiltersFrom(args) {
    const selectedItems = args.selected_items || args.values;
    if (!Array.isArray(selectedItems) || selectedItems.length === 0) {
      throw Object.assign(new Error('excel.update_pivot_table filter requires selected_items.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return { manualFilter: { selectedItems: selectedItems.map(String) } };
  }

  function pivotField(pivot, args) {
    const hierarchy = requiredString(args, 'hierarchy', 'excel.update_pivot_table filter requires hierarchy.');
    const field = requiredString(args, 'field', 'excel.update_pivot_table filter requires field.');
    return pivot.hierarchies.getItem(hierarchy).fields.getItem(field);
  }

  function sheetInfo(sheet, active) {
    return {
      id: sheet.id || null,
      name: sheet.name,
      position: Number.isInteger(sheet.position) ? sheet.position : null,
      visibility: sheet.visibility || null,
      tab_color: sheet.tabColor || null,
      active: Boolean(active)
    };
  }

  function optionalTrimmedString(value) {
    const text = String(value || '').trim();
    return text || null;
  }

  function clearApplyToFrom(value) {
    const modes = {
      all: Excel.ClearApplyTo.all,
      formats: Excel.ClearApplyTo.formats,
      contents: Excel.ClearApplyTo.contents
    };
    const key = String(value || 'contents').trim().toLowerCase();
    if (!modes[key]) {
      throw Object.assign(new Error(`Unsupported clear mode ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return modes[key];
  }

  function deleteShiftDirectionFrom(value) {
    const shifts = {
      up: Excel.DeleteShiftDirection.up,
      left: Excel.DeleteShiftDirection.left
    };
    const key = String(value || '').trim().toLowerCase();
    if (!shifts[key]) {
      throw Object.assign(new Error(`Unsupported delete shift ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return shifts[key];
  }

  function searchCriteriaFrom(args) {
    return {
      completeMatch: Boolean(args.complete_match),
      matchCase: Boolean(args.match_case),
      searchDirection: String(args.search_direction || 'Forward')
    };
  }

  function replaceCriteriaFrom(args) {
    return {
      completeMatch: Boolean(args.complete_match),
      matchCase: Boolean(args.match_case)
    };
  }

  function sortFieldsFrom(fields) {
    if (!Array.isArray(fields) || fields.length === 0) {
      throw Object.assign(new Error('excel.sort_range requires at least one sort field.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return fields.map((field) => {
      if (!field || typeof field !== 'object' || Array.isArray(field)) {
        throw Object.assign(new Error('Sort field must be an object.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
      }
      if (!Number.isInteger(Number(field.key))) {
        throw Object.assign(new Error('Sort field key must be an integer offset.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
      }
      const result = {
        key: Number(field.key),
        ascending: field.ascending !== false
      };
      if (field.sort_on !== undefined) result.sortOn = sortOnFrom(field.sort_on);
      if (field.data_option !== undefined) result.dataOption = sortDataOptionFrom(field.data_option);
      if (field.color !== undefined) result.color = String(field.color);
      if (field.sub_field !== undefined) result.subField = String(field.sub_field);
      return result;
    });
  }

  function filterCriteriaFrom(criteria) {
    if (!criteria || typeof criteria !== 'object' || Array.isArray(criteria)) {
      throw Object.assign(new Error('excel.apply_filter requires criteria.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    const result = { filterOn: filterOnFrom(criteria.filter_on || criteria.filterOn || 'values') };
    if (criteria.criterion1 !== undefined) result.criterion1 = String(criteria.criterion1);
    if (criteria.criterion2 !== undefined) result.criterion2 = String(criteria.criterion2);
    if (criteria.color !== undefined) result.color = String(criteria.color);
    if (criteria.operator !== undefined) result.operator = filterOperatorFrom(criteria.operator);
    if (criteria.dynamic_criteria !== undefined) result.dynamicCriteria = dynamicFilterCriteriaFrom(criteria.dynamic_criteria);
    if (criteria.values !== undefined) result.values = criteria.values;
    if (criteria.sub_field !== undefined) result.subField = String(criteria.sub_field);
    return result;
  }

  function supportsRequirementSet(name, version) {
    return Office.context?.requirements?.isSetSupported?.(name, version) === true;
  }

  function requireRequirementSet(name, version, feature) {
    if (!supportsRequirementSet(name, version)) {
      throw Object.assign(new Error(`${feature} requires ${name} ${version}.`), {
        officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE',
        partialEffect: 'none'
      });
    }
  }

  function requiredString(args, key, message) {
    const value = String(args[key] || '').trim();
    if (!value) throw Object.assign(new Error(message), { officeMcpCode: 'INVALID_ARGUMENT' });
    return value;
  }

  function formulaMatrixFrom(args, rows, columns) {
    if (args.formulas !== undefined) {
      validateMatrixShape(args.formulas, rows, columns, 'formulas');
      return args.formulas.map((row) => row.map((value) => String(value ?? '')));
    }
    const formula = requiredString(args, 'formula', 'excel.set_formula requires formula or formulas.');
    return matrixFromScalar(formula, rows, columns);
  }

  function numberFormatMatrixFrom(args, rows, columns) {
    validateMatrixShape(args.number_formats, rows, columns, 'number_formats');
    return args.number_formats.map((row) => row.map((value) => String(value ?? 'General')));
  }

  function validateMatrixShape(matrix, rows, columns, label) {
    if (!Array.isArray(matrix) || matrix.length !== rows) {
      throw Object.assign(new Error(`${label} must be a ${rows} by ${columns} matrix.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    for (const row of matrix) {
      if (!Array.isArray(row) || row.length !== columns) {
        throw Object.assign(new Error(`${label} must be a ${rows} by ${columns} matrix.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
      }
    }
  }

  function matrixFromScalar(value, rows, columns) {
    return Array.from({ length: rows }, () => Array.from({ length: columns }, () => value));
  }

  function alignmentFrom(value, axis) {
    const horizontal = {
      general: Excel.HorizontalAlignment.general,
      left: Excel.HorizontalAlignment.left,
      center: Excel.HorizontalAlignment.center,
      right: Excel.HorizontalAlignment.right,
      fill: Excel.HorizontalAlignment.fill,
      justify: Excel.HorizontalAlignment.justify,
      centeracrossselection: Excel.HorizontalAlignment.centerAcrossSelection,
      distributed: Excel.HorizontalAlignment.distributed
    };
    const vertical = {
      top: Excel.VerticalAlignment.top,
      center: Excel.VerticalAlignment.center,
      bottom: Excel.VerticalAlignment.bottom,
      justify: Excel.VerticalAlignment.justify,
      distributed: Excel.VerticalAlignment.distributed
    };
    const values = axis === 'horizontal' ? horizontal : vertical;
    const key = String(value || '').replace(/[_\s-]/g, '').toLowerCase();
    if (!values[key]) {
      throw Object.assign(new Error(`Unsupported ${axis} alignment ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return values[key];
  }

  function applyBorders(range, borders) {
    if (!Array.isArray(borders)) {
      throw Object.assign(new Error('borders must be an array.'), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    for (const border of borders) {
      const rangeBorder = range.format.borders.getItem(borderIndexFrom(border.side));
      if (border.color !== undefined) rangeBorder.color = String(border.color);
      if (border.style !== undefined) rangeBorder.style = borderLineStyleFrom(border.style);
      if (border.weight !== undefined) rangeBorder.weight = borderWeightFrom(border.weight);
    }
  }

  function borderIndexFrom(value) {
    const indexes = {
      top: Excel.BorderIndex.edgeTop,
      edgeTop: Excel.BorderIndex.edgeTop,
      bottom: Excel.BorderIndex.edgeBottom,
      edgeBottom: Excel.BorderIndex.edgeBottom,
      left: Excel.BorderIndex.edgeLeft,
      edgeLeft: Excel.BorderIndex.edgeLeft,
      right: Excel.BorderIndex.edgeRight,
      edgeRight: Excel.BorderIndex.edgeRight,
      insideVertical: Excel.BorderIndex.insideVertical,
      insideHorizontal: Excel.BorderIndex.insideHorizontal,
      diagonalDown: Excel.BorderIndex.diagonalDown,
      diagonalUp: Excel.BorderIndex.diagonalUp
    };
    return enumValueFrom(indexes, value, 'border side');
  }

  function borderLineStyleFrom(value) {
    const styles = {
      none: Excel.BorderLineStyle.none,
      continuous: Excel.BorderLineStyle.continuous,
      dash: Excel.BorderLineStyle.dash,
      dashDot: Excel.BorderLineStyle.dashDot,
      dashDotDot: Excel.BorderLineStyle.dashDotDot,
      dot: Excel.BorderLineStyle.dot,
      double: Excel.BorderLineStyle.double,
      slantDashDot: Excel.BorderLineStyle.slantDashDot
    };
    return enumValueFrom(styles, value, 'border style');
  }

  function borderWeightFrom(value) {
    const weights = {
      hairline: Excel.BorderWeight.hairline,
      thin: Excel.BorderWeight.thin,
      medium: Excel.BorderWeight.medium,
      thick: Excel.BorderWeight.thick
    };
    return enumValueFrom(weights, value, 'border weight');
  }

  function sortOnFrom(value) {
    const values = {
      value: Excel.SortOn.value,
      cellColor: Excel.SortOn.cellColor,
      fontColor: Excel.SortOn.fontColor,
      icon: Excel.SortOn.icon
    };
    return enumValueFrom(values, value, 'sort type');
  }

  function sortDataOptionFrom(value) {
    const values = {
      normal: Excel.SortDataOption.normal,
      textAsNumber: Excel.SortDataOption.textAsNumber
    };
    return enumValueFrom(values, value, 'sort data option');
  }

  function sortOrientationFrom(value) {
    const values = {
      rows: Excel.SortOrientation.rows,
      columns: Excel.SortOrientation.columns
    };
    return enumValueFrom(values, value || 'rows', 'sort orientation');
  }

  function sortMethodFrom(value) {
    if (value === undefined || value === null || value === '') return undefined;
    const values = {
      pinYin: Excel.SortMethod.pinYin,
      strokeCount: Excel.SortMethod.strokeCount
    };
    return enumValueFrom(values, value, 'sort method');
  }

  function filterOnFrom(value) {
    const values = {
      bottomItems: Excel.FilterOn.bottomItems,
      bottomPercent: Excel.FilterOn.bottomPercent,
      cellColor: Excel.FilterOn.cellColor,
      dynamic: Excel.FilterOn.dynamic,
      fontColor: Excel.FilterOn.fontColor,
      values: Excel.FilterOn.values,
      topItems: Excel.FilterOn.topItems,
      topPercent: Excel.FilterOn.topPercent,
      icon: Excel.FilterOn.icon,
      custom: Excel.FilterOn.custom
    };
    return enumValueFrom(values, value, 'filter type');
  }

  function filterOperatorFrom(value) {
    const values = {
      and: Excel.FilterOperator.and,
      or: Excel.FilterOperator.or
    };
    return enumValueFrom(values, value, 'filter operator');
  }

  function dynamicFilterCriteriaFrom(value) {
    const key = String(value || '').replace(/[_\s-]/g, '').toLowerCase();
    const match = Object.entries(Excel.DynamicFilterCriteria || {}).find(([name]) => name.toLowerCase() === key);
    if (!match) {
      throw Object.assign(new Error(`Unsupported dynamic filter criteria ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return match[1];
  }

  function enumValueFrom(values, value, label) {
    const key = String(value || '').replace(/[_\s-]/g, '').toLowerCase();
    const match = Object.entries(values).find(([name]) => name.toLowerCase() === key);
    if (!match) {
      throw Object.assign(new Error(`Unsupported ${label} ${value}.`), { officeMcpCode: 'INVALID_ARGUMENT', partialEffect: 'none' });
    }
    return match[1];
  }

  function chartTypeFrom(value) {
    const chartTypes = {
      area: Excel.ChartType.area,
      barClustered: Excel.ChartType.barClustered,
      columnClustered: Excel.ChartType.columnClustered,
      doughnut: Excel.ChartType.doughnut,
      line: Excel.ChartType.line,
      pie: Excel.ChartType.pie,
      scatter: Excel.ChartType.xyscatter
    };
    return chartTypes[String(value)] || Excel.ChartType.columnClustered;
  }

  function targetRange(context, args) {
    const address = requiredString(args, 'address', 'Range address is required.');
    const worksheet = args.sheet
      ? context.workbook.worksheets.getItem(String(args.sheet))
      : context.workbook.worksheets.getActiveWorksheet();
    return worksheet.getRange(address);
  }

  async function getWorkbookInfo() {
    const url = Office.context.document?.url || '';
    let title = fileName(url) || 'Excel Workbook';
    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        await context.sync();
        if (!title && worksheets.items[0]?.name) title = worksheets.items[0].name;
      });
    } catch (error) {
      logger.warn('excel.probe.failed', error);
    }
    return {
      title,
      url: url || null,
      filename: fileName(url) || null,
      is_dirty: null,
      is_read_only: false,
      is_protected: false,
      protection: { kind: null, rights: null, rights_source: 'unavailable' }
    };
  }

  function probeRequirementSets() {
    const requirements = Office.context.requirements;
    return {
      excel_api_1_1: requirements.isSetSupported('ExcelApi', '1.1'),
      excel_api_1_2: requirements.isSetSupported('ExcelApi', '1.2'),
      excel_api_1_3: requirements.isSetSupported('ExcelApi', '1.3'),
      excel_api_1_4: requirements.isSetSupported('ExcelApi', '1.4'),
      excel_api_1_7: requirements.isSetSupported('ExcelApi', '1.7'),
      excel_api_1_8: requirements.isSetSupported('ExcelApi', '1.8'),
      excel_api_1_9: requirements.isSetSupported('ExcelApi', '1.9'),
      excel_api_1_12: requirements.isSetSupported('ExcelApi', '1.12'),
      excel_api_1_13: requirements.isSetSupported('ExcelApi', '1.13')
    };
  }

  function scheduleReconnect() {
    setConnectionState('reconnecting', 'Reconnecting…');
    const delay = reconnectDelay(reconnectAttempt);
    reconnectAttempt += 1;
    connectionDetailEl.textContent = `Disconnected. Reconnecting in ${formatDuration(delay)}.`;
    reconnectTimer = setTimeout(connect, delay);
  }

  function send(message) {
    return sendJsonRpc(socket, message);
  }

  function reply(id, result) {
    return replyJsonRpc(socket, id, result);
  }

  function renderStaticState() {
    renderStaticMetadata({ session: sessionEl, daemon: daemonEl, serverVersion: serverVersionEl, protocolVersion: protocolVersionEl, hostPlatform: hostPlatformEl }, { sessionId, endpoint: configuredEndpoint(), serverInfo, protocolVersion: PROTOCOL_VERSION, defaultHost: 'Excel' });
    renderToolModeControl();
    renderToolSummary();
    renderCurrentTask();
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
      const allowedInGroup = tools.filter((tool) => isToolSupported(tool) && isToolAllowedByMode(tool));
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
    const supported = isToolSupported(tool);
    const allowed = modeAllowed && supported;
    const checked = isToolEnabled(tool) && allowed;
    return [
      `<label class="tool-permission-row${metadata.sideEffect === 'mutating' || metadata.sideEffect === 'destructive' ? ' is-mutating' : ''}${allowed ? '' : ' is-disabled'}" for="${id}">`,
      '<span class="tool-permission-main">',
      '<span class="tool-permission-title">',
      `<span class="tool-permission-name">${escapeHtml(tool)}</span>`,
      `<span class="side-effect-pill ${metadata.sideEffect === 'mutating' || metadata.sideEffect === 'destructive' ? 'mutating' : 'read'}">${escapeHtml(metadata.sideEffect)}</span>`,
      '</span>',
      `<span class="tool-permission-meta">${escapeHtml(metadata.description)}</span>`,
      '</span>',
      `<input id="${id}" class="tool-toggle" type="checkbox" role="switch" data-tool="${escapeHtml(tool)}" aria-label="Toggle ${escapeHtml(tool)}" ${checked ? 'checked' : ''} ${allowed ? '' : 'disabled aria-disabled="true"'} />`,
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
      if (AVAILABLE_TOOLS.includes(tool) && isToolSupported(tool) && isToolAllowedByMode(tool)) toolPermissions[tool] = enabled;
    }
    saveToolPermissions();
    renderToolSummary();
    sendSessionToolUpdate();
  }

  function effectiveTools() {
    return AVAILABLE_TOOLS.filter((tool) => isToolSupported(tool) && isToolEnabled(tool) && isToolAllowedByMode(tool));
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

  function isToolSupported(tool) {
    if (tool === 'excel.find_replace_cells') return supportsRequirementSet('ExcelApi', '1.9');
    if (tool === 'excel.sort_range') return supportsRequirementSet('ExcelApi', '1.2');
    if (tool === 'excel.apply_filter') return supportsRequirementSet('ExcelApi', '1.9');
    return true;
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
    const workbook = documentInfo || {};
    documentTitleEl.textContent = workbook.title || workbook.filename || 'Unknown Excel Workbook';
    protectionEl.textContent = protectionLabel(workbook);
    documentStateEl.textContent = documentStateLabel(workbook);
    hostPlatformEl.textContent = `Excel / ${window.Office?.context?.platform || 'Unknown'}`;
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
    currentTaskEl.innerHTML = taskMarkup({
      tool: currentTask.tool,
      requestId: currentTask.requestId,
      status: 'running',
      elapsedMs: Math.max(0, Date.now() - currentTask.startedAt),
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
      `<div class="task-meta">${escapeHtml(formatDuration(task.elapsedMs || 0))}</div>`,
      metadata
    ].join('');
  }

  function setStatus(text) {
    connectionDetailEl.textContent = text;
    setConnectionState('failed', text);
  }

  function setConnectionState(state, label) {
    setSharedConnectionState({ badge: connectionBadgeEl, detail: connectionDetailEl, announcer: announcerEl }, state, label);
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

  function mapError(error, tool) {
    return {
      office_mcp_code: error.officeMcpCode || error.code || 'OFFICE_JS_ERROR',
      message: redactText(error.message || String(error)),
      tool,
      retriable: false,
      partial_effect: error.partialEffect || 'unknown'
    };
  }

  async function handleMetadataCopy(event) {
    await copyMetadataValue(event, { document, navigator, announcer: announcerEl, logger });
  }
})();
