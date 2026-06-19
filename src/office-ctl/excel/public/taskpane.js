(() => {
  const ADDIN_VERSION = '0.1.10';
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
  const {
    bindDetailsControl,
    middleTruncate,
    officeHostSummary,
    renderRuntimeVersions,
    setCopyableMetadata
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
    'excel.create_table',
    'excel.create_chart'
  ];
  const TOOL_GROUPS = [
    { label: 'Workbook', tools: ['excel.get_workbook_info'] },
    { label: 'Worksheet', tools: ['excel.list_sheets', 'excel.add_sheet', 'excel.update_sheet', 'excel.delete_sheet'] },
    { label: 'Range', tools: ['excel.get_used_range', 'excel.read_range', 'excel.write_range', 'excel.clear_range', 'excel.find_replace_cells'] },
    { label: 'Formula', tools: ['excel.set_formula'] },
    { label: 'Format', tools: ['excel.format_range'] },
    { label: 'Table', tools: ['excel.create_table'] },
    { label: 'Chart', tools: ['excel.create_chart'] }
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
    ['excel.create_table', { category: 'Table', sideEffect: 'mutating', description: 'Create a table from a range.' }],
    ['excel.create_chart', { category: 'Chart', sideEffect: 'mutating', description: 'Create a chart from a range.' }]
  ]);
  const { instanceId, sessionId } = runtimeIds();
  const TOOL_PERMISSION_STORAGE_KEY = `office-mcp.excel.tool-permissions.${sessionId}`;
  const logger = new AddinLogger({ redactText });
  const taskStore = new TaskHistoryStore({ redactText });
  let socket;
  let reconnectTimer;
  let reconnectAttempt = 0;
  let endpointDirty = false;
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
  const settingsFormEl = document.getElementById('settingsForm');
  const endpointInputEl = document.getElementById('endpointInput');
  const endpointErrorEl = document.getElementById('endpointError');
  const saveEndpointEl = document.getElementById('saveEndpoint');
  const announcerEl = document.getElementById('announcer');

  settingsFormEl.addEventListener('submit', saveEndpointOverride);
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
    setConnectionState('reconnecting', 'Reconnecting…');
    reconnectAttempt = 0;
    clearTimeout(reconnectTimer);
    reconnectTimer = setTimeout(connect, 250);
    return true;
  }

  function register() {
    reconnectAttempt = 0;
    setConnectionState('connecting', 'Registering…');
    const requestId = crypto.randomUUID();
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
        case 'excel.create_table':
          data = await createTable(args);
          break;
        case 'excel.create_chart':
          data = await createChart(args);
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
      worksheets.load('items/id,items/name,items/visibility');
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
    const formula = requiredString(args, 'formula', 'excel.set_formula requires formula.');
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      range.load('rowCount,columnCount');
      await context.sync();
      range.formulas = matrixFromScalar(formula, range.rowCount, range.columnCount);
      await context.sync();
      return { address: args.address, formula, wrote_formula: true };
    });
  }

  async function formatRange(args) {
    return Excel.run(async (context) => {
      const range = targetRange(context, args);
      if (args.number_format) {
        range.load('rowCount,columnCount');
        await context.sync();
      }
      if (args.bold !== undefined) range.format.font.bold = Boolean(args.bold);
      if (args.italic !== undefined) range.format.font.italic = Boolean(args.italic);
      if (args.font_color) range.format.font.color = String(args.font_color);
      if (args.fill_color) range.format.fill.color = String(args.fill_color);
      if (args.number_format) range.numberFormat = matrixFromScalar(String(args.number_format), range.rowCount, range.columnCount);
      await context.sync();
      return { address: args.address, formatted: true };
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

  function targetWorksheet(context, args) {
    return args.sheet
      ? context.workbook.worksheets.getItem(String(args.sheet))
      : context.workbook.worksheets.getActiveWorksheet();
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

  function supportsRequirementSet(name, version) {
    return Office.context?.requirements?.isSetSupported?.(name, version) === true;
  }
  function requiredString(args, key, message) {
    const value = String(args[key] || '').trim();
    if (!value) throw Object.assign(new Error(message), { officeMcpCode: 'INVALID_ARGUMENT' });
    return value;
  }

  function matrixFromScalar(value, rows, columns) {
    return Array.from({ length: rows }, () => Array.from({ length: columns }, () => value));
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
      excel_api_1_4: requirements.isSetSupported('ExcelApi', '1.4')
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
    setCopyableMetadata(sessionEl, sessionId);
    setCopyableMetadata(daemonEl, configuredEndpoint());
    renderRuntimeVersions(serverVersionEl, protocolVersionEl, serverInfo, PROTOCOL_VERSION);
    hostPlatformEl.textContent = officeHostSummary('Excel');
    renderToolSummary();
    renderCurrentTask();
    renderHistory();
  }

  function renderToolSummary() {
    const effective = effectiveTools();
    const openGroups = new Set([...toolListEl.querySelectorAll('[data-tool-group]')]
      .filter((input) => input.closest('details')?.open)
      .map((input) => input.dataset.toolGroup));
    toolCountEl.textContent = `Enabled ${effective.length} of ${AVAILABLE_TOOLS.length}`;
    toolListEl.textContent = '';
    for (const group of TOOL_GROUPS) {
      const tools = group.tools.filter((tool) => AVAILABLE_TOOLS.includes(tool));
      if (tools.length === 0) continue;
      const enabledInGroup = tools.filter((tool) => effective.includes(tool));
      const groupEl = document.createElement('details');
      groupEl.className = 'tool-group';
      groupEl.open = openGroups.has(group.label);
      groupEl.innerHTML = [
        '<summary class="tool-group-title">',
        `<span>${escapeHtml(group.label)}</span>`,
        `<span class="tool-group-count">Enabled ${enabledInGroup.length} of ${tools.length}</span>`,
        `<input class="group-toggle" type="checkbox" role="switch" data-tool-group="${escapeHtml(group.label)}" aria-label="Toggle ${escapeHtml(group.label)} tools" ${enabledInGroup.length === tools.length ? 'checked' : ''} />`,
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
    const checked = isToolEnabled(tool);
    return [
      `<label class="tool-permission-row${metadata.sideEffect === 'mutating' ? ' is-mutating' : ''}" for="${id}">`,
      '<span class="tool-permission-main">',
      '<span class="tool-permission-title">',
      `<span class="tool-permission-name">${escapeHtml(tool)}</span>`,
      `<span class="side-effect-pill ${metadata.sideEffect === 'mutating' ? 'mutating' : 'read'}">${escapeHtml(metadata.sideEffect)}</span>`,
      '</span>',
      `<span class="tool-permission-meta">${escapeHtml(metadata.description)}</span>`,
      '</span>',
      `<input id="${id}" class="tool-toggle" type="checkbox" role="switch" data-tool="${escapeHtml(tool)}" aria-label="Toggle ${escapeHtml(tool)}" ${checked ? 'checked' : ''} />`,
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
      if (AVAILABLE_TOOLS.includes(tool)) toolPermissions[tool] = enabled;
    }
    saveToolPermissions();
    renderToolSummary();
    sendSessionToolUpdate();
  }

  function effectiveTools() {
    return AVAILABLE_TOOLS.filter((tool) => isToolSupported(tool) && isToolEnabled(tool));
  }

  function isToolSupported(tool) {
    if (tool === 'excel.find_replace_cells') return supportsRequirementSet('ExcelApi', '1.9');
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

  function saveToolPermissions() {
    window.localStorage?.setItem(TOOL_PERMISSION_STORAGE_KEY, JSON.stringify(toolPermissions));
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
    const error = task.error ? `<div class="task-meta">${escapeHtml(task.error.office_mcp_code)}: ${escapeHtml(task.error.message)} · Retriable: ${valueLabel(task.error.retriable)} · Partial effect: ${escapeHtml(task.error.partial_effect || 'unknown')}</div>` : '';
    const intent = task.userIntent ? `<div class="task-meta">${escapeHtml(redactText(task.userIntent))}</div>` : '';
    const deadline = task.deadlineAt ? `<div class="task-meta">Deadline ${escapeHtml(formatTime(task.deadlineAt))}</div>` : '';
    const cancel = task.cancelRequested ? '<div class="task-meta">Cancel requested</div>' : '';
    const commandId = task.requestId ? `<div class="task-meta task-command-id">Command <button type="button" class="inline-copy" data-copy-value="${escapeHtml(task.requestId)}" aria-label="Copy command ID" title="${escapeHtml(task.requestId)}"><code>${escapeHtml(middleTruncate(task.requestId))}</code></button></div>` : '';
    return [
      '<div class="task-title">',
      `<span>${escapeHtml(task.tool)}</span>`,
      `<span class="status-badge ${tone}">${escapeHtml(titleCase(task.status))}</span>`,
      '</div>',
      commandId,
      `<div class="task-meta">${escapeHtml(formatDuration(task.elapsedMs || 0))}</div>`,
      deadline,
      cancel,
      intent,
      error
    ].join('');
  }

  function setStatus(text) {
    connectionBadgeEl.textContent = text;
    connectionDetailEl.textContent = text;
  }

  function setConnectionState(state, label) {
    connectionBadgeEl.textContent = label;
    connectionBadgeEl.className = `status-badge ${statusClass(state)}`;
  }

  function statusClass(state) {
    if (state === 'connected') return 'status-success';
    if (state === 'failed' || state === 'unsupported') return 'status-danger';
    if (state === 'reconnecting') return 'status-warning';
    return 'status-neutral';
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

  function valueLabel(value) {
    if (value === true) return 'yes';
    if (value === false) return 'no';
    return 'unknown';
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
    const button = event.target.closest('[data-copy-target], [data-copy-value]');
    if (!button) return;
    const target = button.dataset.copyTarget ? document.getElementById(button.dataset.copyTarget) : null;
    const value = button.dataset.copyValue || target?.textContent?.trim();
    if (!value || value === '-') return;
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(value);
      else fallbackCopy(value);
      announcerEl.textContent = `Copied ${button.getAttribute('aria-label') || 'value'}`;
    } catch (error) {
      logger.warn('metadata_copy.failed', error);
      announcerEl.textContent = 'Copy failed';
    }
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
