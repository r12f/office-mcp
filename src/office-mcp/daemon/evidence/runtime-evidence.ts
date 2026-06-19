import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

type GateStatus = 'passed' | 'failed' | 'skipped' | 'blocked_by_runtime';

type EvidenceGate = {
  name: string;
  status: GateStatus;
  started_at: string;
  finished_at: string;
  details: Record<string, unknown>;
};

type EvidenceReport = {
  schema_version: 1;
  generated_at: string;
  endpoint: string;
  session_id?: string;
  gates: EvidenceGate[];
};

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const tsxCli = resolve(evidenceRoot, 'node_modules/tsx/dist/cli.mjs');
const endpoint = readOption('--endpoint') ?? process.env.OFFICE_MCP_MCP_ENDPOINT ?? 'http://127.0.0.1:8800/mcp';
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/runtime-evidence.json'));
const requestedSessionId = readOption('--session-id');
const includeMutation = hasFlag('--include-mutation');
const includeFullWordSmoke = hasFlag('--include-full-word-smoke');
const includeExcelSmoke = hasFlag('--include-excel-smoke');
const includePowerPointSmoke = hasFlag('--include-powerpoint-smoke');
const includeComTrackedChanges = hasFlag('--include-com-tracked-changes');
const includeTrackedChanges = hasFlag('--include-tracked-changes');
const irmMode = readOption('--irm-mode') ?? 'none';
const irmDocumentPath = readOption('--irm-document-path') ?? process.env.OFFICE_MCP_IRM_DOCUMENT_PATH;
const waitForSessionMs = intOption('--wait-for-session-ms', 0);
const agentClientEvidencePath = readOption('--agent-client-evidence-path') ?? process.env.OFFICE_MCP_AGENT_CLIENT_EVIDENCE_PATH;
const wantsWordRuntime = Boolean(
  requestedSessionId ||
  irmDocumentPath ||
  includeMutation ||
  includeFullWordSmoke ||
  includeComTrackedChanges ||
  includeTrackedChanges ||
  irmMode !== 'none' ||
  agentClientEvidencePath
);
const wantsWordBaseline = !(includeExcelSmoke || includePowerPointSmoke) || wantsWordRuntime;

const report: EvidenceReport = {
  schema_version: 1,
  generated_at: new Date().toISOString(),
  endpoint,
  session_id: requestedSessionId,
  gates: []
};

const client = new Client({ name: 'office-mcp-runtime-evidence', version: '0.1.0' });
const transport = new StreamableHTTPClientTransport(new URL(endpoint));

try {
  await client.connect(transport);
  const sessionsResult = await runSessionDiscoveryGate();
  let sessions = Array.isArray(sessionsResult.sessions) ? sessionsResult.sessions as Array<Record<string, unknown>> : [];
  if (!requestedSessionId && irmDocumentPath && waitForSessionMs > 0 && !selectWordSessionId(sessions, irmDocumentPath)) {
    const waitResult = await runWaitForSessionGate(irmDocumentPath, waitForSessionMs);
    sessions = Array.isArray(waitResult.sessions) ? waitResult.sessions as Array<Record<string, unknown>> : sessions;
  }
  if (includeExcelSmoke && waitForSessionMs > 0 && !selectHostSessionId(sessions, 'excel')) {
    const waitResult = await runWaitForHostSessionGate('excel', waitForSessionMs);
    sessions = Array.isArray(waitResult.sessions) ? waitResult.sessions as Array<Record<string, unknown>> : sessions;
  }
  if (includePowerPointSmoke && waitForSessionMs > 0 && !selectHostSessionId(sessions, 'powerpoint')) {
    const waitResult = await runWaitForHostSessionGate('powerpoint', waitForSessionMs);
    sessions = Array.isArray(waitResult.sessions) ? waitResult.sessions as Array<Record<string, unknown>> : sessions;
  }
  const sessionId = wantsWordBaseline ? requestedSessionId ?? selectWordSessionId(sessions, irmDocumentPath) : '';
  const excelSessionId = selectHostSessionId(sessions, 'excel');
  const powerPointSessionId = selectHostSessionId(sessions, 'powerpoint');
  if (sessionId) report.session_id = sessionId;

  if (!sessionId && wantsWordBaseline) {
    const reason = irmDocumentPath
      ? `No connected Word add-in session matched the requested IRM document path: ${irmDocumentPath}. Open that document and open Office MCP Control, then rerun this script.`
      : 'No connected Word add-in session. Open Word, open Office MCP Control, then rerun this script.';
    addGate('word.runtime_smoke', 'blocked_by_runtime', { reason });
    addGate('agent_client_stdio_bridge', 'blocked_by_runtime', {
      reason: 'No connected Word add-in session to prove an agent client can call MCP.'
    });
    if (irmDocumentPath) await runIrmDocumentPreflightGate(irmDocumentPath);
    await runAgentClientPromptGate(agentClientEvidencePath);
    addGate('irm_rights_matrix', 'blocked_by_runtime', {
      reason: irmDocumentPath ? reason : 'No connected IRM-protected Word session was provided.'
    });
  } else if (sessionId) {
    await runWordReadGate(sessionId);
    if (includeMutation) await runWordMutationGate(sessionId);
    if (includeFullWordSmoke) await runFullWordSmokeGate(sessionId);
    await runAgentClientBridgeGate(sessionId);
    await runClaudeDesktopInstallationGate();
    await runAgentClientPromptGate(agentClientEvidencePath);
    if (includeTrackedChanges) await runTrackedChangeGate(sessionId);
    if (includeComTrackedChanges) await runComTrackedChangeGate(sessionId);
    await runIrmGate(sessionId, irmMode, undefined, irmDocumentPath);
    if (irmDocumentPath) await runIrmDocumentPreflightGate(irmDocumentPath);
  }
  if (includeExcelSmoke) {
    if (excelSessionId) await runExcelSmokeGate(excelSessionId);
    else addGate('excel.runtime_smoke', 'blocked_by_runtime', {
      reason: 'No connected Excel add-in session. Open Excel, open Office MCP Control, then rerun this script.'
    });
  }
  if (includePowerPointSmoke) {
    if (powerPointSessionId) await runPowerPointSmokeGate(powerPointSessionId);
    else addGate('powerpoint.runtime_smoke', 'blocked_by_runtime', {
      reason: 'No connected PowerPoint add-in session. Open PowerPoint, open Office MCP Control, then rerun this script.'
    });
  }
} catch (error) {
  addGate('runtime_evidence_harness', 'failed', { error: errorMessage(error) });
} finally {
  await client.close().catch(() => undefined);
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
}

async function runFullWordSmokeGate(sessionId: string): Promise<void> {
  const modes = ['word-core', 'word-formatting', 'word-review', 'word-resources', 'word-spec-args'];
  for (const mode of modes) {
    await runGate(`word.full_smoke.${mode}`, async () => {
      const output = runSmokeMode(mode, sessionId);
      const parsed = JSON.parse(output) as Record<string, unknown>;
      return { mode, output_bytes: Buffer.byteLength(output, 'utf8'), summary_keys: Object.keys(parsed) };
    });
  }
}

async function runSessionDiscoveryGate(): Promise<Record<string, unknown>> {
  return await runGate('word.session_discovery', async () => listSessionsDetails());
}

async function runWaitForSessionGate(documentPath: string, timeoutMs: number): Promise<Record<string, unknown>> {
  return await runGate('word.wait_for_requested_session', async () => {
    const started = Date.now();
    let latest: Record<string, unknown> = { sessions: [], session_count: 0 };
    while (Date.now() - started <= timeoutMs) {
      latest = await listSessionsDetails();
      const sessions = Array.isArray(latest.sessions) ? latest.sessions as Array<Record<string, unknown>> : [];
      const sessionId = selectWordSessionId(sessions, documentPath);
      if (sessionId) return { ...latest, matched_session_id: sessionId, waited_ms: Date.now() - started };
      await sleep(1000);
    }
    throw new Error(`Timed out waiting ${timeoutMs} ms for a session matching ${documentPath}`);
  });
}

async function listSessionsDetails(): Promise<Record<string, unknown>> {
  const result = await client.callTool({ name: 'office.list_sessions', arguments: {} });
  const data = toolData(result);
  const sessions = Array.isArray(data.sessions) ? data.sessions : [];
  return { sessions, session_count: sessions.length };
}

async function runComTrackedChangeGate(sessionId: string): Promise<void> {
  for (const action of ['accept', 'reject']) {
    await runGate(`word.tracked_change_com.${action}`, async () => {
      const output = runSmokeMode('word-track-change-com', sessionId, action);
      const parsed = JSON.parse(output) as { mutation?: { skipped?: boolean; action?: string } };
      if (parsed.mutation?.skipped) throw new Error(`Tracked-change COM ${action} smoke skipped mutation.`);
      return { action, output_bytes: Buffer.byteLength(output, 'utf8'), mutation_action: parsed.mutation?.action };
    });
  }
}

async function runWordReadGate(sessionId: string): Promise<void> {
  await runGate('word.runtime_read_smoke', async () => {
    const info = toolData(await client.callTool({ name: 'office.get_session_info', arguments: { session_id: sessionId } }));
    const paragraph = toolData(await client.callTool({ name: 'word.get_paragraph', arguments: { session_id: sessionId, index: 0 } }));
    const document = resourceData(await client.readResource({ uri: `office://word/${sessionId}/document?offset=0&limit=5` }));
    return {
      document_title: (info.document as { title?: string } | undefined)?.title,
      available_tool_count: Array.isArray(info.available_tools) ? info.available_tools.length : undefined,
      paragraph_0_text_length: String(paragraph.text ?? '').length,
      document_text_length: String(document.text ?? '').length
    };
  });
}

async function runWordMutationGate(sessionId: string): Promise<void> {
  await runGate('word.runtime_mutation_smoke', async () => {
    const marker = `office-mcp runtime evidence ${Date.now()}`;
    const insert = await client.callTool({
      name: 'word.insert_paragraph',
      arguments: { session_id: sessionId, text: marker, anchor: { kind: 'end_of_document' }, style: 'Normal' }
    });
    const find = toolData(await client.callTool({ name: 'word.find_text', arguments: { session_id: sessionId, query: marker, limit: 5 } }));
    if (Number(find.count ?? 0) < 1) throw new Error('Inserted mutation marker was not found after insertion.');
    return { marker, insert: toolData(insert), find_count: find.count };
  });
}

async function runWaitForHostSessionGate(hostApp: string, timeoutMs: number): Promise<Record<string, unknown>> {
  return await runGate(`${hostApp}.wait_for_session`, async () => {
    const started = Date.now();
    let latest: Record<string, unknown> = { sessions: [], session_count: 0 };
    while (Date.now() - started <= timeoutMs) {
      latest = await listSessionsDetails();
      const sessions = Array.isArray(latest.sessions) ? latest.sessions as Array<Record<string, unknown>> : [];
      const sessionId = selectHostSessionId(sessions, hostApp);
      if (sessionId) return { ...latest, matched_session_id: sessionId, waited_ms: Date.now() - started };
      await sleep(1000);
    }
    throw new Error(`Timed out waiting ${timeoutMs} ms for an active ${hostApp} session.`);
  });
}

async function runExcelSmokeGate(sessionId: string): Promise<void> {
  await runGate('excel.runtime_smoke', async () => {
    const marker = `OfficeMCP${Date.now()}`;
    const sheetName = `OfficeMcpSmoke${Date.now()}`;
    const renamedSheetName = `${sheetName}Renamed`;
    const cleanupSheetName = `${sheetName}Cleanup`;
    const tableName = `OfficeMcpTable${Date.now()}`;
    const chartTitle = 'Office MCP Smoke Updated';
    const pivotName = `OfficeMcpPivot${Date.now()}`;
    const info = await callToolData('office.get_session_info', { session_id: sessionId });
    const workbookInfo = await callToolData('excel.get_workbook_info', { session_id: sessionId });
    const sheetListBefore = await callToolData('excel.list_sheets', { session_id: sessionId });
    const sheet = await callToolData('excel.add_sheet', {
      session_id: sessionId,
      name: sheetName,
      activate: true
    });
    const updatedSheet = await callToolData('excel.update_sheet', {
      session_id: sessionId,
      sheet: sheetName,
      name: renamedSheetName,
      tab_color: '#4472C4',
      activate: true
    });
    const cleanupSheet = await callToolData('excel.add_sheet', {
      session_id: sessionId,
      name: cleanupSheetName,
      activate: false
    });
    const deletedSheet = await callToolData('excel.delete_sheet', {
      session_id: sessionId,
      sheet: cleanupSheetName
    });
    const readBefore = await callToolData('excel.read_range', { session_id: sessionId, sheet: renamedSheetName, address: 'A1:B2' });
    const write = await callToolData('excel.write_range', {
      session_id: sessionId,
      sheet: renamedSheetName,
      address: 'A1:C5',
      values: [
        ['Label', 'Region', 'Value'],
        [marker, 'West', 42],
        [`${marker}B`, 'East', 24],
        [`${marker}C`, 'West', 12],
        ['Total', 'All', 0]
      ]
    });
    const formula = await callToolData('excel.set_formula', { session_id: sessionId, sheet: renamedSheetName, address: 'C5', formula: '=SUM(C2:C4)' });
    const format = await callToolData('excel.format_range', {
      session_id: sessionId,
      sheet: renamedSheetName,
      address: 'A1:C5',
      bold: true,
      fill_color: '#DDEEFF',
      number_format: 'General',
      horizontal_alignment: 'center',
      wrap_text: true
    });
    const usedRange = await callToolData('excel.get_used_range', {
      session_id: sessionId,
      sheet: renamedSheetName
    });
    const findReplace = await callToolData('excel.find_replace_cells', {
      session_id: sessionId,
      sheet: renamedSheetName,
      address: 'A1:C5',
      query: marker,
      replacement: `${marker}Updated`,
      complete_match: false,
      match_case: true
    });
    const clear = await callToolData('excel.clear_range', {
      session_id: sessionId,
      sheet: renamedSheetName,
      address: 'D1:D1',
      apply_to: 'contents'
    });
    const table = await callToolData('excel.create_table', {
      session_id: sessionId,
      sheet: renamedSheetName,
      address: 'A1:C5',
      has_headers: true,
      name: tableName
    });
    const tableUpdate = await callToolData('excel.update_table', {
      session_id: sessionId,
      table: tableName,
      action: 'options',
      show_filter_button: true,
      show_banded_rows: true,
      style: 'TableStyleMedium2'
    });
    const sort = await callToolData('excel.sort_range', {
      session_id: sessionId,
      target_type: 'table',
      table: tableName,
      fields: [{ key: 2, ascending: false }]
    });
    const filter = await callToolData('excel.apply_filter', {
      session_id: sessionId,
      target_type: 'table',
      table: tableName,
      column: 'Region',
      criteria: { filter_on: 'values', values: ['West', 'All'] }
    });
    const chart = await callToolData('excel.create_chart', {
      session_id: sessionId,
      sheet: renamedSheetName,
      address: 'A1:C5',
      type: 'columnClustered',
      title: 'Office MCP Smoke'
    });
    const chartName = String(chart.chart ?? '');
    const chartUpdate = await callToolData('excel.update_chart', {
      session_id: sessionId,
      sheet: renamedSheetName,
      chart: chartName,
      action: 'title',
      title: chartTitle,
      visible: true
    });
    const pivotTable = await callToolData('excel.create_pivot_table', {
      session_id: sessionId,
      name: pivotName,
      table: tableName,
      destination: `${renamedSheetName}!E3`
    });
    const pivotUpdate = await callToolData('excel.update_pivot_table', {
      session_id: sessionId,
      pivot_table: pivotName,
      action: 'refresh'
    });
    const readAfter = await callToolData('excel.read_range', { session_id: sessionId, sheet: renamedSheetName, address: 'A1:C5' });
    const values = readAfter.values as unknown[][] | undefined;
    if (!Array.isArray(values) || !String(values.flat().join(' ')).includes(`${marker}Updated`)) {
      throw new Error('Excel smoke marker was not found after write_range.');
    }
    return {
      session_id: sessionId,
      sheet_name: renamedSheetName,
      document_title: (info.document as { title?: string } | undefined)?.title,
      available_tool_count: Array.isArray(info.available_tools) ? info.available_tools.length : undefined,
      workbook_info: {
        sheet_count: workbookInfo.sheet_count,
        table_count: workbookInfo.table_count,
        active_sheet: workbookInfo.active_sheet
      },
      sheet_list_count: Array.isArray(sheetListBefore.sheets) ? sheetListBefore.sheets.length : undefined,
      read_before_address: readBefore.address,
      write,
      formula,
      format,
      used_range: usedRange,
      find_replace: findReplace,
      clear,
      table,
      table_update: tableUpdate,
      sort,
      filter,
      chart,
      chart_update: chartUpdate,
      pivot_table: pivotTable,
      pivot_update: pivotUpdate,
      sheet,
      updated_sheet: updatedSheet,
      cleanup_sheet: cleanupSheet,
      deleted_sheet: deletedSheet,
      marker_found: true
    };
  });
}


async function runPowerPointSmokeGate(sessionId: string): Promise<void> {
  await runGate('powerpoint.runtime_smoke', async () => {
    const marker = `Office MCP PowerPoint smoke ${Date.now()}`;
    const replacement = `${marker} updated`;
    const info = await callToolData('office.get_session_info', { session_id: sessionId });
    const availableTools = Array.isArray(info.available_tools) ? info.available_tools.map(String) : [];
    const presentationInfo = await callToolData('powerpoint.get_presentation_info', { session_id: sessionId, include_selection: true });
    const activeView = await callToolData('powerpoint.get_active_view', { session_id: sessionId });
    const slidesBefore = await callToolData('powerpoint.list_slides', { session_id: sessionId, include_tags: true });
    const layouts = await callToolData('powerpoint.list_layouts', { session_id: sessionId });
    const addSlide = await callToolData('powerpoint.add_slide', {
      session_id: sessionId,
      title: marker,
      content: 'PowerPoint runtime smoke content'
    });
    const addTextBox = await callToolData('powerpoint.add_text_box', {
      session_id: sessionId,
      slide_id: String(addSlide.slide_id ?? ''),
      text: `Shape ${marker}`,
      left: 72,
      top: 220,
      width: 360,
      height: 64
    });
    const shapeId = String((addTextBox.shape as Record<string, unknown> | undefined)?.shape_id ?? '');
    const shapes = await callToolData('powerpoint.list_shapes', { session_id: sessionId, slide_id: String(addSlide.slide_id ?? '') });
    const readText = await callToolData('powerpoint.read_text', { session_id: sessionId, slide_id: String(addSlide.slide_id ?? '') });
    const replaceText = await callToolData('powerpoint.replace_text', {
      session_id: sessionId,
      search: marker,
      replacement,
      match_case: true
    });
    const formatText = shapeId ? await callToolData('powerpoint.format_text', {
      session_id: sessionId,
      slide_id: String(addSlide.slide_id ?? ''),
      shape_id: shapeId,
      bold: true,
      font_size: 18
    }) : { skipped: true, reason: 'No text box shape id returned.' };
    const layout = await callToolData('powerpoint.apply_layout', {
      session_id: sessionId,
      slide_id: String(addSlide.slide_id ?? ''),
      layout: 'TitleOnly'
    });
    const table = await callToolResult('powerpoint.add_table', {
      session_id: sessionId,
      slide_id: String(addSlide.slide_id ?? ''),
      rows: 2,
      columns: 2,
      values: [['Office', 'MCP'], ['PowerPoint', 'Smoke']]
    });
    const tableData = toolData(table);
    const tableSupported = !isToolError(table);
    const tableHostRejection = isExplicitHostRejection(table, tableData);
    if (!tableSupported && !tableHostRejection) throw new Error(`PowerPoint table creation failed without explicit host-capability rejection: ${JSON.stringify(tableData)}`);
    const tableShapeId = String(tableData.shape_id ?? '');
    const readTable = tableSupported && tableShapeId ? await callToolData('powerpoint.read_table', {
      session_id: sessionId,
      slide_id: String(addSlide.slide_id ?? ''),
      shape_id: tableShapeId
    }) : { host_rejection: tableHostRejection };
    const file = await callToolResult('powerpoint.export_file', { session_id: sessionId, format: 'pdf', slice_size: 1048576 });
    const fileData = toolData(file);
    const exportSupported = !isToolError(file);
    const exportHostRejection = isExplicitHostRejection(file, fileData);
    if (!exportSupported && !exportHostRejection) throw new Error(`PowerPoint file export failed without explicit host-capability rejection: ${JSON.stringify(fileData)}`);
    const categoryProofs = {
      presentation: Boolean(presentationInfo && activeView && (exportSupported || exportHostRejection)),
      slides: typeof addSlide.slide_id === 'string' && Array.isArray(slidesBefore.slides),
      layout: typeof layout.slide_id === 'string' && Array.isArray(layouts.masters),
      shapes: shapeId.length > 0 && Array.isArray(shapes.shapes),
      text: Number(replaceText.replacements ?? 0) >= 1 && Array.isArray(readText.items),
      tables: tableSupported ? isRecord(readTable) : tableHostRejection
    };
    return {
      session_id: sessionId,
      document_title: (info.document as { title?: string } | undefined)?.title,
      available_tool_count: availableTools.length,
      available_tools: availableTools,
      presentation_info: presentationInfo,
      active_view: activeView,
      list_slides: slidesBefore,
      list_layouts: layouts,
      add_slide: addSlide,
      add_text_box: addTextBox,
      list_shapes: shapes,
      read_text: readText,
      replace_text: replaceText,
      format_text: formatText,
      layout,
      add_table: tableData,
      read_table: readTable,
      marker,
      replacement,
      mutation_proved: typeof addSlide.slide_id === 'string' && Number(replaceText.replacements ?? 0) >= 1,
      tool_category_proofs: categoryProofs,
      export_supported: exportSupported,
      export_host_rejection: exportHostRejection,
      export_mime_type: exportSupported ? fileData.mime_type : undefined,
      export_size: exportSupported ? fileData.size : undefined,
      table_supported: tableSupported,
      table_host_rejection: tableHostRejection
    };
  });
}


async function callToolResult(name: string, args: Record<string, unknown>): Promise<unknown> {
  return await client.callTool({ name, arguments: args });
}
async function callToolData(name: string, args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const result = await client.callTool({ name, arguments: args });
  if (isToolError(result)) throw new Error(`${name} returned an MCP tool error: ${JSON.stringify(toolData(result))}`);
  return toolData(result);
}

async function runAgentClientBridgeGate(sessionId: string): Promise<void> {
  await runGate('agent_client_stdio_bridge', async () => {
    const stdioClient = new Client({ name: 'office-mcp-agent-client-evidence', version: '0.1.0' });
    const stdioTransport = new StdioClientTransport({ command: 'cargo', args: ['run', '-q', '-p', 'office-mcp-daemon', '--', 'stdio'], cwd: repoRoot, stderr: 'pipe' });
    await stdioClient.connect(stdioTransport);
    try {
      const data = toolData(await stdioClient.callTool({ name: 'office.list_sessions', arguments: {} }));
      const sessions = Array.isArray(data.sessions) ? data.sessions : [];
      if (sessions.length < 1) throw new Error('stdio bridge connected but returned no sessions.');
      const paragraph = toolData(await stdioClient.callTool({ name: 'word.get_paragraph', arguments: { session_id: sessionId, index: 0 } }));
      if (typeof paragraph.text !== 'string') throw new Error('stdio bridge could not read paragraph 0.');
      return { session_count: sessions.length, paragraph_0_text: paragraph.text, paragraph_0_length: paragraph.text.length };
    } finally {
      await stdioClient.close();
    }
  });
}

async function runClaudeDesktopInstallationGate(): Promise<void> {
  await runGate('claude_desktop_installation', async () => {
    const appData = process.env.APPDATA ?? '';
    const localAppData = process.env.LOCALAPPDATA ?? '';
    const configPath = appData ? join(appData, 'Claude', 'claude_desktop_config.json') : '';
    const installCandidates = [
      localAppData ? join(localAppData, 'Programs', 'Claude') : '',
      localAppData ? join(localAppData, 'AnthropicClaude') : ''
    ].filter(Boolean);
    const processOutput = execFileSync('powershell.exe', [
      '-NoProfile',
      '-Command',
      "Get-Process | Where-Object { $_.ProcessName -match 'Claude|claude' } | Select-Object -ExpandProperty ProcessName"
    ], { encoding: 'utf8' }).trim();
    const runningProcesses = processOutput ? processOutput.split(/\r?\n/).filter(Boolean) : [];
    const configHasOfficeMcp = configPath ? claudeConfigHasOfficeMcp(configPath) : false;
    const appDetected = installCandidates.some((path) => existsSync(path)) || runningProcesses.length > 0;
    return {
      config_path: configPath,
      config_exists: configPath ? existsSync(configPath) : false,
      config_has_office_mcp: configHasOfficeMcp,
      install_candidates: installCandidates.map((path) => ({ path, exists: existsSync(path) })),
      running_processes: runningProcesses,
      app_detected: appDetected,
      ui_validation_ready: configHasOfficeMcp && appDetected
    };
  });
}

function claudeConfigHasOfficeMcp(configPath: string): boolean {
  if (!existsSync(configPath)) return false;
  try {
    const parsed = JSON.parse(readFileSync(configPath, 'utf8')) as { mcpServers?: Record<string, unknown> };
    return Boolean(parsed.mcpServers?.['office-mcp']);
  } catch {
    return false;
  }
}

async function runAgentClientPromptGate(evidencePath?: string): Promise<void> {
  if (!evidencePath) {
    addGate('agent_client_prompt', 'skipped', {
      reason: 'Pass --agent-client-evidence-path after recording a successful agent client prompt.'
    });
    return;
  }
  await runGate('agent_client_prompt', async () => {
    const absolutePath = resolve(evidencePath);
    if (!existsSync(absolutePath)) throw new Error('agent client evidence file does not exist: ' + absolutePath);
    const evidence = JSON.parse(readFileSync(absolutePath, 'utf8')) as Record<string, unknown>;
    if (evidence.schema_version !== 1) throw new Error('Unsupported agent client evidence schema_version.');
    if (evidence.kind !== 'agent_client_prompt') throw new Error('Unsupported agent client evidence kind.');
    if (evidence.passed !== true) throw new Error('agent client evidence did not pass.');
    if (typeof evidence.prompt !== 'string' || evidence.prompt.length === 0) throw new Error('agent client evidence is missing prompt.');
    if (typeof evidence.observed_answer !== 'string' || evidence.observed_answer.length === 0) throw new Error('agent client evidence is missing observed_answer.');
    return { evidence_path: absolutePath, ...evidence };
  });
}

async function runIrmDocumentPreflightGate(documentPath: string): Promise<Record<string, unknown>> {
  return await runGate('irm_document_preflight', async () => {
    if (!existsSync(documentPath)) throw new Error(`IRM document path does not exist: ${documentPath}`);
    const script = [
      `$docPath = ${psString(documentPath)}`,
      '$word = New-Object -ComObject Word.Application',
      '$word.Visible = $false',
      'try {',
      '  $doc = $word.Documents.Open($docPath, $false, $true)',
      '  try {',
      '    $permissionEnabled = $false',
      '    $permissionCount = $null',
      '    $enabledError = $null',
      '    $countError = $null',
      '    try { $permissionEnabled = [bool]$doc.Permission.Enabled } catch { $enabledError = $_.Exception.Message }',
      '    try { $permissionCount = [int]$doc.Permission.Count } catch { $countError = $_.Exception.Message }',
      '    [pscustomobject]@{',
      '      FullName = $doc.FullName',
      '      Name = $doc.Name',
      '      ReadOnly = [bool]$doc.ReadOnly',
      '      ProtectionType = [int]$doc.ProtectionType',
      '      PermissionEnabled = $permissionEnabled',
      '      PermissionCount = $permissionCount',
      '      PermissionEnabledError = $enabledError',
      '      PermissionCountError = $countError',
      '    } | ConvertTo-Json -Depth 4',
      '  } finally {',
      '    $doc.Close($false)',
      '  }',
      '} finally {',
      '  $word.Quit()',
      '  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null',
      '}'
    ].join('\n');
    const raw = execFileSync('powershell.exe', ['-NoProfile', '-Command', script], { encoding: 'utf8', maxBuffer: 1024 * 1024 });
    const data = JSON.parse(raw) as Record<string, unknown>;
    if (data.PermissionEnabled !== true) throw new Error(`Document does not report IRM permissions: ${raw}`);
    return { document_path: documentPath, ...data };
  });
}

async function runIrmGate(sessionId: string, mode: string, irmPreflight?: Record<string, unknown>, irmDocumentPath?: string): Promise<void> {
  if (mode === 'none') {
    addGate('irm_rights_matrix', 'skipped', {
      reason: 'Run with --irm-mode protected-read or --irm-mode protected-edit against a representative IRM document.'
    });
    return;
  }
  await runGate('irm_rights_matrix', async () => {
    const info = toolData(await client.callTool({ name: 'office.get_session_info', arguments: { session_id: sessionId } }));
    const document = info.document as { is_protected?: boolean; protection_kind?: string; rights?: unknown; rights_source?: string } | undefined;
    if (irmDocumentPath && !sessionInfoMatchesDocumentPath(info, irmDocumentPath)) {
      throw new Error(`Selected session does not match IRM document path: ${irmDocumentPath}`);
    }
    const preflightProtected = irmPreflight?.PermissionEnabled === true;
    if (!document?.is_protected && !preflightProtected && !irmDocumentPath) throw new Error('Selected session does not report a protected document and IRM preflight did not prove permissions.');
    const read = await client.callTool({ name: 'word.get_text', arguments: { session_id: sessionId, offset: 0, limit: 5 } });
    const mutation = mode === 'protected-edit'
      ? await client.callTool({ name: 'word.insert_paragraph', arguments: { session_id: sessionId, text: `IRM edit evidence ${Date.now()}`, anchor: { kind: 'end_of_document' }, style: 'Normal' } })
      : undefined;
    return {
      mode,
      protection_kind: document?.protection_kind ?? (irmDocumentPath ? 'document_path_preflight_pending' : preflightProtected ? 'irm_preflight' : undefined),
      rights_source: document?.rights_source ?? (irmDocumentPath ? 'document_path' : preflightProtected ? 'com_preflight' : undefined),
      rights: document?.rights,
      preflight: irmPreflight,
      read_ok: !isToolError(read),
      mutation_ok: mutation ? !isToolError(mutation) : undefined,
      mutation_error: mutation && isToolError(mutation) ? toolData(mutation) : undefined
    };
  });
}

async function runTrackedChangeGate(sessionId: string): Promise<void> {
  await runGate('word.tracked_change_resource_smoke', async () => {
    const data = resourceData(await client.readResource({ uri: `office://word/${sessionId}/track_changes` }));
    return { change_count: Array.isArray(data.changes) ? data.changes.length : undefined };
  });
}

async function runGate(name: string, run: () => Promise<Record<string, unknown>>): Promise<Record<string, unknown>> {
  const started = new Date().toISOString();
  try {
    const details = await run();
    report.gates.push({ name, status: 'passed', started_at: started, finished_at: new Date().toISOString(), details });
    return details;
  } catch (error) {
    report.gates.push({ name, status: 'failed', started_at: started, finished_at: new Date().toISOString(), details: { error: errorMessage(error) } });
    return {};
  }
}

function addGate(name: string, status: GateStatus, details: Record<string, unknown>): void {
  const now = new Date().toISOString();
  report.gates.push({ name, status, started_at: now, finished_at: now, details });
}

function toolData(result: unknown): Record<string, unknown> {
  const structured = (result as { structuredContent?: { data?: unknown; error?: unknown } }).structuredContent;
  if (structured?.data && typeof structured.data === 'object') return structured.data as Record<string, unknown>;
  if (structured?.error && typeof structured.error === 'object') return structured.error as Record<string, unknown>;
  const text = (result as { content?: Array<{ text?: string }> }).content?.[0]?.text;
  if (!text) return {};
  const parsed = JSON.parse(text);
  return parsed?.data ?? parsed;
}

function resourceData(result: unknown): Record<string, unknown> {
  const text = (result as { contents?: Array<{ text?: string }> }).contents?.[0]?.text;
  return text ? JSON.parse(text) : {};
}

function isToolError(result: unknown): boolean {
  return (result as { isError?: boolean }).isError === true;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isExplicitHostRejection(result: unknown, data: Record<string, unknown>): boolean {
  return isToolError(result) && ['HOST_CAPABILITY_UNAVAILABLE', 'HOST_ERROR'].includes(String(data.office_mcp_code ?? data.code ?? ''));
}

function runSmokeMode(mode: string, sessionId: string, extraArg?: string): string {
  const args = [resolve(evidenceRoot, 'mcp-smoke.ts'), endpoint, mode, sessionId];
  if (extraArg) args.push(extraArg);
  const raw = execFileSync(process.execPath, [tsxCli, ...args], {
    cwd: evidenceRoot,
    encoding: 'utf8',
    maxBuffer: 20 * 1024 * 1024
  });
  const jsonStart = raw.indexOf('{');
  if (jsonStart === -1) throw new Error(`Smoke mode ${mode} did not emit JSON.`);
  return raw.slice(jsonStart);
}

function selectWordSessionId(sessions: Array<Record<string, unknown>>, documentPath?: string): string {
  const wordSessions = sessions.filter((session) => {
    const host = session.host as { app?: string } | undefined;
    return String(host?.app ?? '').toLowerCase() === 'word';
  });
  if (documentPath) {
    const match = wordSessions.find((session) => sessionInfoMatchesDocumentPath(session, documentPath) && session.status === 'active');
    if (match?.session_id) return String(match.session_id);
    return '';
  }
  return String(wordSessions.find((session) => session.status === 'active')?.session_id ?? wordSessions[0]?.session_id ?? '');
}

function selectHostSessionId(sessions: Array<Record<string, unknown>>, hostApp: string): string {
  const match = sessions.find((session) => {
    const host = session.host as { app?: string } | undefined;
    return String(host?.app ?? '').toLowerCase() === hostApp && session.status === 'active';
  });
  return String(match?.session_id ?? '');
}

function sessionInfoMatchesDocumentPath(sessionOrInfo: Record<string, unknown>, documentPath: string): boolean {
  const document = sessionOrInfo.document as { title?: string; filename?: string; url?: string } | undefined;
  const expectedName = basename(documentPath).toLowerCase();
  const values = [document?.title, document?.filename, document?.url].filter((value): value is string => typeof value === 'string');
  return values.some((value) => value.toLowerCase().includes(expectedName));
}

function readOption(name: string): string | undefined {
  const index = process.argv.lastIndexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}

function hasFlag(name: string): boolean {
  return process.argv.includes(name);
}

function intOption(name: string, fallback: number): number {
  const value = readOption(name);
  if (value === undefined) return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 0) throw new Error(`${name} must be a non-negative integer`);
  return parsed;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function psString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
