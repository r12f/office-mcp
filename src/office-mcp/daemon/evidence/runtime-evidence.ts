import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
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
const endpoint = readOption('--endpoint') ?? process.env.OFFICE_MCP_MCP_ENDPOINT ?? 'http://127.0.0.1:8800/mcp';
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/runtime-evidence.json'));
const requestedSessionId = readOption('--session-id');
const irmMode = readOption('--irm-mode') ?? 'none';
const irmDocumentPath = readOption('--irm-document-path') ?? process.env.OFFICE_MCP_IRM_DOCUMENT_PATH;
const waitForSessionMs = intOption('--wait-for-session-ms', 0);
const agentClientEvidencePath = readOption('--agent-client-evidence-path') ?? process.env.OFFICE_MCP_AGENT_CLIENT_EVIDENCE_PATH;

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

  const sessionId = requestedSessionId ?? selectWordSessionId(sessions, irmDocumentPath);
  if (sessionId) report.session_id = sessionId;

  if (!sessionId) {
    const reason = irmDocumentPath
      ? `No connected Word add-in session matched the requested IRM document path: ${irmDocumentPath}. Open that document and open Office MCP Control, then rerun this script.`
      : 'No connected Word add-in session. Open Word, open Office MCP Control, then rerun this script.';
    addGate('agent_client_stdio_bridge', 'blocked_by_runtime', {
      reason: 'No connected Word add-in session to prove an agent client can call MCP.'
    });
    if (irmDocumentPath) await runIrmDocumentPreflightGate(irmDocumentPath);
    await runAgentClientPromptGate(agentClientEvidencePath);
    addGate('irm_rights_matrix', 'blocked_by_runtime', {
      reason: irmDocumentPath ? reason : 'No connected IRM-protected Word session was provided.'
    });
  } else {
    await runAgentClientBridgeGate(sessionId);
    await runClaudeDesktopInstallationGate();
    await runAgentClientPromptGate(agentClientEvidencePath);
    await runIrmGate(sessionId, irmMode, undefined, irmDocumentPath);
    if (irmDocumentPath) await runIrmDocumentPreflightGate(irmDocumentPath);
  }
} catch (error) {
  addGate('runtime_evidence_harness', 'failed', { error: errorMessage(error) });
} finally {
  await client.close().catch(() => undefined);
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
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

async function runAgentClientBridgeGate(sessionId: string): Promise<void> {
  await runGate('agent_client_stdio_bridge', async () => {
    const stdioClient = new Client({ name: 'office-mcp-agent-client-evidence', version: '0.1.0' });
    const stdioTransport = new StdioClientTransport({ command: 'cargo', args: ['run', '-q', '-p', 'office-mcp-daemon', '--', 'stdio'], cwd: repoRoot, stderr: 'pipe' });
    await stdioClient.connect(stdioTransport);
    try {
      const data = toolData(await stdioClient.callTool({ name: 'office.list_sessions', arguments: {} }));
      const sessions = Array.isArray(data.sessions) ? data.sessions : [];
      if (sessions.length < 1) throw new Error('stdio bridge connected but returned no sessions.');
      const info = toolData(await stdioClient.callTool({ name: 'office.get_session_info', arguments: { session_id: sessionId } }));
      if (!isRecord(info.document)) throw new Error('stdio bridge could not read session metadata.');
      return { session_count: sessions.length, document: info.document, host: info.host };
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

function isToolError(result: unknown): boolean {
  return (result as { isError?: boolean }).isError === true;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
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
