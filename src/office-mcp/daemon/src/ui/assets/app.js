const state = { snapshot: null, search: '', app: 'all', client: 'all', result: 'all', previousStatus: null, selectedRowKey: null, toolAccessRenderKey: null };
const $ = (id) => document.getElementById(id);

$('search').addEventListener('input', (event) => { state.search = event.target.value.toLowerCase(); render(); });
$('appFilter').addEventListener('change', (event) => { state.app = event.target.value; render(); });
$('clientFilter').addEventListener('change', (event) => { state.client = event.target.value; render(); });
$('resultFilter').addEventListener('change', (event) => { state.result = event.target.value; render(); });
$('clearInspector').addEventListener('click', clearInspector);
$('refreshLogTail').addEventListener('click', refreshLogTail);
$('toolAccessMode').addEventListener('click', (event) => {
  const button = event.target.closest('[data-access-mode]');
  if (!button) return;
  updateToolAccessPolicy({ ...currentToolAccessPolicy(), access_mode: button.dataset.accessMode });
});
$('toolAccessList').addEventListener('click', (event) => {
  const toggle = event.target.closest('.tool-access-toggle');
  if (!toggle) return;
  event.preventDefault();
  event.stopPropagation();
  toggleToolAccess(toggle);
});

document.addEventListener('click', async (event) => {
  const copy = event.target.closest('[data-copy], [data-copy-value]');
  if (copy) {
    event.stopPropagation();
    const target = copy.dataset.copy ? $(copy.dataset.copy) : null;
    const value = copy.dataset.copyValue || target?.value || target?.textContent;
    await copyText(value, copy);
    return;
  }
  const diagnostic = event.target.closest('[data-open-diagnostic]');
  if (diagnostic) {
    event.stopPropagation();
    await openDiagnostic(diagnostic.dataset.openDiagnostic, diagnostic);
    return;
  }
  const inspect = event.target.closest('[data-inspect]');
  if (inspect) inspectRow(inspect);
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && clearInspector()) {
    event.preventDefault();
    return;
  }
  if (handleRowNavigation(event)) return;
  if (event.key !== 'Enter' && event.key !== ' ') return;
  const target = event.target.closest('[data-key-activate]');
  if (!target) return;
  event.preventDefault();
  target.click();
});

refresh();
refreshLogTail();
setInterval(refresh, 2000);

async function refresh() {
  try {
    const response = await fetch('/ui/state', { cache: 'no-store' });
    if (!response.ok) throw new Error('UI state returned ' + response.status);
    state.snapshot = await response.json();
    window.__OFFICE_MCP_UI__ = state.snapshot;
    announceStatus(snapshotStatus(state.snapshot));
    render();
  } catch (error) {
    $('healthBadge').textContent = 'Down';
    $('healthBadge').className = 'badge danger';
    $('daemonMeta').textContent = error.message || 'UI state unavailable';
    announce('Daemon UI state is unavailable');
  }
}

function render() {
  const snapshot = state.snapshot;
  if (!snapshot) return;
  const focusKey = focusedRowKey();
  const docs = Object.values(snapshot.documents || {}).flat();
  const health = title(snapshot.daemon?.status || 'down');
  $('healthBadge').textContent = health;
  $('healthBadge').className = 'badge ' + tone(snapshot.daemon?.status);
  $('daemonMeta').textContent = 'Daemon state is live';
  $('clientCount').textContent = fmt(snapshot.clients?.length || 0);
  $('documentCount').textContent = fmt(docs.length);
  $('taskCount').textContent = fmt(snapshot.current_tasks?.length || 0);
  $('mcpEndpoint').textContent = snapshot.daemon?.mcp_endpoint || '-';
  $('addinEndpoint').textContent = snapshot.daemon?.addin_endpoint || '-';
  $('daemonVersion').textContent = snapshot.daemon?.version || '-';
  $('daemonUptime').textContent = duration(snapshot.daemon?.uptime_ms || 0);
  $('configPath').textContent = snapshot.daemon?.config_path || '-';
  $('logPath').textContent = snapshot.daemon?.log_path || '-';
  setTextareaValue($('lastError'), snapshot.daemon?.last_error || 'None');
  renderToolAccessIfNeeded(snapshot.daemon?.tool_catalog || [], snapshot.daemon?.tool_access_policy || {});
  renderClientFilter(snapshot);
  renderDocuments(snapshot.documents || {});
  renderClients(snapshot.clients || []);
  renderCommands('currentTasks', filterCommands(snapshot.current_tasks || [], true), true);
  const history = filterCommands(snapshot.recent_commands || [], false);
  renderCommands('history', history, false);
  restoreSelectedRow();
  restoreRowFocus(focusKey);
}

function renderClientFilter(snapshot) {
  const options = clientFilterOptions(snapshot);
  if (state.client !== 'all' && !options.some((option) => option.value === state.client)) state.client = 'all';
  $('clientFilter').innerHTML = options.map((option) => `<option value="${esc(option.value)}"${option.value === state.client ? ' selected' : ''}>${esc(option.label)}</option>`).join('');
}

function clientFilterOptions(snapshot) {
  const clients = new Map();
  for (const client of snapshot.clients || []) clients.set(client.client_id, client.name || client.client_id);
  for (const command of [...(snapshot.current_tasks || []), ...(snapshot.recent_commands || [])]) {
    const id = command.client_id || command.client_name;
    if (id && !clients.has(id)) clients.set(id, command.client_name || id);
  }
  return [{ value: 'all', label: 'All clients' }, ...[...clients.entries()].map(([value, label]) => ({ value, label }))];
}

function renderToolAccessIfNeeded(catalog, policy, options = {}) {
  const renderKey = toolAccessRenderKey(catalog);
  if (!options.force && state.toolAccessRenderKey === renderKey) {
    return;
  }
  state.toolAccessRenderKey = renderKey;
  renderToolAccess(catalog, policy);
}

function renderToolAccess(catalog, policy) {
  renderToolAccessMode(policy.access_mode || 'all');
  const disabledApps = new Set(policy.disabled_apps || []);
  const disabledCategories = new Set((policy.disabled_categories || []).map(categoryKey));
  const disabledTools = new Set(policy.disabled_tools || []);
  $('toolAccessList').innerHTML = groupedToolAccessCatalog(catalog).map((appGroup) => {
    const appEnabled = !disabledApps.has(appGroup.app);
    const categoryMarkup = appGroup.categories.map((categoryGroup) => {
      const categoryEnabled = !disabledCategories.has(`${appGroup.app}:${categoryGroup.category}`);
      const tools = categoryGroup.tools.map((tool) => `<div class="tool-access-tool" data-tool-name="${esc(tool.name)}"><code title="${esc(tool.name)}">${esc(tool.name)}</code><span class="tool-access-effect">${esc(tool.side_effect)}</span>${toolAccessToggle(!disabledTools.has(tool.name), 'Toggle ' + tool.name, { tool: tool.name })}</div>`).join('');
      return `<details open data-tool-category="${esc(categoryGroup.category)}"><summary><span>${esc(categoryGroup.category)}</span><span class="tool-access-count">${esc(enabledToolCount(categoryGroup.tools, disabledTools))}/${esc(categoryGroup.tools.length)}</span>${toolAccessToggle(categoryEnabled, 'Toggle ' + categoryGroup.category, { app: appGroup.app, category: categoryGroup.category })}</summary><div class="tool-access-tools">${tools}</div></details>`;
    }).join('');
    return `<details open data-tool-app="${esc(appGroup.app)}"><summary><span>${esc(title(appGroup.app))}</span><span class="tool-access-count">${esc(enabledToolCount(appGroup.tools, disabledTools))}/${esc(appGroup.tools.length)}</span>${toolAccessToggle(appEnabled, 'Toggle ' + title(appGroup.app), { app: appGroup.app })}</summary>${categoryMarkup}</details>`;
  }).join('') || emptyState('No tool catalog available', 'The daemon did not publish tool metadata.');
}

function renderToolAccessMode(mode) {
  document.querySelectorAll('#toolAccessMode [data-access-mode]').forEach((button) => {
    button.setAttribute('aria-checked', button.dataset.accessMode === mode ? 'true' : 'false');
  });
}

function groupedToolAccessCatalog(catalog) {
  const apps = new Map();
  for (const tool of catalog) {
    if (!tool?.name || !tool?.app || !tool?.category) continue;
    if (!apps.has(tool.app)) apps.set(tool.app, { app: tool.app, categories: new Map(), tools: [] });
    const appGroup = apps.get(tool.app);
    appGroup.tools.push(tool);
    if (!appGroup.categories.has(tool.category)) appGroup.categories.set(tool.category, { category: tool.category, tools: [] });
    appGroup.categories.get(tool.category).tools.push(tool);
  }
  return [...apps.values()].map((appGroup) => ({ ...appGroup, categories: [...appGroup.categories.values()] }));
}

function toolAccessRenderKey(catalog) {
  return groupedToolAccessCatalog(catalog).map((appGroup) => {
    const categories = appGroup.categories.map((categoryGroup) => `${categoryGroup.category}:${categoryGroup.tools.map((tool) => tool.name).join(',')}`).join('|');
    return `${appGroup.app}:${categories}`;
  }).join(';');
}

function toolAccessToggle(enabled, label, data) {
  const dataAttributes = Object.entries(data).map(([key, value]) => `data-${key}="${esc(value)}"`).join(' ');
  return `<button type="button" class="tool-access-toggle" ${dataAttributes} aria-label="${esc(label)}" aria-pressed="${enabled ? 'true' : 'false'}"></button>`;
}

function enabledToolCount(tools, disabledTools) {
  return tools.filter((tool) => !disabledTools.has(tool.name)).length;
}

function currentToolAccessPolicy() {
  const policy = state.snapshot?.daemon?.tool_access_policy || {};
  return {
    access_mode: policy.access_mode || 'all',
    disabled_apps: [...(policy.disabled_apps || [])],
    disabled_categories: [...(policy.disabled_categories || [])],
    disabled_tools: [...(policy.disabled_tools || [])]
  };
}

function toggleToolAccess(toggle) {
  const policy = currentToolAccessPolicy();
  if (toggle.dataset.tool) {
    policy.disabled_tools = toggledList(policy.disabled_tools, toggle.dataset.tool);
  } else if (toggle.dataset.category) {
    const key = { app: toggle.dataset.app, category: toggle.dataset.category };
    policy.disabled_categories = toggledCategoryList(policy.disabled_categories, key);
  } else if (toggle.dataset.app) {
    policy.disabled_apps = toggledList(policy.disabled_apps, toggle.dataset.app);
  }
  updateToolAccessPolicy(policy);
}

async function updateToolAccessPolicy(policy) {
  try {
    const response = await fetch('/ui/tool-access-policy', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(policy)
    });
    if (!response.ok) throw new Error('Tool access update returned ' + response.status);
    state.snapshot = await response.json();
    window.__OFFICE_MCP_UI__ = state.snapshot;
    renderToolAccessIfNeeded(state.snapshot.daemon?.tool_catalog || [], state.snapshot.daemon?.tool_access_policy || {}, { force: true });
    render();
    announce('Updated global tool access');
  } catch (error) {
    announce(error.message || 'Tool access update failed');
  }
}

async function refreshLogTail() {
  const button = $('refreshLogTail');
  if (button.disabled) return;
  button.disabled = true;
  try {
    const response = await fetch('/ui/log-tail', { cache: 'no-store' });
    if (!response.ok) throw new Error(await response.text() || 'Log tail returned ' + response.status);
    const body = await response.json();
    setTextareaValue($('logTail'), body.text || 'No log output yet.');
    $('logTailMeta').textContent = `${body.truncated ? 'Showing latest' : 'Showing'} ${fmt(body.bytes_read || 0)} bytes from ${body.path || 'daemon log'}.`;
    announce('Daemon log tail refreshed');
  } catch (error) {
    setTextareaValue($('logTail'), error.message || 'Daemon log tail is unavailable.');
    $('logTailMeta').textContent = 'Daemon log tail is unavailable.';
  } finally {
    button.disabled = false;
  }
}

function toggledList(values, value) {
  return values.includes(value) ? values.filter((item) => item !== value) : [...values, value];
}

function toggledCategoryList(values, value) {
  const key = categoryKey(value);
  return values.some((item) => categoryKey(item) === key) ? values.filter((item) => categoryKey(item) !== key) : [...values, value];
}

function categoryKey(value) { return `${value.app}:${value.category}`; }

function renderDocuments(groups) {
  const rows = [];
  for (const [app, docs] of Object.entries(groups)) {
    if (state.app !== 'all' && app !== state.app) continue;
    const visible = docs.filter((doc) => matches(JSON.stringify(doc)));
    if (!visible.length) continue;
    rows.push(`<h3>${esc(title(app))}</h3>`);
    for (const doc of visible) {
      rows.push(renderDocumentCard(doc, app));
    }
  }
  const filtered = state.search || state.app !== 'all';
  $('documents').innerHTML = rows.join('') || (filtered ? emptyState('No matching documents', 'Adjust the app or search filter.') : emptyState('No documents connected', 'Open Word, Excel, or PowerPoint, then open Office MCP Control.', state.snapshot?.daemon?.addin_endpoint, 'Copy add-in endpoint'));
  annotateInspectableRows($('documents'));
}

function renderDocumentCard(doc, app) {
  const label = doc.document?.title || doc.document?.filename || 'Untitled';
  const status = documentConnectionLabel(doc.status);
  const metrics = documentTaskMetrics(doc.session_id);
  const sessionId = doc.session_id || '-';
  const hostVersion = `${title(doc.host?.app || app)} ${doc.host?.version || '-'}`;
  return `<button class="row document-card ${esc(app)}" type="button" data-key-activate data-focus-key="document:${esc(doc.session_id || label)}" data-inspect='${attr(doc)}' aria-label="Inspect ${esc(label)} ${esc(status)}"><span class="document-card-title"><span class="state-dot ${esc(documentStateTone(doc.status))}" aria-hidden="true"></span><strong title="${esc(label)}">${esc(label)}</strong><span class="pill ${esc(documentStateTone(doc.status))}">${esc(title(status))}</span></span><span class="document-card-session" data-copy-value="${esc(sessionId)}" title="${esc(sessionId)}"><code>${esc(middleTruncate(sessionId, 24))}</code></span><span class="document-card-meta"><span>${esc(hostVersion)}</span><span>${esc(doc.available_tool_count || 0)} tools</span><span>Queue ${esc(doc.queue_depth || 0)}</span></span><span class="document-card-meta document-card-footer"><span>Finished ${esc(metrics.finished)}</span><span>Failed ${esc(metrics.failed)}</span></span><span class="document-card-uptime" title="Session uptime">${esc(sessionUptime(doc))}</span></button>`;
}

function sessionUptime(doc) {
  const registeredAt = sessionRegisteredAtMillis(doc.registered_at);
  if (!Number.isFinite(registeredAt)) return '0s';
  return duration(Math.max(0, Date.now() - registeredAt));
}

function sessionRegisteredAtMillis(value) {
  const text = String(value || '');
  if (text.startsWith('unix:')) return Number(text.slice(5)) * 1000;
  return Date.parse(text);
}

function documentTaskMetrics(sessionId) {
  const commands = state.snapshot?.document_command_history?.[sessionId] || [];
  return commands.reduce((metrics, command) => {
    if (command.status === 'failure' || command.status === 'timeout') metrics.failed += 1;
    else if (command.status === 'success') metrics.finished += 1;
    return metrics;
  }, { finished: 0, failed: 0 });
}

function renderClients(clients) {
  $('clients').innerHTML = clients.map((client) => `<button class="row" type="button" data-key-activate data-focus-key="client:${esc(client.client_id || client.name)}" data-inspect='${attr(client)}'><strong>${esc(client.name || client.client_id)}</strong><span>${esc(client.transport)} | in flight ${esc(client.in_flight_request_count || 0)}</span></button>`).join('') || clientEmptyState(state.snapshot?.daemon?.mcp_endpoint);
  annotateInspectableRows($('clients'));
}

function clientEmptyState(mcpEndpoint) {
  return `<p class="empty"><strong>No MCP clients connected</strong>Connect an MCP client using either local transport.${emptyCopy(mcpEndpoint, 'Copy MCP endpoint', 'MCP endpoint')}${emptyCopy('office-mcp-daemon stdio', 'Copy stdio bridge command', 'Stdio bridge')}</p>`;
}

function renderCommands(target, commands, running) {
  if (!commands.length) {
    $(target).innerHTML = running ? emptyState('No command is running', 'New tool calls appear here while they are in flight.') : emptyState('No command history yet', 'Completed, failed, cancelled, and timed-out commands appear here.');
    return;
  }
  const rows = commands.map((command) => `<tr tabindex="0" role="button" aria-label="Inspect ${esc(command.tool)} ${esc(statusLabel(command.status))}" data-key-activate data-focus-key="command:${esc(command.command_id || command.mcp_request_id || command.tool)}" data-inspect='${attr(command)}'><td><strong>${esc(command.tool)}</strong><br><small>${copyableId(command.command_id || command.mcp_request_id, 'Copy command ID')} / ${copyableId(command.session_id, 'Copy session ID')}</small></td><td>${esc(command.client_name || command.client_id || '-')}</td><td><span class="pill ${tone(command.status)}">${esc(statusLabel(command.status))}</span></td><td>${running ? duration(Date.now() - (command.started_at || Date.now())) : duration(command.elapsed_ms || 0)}</td><td>${esc(command.error?.office_mcp_code || '')}<br><small>${esc(command.error?.message || '')}</small></td></tr>`).join('');
  $(target).innerHTML = `<table><thead><tr><th>Tool</th><th>Client</th><th>Status</th><th>Time</th><th>Error</th></tr></thead><tbody>${rows}</tbody></table>`;
  annotateInspectableRows($(target));
}

function filterCommands(commands, running) {
  return commands.filter((command) => clientMatches(command) && (running || state.result === 'all' || command.status === state.result) && matches(JSON.stringify(command)));
}

function clientMatches(command) { return state.client === 'all' || command.client_id === state.client || command.client_name === state.client; }

function annotateInspectableRows(scope) {
  const rows = [...(scope?.querySelectorAll?.('.row, tr[data-inspect]') || [])].filter((item) => !item.disabled);
  rows.forEach((row, index) => {
    row.setAttribute('aria-posinset', String(index + 1));
    row.setAttribute('aria-setsize', String(rows.length));
    if (!row.hasAttribute('aria-selected')) row.setAttribute('aria-selected', 'false');
  });
}

function handleRowNavigation(event) {
  const keys = ['ArrowDown', 'ArrowUp', 'Home', 'End', 'PageDown', 'PageUp'];
  if (!keys.includes(event.key)) return false;
  const target = event.target instanceof Element ? event.target : document.activeElement;
  const row = target?.closest('.row, tr[data-inspect]');
  if (!row) return false;
  const scope = row.closest('#documents, #clients, #currentTasks, #history');
  if (!scope) return false;
  const rows = [...scope.querySelectorAll('.row, tr[data-inspect]')].filter((item) => !item.disabled);
  const index = rows.indexOf(row);
  if (index < 0) return false;
  const pageStep = Math.max(1, Math.min(5, rows.length - 1));
  const nextIndex = rowNavigationIndex(event.key, index, rows.length, pageStep);
  event.preventDefault();
  rows[nextIndex].focus();
  return true;
}

function focusedRowKey() {
  return document.activeElement?.closest?.('.row, tr[data-inspect]')?.dataset.focusKey || null;
}

function restoreRowFocus(focusKey) {
  if (!focusKey) return;
  const row = document.querySelector(`[data-focus-key="${cssEscape(focusKey)}"]`);
  if (row) row.focus();
}

function rowNavigationIndex(key, index, count, pageStep) {
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  if (key === 'ArrowDown') return Math.min(count - 1, index + 1);
  if (key === 'ArrowUp') return Math.max(0, index - 1);
  if (key === 'PageDown') return Math.min(count - 1, index + pageStep);
  if (key === 'PageUp') return Math.max(0, index - pageStep);
  return index;
}

function inspectRow(element) {
  state.selectedRowKey = element.dataset.focusKey || null;
  setSelectedInspectableRow(element);
  $('inspectorLog').value = JSON.stringify(JSON.parse(element.dataset.inspect), null, 2);
}

function setSelectedInspectableRow(element) {
  const scope = element.closest('#documents, #clients, #currentTasks, #history');
  scope?.querySelectorAll?.('.row, tr[data-inspect]').forEach((row) => row.setAttribute('aria-selected', row === element ? 'true' : 'false'));
}

function restoreSelectedRow() {
  document.querySelectorAll('.row, tr[data-inspect]').forEach((row) => {
    row.setAttribute('aria-selected', row.dataset.focusKey === state.selectedRowKey ? 'true' : 'false');
  });
}

function clearInspector() {
  const inspector = $('inspectorLog');
  if (!inspector || inspector.value.trim() === 'Select a row.') return false;
  inspector.value = 'Select a row.';
  state.selectedRowKey = null;
  document.querySelectorAll('.row[aria-selected="true"], tr[data-inspect][aria-selected="true"]').forEach((row) => row.setAttribute('aria-selected', 'false'));
  announce('Inspector cleared');
  return true;
}

function setTextareaValue(textarea, value) {
  if (!textarea || textarea.value === value) return;
  const focused = document.activeElement === textarea;
  const selectionStart = textarea.selectionStart;
  const selectionEnd = textarea.selectionEnd;
  textarea.value = value;
  if (focused) textarea.setSelectionRange(Math.min(selectionStart, value.length), Math.min(selectionEnd, value.length));
}

async function copyText(text, button) {
  if (!text || text === '-') return;
  if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
  else fallbackCopy(text);
  announce('Copied ' + (button.querySelector('span')?.textContent || 'value'));
}

async function openDiagnostic(target, button) {
  if (!target || button.disabled) return;
  button.disabled = true;
  try {
    const response = await fetch('/ui/open-diagnostic', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ target })
    });
    if (!response.ok) throw new Error(await response.text() || 'Failed to open diagnostic path');
    announce('Opened ' + target + ' path');
  } catch (error) {
    announce(error.message || 'Failed to open diagnostic path');
  } finally {
    button.disabled = false;
  }
}

function emptyState(titleText, bodyText, codeText, copyLabel = 'Copy endpoint') {
  const label = copyLabel.replace(/^Copy\s+/i, '');
  const copy = emptyCopy(codeText, copyLabel, label);
  return `<p class="empty"><strong>${esc(titleText)}</strong>${esc(bodyText)}${copy}</p>`;
}

function emptyCopy(codeText, copyLabel, label) {
  return codeText ? `<button type="button" class="empty-copy" data-copy-value="${esc(codeText)}" aria-label="${esc(copyLabel)}" title="${esc(codeText)}"><span>${esc(label)}</span><code>${esc(middleTruncate(codeText, 46))}</code></button>` : '';
}

function copyableId(value, label) {
  const text = String(value || '-');
  if (text === '-') return '-';
  return `<button type="button" class="id-copy" data-copy-value="${esc(text)}" aria-label="${esc(label)}" title="${esc(text)}"><code>${esc(middleTruncate(text))}</code></button>`;
}

function middleTruncate(value, maxLength = 30) {
  const text = String(value || '');
  if (text.length <= maxLength) return text;
  const marker = '...';
  const available = maxLength - marker.length;
  const head = Math.ceil(available / 2);
  const tail = Math.floor(available / 2);
  return `${text.slice(0, head)}${marker}${text.slice(text.length - tail)}`;
}

function fallbackCopy(text) {
  const area = document.createElement('textarea');
  area.value = text;
  area.setAttribute('readonly', '');
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.appendChild(area);
  area.select();
  document.execCommand('copy');
  area.remove();
}

function matches(text) { return !state.search || text.toLowerCase().includes(state.search); }
function documentConnectionLabel(status) { return status === 'active' || !status ? 'active' : 'dead'; }
function documentStateTone(status) { return status === 'active' || !status ? 'success' : 'danger'; }
function fmt(value) { return new Intl.NumberFormat().format(value); }
function duration(ms) { if (!ms) return '0s'; const seconds = ms / 1000; if (seconds < 60) return new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 }).format(seconds) + 's'; return new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 }).format(seconds / 60) + 'm'; }
function relative(value) { if (!value) return 'now'; const delta = Math.round((Number(value) - Date.now()) / 1000); const abs = Math.abs(delta); const unit = abs < 60 ? 'second' : abs < 3600 ? 'minute' : 'hour'; const divisor = unit === 'second' ? 1 : unit === 'minute' ? 60 : 3600; return new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' }).format(Math.round(delta / divisor), unit); }
function statusLabel(value) { if (value === 'success') return 'Succeeded'; if (value === 'failure') return 'Failed'; if (value === 'timeout') return 'Timed Out'; if (value === 'cancelled') return 'Cancelled'; return title(value); }
function title(value) { return String(value || '').replace(/_/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase()); }
function tone(value) { return value === 'up' || value === 'success' ? 'success' : value === 'running' ? 'accent' : value === 'timeout' ? 'warning' : value === 'degraded' ? 'warning' : value === 'down' || value === 'failure' ? 'danger' : 'neutral'; }
function snapshotStatus(snapshot) { return snapshot?.daemon?.status || 'down'; }
function announceStatus(status) { if (state.previousStatus === status) return; state.previousStatus = status; announce('Daemon status ' + title(status)); }
function announce(message) { $('announcer').textContent = message; }
function esc(value) { return String(value ?? '').replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char]); }
function attr(value) { return esc(JSON.stringify(value)).replace(/'/g, '&#39;'); }
function cssEscape(value) { return String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"'); }
