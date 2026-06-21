const state = { snapshot: null, search: '', app: 'all', result: 'all', previousStatus: null };
const $ = (id) => document.getElementById(id);

$('search').addEventListener('input', (event) => { state.search = event.target.value.toLowerCase(); render(); });
$('appFilter').addEventListener('change', (event) => { state.app = event.target.value; render(); });
$('resultFilter').addEventListener('change', (event) => { state.result = event.target.value; render(); });
$('clearInspector').addEventListener('click', () => { $('inspector').textContent = 'Select a row.'; announce('Inspector cleared'); });
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
    const value = copy.dataset.copyValue || $(copy.dataset.copy)?.textContent;
    await copyText(value, copy);
    return;
  }
  const inspect = event.target.closest('[data-inspect]');
  if (inspect) inspectRow(inspect);
});

document.addEventListener('keydown', (event) => {
  if (event.key !== 'Enter' && event.key !== ' ') return;
  const target = event.target.closest('[data-key-activate]');
  if (!target) return;
  event.preventDefault();
  target.click();
});

refresh();
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
  const docs = Object.values(snapshot.documents || {}).flat();
  const health = title(snapshot.daemon?.status || 'down');
  $('healthBadge').textContent = health;
  $('healthBadge').className = 'badge ' + tone(snapshot.daemon?.status);
  $('daemonMeta').textContent = snapshot.daemon?.last_error || 'Daemon state is live';
  $('clientCount').textContent = fmt(snapshot.clients?.length || 0);
  $('documentCount').textContent = fmt(docs.length);
  $('taskCount').textContent = fmt(snapshot.current_tasks?.length || 0);
  $('mcpEndpoint').textContent = snapshot.daemon?.mcp_endpoint || '-';
  $('addinEndpoint').textContent = snapshot.daemon?.addin_endpoint || '-';
  $('daemonVersion').textContent = snapshot.daemon?.version || '-';
  $('daemonUptime').textContent = duration(snapshot.daemon?.uptime_ms || 0);
  $('configPath').textContent = snapshot.daemon?.config_path || '-';
  $('logPath').textContent = snapshot.daemon?.log_path || '-';
  $('lastError').textContent = snapshot.daemon?.last_error || 'None';
  renderToolAccess(snapshot.daemon?.tool_catalog || [], snapshot.daemon?.tool_access_policy || {});
  renderDocuments(snapshot.documents || {});
  renderClients(snapshot.clients || []);
  renderCommands('currentTasks', snapshot.current_tasks || [], true);
  const history = (snapshot.recent_commands || []).filter((command) => state.result === 'all' || command.status === state.result);
  renderCommands('history', history, false);
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
    render();
    announce('Updated global tool access');
  } catch (error) {
    announce(error.message || 'Tool access update failed');
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
}

function renderDocumentCard(doc, app) {
  const label = doc.document?.title || doc.document?.filename || 'Untitled';
  const status = documentConnectionLabel(doc.status);
  const metrics = documentTaskMetrics(doc.session_id);
  return `<button class="row document-card ${esc(app)}" type="button" data-key-activate data-inspect='${attr(doc)}' aria-label="Inspect ${esc(label)} ${esc(status)}"><span class="document-card-title"><span class="state-dot ${esc(documentStateTone(doc.status))}" aria-hidden="true"></span><strong title="${esc(label)}">${esc(label)}</strong><span class="pill ${esc(documentStateTone(doc.status))}">${esc(title(status))}</span></span><span class="document-card-session" data-copy-value="${esc(doc.session_id || '-')}" title="${esc(doc.session_id || '-')}"><span>Session ID</span><code>${esc(doc.session_id || '-')}</code></span><span class="document-card-meta"><span>Version ${esc(doc.host?.version || '-')}</span><span>${esc(doc.available_tool_count || 0)} tools</span><span>Queue ${esc(doc.queue_depth || 0)}</span></span><span class="document-card-meta"><span>Finished ${esc(metrics.finished)}</span><span>Failed ${esc(metrics.failed)}</span></span></button>`;
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
  $('clients').innerHTML = clients.map((client) => `<button class="row" type="button" data-inspect='${attr(client)}'><strong>${esc(client.name || client.client_id)}</strong><span>${esc(client.transport)} | in flight ${esc(client.in_flight_request_count || 0)}</span></button>`).join('') || emptyState('No MCP clients connected', 'Connect an MCP client using this endpoint.', state.snapshot?.daemon?.mcp_endpoint, 'Copy MCP endpoint');
}

function renderCommands(target, commands, running) {
  if (!commands.length) {
    $(target).innerHTML = running ? emptyState('No command is running', 'New tool calls appear here while they are in flight.') : emptyState('No command history yet', 'Completed, failed, cancelled, and timed-out commands appear here.');
    return;
  }
  const rows = commands.map((command) => `<tr tabindex="0" role="button" aria-label="Inspect ${esc(command.tool)} ${esc(title(command.status))}" data-key-activate data-inspect='${attr(command)}'><td><strong>${esc(command.tool)}</strong><br><small>${copyableId(command.command_id || command.mcp_request_id, 'Copy command ID')} / ${copyableId(command.session_id, 'Copy session ID')}</small></td><td>${esc(command.client_name || command.client_id || '-')}</td><td><span class="pill ${tone(command.status)}">${esc(title(command.status))}</span></td><td>${running ? duration(Date.now() - (command.started_at || Date.now())) : duration(command.elapsed_ms || 0)}</td><td>${esc(command.error?.office_mcp_code || '')}<br><small>${esc(command.error?.message || '')}</small></td></tr>`).join('');
  $(target).innerHTML = `<table><thead><tr><th>Tool</th><th>Client</th><th>Status</th><th>Time</th><th>Error</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function inspectRow(element) { $('inspector').textContent = JSON.stringify(JSON.parse(element.dataset.inspect), null, 2); }

async function copyText(text, button) {
  if (!text || text === '-') return;
  if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
  else fallbackCopy(text);
  announce('Copied ' + (button.querySelector('span')?.textContent || 'value'));
}

function emptyState(titleText, bodyText, codeText, copyLabel = 'Copy endpoint') {
  const label = copyLabel.replace(/^Copy\s+/i, '');
  const copy = codeText ? `<button type="button" class="empty-copy" data-copy-value="${esc(codeText)}" aria-label="${esc(copyLabel)}" title="${esc(codeText)}"><span>${esc(label)}</span><code>${esc(middleTruncate(codeText, 46))}</code></button>` : '';
  return `<p class="empty"><strong>${esc(titleText)}</strong>${esc(bodyText)}${copy}</p>`;
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
function title(value) { return String(value || '').replace(/_/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase()); }
function tone(value) { return value === 'up' || value === 'success' ? 'success' : value === 'degraded' || value === 'running' ? 'warning' : value === 'down' || value === 'failure' || value === 'timeout' ? 'danger' : 'neutral'; }
function snapshotStatus(snapshot) { return snapshot?.daemon?.status || 'down'; }
function announceStatus(status) { if (state.previousStatus === status) return; state.previousStatus = status; announce('Daemon status ' + title(status)); }
function announce(message) { $('announcer').textContent = message; }
function esc(value) { return String(value ?? '').replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char]); }
function attr(value) { return esc(JSON.stringify(value)).replace(/'/g, '&#39;'); }
