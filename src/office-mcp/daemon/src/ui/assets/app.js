const state = { snapshot: null, search: '', result: 'all', expandedDocuments: new Set(), previousStatus: null };
const $ = (id) => document.getElementById(id);

$('search').addEventListener('input', (event) => { state.search = event.target.value.toLowerCase(); render(); });
$('resultFilter').addEventListener('change', (event) => { state.result = event.target.value; render(); });
$('clearInspector').addEventListener('click', () => { $('inspector').textContent = 'Select a row.'; announce('Inspector cleared'); });

document.addEventListener('click', async (event) => {
  const copy = event.target.closest('[data-copy], [data-copy-value]');
  if (copy) {
    event.stopPropagation();
    const value = copy.dataset.copyValue || $(copy.dataset.copy)?.textContent;
    await copyText(value, copy);
    return;
  }
  const toggle = event.target.closest('[data-document-toggle]');
  if (toggle) toggleDocument(toggle.dataset.documentToggle);
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
  $('configPath').textContent = snapshot.daemon?.config_path || 'Not configured';
  $('logPath').textContent = snapshot.daemon?.log_path || 'Not configured';
  $('lastError').textContent = snapshot.daemon?.last_error || 'None';
  renderDocuments(snapshot.documents || {});
  renderClients(snapshot.clients || []);
  renderCommands('currentTasks', snapshot.current_tasks || [], true);
  const history = (snapshot.recent_commands || []).filter((command) => state.result === 'all' || command.status === state.result);
  renderCommands('history', history, false);
}

function renderDocuments(groups) {
  const rows = [];
  const documentCommandHistory = state.snapshot?.document_command_history || {};
  for (const [app, docs] of Object.entries(groups)) {
    const visible = docs.filter((doc) => matches(JSON.stringify(doc)));
    if (!visible.length) continue;
    rows.push(`<h3>${esc(title(app))}</h3>`);
    for (const doc of visible) {
      const label = doc.document?.title || doc.document?.filename || 'Untitled';
      const connection = documentConnectionLabel(doc.status);
      const expanded = state.expandedDocuments.has(doc.session_id);
      const detailId = `document-detail-${safeId(doc.session_id)}`;
      rows.push(`<button class="row ${esc(app)}" type="button" data-key-activate data-document-toggle="${esc(doc.session_id)}" data-inspect='${attr(doc)}' aria-expanded="${expanded}" aria-controls="${detailId}"><strong>${esc(label)}</strong><span>${esc(connection)} | ${copyableId(doc.session_id, 'Copy session ID')}</span><span class="meta">${expanded ? 'Hide details' : 'Show details'} | ${esc(doc.host?.version || '-')} | ${esc(doc.available_tool_count || 0)} tools | queue ${esc(doc.queue_depth || 0)}</span></button>`);
      rows.push(renderDocumentHistory(doc.session_id, documentCommandHistory[doc.session_id] || [], expanded, detailId));
    }
  }
  $('documents').innerHTML = rows.join('') || emptyState('No documents connected', 'Open Word, Excel, or PowerPoint, then open Office MCP Control.', state.snapshot?.daemon?.addin_endpoint);
}

function renderDocumentHistory(sessionId, commands, expanded, detailId) {
  const hidden = expanded ? '' : ' hidden';
  if (!commands.length) return `<div id="${detailId}" class="doc-history" aria-label="Command history for ${esc(sessionId)}"${hidden}>${emptyState('No recent commands for this document', 'Completed commands for this document appear here.')}</div>`;
  const rows = commands.slice(0, 10).map((command) => `<button class="history-row" type="button" data-inspect='${attr(command)}'><span><strong>${esc(command.tool)}</strong><small>${copyableId(command.command_id || command.mcp_request_id, 'Copy command ID')} | ${esc(relative(command.completed_at || command.started_at))}</small></span><span class="pill ${tone(command.status)}">${esc(title(command.status))}</span><small>${esc(command.error?.office_mcp_code || '')}</small></button>`).join('');
  return `<div id="${detailId}" class="doc-history" aria-label="Command history for ${esc(sessionId)}"${hidden}>${rows}</div>`;
}

function renderClients(clients) {
  $('clients').innerHTML = clients.map((client) => `<button class="row" type="button" data-inspect='${attr(client)}'><strong>${esc(client.name || client.client_id)}</strong><span>${esc(client.transport)} | in flight ${esc(client.in_flight_request_count || 0)}</span></button>`).join('') || emptyState('No MCP clients connected', 'Connect an MCP client using this endpoint.', state.snapshot?.daemon?.mcp_endpoint);
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

function emptyState(titleText, bodyText, codeText) {
  const code = codeText ? `<code>${esc(codeText)}</code>` : '';
  return `<p class="empty"><strong>${esc(titleText)}</strong>${esc(bodyText)}${code}</p>`;
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
function documentConnectionLabel(status) { return status === 'stale' ? 'stale | reconnecting' : (status || 'active'); }
function toggleDocument(sessionId) {
  if (!sessionId) return;
  if (state.expandedDocuments.has(sessionId)) state.expandedDocuments.delete(sessionId);
  else state.expandedDocuments.add(sessionId);
  render();
  document.querySelector(`[data-document-toggle="${cssEscape(sessionId)}"]`)?.focus();
}
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
function safeId(value) { return String(value || 'unknown').replace(/[^a-zA-Z0-9_-]/g, '-'); }
function cssEscape(value) { return String(value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"'); }
