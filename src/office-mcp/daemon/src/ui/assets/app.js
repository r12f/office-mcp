const state = { snapshot: null, search: '', result: 'all', expandedDocuments: new Set() };
const $ = (id) => document.getElementById(id);
$('search').addEventListener('input', (event) => { state.search = event.target.value.toLowerCase(); render(); });
$('resultFilter').addEventListener('change', (event) => { state.result = event.target.value; render(); });
$('clearInspector').addEventListener('click', () => { $('inspector').textContent = 'Select a row.'; });
document.addEventListener('click', async (event) => {
  const copy = event.target.closest('[data-copy]');
  if (copy) await copyText($(copy.dataset.copy).textContent, copy);
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
    render();
  } catch (error) {
    $('healthBadge').textContent = 'Down';
    $('healthBadge').className = 'badge danger';
    $('daemonMeta').textContent = error.message || 'UI state unavailable';
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
  const document_command_history = state.snapshot?.document_command_history || {};
  for (const [app, docs] of Object.entries(groups)) {
    const visible = docs.filter((doc) => matches(JSON.stringify(doc)));
    if (!visible.length) continue;
    rows.push(`<h3>${esc(title(app))}</h3>`);
    for (const doc of visible) {
      const label = doc.document?.title || doc.document?.filename || 'Untitled';
      const connection = documentConnectionLabel(doc.status);
      const expanded = state.expandedDocuments.has(doc.session_id);
      const detailId = `document-detail-${safeId(doc.session_id)}`;
      rows.push(`<button class="row ${esc(app)}" type="button" data-key-activate data-document-toggle="${esc(doc.session_id)}" data-inspect='${attr(doc)}' aria-expanded="${expanded}" aria-controls="${detailId}"><strong>${esc(label)}</strong><span>${esc(connection)} | ${esc(doc.session_id)}</span><span class="meta">${expanded ? 'Hide details' : 'Show details'} | ${esc(doc.host?.version || '-')} | ${esc(doc.available_tool_count || 0)} tools | queue ${esc(doc.queue_depth || 0)}</span></button>`);
      rows.push(renderDocumentHistory(doc.session_id, document_command_history[doc.session_id] || [], expanded, detailId));
    }
  }
  $('documents').innerHTML = rows.join('') || '<p class="empty">No documents connected. Open Word and load the add-in.</p>';
}
function renderDocumentHistory(sessionId, commands, expanded, detailId) {
  const hidden = expanded ? '' : ' hidden';
  if (!commands.length) return `<div id="${detailId}" class="doc-history" aria-label="Command history for ${esc(sessionId)}"${hidden}><p class="empty">No recent commands for this document.</p></div>`;
  const rows = commands.slice(0, 10).map((command) => `<button class="history-row" type="button" data-inspect='${attr(command)}'><span><strong>${esc(command.tool)}</strong><small>${esc(relative(command.completed_at || command.started_at))}</small></span><span class="pill ${tone(command.status)}">${esc(title(command.status))}</span><small>${esc(command.error?.office_mcp_code || '')}</small></button>`).join('');
  return `<div id="${detailId}" class="doc-history" aria-label="Command history for ${esc(sessionId)}"${hidden}>${rows}</div>`;
}
function renderClients(clients) {
  $('clients').innerHTML = clients.map((client) => `<button class="row" type="button" data-inspect='${attr(client)}'><strong>${esc(client.name || client.client_id)}</strong><span>${esc(client.transport)} | in flight ${esc(client.in_flight_request_count || 0)}</span></button>`).join('') || '<p class="empty">No MCP clients connected.</p>';
}
function renderCommands(target, commands, running) {
  if (!commands.length) { $(target).innerHTML = `<p class="empty">${running ? 'No command is running.' : 'No command history yet.'}</p>`; return; }
  const rows = commands.map((command) => `<tr tabindex="0" role="button" aria-label="Inspect ${esc(command.tool)} ${esc(title(command.status))}" data-key-activate data-inspect='${attr(command)}'><td><strong>${esc(command.tool)}</strong><br><small>${esc(command.session_id || '-')}</small></td><td>${esc(command.client_name || command.client_id || '-')}</td><td><span class="pill ${tone(command.status)}">${esc(title(command.status))}</span></td><td>${running ? duration(Date.now() - (command.started_at || Date.now())) : duration(command.elapsed_ms || 0)}</td><td>${esc(command.error?.office_mcp_code || '')}<br><small>${esc(command.error?.message || '')}</small></td></tr>`).join('');
  $(target).innerHTML = `<table><thead><tr><th>Tool</th><th>Client</th><th>Status</th><th>Time</th><th>Error</th></tr></thead><tbody>${rows}</tbody></table>`;
}
function inspectRow(element) { $('inspector').textContent = JSON.stringify(JSON.parse(element.dataset.inspect), null, 2); }
async function copyText(text, button) {
  await navigator.clipboard?.writeText(text);
  $('announcer').textContent = 'Copied ' + (button.querySelector('span')?.textContent || 'value');
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
function esc(value) { return String(value ?? '').replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char]); }
function attr(value) { return esc(JSON.stringify(value)).replace(/'/g, '&#39;'); }
function safeId(value) { return String(value || 'unknown').replace(/[^a-zA-Z0-9_-]/g, '-'); }
function cssEscape(value) { return String(value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"'); }
