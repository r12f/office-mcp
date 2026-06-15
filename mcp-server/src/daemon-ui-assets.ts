import type { DaemonConfig } from './config.js';

export function renderDaemonConsoleShell(config: DaemonConfig, uiToken: string): string {
  const boot = {
    token: uiToken,
    stateUrl: `${config.addin.origin}/ui/state`,
    eventsUrl: `${config.addin.origin}/ui/events`
  };
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta name="theme-color" content="#F7F8FA" />
    <title>Office MCP</title>
    <link rel="stylesheet" href="/ui/app.css" />
    <script>window.__OFFICE_MCP_UI__ = ${JSON.stringify(boot)};</script>
    <script defer src="/ui/app.js"></script>
  </head>
  <body>
    <a class="skip-link" href="#main">Skip to Activity</a>
    <main id="main" class="shell">
      <header class="status-strip" aria-label="Daemon Status">
        <div class="brand-lockup"><span id="health" class="badge neutral">Loading...</span><h1>Office MCP</h1></div>
        <div class="strip-metrics" aria-label="Live Counts"><span id="clientCount">0 Clients</span><span id="documentCount">0 Documents</span><span id="taskCount">0 Running</span></div>
        <div class="endpoint-strip" aria-label="Endpoints"><span><code id="mcpEndpoint">-</code><button class="icon-button" type="button" data-copy="mcpEndpoint" aria-label="Copy MCP Endpoint">Copy</button></span><span><code id="addinEndpoint">-</code><button class="icon-button" type="button" data-copy="addinEndpoint" aria-label="Copy Add-in Endpoint">Copy</button></span></div>
      </header>
      <section class="layout">
        <nav class="panel navigation" aria-labelledby="navigationHeading"><div class="section-heading"><h2 id="navigationHeading">Connections</h2><span id="filterLabel" class="meta-pill">All Apps</span></div><label class="search-label" for="searchInput">Search</label><input id="searchInput" name="search" type="search" autocomplete="off" spellcheck="false" placeholder="session, tool, client..." /><div id="documents" class="document-groups"></div><section aria-labelledby="clientsHeading"><h3 id="clientsHeading">Clients</h3><div id="clients"></div></section></nav>
        <section class="panel activity" aria-labelledby="activityHeading"><div class="section-heading"><h2 id="activityHeading">Activity</h2><select id="resultFilter" name="resultFilter" aria-label="Filter Command Results"><option value="all">All Results</option><option value="success">Succeeded</option><option value="failure">Failed</option><option value="timeout">Timed Out</option><option value="cancelled">Cancelled</option></select></div><h3>Current Tasks</h3><div id="currentTasks"></div><h3>Recent Command History</h3><div id="history"></div></section>
        <aside class="panel inspector" aria-labelledby="inspectorHeading"><div class="section-heading"><h2 id="inspectorHeading">Inspector</h2><button id="closeInspector" class="icon-button" type="button" aria-label="Close Inspector">Close</button></div><div id="inspector">Select a row.</div></aside>
      </section>
      <div id="announcer" class="sr-only" aria-live="polite"></div>
    </main>
  </body>
</html>`;
}

export const DAEMON_CONSOLE_JS = String.raw`(() => {
  const boot = window.__OFFICE_MCP_UI__;
  const params = new URLSearchParams(location.search);
  const ui = { snapshot: null, selected: params.get('selected') || '', search: params.get('q') || '', result: params.get('result') || 'all' };
  const nf = new Intl.NumberFormat(undefined);
  const df = new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 });
  const tf = new Intl.DateTimeFormat(undefined, { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const el = (id) => document.getElementById(id);
  const search = el('searchInput');
  const result = el('resultFilter');
  search.value = ui.search;
  result.value = ui.result;
  search.addEventListener('input', () => { ui.search = search.value; syncUrl(); render(ui.snapshot); });
  result.addEventListener('change', () => { ui.result = result.value; syncUrl(); render(ui.snapshot); });
  el('closeInspector').addEventListener('click', () => { ui.selected = ''; syncUrl(); renderInspector(); });
  document.addEventListener('click', async (event) => {
    const copy = event.target.closest('[data-copy]');
    if (copy) await copyText(el(copy.dataset.copy)?.textContent || '', copy);
    const row = event.target.closest('[data-select]');
    if (row) { ui.selected = row.dataset.select; syncUrl(); renderInspector(); }
  });
  refresh();
  setInterval(refresh, 2000);
  streamEvents();
  async function refresh() {
    try {
      const response = await fetch(boot.stateUrl, { headers: { 'x-office-mcp-ui-token': boot.token } });
      if (!response.ok) throw new Error('UI state returned ' + response.status);
      render(await response.json());
    } catch (error) {
      el('health').textContent = 'Down';
      el('health').className = 'badge danger';
      announce(error.message || 'UI state unavailable.');
    }
  }
  async function streamEvents() {
    if (!boot.eventsUrl || !window.TextDecoder) return;
    try {
      const response = await fetch(boot.eventsUrl, { headers: { 'x-office-mcp-ui-token': boot.token } });
      if (!response.ok || !response.body) return;
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      for (;;) {
        const next = await reader.read();
        if (next.done) break;
        buffer += decoder.decode(next.value, { stream: true });
        const events = buffer.split('\n\n');
        buffer = events.pop() || '';
        for (const event of events) {
          const data = event.split('\n').find((line) => line.startsWith('data: '));
          if (data) render(JSON.parse(data.slice(6)));
        }
      }
    } catch {
      // Polling remains active when the live event stream is unavailable.
    }
  }
  function render(snapshot) {
    if (!snapshot) return;
    ui.snapshot = snapshot;
    const docList = Object.values(snapshot.documents || {}).flat();
    el('health').textContent = title(snapshot.daemon.status);
    el('health').className = 'badge ' + tone(snapshot.daemon.status);
    el('clientCount').textContent = count(snapshot.clients.length, 'Client');
    el('documentCount').textContent = count(docList.length, 'Document');
    el('taskCount').textContent = count(snapshot.current_tasks.length, 'Running');
    el('mcpEndpoint').textContent = snapshot.daemon.mcp_endpoint;
    el('addinEndpoint').textContent = snapshot.daemon.addin_endpoint;
    el('documents').innerHTML = renderDocumentGroups(snapshot.documents || {});
    el('clients').innerHTML = snapshot.clients.length ? table(['Client', 'Transport', 'Last Activity', 'In Flight'], snapshot.clients.map(clientRow)) : '<p class="empty">No clients connected. Copy the MCP endpoint or use the stdio bridge command.</p>';
    el('currentTasks').innerHTML = snapshot.current_tasks.length ? table(['Tool', 'Target', 'Client', 'Elapsed', 'Result'], snapshot.current_tasks.map(commandRow)) : '<p class="empty">No command is running.</p>';
    const filtered = (snapshot.recent_commands || []).filter((cmd) => (ui.result === 'all' || cmd.status === ui.result) && matches(cmd));
    el('history').innerHTML = filtered.length ? table(['Tool', 'Target', 'Client', 'Duration', 'Result'], filtered.map(commandRow)) : '<p class="empty">No command history matches the current filter.</p>';
    renderInspector();
  }
  function renderDocumentGroups(groups) {
    const html = ['word', 'excel', 'powerpoint', 'outlook', 'other'].map((app) => {
      const rows = (groups[app] || []).filter(matches);
      return rows.length ? '<section class="app-group ' + app + '"><h3>' + appLabel(app) + '</h3>' + rows.map(documentRow).join('') + '</section>' : '';
    }).join('');
    return html || '<p class="empty">No documents connected. Open the add-in in Word and keep the task pane running.</p>';
  }
  function documentRow(doc) {
    const id = 'document:' + doc.session_id;
    return '<button class="row ' + esc(doc.app) + selected(id) + '" type="button" data-select="' + attr(id) + '"><span>' + esc(doc.document.title || doc.document.filename || 'Untitled') + '</span><small>' + esc(doc.status) + ' - ' + shortId(doc.session_id) + ' - ' + nf.format(doc.available_tool_count) + ' tools</small></button>';
  }
  function commandRow(command) {
    const id = 'command:' + command.command_id;
    return '<tr><td><button class="table-select" type="button" data-select="' + attr(id) + '">' + esc(command.tool) + '</button></td><td>' + esc(shortId(command.session_id || '-')) + '</td><td>' + esc(command.client_name || command.client_id || '-') + '</td><td class="num">' + elapsed(command) + '</td><td><span class="badge ' + tone(command.status) + '">' + statusText(command.status) + '</span></td></tr>';
  }
  function clientRow(client) {
    const id = 'client:' + client.client_id;
    return '<tr><td><button class="table-select" type="button" data-select="' + attr(id) + '">' + esc(client.name || shortId(client.client_id)) + '</button></td><td>' + esc(client.transport) + '</td><td>' + time(client.last_activity_at) + '</td><td class="num">' + nf.format(client.in_flight_request_count) + '</td></tr>';
  }
  function renderInspector() {
    if (!ui.snapshot || !ui.selected) { el('inspector').innerHTML = '<p class="empty">Select a client, document, task, or command.</p>'; return; }
    const split = ui.selected.indexOf(':');
    const kind = ui.selected.slice(0, split);
    const id = ui.selected.slice(split + 1);
    if (kind === 'document') return renderDocumentInspector(id);
    if (kind === 'client') return renderJsonInspector('Client', (ui.snapshot.clients || []).find((client) => client.client_id === id));
    if (kind === 'command') return renderJsonInspector('Command', [...(ui.snapshot.current_tasks || []), ...(ui.snapshot.recent_commands || [])].find((cmd) => cmd.command_id === id));
    el('inspector').innerHTML = '<p class="empty">Selection is no longer available.</p>';
  }
  function renderDocumentInspector(sessionId) {
    const doc = Object.values(ui.snapshot.documents || {}).flat().find((item) => item.session_id === sessionId);
    if (!doc) { el('inspector').innerHTML = '<p class="empty">Document disconnected.</p>'; return; }
    const items = ui.snapshot.document_command_history?.[sessionId] || [];
    el('inspector').innerHTML = '<dl class="details"><dt>Title</dt><dd>' + esc(doc.document.title || '-') + '</dd><dt>Session</dt><dd><code>' + esc(doc.session_id) + '</code></dd><dt>Host</dt><dd>' + esc([doc.host?.app, doc.host?.version, doc.host?.platform].filter(Boolean).join(' - ') || '-') + '</dd><dt>Protection</dt><dd>' + esc(doc.document.protection_kind || 'unknown') + '</dd><dt>Dirty / Read-only</dt><dd>' + bool(doc.document.is_dirty) + ' / ' + bool(doc.document.is_read_only) + '</dd><dt>Queue Depth</dt><dd>' + nf.format(doc.queue_depth || 0) + '</dd></dl><h3>Latest 10 Commands</h3>' + (items.length ? table(['Tool', 'Duration', 'Result'], items.map((cmd) => '<tr><td>' + esc(cmd.tool) + '</td><td class="num">' + elapsed(cmd) + '</td><td>' + statusText(cmd.status) + '</td></tr>')) : '<p class="empty">No commands have completed for this document.</p>');
  }
  function renderJsonInspector(label, value) { el('inspector').innerHTML = value ? '<h3>' + esc(label) + '</h3><pre>' + esc(JSON.stringify(value, null, 2)) + '</pre>' : '<p class="empty">Selection is no longer available.</p>'; }
  function table(headings, rows) { return '<table><thead><tr>' + headings.map((heading) => '<th scope="col">' + esc(heading) + '</th>').join('') + '</tr></thead><tbody>' + rows.join('') + '</tbody></table>'; }
  function syncUrl() { const p = new URLSearchParams(); if (ui.selected) p.set('selected', ui.selected); if (ui.search) p.set('q', ui.search); if (ui.result !== 'all') p.set('result', ui.result); history.replaceState(null, '', p.toString() ? '?' + p.toString() : location.pathname); }
  async function copyText(text, button) { await navigator.clipboard.writeText(text); announce('Copied ' + text); button.textContent = 'Copied'; setTimeout(() => { button.textContent = 'Copy'; }, 1200); }
  function matches(item) { const q = ui.search.trim().toLowerCase(); return !q || JSON.stringify(item).toLowerCase().includes(q); }
  function count(value, noun) { return value + ' ' + noun + (value === 1 ? '' : 's'); }
  function title(value) { return String(value || '').replace(/_/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase()); }
  function tone(value) { return value === 'up' || value === 'success' ? 'success' : value === 'degraded' || value === 'running' || value === 'timeout' ? 'warning' : value === 'down' || value === 'failure' ? 'danger' : 'neutral'; }
  function statusText(value) { return value === 'success' ? 'Succeeded' : value === 'failure' ? 'Failed' : value === 'timeout' ? 'Timed Out' : value === 'cancelled' ? 'Cancelled' : title(value); }
  function elapsed(command) { return df.format(((command.elapsed_ms ?? (command.started_at ? Date.now() - Date.parse(command.started_at) : 0)) || 0) / 1000) + 's'; }
  function time(value) { return value ? tf.format(new Date(value)) : '-'; }
  function shortId(value) { return String(value || '').length > 13 ? String(value).slice(0, 6) + '...' + String(value).slice(-6) : String(value || '-'); }
  function selected(id) { return ui.selected === id ? ' selected' : ''; }
  function appLabel(app) { return app === 'powerpoint' ? 'PowerPoint' : title(app); }
  function bool(value) { return value === true ? 'yes' : value === false ? 'no' : 'unknown'; }
  function announce(text) { el('announcer').textContent = text; }
  function esc(value) { return String(value).replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char]); }
  function attr(value) { return esc(value).replace(/'/g, '&#39;'); }
})();`;

export const DAEMON_CONSOLE_CSS = `:root{color-scheme:light;--canvas:#f7f8fa;--surface:#fff;--raised:#f2f5f8;--text:#17202a;--muted:#5a6673;--border:#d8dee6;--word:#2b579a;--excel:#217346;--powerpoint:#b7472a;--outlook:#0078d4;--success:#168a45;--warning:#8a5a00;--danger:#c9352b;--focus:#4c8dff}*{box-sizing:border-box}html,body{min-width:320px;margin:0;overflow-x:hidden;background:var(--canvas);color:var(--text);font:13px/1.45 "Segoe UI",system-ui,sans-serif;-webkit-tap-highlight-color:transparent}button,input,select{font:inherit}button{touch-action:manipulation}.skip-link{position:fixed;top:8px;left:8px;z-index:10;transform:translateY(-150%);background:var(--surface);color:var(--text);padding:6px 10px;border:1px solid var(--border);border-radius:6px;transition:transform 120ms ease-out}.skip-link:focus-visible{transform:translateY(0)}button:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid var(--focus);outline-offset:2px}.shell{display:grid;gap:12px;min-height:100vh;padding:12px}.status-strip{display:grid;grid-template-columns:auto minmax(220px,1fr) minmax(280px,1.4fr);align-items:center;gap:16px;padding:12px;border:1px solid var(--border);border-radius:8px;background:var(--surface)}.brand-lockup{display:grid;gap:4px}.strip-metrics,.endpoint-strip{display:flex;gap:10px;min-width:0;color:var(--muted);font-variant-numeric:tabular-nums}.endpoint-strip{flex-wrap:wrap;justify-content:flex-end}.endpoint-strip span{display:flex;align-items:center;min-width:0;gap:6px}h1,h2,h3{margin:0;letter-spacing:0;text-wrap:balance}h1{font-size:18px}h2{font-size:15px}h3{margin-top:12px;font-size:12px;color:var(--muted);text-transform:uppercase}.layout{display:grid;grid-template-columns:minmax(240px,310px) minmax(430px,1fr) minmax(280px,380px);gap:12px;min-height:0}.panel{min-width:0;padding:12px;border:1px solid var(--border);border-radius:8px;background:var(--surface)}.section-heading{display:flex;align-items:center;justify-content:space-between;gap:8px}.activity{display:grid;align-content:start;gap:10px}.badge,.meta-pill{display:inline-flex;gap:6px;align-items:center;width:fit-content;max-width:100%;padding:2px 7px;border:1px solid var(--border);border-radius:999px;background:var(--raised);font-variant-numeric:tabular-nums;white-space:nowrap}.badge:before{width:7px;height:7px;border-radius:50%;background:currentColor;content:""}.success{color:var(--success)}.warning{color:var(--warning)}.danger{color:var(--danger)}.neutral{color:var(--muted)}.search-label{display:block;margin-top:10px;font-weight:600}input,select{width:100%;min-width:0;margin-top:5px;padding:7px 8px;border:1px solid var(--border);border-radius:7px;background:var(--surface);color:var(--text)}select{width:auto}.icon-button{min-height:32px;border:1px solid var(--border);border-radius:7px;background:var(--surface);color:var(--text);cursor:pointer}.icon-button:hover,button:hover{border-color:var(--focus)}.row{display:grid;width:100%;min-width:0;gap:2px;margin-top:8px;padding:10px;border:1px solid var(--border);border-left:3px solid var(--border);border-radius:7px;background:var(--raised);color:var(--text);text-align:left;cursor:pointer}.row.word{border-left-color:var(--word)}.row.excel{border-left-color:var(--excel)}.row.powerpoint{border-left-color:var(--powerpoint)}.row.outlook{border-left-color:var(--outlook)}.row.selected{border-color:var(--focus)}.row span,.row small{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row small,.empty{color:var(--muted)}table{width:100%;border-collapse:collapse;table-layout:fixed;font-variant-numeric:tabular-nums}th,td{min-width:0;padding:7px 8px;border-bottom:1px solid var(--border);overflow:hidden;text-align:left;text-overflow:ellipsis;white-space:nowrap}th{color:var(--muted);font-size:12px;font-weight:650}.num{text-align:right}.table-select{max-width:100%;border:0;background:transparent;color:var(--text);text-align:left;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:pointer}code,pre{font-family:"Cascadia Mono",Consolas,monospace;font-size:12px}pre{overflow:auto;white-space:pre-wrap;overflow-wrap:anywhere}.details{display:grid;grid-template-columns:96px minmax(0,1fr);gap:7px 10px;margin:0 0 12px}.details dt{color:var(--muted)}.details dd{min-width:0;margin:0;overflow-wrap:anywhere}.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap}@media(max-width:1000px){.status-strip,.layout{grid-template-columns:1fr}.endpoint-strip{justify-content:flex-start}.inspector{order:3}}@media(prefers-reduced-motion:reduce){*,*:before,*:after{scroll-behavior:auto!important;transition-duration:.01ms!important;animation-duration:.01ms!important;animation-iteration-count:1!important}}@media(prefers-color-scheme:dark){:root{color-scheme:dark;--canvas:#11161c;--surface:#18212b;--raised:#202b36;--text:#f2f5f8;--muted:#b7c0ca;--border:#344250;--focus:#78a8ff}}@media(forced-colors:active){.status-strip,.panel,.row,.badge,.meta-pill,.icon-button,input,select{border:1px solid CanvasText}}`;
