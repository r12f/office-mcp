(function attachOfficeCtlMainUi(global) {
  function middleTruncate(value, maxLength = 30) {
    const text = String(value || '');
    if (text.length <= maxLength) return text;
    const marker = '...';
    const available = Math.max(4, maxLength - marker.length);
    const head = Math.ceil(available / 2);
    const tail = Math.floor(available / 2);
    return `${text.slice(0, head)}${marker}${text.slice(text.length - tail)}`;
  }

  function setCopyableMetadata(element, value) {
    if (!element) return;
    const text = value ? String(value) : '-';
    element.textContent = element.id === 'session' ? text : middleTruncate(text);
    const button = element.closest('button');
    if (button) {
      button.dataset.copyValue = text;
      button.title = text === '-' ? button.getAttribute('aria-label') || '' : text;
    }
  }

  async function copyMetadataValue(event, options = {}) {
    const button = event?.target?.closest?.('[data-copy-target], [data-copy-value]');
    if (!button) return false;
    const documentRef = options.document || global.document;
    const navigatorRef = options.navigator || global.navigator;
    const target = button.dataset.copyTarget ? documentRef?.getElementById?.(button.dataset.copyTarget) : null;
    const value = button.dataset.copyValue || target?.textContent?.trim();
    if (!value || value === '-') return false;
    try {
      if (navigatorRef?.clipboard?.writeText) await navigatorRef.clipboard.writeText(value);
      else (options.fallbackCopy || fallbackCopy)(value, documentRef);
      if (options.announcer) options.announcer.textContent = `Copied ${button.getAttribute?.('aria-label') || 'value'}`;
      return true;
    } catch (error) {
      options.logger?.warn?.('metadata_copy.failed', error);
      if (options.announcer) options.announcer.textContent = 'Copy failed';
      return false;
    }
  }

  function fallbackCopy(value, documentRef = global.document) {
    const area = documentRef.createElement('textarea');
    area.value = value;
    area.setAttribute('readonly', '');
    area.style.position = 'fixed';
    area.style.opacity = '0';
    documentRef.body.appendChild(area);
    area.select();
    documentRef.execCommand('copy');
    area.remove();
  }

  function renderRuntimeVersions(serverVersionElement, protocolVersionElement, serverInfo, fallbackProtocolVersion) {
    const info = serverInfo || {};
    if (serverVersionElement) serverVersionElement.textContent = `Server ${info.serverVersion || 'Unknown'}`;
    if (protocolVersionElement) protocolVersionElement.textContent = `Protocol ${info.protocolVersion || fallbackProtocolVersion || 'Unknown'}`;
  }

  function renderStaticMetadata(elements, options = {}) {
    setCopyableMetadata(elements?.session, options.sessionId);
    setCopyableMetadata(elements?.daemon, options.endpoint);
    renderRuntimeVersions(elements?.serverVersion, elements?.protocolVersion, options.serverInfo, options.protocolVersion);
    if (elements?.hostPlatform) elements.hostPlatform.textContent = officeHostSummary(options.defaultHost);
  }

  function renderToolModeControl(control, selectedMode) {
    control?.querySelectorAll('[data-tool-mode]').forEach((button) => {
      const selected = button.dataset.toolMode === selectedMode;
      button.setAttribute('aria-checked', selected ? 'true' : 'false');
    });
  }

  function isToolAllowedByCapabilityMode(mode, sideEffect = 'read') {
    if (mode === 'read') return sideEffect === 'read';
    if (mode === 'write') return sideEffect !== 'destructive';
    return true;
  }

  function officeHostSummary(defaultHost) {
    const diagnostics = global.Office?.context?.diagnostics || {};
    const host = diagnostics.host || defaultHost || 'Office';
    const version = diagnostics.version || 'Unknown';
    const platform = global.Office?.context?.platform || 'Unknown';
    return `${host} ${version} / ${platform}`;
  }

  function protectionLabel(info = {}) {
    const label = info.protection?.label || info.protection?.kind;
    if (!label || String(label).toLowerCase() === 'none') return info.is_protected === true ? 'Protected' : 'Not protected';
    return String(label);
  }

  function documentStateLabel(info = {}) {
    if (info.is_read_only === true) return 'Read-only';
    const label = info.protection?.label || info.protection?.kind;
    if (info.is_protected === true || (label && String(label).toLowerCase() !== 'none')) return `Protected${label && String(label).toLowerCase() !== 'none' ? `: ${label}` : ''}`;
    if (info.is_dirty === true) return 'Editable, unsaved changes';
    return 'Editable';
  }

  function stopDetailsToggle(event) {
    event.stopPropagation();
    if (event.type === 'keydown' && event.key !== ' ' && event.key !== 'Enter') return;
    if (event.type === 'keydown') event.stopPropagation();
  }

  function bindDetailsControl(control, onChange) {
    if (!control) return;
    for (const eventName of ['pointerdown', 'mousedown', 'touchstart', 'click', 'keydown']) {
      control.addEventListener(eventName, stopDetailsToggle);
    }
    if (onChange) control.addEventListener('change', onChange);
  }

  function statusClass(state) {
    if (state === 'connected' || state === 'success') return 'status-success';
    if (state === 'connecting' || state === 'reconnecting' || state === 'running') return 'status-warning';
    if (state === 'failed' || state === 'failure' || state === 'disconnected' || state === 'unsupported') return 'status-danger';
    return 'status-neutral';
  }

  function setConnectionState(elements, state, label) {
    const badge = elements?.badge;
    if (badge) {
      badge.textContent = label;
      badge.className = `status-badge ${statusClass(state)}`;
    }
    if (state === 'connected' && elements?.detail) elements.detail.textContent = 'None';
    if (elements?.announcer) elements.announcer.textContent = label;
  }

  function commandIdMarkup(requestId, options = {}) {
    if (!requestId) return '';
    const escapeHtml = options.escapeHtml || ((value) => String(value));
    const escaped = escapeHtml(requestId);
    return `<div class="task-meta task-command-id">Command <button type="button" class="inline-copy" data-copy-value="${escaped}" aria-label="Copy command ID" title="${escaped}"><code>${escapeHtml(middleTruncate(requestId))}</code></button></div>`;
  }

  function taskMetadataMarkup(task, options = {}) {
    const escapeHtml = options.escapeHtml || ((value) => String(value));
    const formatTime = options.formatTime || ((value) => String(value));
    const redactText = options.redactText || ((value) => String(value));
    const valueLabel = options.valueLabel || ((value) => String(value));
    const error = task.error ? `<div class="task-meta">${escapeHtml(task.error.office_mcp_code)}: ${escapeHtml(task.error.message)} · Retriable: ${valueLabel(task.error.retriable)} · Partial effect: ${escapeHtml(task.error.partial_effect || 'unknown')}</div>` : '';
    const intent = task.userIntent ? `<div class="task-meta">${escapeHtml(redactText(task.userIntent))}</div>` : '';
    const deadline = task.deadlineAt ? `<div class="task-meta">Deadline ${escapeHtml(formatTime(task.deadlineAt))}</div>` : '';
    const cancel = task.cancelRequested ? '<div class="task-meta">Cancel requested</div>' : '';
    return `${deadline}${cancel}${intent}${error}`;
  }

  global.OfficeCtlMainUi = Object.freeze({
    bindDetailsControl,
    commandIdMarkup,
    copyMetadataValue,
    isToolAllowedByCapabilityMode,
    middleTruncate,
    officeHostSummary,
    protectionLabel,
    renderRuntimeVersions,
    renderStaticMetadata,
    renderToolModeControl,
    setConnectionState,
    setCopyableMetadata,
    statusClass,
    documentStateLabel,
    taskMetadataMarkup
  });
})(globalThis);
