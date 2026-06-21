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
      else options.fallbackCopy?.(value);
      if (options.announcer) options.announcer.textContent = `Copied ${button.getAttribute?.('aria-label') || 'value'}`;
      return true;
    } catch (error) {
      options.logger?.warn?.('metadata_copy.failed', error);
      if (options.announcer) options.announcer.textContent = 'Copy failed';
      return false;
    }
  }

  function renderRuntimeVersions(serverVersionElement, protocolVersionElement, serverInfo, fallbackProtocolVersion) {
    const info = serverInfo || {};
    if (serverVersionElement) serverVersionElement.textContent = `Server ${info.serverVersion || 'Unknown'}`;
    if (protocolVersionElement) protocolVersionElement.textContent = `Protocol ${info.protocolVersion || fallbackProtocolVersion || 'Unknown'}`;
  }

  function officeHostSummary(defaultHost) {
    const diagnostics = global.Office?.context?.diagnostics || {};
    const host = diagnostics.host || defaultHost || 'Office';
    const version = diagnostics.version || 'Unknown';
    const platform = global.Office?.context?.platform || 'Unknown';
    return `${host} ${version} / ${platform}`;
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

  function commandIdMarkup(requestId, options = {}) {
    if (!requestId) return '';
    const escapeHtml = options.escapeHtml || ((value) => String(value));
    const escaped = escapeHtml(requestId);
    return `<div class="task-meta task-command-id">Command <button type="button" class="inline-copy" data-copy-value="${escaped}" aria-label="Copy command ID" title="${escaped}"><code>${escapeHtml(middleTruncate(requestId))}</code></button></div>`;
  }

  global.OfficeCtlMainUi = Object.freeze({
    bindDetailsControl,
    commandIdMarkup,
    copyMetadataValue,
    middleTruncate,
    officeHostSummary,
    renderRuntimeVersions,
    setCopyableMetadata,
    statusClass
  });
})(globalThis);
