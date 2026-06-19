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

  global.OfficeCtlMainUi = Object.freeze({
    bindDetailsControl,
    middleTruncate,
    officeHostSummary,
    renderRuntimeVersions,
    setCopyableMetadata
  });
})(globalThis);
