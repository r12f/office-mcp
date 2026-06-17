(function attachOfficeCtlCommon(global) {
  const OfficeCtlCommon = {
    boolLabel(value) {
      return value === true ? 'yes' : value === false ? 'no' : 'unknown';
    },

    escapeHtml(value) {
      return String(value).replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[char]);
    },

    fileName(path) {
      return String(path || '').split(/[\\/]/).pop() || String(path || '');
    },

    formatDuration(ms) {
      return `${new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 }).format(Number(ms || 0) / 1000)}s`;
    },

    formatTime(value) {
      return new Intl.DateTimeFormat(undefined, { hour: '2-digit', minute: '2-digit', second: '2-digit' }).format(new Date(value));
    },

    redactText(text) {
      return String(text || '')
        .replace(/Bearer\s+[A-Za-z0-9._~+\/-]+=*/gi, 'Bearer [redacted]')
        .replace(/(password|passphrase|token)=([^\s&]+)/gi, '$1=[redacted]')
        .replace(/base64,[A-Za-z0-9+/=]+/g, 'base64,[redacted]')
        .slice(0, 300);
    },

    titleCase(value) {
      return String(value || '').replace(/_/g, ' ').replace(/\b\w/g, (char) => char.toUpperCase());
    }
  };

  global.OfficeCtlCommon = Object.freeze(OfficeCtlCommon);
})(globalThis);
