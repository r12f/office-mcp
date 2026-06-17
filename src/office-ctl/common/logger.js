(function attachOfficeCtlLogger(global) {
  const DEFAULT_SCOPE = 'office-mcp';

  class AddinLogger {
    constructor(options = {}) {
      this.scope = options.scope || DEFAULT_SCOPE;
      this.console = options.console || global.console;
      this.redactText = options.redactText || ((value) => String(value || '').slice(0, 300));
    }

    info(event, fields = {}) {
      this.write('info', event, fields);
    }

    warn(event, fields = {}) {
      this.write('warn', event, fields);
    }

    error(event, fields = {}) {
      this.write('error', event, fields);
    }

    write(level, event, fields) {
      const record = {
        scope: this.scope,
        event: this.redactText(event),
        ...this.redactFields(fields)
      };
      const sink = this.console?.[level] || this.console?.log;
      if (typeof sink === 'function') sink.call(this.console, `${this.scope} ${record.event}`, record);
    }

    redactFields(fields) {
      const safe = {};
      for (const [key, value] of Object.entries(fields || {})) {
        safe[key] = this.redactValue(value);
      }
      return safe;
    }

    redactValue(value) {
      if (value == null || typeof value === 'number' || typeof value === 'boolean') return value;
      if (Array.isArray(value)) return value.map((item) => this.redactValue(item));
      if (typeof value === 'object') {
        const result = {};
        for (const [key, nested] of Object.entries(value)) {
          if (/^(text|body|content|base64)$/i.test(key)) result[key] = '[redacted]';
          else result[key] = this.redactValue(nested);
        }
        return result;
      }
      return this.redactText(value);
    }
  }

  global.OfficeCtlLogger = Object.freeze({ AddinLogger });
})(globalThis);
