import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { assertBoundaryAuthConfig, loadConfig, parseToml } from '../src/config.js';
import type { DaemonConfig } from '../src/config.js';

test('parses the supported daemon config TOML subset', () => {
  assert.deepEqual(parseToml(`
[addin_channel]
bind = "localhost"
port = 8765 # inline comments are ignored
shared_secret = "not#comment"

[mcp_http]
bind = "127.0.0.1"
port = 8800

[limits]
requests_per_minute = 120
`), {
    addin_channel: { bind: 'localhost', port: 8765, shared_secret: 'not#comment' },
    mcp_http: { bind: '127.0.0.1', port: 8800 },
    limits: { requests_per_minute: 120 }
  });
});

test('loads daemon config from config.toml', () => {
  withTempConfig(`
[addin_channel]
bind = "localhost"
port = 9443
certificate_path = "C:\\\\certs\\\\office-mcp.pfx"
certificate_passphrase = "secret"
heartbeat_interval_sec = 11
heartbeat_timeout_sec = 5
session_grace_sec = 22
max_pending_per_session = 2

[mcp_http]
bind = "127.0.0.1"
port = 9900

[limits]
max_response_bytes = 1234
max_request_bytes = 4567
max_ws_frame_bytes = 6789
default_tool_timeout_ms = 321
requests_per_minute = 42

[audit]
enabled = true
path = "C:\\\\logs\\\\office-mcp-audit.jsonl"

[logging]
level = "debug"
file = "C:\\\\logs\\\\office-mcp.log"
`, (configPath) => {
    const loaded = loadConfig({ configPath });

    assert.equal(loaded.addin.port, 9443);
    assert.equal(loaded.addin.origin, 'https://localhost:9443');
    assert.equal(loaded.addin.pfxPath, 'C:\\certs\\office-mcp.pfx');
    assert.equal(loaded.addin.pfxPassphrase, 'secret');
    assert.equal(loaded.addin.heartbeatIntervalSec, 11);
    assert.equal(loaded.addin.heartbeatTimeoutSec, 5);
    assert.equal(loaded.addin.sessionGraceSec, 22);
    assert.equal(loaded.addin.maxPendingPerSession, 2);
    assert.equal(loaded.mcp.port, 9900);
    assert.equal(loaded.limits.maxResponseBytes, 1234);
    assert.equal(loaded.limits.maxRequestBytes, 4567);
    assert.equal(loaded.limits.maxWsFrameBytes, 6789);
    assert.equal(loaded.limits.defaultToolTimeoutMs, 321);
    assert.equal(loaded.limits.requestsPerMinute, 42);
    assert.equal(loaded.audit.enabled, true);
    assert.equal(loaded.audit.path, 'C:\\logs\\office-mcp-audit.jsonl');
    assert.equal(loaded.logging.level, 'debug');
    assert.equal(loaded.logging.file, 'C:\\logs\\office-mcp.log');
  });
});

test('empty audit path falls back to the platform default audit path', () => {
  withTempConfig(`
[audit]
enabled = true
path = ""
`, (configPath) => {
    const loaded = loadConfig({ configPath });

    assert.equal(loaded.audit.enabled, true);
    assert.match(loaded.audit.path.replace(/\\/g, '/'), /office-mcp\/audit\.jsonl$/);
  });
});

test('empty logging file falls back to the platform default log file', () => {
  withTempConfig(`
[logging]
level = "warn"
file = ""
`, (configPath) => {
    const loaded = loadConfig({ configPath });

    assert.equal(loaded.logging.level, 'warn');
    assert.match(loaded.logging.file.replace(/\\/g, '/'), /office-mcp\/office-mcp\.log$/);
  });
});

test('environment variables override daemon config file values', () => {
  withTempConfig(`
[addin_channel]
bind = "0.0.0.0"
port = 9443
shared_secret = "file-secret"

[mcp_http]
bind = "0.0.0.0"
port = 9900
api_key = "file-key"
`, (configPath) => {
    withEnv({
      OFFICE_MCP_ADDIN_HOST: 'localhost',
      OFFICE_MCP_ADDIN_PORT: '8765',
      OFFICE_MCP_ADDIN_SHARED_SECRET: '',
      OFFICE_MCP_MCP_HOST: '127.0.0.1',
      OFFICE_MCP_MCP_PORT: '8800',
      OFFICE_MCP_MCP_API_KEY: ''
    }, () => {
      const loaded = loadConfig({ configPath });

      assert.equal(loaded.addin.host, 'localhost');
      assert.equal(loaded.addin.port, 8765);
      assert.equal(loaded.addin.sharedSecret, '');
      assert.equal(loaded.mcp.host, '127.0.0.1');
      assert.equal(loaded.mcp.port, 8800);
      assert.equal(loaded.mcp.apiKey, '');
      assert.doesNotThrow(() => assertBoundaryAuthConfig(loaded));
    });
  });
});

test('section-style environment variables override legacy environment names', () => {
  withEnv({
    OFFICE_MCP_ADDIN_HOST: '127.0.0.1',
    OFFICE_MCP_ADDIN_CHANNEL__BIND: 'localhost',
    OFFICE_MCP_ADDIN_PORT: '9999',
    OFFICE_MCP_ADDIN_CHANNEL__PORT: '8766',
    OFFICE_MCP_ADDIN_PFX_PATH: 'legacy.pfx',
    OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH: 'section.pfx',
    OFFICE_MCP_ADDIN_PFX_PASSPHRASE: 'legacy-pass',
    OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE: 'section-pass',
    OFFICE_MCP_ADDIN_HEARTBEAT_INTERVAL_SEC: '99',
    OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_INTERVAL_SEC: '31',
    OFFICE_MCP_ADDIN_HEARTBEAT_TIMEOUT_SEC: '98',
    OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_TIMEOUT_SEC: '12',
    OFFICE_MCP_ADDIN_SESSION_GRACE_SEC: '97',
    OFFICE_MCP_ADDIN_CHANNEL__SESSION_GRACE_SEC: '61',
    OFFICE_MCP_ADDIN_MAX_PENDING_PER_SESSION: '96',
    OFFICE_MCP_ADDIN_CHANNEL__MAX_PENDING_PER_SESSION: '5',
    OFFICE_MCP_ADDIN_SHARED_SECRET: 'legacy-secret',
    OFFICE_MCP_ADDIN_CHANNEL__SHARED_SECRET: 'section-secret',
    OFFICE_MCP_MCP_HOST: '127.0.0.2',
    OFFICE_MCP_MCP_HTTP__BIND: '127.0.0.1',
    OFFICE_MCP_MCP_PORT: '9998',
    OFFICE_MCP_MCP_HTTP__PORT: '8801',
    OFFICE_MCP_MCP_API_KEY: 'legacy-key',
    OFFICE_MCP_MCP_HTTP__API_KEY: 'section-key',
    OFFICE_MCP_MAX_RESPONSE_BYTES: '9997',
    OFFICE_MCP_LIMITS__MAX_RESPONSE_BYTES: '1001',
    OFFICE_MCP_MAX_REQUEST_BYTES: '9996',
    OFFICE_MCP_LIMITS__MAX_REQUEST_BYTES: '1002',
    OFFICE_MCP_MAX_WS_FRAME_BYTES: '9995',
    OFFICE_MCP_LIMITS__MAX_WS_FRAME_BYTES: '1003',
    OFFICE_MCP_DEFAULT_TOOL_TIMEOUT_MS: '9994',
    OFFICE_MCP_LIMITS__DEFAULT_TOOL_TIMEOUT_MS: '1004',
    OFFICE_MCP_REQUESTS_PER_MINUTE: '9993',
    OFFICE_MCP_LIMITS__REQUESTS_PER_MINUTE: '1005',
    OFFICE_MCP_AUDIT_ENABLED: 'false',
    OFFICE_MCP_AUDIT__ENABLED: 'true',
    OFFICE_MCP_AUDIT_PATH: 'legacy-audit.jsonl',
    OFFICE_MCP_AUDIT__PATH: 'section-audit.jsonl',
    OFFICE_MCP_LOG_LEVEL: 'error',
    OFFICE_MCP_LOGGING__LEVEL: 'trace',
    OFFICE_MCP_LOG_FILE: 'legacy.log',
    OFFICE_MCP_LOGGING__FILE: 'section.log'
  }, () => {
    const loaded = loadConfig({ configPath: 'missing-config.toml' });

    assert.equal(loaded.addin.host, 'localhost');
    assert.equal(loaded.addin.port, 8766);
    assert.equal(loaded.addin.pfxPath, 'section.pfx');
    assert.equal(loaded.addin.pfxPassphrase, 'section-pass');
    assert.equal(loaded.addin.heartbeatIntervalSec, 31);
    assert.equal(loaded.addin.heartbeatTimeoutSec, 12);
    assert.equal(loaded.addin.sessionGraceSec, 61);
    assert.equal(loaded.addin.maxPendingPerSession, 5);
    assert.equal(loaded.addin.sharedSecret, 'section-secret');
    assert.equal(loaded.mcp.host, '127.0.0.1');
    assert.equal(loaded.mcp.port, 8801);
    assert.equal(loaded.mcp.apiKey, 'section-key');
    assert.equal(loaded.limits.maxResponseBytes, 1001);
    assert.equal(loaded.limits.maxRequestBytes, 1002);
    assert.equal(loaded.limits.maxWsFrameBytes, 1003);
    assert.equal(loaded.limits.defaultToolTimeoutMs, 1004);
    assert.equal(loaded.limits.requestsPerMinute, 1005);
    assert.equal(loaded.audit.enabled, true);
    assert.equal(loaded.audit.path, 'section-audit.jsonl');
    assert.equal(loaded.logging.level, 'trace');
    assert.equal(loaded.logging.file, 'section.log');
  });
});

test('refuses non-loopback add-in bind without a shared secret', () => {
  const testConfig = config();
  testConfig.addin.host = '0.0.0.0';

  assert.throws(() => assertBoundaryAuthConfig(testConfig), /OFFICE_MCP_ADDIN_SHARED_SECRET/);
});

test('refuses non-loopback MCP bind without an API key', () => {
  const testConfig = config();
  testConfig.mcp.host = '0.0.0.0';

  assert.throws(() => assertBoundaryAuthConfig(testConfig), /OFFICE_MCP_MCP_API_KEY/);
});

test('allows loopback binds without secrets', () => {
  assert.doesNotThrow(() => assertBoundaryAuthConfig(config()));
});

test('allows non-loopback binds when required secrets are configured', () => {
  const testConfig = config();
  testConfig.addin.host = '0.0.0.0';
  testConfig.addin.sharedSecret = 'addin-secret';
  testConfig.mcp.host = '0.0.0.0';
  testConfig.mcp.apiKey = 'mcp-secret';

  assert.doesNotThrow(() => assertBoundaryAuthConfig(testConfig));
});

function config(): DaemonConfig {
  return {
    addin: {
      host: 'localhost',
      port: 8765,
      origin: 'https://localhost:8765',
      pfxPath: '.office-mcp-localhost.pfx',
      pfxPassphrase: 'office-mcp-localhost',
      heartbeatIntervalSec: 30,
      heartbeatTimeoutSec: 10,
      sessionGraceSec: 60,
      maxPendingPerSession: 4,
      sharedSecret: ''
    },
    mcp: { host: '127.0.0.1', port: 8800, apiKey: '' },
    limits: { maxResponseBytes: 1024 * 1024, maxRequestBytes: 16 * 1024 * 1024, maxWsFrameBytes: 16 * 1024 * 1024, defaultToolTimeoutMs: 30000, requestsPerMinute: 120 },
    audit: { enabled: false, path: 'audit.jsonl' },
    logging: { level: 'info', file: '' }
  };
}

function withTempConfig(contents: string, callback: (path: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-config-'));
  try {
    const path = join(dir, 'config.toml');
    writeFileSync(path, contents, 'utf8');
    callback(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function withEnv(values: Record<string, string>, callback: () => void): void {
  const previous = new Map<string, string | undefined>();
  for (const [key, value] of Object.entries(values)) {
    previous.set(key, process.env[key]);
    process.env[key] = value;
  }
  try {
    callback();
  } finally {
    for (const [key, value] of previous) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}
