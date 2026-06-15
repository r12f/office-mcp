import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import type { DaemonConfig, LogLevel } from '../src/config.js';
import { createLogger } from '../src/logger.js';

test('logger writes records at or above the configured level', () => {
  withLogFile((logPath) => {
    const logger = createLogger(config('warn', logPath));

    withMutedConsole(() => {
      logger.info('hidden info');
      logger.warn('visible warning', { event: 'warn_event' });
      logger.error('visible error', { event: 'error_event' });
    });

    const records = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line) as Record<string, unknown>);
    assert.equal(records.length, 2);
    assert.equal(records[0].level, 'warn');
    assert.equal(records[0].message, 'visible warning');
    assert.equal(records[0].event, 'warn_event');
    assert.equal(records[1].level, 'error');
    assert.equal(records[1].message, 'visible error');
  });
});

function withLogFile(callback: (path: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-log-'));
  try {
    callback(join(dir, 'office-mcp.log'));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function withMutedConsole(callback: () => void): void {
  const previousWarn = console.warn;
  const previousError = console.error;
  console.warn = () => undefined;
  console.error = () => undefined;
  try {
    callback();
  } finally {
    console.warn = previousWarn;
    console.error = previousError;
  }
}

function config(level: LogLevel, file: string): DaemonConfig {
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
    logging: { level, file }
  };
}
