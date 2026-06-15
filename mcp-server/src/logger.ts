import { appendFileSync, mkdirSync } from 'node:fs';
import { dirname } from 'node:path';
import type { DaemonConfig, LogLevel } from './config.js';

const LEVELS: Record<LogLevel, number> = { trace: 10, debug: 20, info: 30, warn: 40, error: 50 };

export type Logger = {
  trace: (message: string, fields?: Record<string, unknown>) => void;
  debug: (message: string, fields?: Record<string, unknown>) => void;
  info: (message: string, fields?: Record<string, unknown>) => void;
  warn: (message: string, fields?: Record<string, unknown>) => void;
  error: (message: string, fields?: Record<string, unknown>) => void;
};

export function createLogger(config: DaemonConfig): Logger {
  const write = (level: LogLevel, message: string, fields: Record<string, unknown> = {}) => {
    if (LEVELS[level] < LEVELS[config.logging.level]) return;
    const record = { ts: new Date().toISOString(), level, message, ...fields };
    const line = JSON.stringify(record);
    if (config.logging.file) {
      mkdirSync(dirname(config.logging.file), { recursive: true });
      appendFileSync(config.logging.file, `${line}\n`, 'utf8');
    }
    if (level === 'error') console.error(message);
    else if (level === 'warn') console.warn(message);
    else console.log(message);
  };

  return {
    trace: (message, fields) => write('trace', message, fields),
    debug: (message, fields) => write('debug', message, fields),
    info: (message, fields) => write('info', message, fields),
    warn: (message, fields) => write('warn', message, fields),
    error: (message, fields) => write('error', message, fields)
  };
}
