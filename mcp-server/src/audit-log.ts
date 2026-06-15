import { mkdirSync, appendFileSync } from 'node:fs';
import { dirname } from 'node:path';
import type { DaemonConfig } from './config.js';

export type AuditRecord = {
  ts: string;
  tool: string;
  session_id?: string;
  duration_ms: number;
  ok: boolean;
  error_code?: string;
  error_message?: string;
};

export function writeAuditRecord(config: DaemonConfig, record: AuditRecord): void {
  if (!config.audit.enabled) return;
  mkdirSync(dirname(config.audit.path), { recursive: true });
  appendFileSync(config.audit.path, `${JSON.stringify(record)}\n`, 'utf8');
}
