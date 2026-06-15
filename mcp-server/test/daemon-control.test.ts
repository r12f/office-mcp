import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { runDaemonControl, runInstalledLauncherCommand, runScheduledTaskCommand, type ExecFileSync } from '../src/daemon-control.js';

test('daemon control uses the Windows Scheduled Task when it succeeds', () => {
  const calls: ExecCall[] = [];
  runDaemonControl('start', fakeExec(calls), 'win32');

  assert.equal(calls.length, 1);
  assert.equal(calls[0].file, 'powershell.exe');
  assert.match(calls[0].args.join(' '), /Start-ScheduledTask -TaskName 'office-mcp'/);
});

test('daemon control falls back to the MSI launcher when no Scheduled Task is available', () => {
  withInstallRoot((installRoot) => {
    withEnv({ OFFICE_MCP_INSTALL_ROOT: installRoot }, () => {
      const calls: ExecCall[] = [];
      runDaemonControl('start', fakeExec(calls, { failFirst: true }), 'win32');

      assert.equal(calls.length, 2);
      assert.match(calls[0].args.join(' '), /Start-ScheduledTask/);
      assert.match(calls[1].args.join(' '), /Start-Process/);
      assert.match(calls[1].args.join(' '), /office-mcp-daemon\.ps1/);
    });
  });
});

test('installed daemon stop targets the installed launcher process', () => {
  withInstallRoot((installRoot) => {
    const calls: ExecCall[] = [];
    runInstalledLauncherCommand('stop', installRoot, fakeExec(calls));

    assert.equal(calls.length, 1);
    assert.match(calls[0].args.join(' '), /Get-CimInstance Win32_Process/);
    assert.match(calls[0].args.join(' '), /office-mcp-daemon\.ps1/);
    assert.match(calls[0].args.join(' '), /Stop-Process/);
  });
});

test('scheduled task daemon stop uses Stop-ScheduledTask', () => {
  const calls: ExecCall[] = [];
  runScheduledTaskCommand('stop', fakeExec(calls));

  assert.equal(calls.length, 1);
  assert.match(calls[0].args.join(' '), /Stop-ScheduledTask -TaskName 'office-mcp'/);
});

type ExecCall = { file: string; args: string[] };

function fakeExec(calls: ExecCall[], options: { failFirst?: boolean } = {}): ExecFileSync {
  return ((file: string, args?: readonly string[]) => {
    calls.push({ file, args: [...(args ?? [])] });
    if (options.failFirst && calls.length === 1) throw new Error('task not found');
    return Buffer.from('');
  }) as ExecFileSync;
}

function withInstallRoot(callback: (installRoot: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-install-'));
  try {
    writeFileSync(join(dir, 'office-mcp-daemon.ps1'), '', 'utf8');
    callback(dir);
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
