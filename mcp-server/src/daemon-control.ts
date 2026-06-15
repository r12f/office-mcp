import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

export type DaemonControlAction = 'start' | 'stop';
export type ExecFileSync = typeof execFileSync;

export function runDaemonControl(action: DaemonControlAction, exec: ExecFileSync = execFileSync, platform = process.platform): void {
  if (platform !== 'win32') {
    throw new Error(`daemon ${action} is currently implemented for Windows only.`);
  }

  try {
    runScheduledTaskCommand(action, exec);
    return;
  } catch (error) {
    if (!process.env.OFFICE_MCP_INSTALL_ROOT) throw error;
  }

  runInstalledLauncherCommand(action, process.env.OFFICE_MCP_INSTALL_ROOT, exec);
}

export function runScheduledTaskCommand(action: DaemonControlAction, exec: ExecFileSync = execFileSync): void {
  const taskName = process.env.OFFICE_MCP_TASK_NAME ?? 'office-mcp';
  const commandName = action === 'start' ? 'Start-ScheduledTask' : 'Stop-ScheduledTask';
  exec('powershell.exe', ['-NoProfile', '-Command', `${commandName} -TaskName '${taskName}'`], { stdio: 'inherit' });
}

export function runInstalledLauncherCommand(action: DaemonControlAction, installRoot: string, exec: ExecFileSync = execFileSync): void {
  const launcherPath = join(installRoot, 'office-mcp-daemon.ps1');
  if (!existsSync(launcherPath)) {
    throw new Error(`Cannot find installed daemon launcher: ${launcherPath}`);
  }

  if (action === 'start') {
    exec('powershell.exe', [
      '-NoProfile',
      '-Command',
      `Start-Process -WindowStyle Hidden powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File','${escapePowerShellSingleQuoted(launcherPath)}')`
    ], { stdio: 'inherit' });
    return;
  }

  exec('powershell.exe', [
    '-NoProfile',
    '-Command',
    `$launcher='${escapePowerShellSingleQuoted(launcherPath)}'; Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like "*$launcher*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }`
  ], { stdio: 'inherit' });
}

function escapePowerShellSingleQuoted(value: string): string {
  return value.replace(/'/g, "''");
}
