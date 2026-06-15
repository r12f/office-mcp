import { join } from 'node:path';

export type ClaudeDesktopMode = 'dev' | 'installed';

export type ClaudeDesktopConfig = {
  mcpServers: {
    'office-mcp': {
      command: string;
      args: string[];
      cwd?: string;
    };
  };
};

export function buildClaudeDesktopConfig(options: {
  mode: ClaudeDesktopMode;
  cwd?: string;
  installRoot?: string;
}): ClaudeDesktopConfig {
  if (options.mode === 'installed') {
    const installRoot = options.installRoot ?? defaultWindowsInstallRoot();
    return {
      mcpServers: {
        'office-mcp': {
          command: 'powershell.exe',
          args: ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(installRoot, 'office-mcp.ps1'), 'stdio']
        }
      }
    };
  }

  return {
    mcpServers: {
      'office-mcp': {
        command: 'node',
        args: ['dist/src/cli.js', 'stdio'],
        cwd: options.cwd ?? process.cwd()
      }
    }
  };
}

function defaultWindowsInstallRoot(): string {
  return process.env.OFFICE_MCP_INSTALL_ROOT
    ?? join(process.env.LOCALAPPDATA ?? join(process.env.USERPROFILE ?? 'C:\\Users\\Default', 'AppData', 'Local'), 'office-mcp');
}
