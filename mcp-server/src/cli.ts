import { loadConfig } from './config.js';
import { startDaemon } from './daemon.js';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { connect } from 'node:net';
import { startStdioBridge } from './stdio-bridge.js';
import { buildClaudeDesktopConfig, type ClaudeDesktopMode } from './client-config.js';
import { runDaemonControl } from './daemon-control.js';

const command = process.argv.slice(2);

if (command[0] === 'daemon' && command[1] === 'run') {
  const daemon = await startDaemon(loadConfig());
  const shutdown = async () => {
    await daemon.close();
    process.exit(0);
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
} else if (command[0] === 'daemon' && command[1] === 'status') {
  const config = loadConfig();
  const [mcpListening, addinListening] = await Promise.all([
    isPortListening(config.mcp.host, config.mcp.port),
    isPortListening(config.addin.host, config.addin.port)
  ]);
  const running = mcpListening && addinListening;
  console.log(JSON.stringify({
    running,
    mcp: { host: config.mcp.host, port: config.mcp.port, listening: mcpListening },
    addin: { host: config.addin.host, port: config.addin.port, origin: config.addin.origin, listening: addinListening }
  }, null, 2));
  process.exit(running ? 0 : 1);
} else if (command[0] === 'daemon' && (command[1] === 'start' || command[1] === 'stop')) {
  runDaemonControl(command[1]);
} else if (command[0] === 'stdio') {
  const config = loadConfig();
  await startStdioBridge(new URL(`http://${config.mcp.host}:${config.mcp.port}/mcp`));
} else if (command[0] === 'config' && (command[1] === 'endpoints' || command[1] === 'show')) {
  const config = loadConfig();
  const endpoints = {
    mcp: `http://${config.mcp.host}:${config.mcp.port}/mcp`,
    addin_origin: config.addin.origin,
    addin_wss: config.addin.origin.replace('https://', 'wss://') + '/addin'
  };
  console.log(JSON.stringify(command[1] === 'show' ? { ...redactConfig(config), endpoints } : endpoints, null, 2));
} else if (command[0] === 'config' && command[1] === 'claude-desktop') {
  const mode: ClaudeDesktopMode = command.includes('--installed') ? 'installed' : 'dev';
  const installRoot = readOption(command, '--install-root');
  console.log(JSON.stringify(buildClaudeDesktopConfig({ mode, cwd: process.cwd(), installRoot }), null, 2));
} else if (command[0] === 'sessions') {
  const config = loadConfig();
  const client = new Client({ name: 'office-mcp-cli', version: '0.1.0' });
  const transport = new StreamableHTTPClientTransport(new URL(`http://${config.mcp.host}:${config.mcp.port}/mcp`));
  await client.connect(transport);
  try {
    const result = await client.callTool({ name: 'office.list_sessions', arguments: {} });
    console.log(JSON.stringify(result.structuredContent ?? result, null, 2));
  } finally {
    await client.close();
  }
} else {
  console.log('Usage:');
  console.log('  office-mcp daemon run');
  console.log('  office-mcp daemon status');
  console.log('  office-mcp daemon start|stop   # Windows autostart integration');
  console.log('  office-mcp stdio               # bridge stdio-only MCP clients to the daemon');
  console.log('  office-mcp config endpoints');
  console.log('  office-mcp config show');
  console.log('  office-mcp config claude-desktop [--installed] [--install-root <path>]');
  console.log('  office-mcp sessions');
  process.exit(command.length === 0 ? 0 : 1);
}

function readOption(args: string[], name: string): string | undefined {
  const index = args.indexOf(name);
  if (index === -1) return undefined;
  const value = args[index + 1];
  if (!value || value.startsWith('--')) {
    throw new Error(`${name} requires a value`);
  }
  return value;
}

function redactConfig(config: ReturnType<typeof loadConfig>): ReturnType<typeof loadConfig> {
  return {
    ...config,
    addin: {
      ...config.addin,
      pfxPassphrase: config.addin.pfxPassphrase ? '<redacted>' : '',
      sharedSecret: config.addin.sharedSecret ? '<redacted>' : ''
    },
    mcp: {
      ...config.mcp,
      apiKey: config.mcp.apiKey ? '<redacted>' : ''
    }
  };
}

function isPortListening(host: string, port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const socket = connect({ host, port });
    socket.setTimeout(1000);
    socket.once('connect', () => {
      socket.destroy();
      resolve(true);
    });
    socket.once('timeout', () => {
      socket.destroy();
      resolve(false);
    });
    socket.once('error', () => resolve(false));
  });
}
