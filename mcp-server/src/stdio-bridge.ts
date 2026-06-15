import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';

export type StdioBridge = {
  close: () => Promise<void>;
};

export async function startStdioBridge(endpoint: URL): Promise<StdioBridge> {
  const stdio = new StdioServerTransport();
  const http = new StreamableHTTPClientTransport(endpoint);

  stdio.onmessage = (message: JSONRPCMessage) => {
    http.send(message).catch((error) => stdio.onerror?.(asError(error)));
  };
  http.onmessage = (message: JSONRPCMessage) => {
    stdio.send(message).catch((error) => http.onerror?.(asError(error)));
  };
  stdio.onerror = (error) => {
    console.error(`office-mcp stdio error: ${error.message}`);
  };
  http.onerror = (error) => {
    console.error(`office-mcp daemon transport error: ${error.message}`);
  };
  stdio.onclose = () => {
    http.close().catch(() => undefined);
  };
  http.onclose = () => {
    stdio.close().catch(() => undefined);
  };

  await http.start();
  await stdio.start();

  return {
    close: async () => {
      await Promise.allSettled([stdio.close(), http.close()]);
    }
  };
}

function asError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}
