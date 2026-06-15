import type { WebSocket } from 'ws';

export const SERVER_VERSION = '0.1.0';
export const ADDIN_PROTOCOL_VERSION = '1.0';

export type OfficeMcpCode =
  | 'GENERIC_FAILURE'
  | 'NO_SESSIONS'
  | 'SESSION_LOST'
  | 'SESSION_NOT_FOUND'
  | 'SESSION_STALE'
  | 'MAX_PENDING_EXCEEDED'
  | 'AUTH_FAILED'
  | 'IRM_DENIED'
  | 'DOCUMENT_READ_ONLY'
  | 'PROTECTION_BLOCKS'
  | 'ANCHOR_NOT_FOUND'
  | 'NO_MATCHES'
  | 'INDEX_OUT_OF_RANGE'
  | 'INVALID_ARGUMENT'
  | 'PATH_REFUSED'
  | 'IMAGE_FETCH_FAILED'
  | 'HOST_CAPABILITY_UNAVAILABLE'
  | 'CANCELLED'
  | 'TIMEOUT'
  | 'MAX_RESPONSE_SIZE'
  | 'STALE_INDEX'
  | 'PROTOCOL_VERSION_MISMATCH'
  | 'HEARTBEAT_MISSED'
  | 'INTERNAL_BUG';

export type JsonRpcId = string | number | null;

export type JsonRpcRequest = {
  jsonrpc: '2.0';
  id?: JsonRpcId;
  method: string;
  params?: unknown;
};

export type JsonRpcResponse = {
  jsonrpc: '2.0';
  id: JsonRpcId;
  result?: unknown;
  error?: {
    code: number;
    message: string;
    data?: unknown;
  };
};

export type DocumentInfo = {
  title?: string | null;
  url?: string | null;
  filename?: string | null;
  is_dirty?: boolean | null;
  is_read_only?: boolean | null;
  is_protected?: boolean | null;
  protection?: {
    kind?: string | null;
    rights?: string[] | null;
    rights_source?: string | null;
  };
  opened_at?: string | null;
};

export type HostInfo = {
  app: string;
  version?: string | null;
  platform?: string | null;
  build?: string | null;
};

export type AddInInfo = {
  version: string;
  protocol_version: string;
  requirement_sets?: Record<string, string | null>;
  supported_features?: string[];
};

export type RuntimeInfo = {
  instance_id: string;
  host: HostInfo;
  add_in: AddInInfo;
  registered_at: string;
};

export type SessionInfo = {
  session_id: string;
  instance_id: string;
  document: DocumentInfo;
  available_tools: string[];
  is_active?: boolean | null;
  status: 'active' | 'stale';
  registered_at: string;
  stale_since?: string;
};

export type SessionDescriptor = {
  session_id: string;
  instance_id: string;
  app: string;
  document: {
    title: string | null;
    url: string | null;
    filename: string | null;
    is_dirty: boolean | null;
    is_protected: boolean | null;
    protection_kind: string | null;
    rights: string[] | null;
    rights_source: string | null;
  };
  is_active: boolean | null;
  capability_tiers: string[];
  available_tool_count: number;
  registered_at: string;
  status: 'active' | 'stale';
};

export type AddinConnection = {
  socket: WebSocket;
  runtime: RuntimeInfo;
  session?: SessionInfo;
  heartbeat?: {
    interval: NodeJS.Timeout;
    timeout?: NodeJS.Timeout;
    pendingPingId?: string;
    missedPongs: number;
  };
  pending: Map<string, {
    resolve: (value: AddinToolResult) => void;
    reject: (error: Error) => void;
    timeout: NodeJS.Timeout;
    sessionId: string;
    tool: string;
  }>;
  queue: Promise<unknown>;
  invokeTool: (sessionId: string, tool: string, args: Record<string, unknown>, timeoutMs: number) => Promise<AddinToolResult>;
};

export type ToolFailure = {
  office_mcp_code: OfficeMcpCode;
  message: string;
  session_id?: string;
  tool?: string;
  retriable: boolean;
  partial_effect?: 'none' | 'possible' | 'unknown';
  max_response_bytes?: number;
};

export type ToolSuccess = {
  ok: true;
  data: unknown;
  elapsed_ms?: number;
};

export type ToolErrorResult = {
  ok: false;
  error: ToolFailure;
  elapsed_ms?: number;
};

export type AddinToolResult = ToolSuccess | ToolErrorResult;


