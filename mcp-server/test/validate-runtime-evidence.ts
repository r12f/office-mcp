import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

type GateStatus = 'passed' | 'failed' | 'skipped' | 'blocked_by_runtime';

type EvidenceGate = {
  name: string;
  status: GateStatus;
  details?: Record<string, unknown>;
};

type EvidenceReport = {
  schema_version: number;
  generated_at: string;
  endpoint: string;
  session_id?: string;
  gates: EvidenceGate[];
};

const evidencePath = resolve(readOption('--input') ?? '../artifacts/runtime-evidence.json');
const requireIrm = hasFlag('--require-irm');
const requireFullWordSmoke = hasFlag('--require-full-word-smoke');
const requireComTrackedChanges = hasFlag('--require-com-tracked-changes');
const requireIrmPreflight = hasFlag('--require-irm-preflight');
const requireClaudeDesktopInstallation = hasFlag('--require-claude-desktop-installation');
const requireAgentClientPrompt = hasFlag('--require-agent-client-prompt');
const requireMutation = hasFlag('--require-mutation');
const report = JSON.parse(readFileSync(evidencePath, 'utf8')) as EvidenceReport;

const requiredPassed = [
  'word.session_discovery',
  'word.runtime_read_smoke',
  'agent_client_stdio_bridge'
];

const fullWordSmokeGates = [
  'word.full_smoke.word-core',
  'word.full_smoke.word-formatting',
  'word.full_smoke.word-review',
  'word.full_smoke.word-resources',
  'word.full_smoke.word-spec-args'
];

const comTrackedChangeGates = [
  'word.tracked_change_com.accept',
  'word.tracked_change_com.reject'
];

const failures: string[] = [];

if (report.schema_version !== 1) failures.push(`Unsupported schema_version: ${report.schema_version}`);
if (!report.session_id) failures.push('Missing session_id.');

for (const name of requiredPassed) {
  requirePassedGate(name);
}

if (requireFullWordSmoke) {
  for (const name of fullWordSmokeGates) requirePassedGate(name);
}

if (requireMutation) requirePassedGate('word.runtime_mutation_smoke');

if (requireComTrackedChanges) {
  for (const name of comTrackedChangeGates) requirePassedGate(name);
}

if (requireIrmPreflight) requirePassedGate('irm_document_preflight');
if (requireClaudeDesktopInstallation) {
  const gate = requirePassedGate('claude_desktop_installation');
  if (gate && gate.details?.ui_validation_ready !== true) {
    failures.push('Claude Desktop installation gate is not UI-validation ready.');
  }
}

if (requireAgentClientPrompt) requirePassedGate('agent_client_prompt');

const irmGate = gateByName('irm_rights_matrix');
if (requireIrm) {
  if (!irmGate) {
    failures.push('Missing required IRM gate: irm_rights_matrix');
  } else if (irmGate.status !== 'passed') {
    failures.push(`IRM gate is ${irmGate.status}, expected passed.`);
  }
}

const failedGate = report.gates.find((gate) => gate.status === 'failed' || gate.status === 'blocked_by_runtime');
if (failedGate) failures.push(`Gate ${failedGate.name} is ${failedGate.status}.`);

const summary = {
  ok: failures.length === 0,
  require_irm: requireIrm,
  require_full_word_smoke: requireFullWordSmoke,
  require_com_tracked_changes: requireComTrackedChanges,
  require_irm_preflight: requireIrmPreflight,
  require_claude_desktop_installation: requireClaudeDesktopInstallation,
  require_agent_client_prompt: requireAgentClientPrompt,
  require_mutation: requireMutation,
  generated_at: report.generated_at,
  endpoint: report.endpoint,
  session_id: report.session_id,
  gates: Object.fromEntries(report.gates.map((gate) => [gate.name, gate.status])),
  failures
};

console.log(JSON.stringify(summary, null, 2));
if (failures.length > 0) process.exit(1);

function gateByName(name: string): EvidenceGate | undefined {
  return report.gates.find((gate) => gate.name === name);
}

function requirePassedGate(name: string): EvidenceGate | undefined {
  const gate = gateByName(name);
  if (!gate) {
    failures.push(`Missing required gate: ${name}`);
  } else if (gate.status !== 'passed') {
    failures.push(`Gate ${name} is ${gate.status}, expected passed.`);
  }
  return gate;
}

function readOption(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}

function hasFlag(name: string): boolean {
  return process.argv.includes(name);
}

