import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/agent-client-evidence.json'));
const prompt = requiredOption('--prompt');
const expectedSubstring = requiredOption('--expected-substring');
const observedAnswer = requiredOption('--observed-answer');
const documentTitle = readOption('--document-title');
const sessionId = readOption('--session-id');
const screenshotPath = readOption('--screenshot-path');
const tester = readOption('--tester') ?? process.env.USERNAME ?? process.env.USER ?? 'unknown';
const passed = observedAnswer.toLowerCase().includes(expectedSubstring.toLowerCase());

const evidence = {
  schema_version: 1,
  kind: 'agent_client_prompt',
  recorded_at: new Date().toISOString(),
  tester,
  prompt,
  expected_substring: expectedSubstring,
  observed_answer: observedAnswer,
  document_title: documentTitle,
  session_id: sessionId,
  screenshot_path: screenshotPath,
  screenshot_exists: screenshotPath ? existsSync(screenshotPath) : undefined,
  passed
};

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, JSON.stringify(evidence, null, 2));
console.log(JSON.stringify(evidence, null, 2));
if (!passed) process.exit(1);

function requiredOption(name: string): string {
  const value = readOption(name);
  if (!value) throw new Error(`${name} is required`);
  return value;
}

function readOption(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}
