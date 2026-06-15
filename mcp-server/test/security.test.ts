import assert from 'node:assert/strict';
import test from 'node:test';
import { constantTimeEquals } from '../src/security.js';

test('constantTimeEquals matches equal secrets', () => {
  assert.equal(constantTimeEquals('secret', 'secret'), true);
});

test('constantTimeEquals rejects unequal and differently sized secrets', () => {
  assert.equal(constantTimeEquals('secret', 'SECRET'), false);
  assert.equal(constantTimeEquals('secret', 's'), false);
});
