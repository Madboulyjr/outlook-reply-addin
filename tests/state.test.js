import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createInitialState, applyAction } from '../public/state.js';

test('initial state has defaults', () => {
  const s = createInitialState();
  assert.equal(s.mode, 'reply');
  assert.equal(s.tone, 'professional');
  assert.equal(s.length, 'medium');
  assert.equal(s.language, 'auto');
  assert.equal(s.useProQuality, false);
  assert.equal(s.draft, null);
  assert.equal(s.loading, false);
  assert.equal(s.error, null);
});

test('SET_MODE switches mode and clears draft', () => {
  const s = applyAction(
    { ...createInitialState(), draft: 'old' },
    { type: 'SET_MODE', mode: 'compose' }
  );
  assert.equal(s.mode, 'compose');
  assert.equal(s.draft, null);
});

test('GENERATE_START sets loading and clears error', () => {
  const s = applyAction(
    { ...createInitialState(), error: 'old', loading: false },
    { type: 'GENERATE_START' }
  );
  assert.equal(s.loading, true);
  assert.equal(s.error, null);
});

test('GENERATE_SUCCESS sets draft and clears loading', () => {
  const s = applyAction(
    { ...createInitialState(), loading: true },
    { type: 'GENERATE_SUCCESS', draft: 'Hey Sara,\n\nBR,', subject: null, missingInfo: null }
  );
  assert.equal(s.draft, 'Hey Sara,\n\nBR,');
  assert.equal(s.loading, false);
});

test('GENERATE_FAIL sets error and clears loading', () => {
  const s = applyAction(
    { ...createInitialState(), loading: true },
    { type: 'GENERATE_FAIL', error: 'rate limited' }
  );
  assert.equal(s.error, 'rate limited');
  assert.equal(s.loading, false);
});

test('SET_TONE updates tone only', () => {
  const s = applyAction(createInitialState(), { type: 'SET_TONE', tone: 'casual' });
  assert.equal(s.tone, 'casual');
});
