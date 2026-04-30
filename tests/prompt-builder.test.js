import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildGeneratePayload, buildRefinePayload } from '../public/prompt-builder.js';

test('buildGeneratePayload reply', () => {
  const state = {
    mode: 'reply', tone: 'casual', length: 'short', language: 'auto',
    useProQuality: false,
    notes: 'logo too small',
    customRules: '',
  };
  const email = { senderName: 'Sara', senderEmail: 's@s.com', subject: 'sub', body: 'body' };
  const payload = buildGeneratePayload(state, email, null);
  assert.equal(payload.mode, 'reply');
  assert.equal(payload.tone, 'casual');
  assert.equal(payload.reply.senderName, 'Sara');
  assert.equal(payload.reply.notes, 'logo too small');
  assert.equal(payload.reply.body, 'body');
});

test('buildGeneratePayload compose', () => {
  const state = {
    mode: 'compose', tone: 'professional', length: 'medium', language: 'english',
    useProQuality: true,
    customRules: 'Mention next steps.',
  };
  const compose = { to: 'Sara', topic: 'v4', notes: 'friday' };
  const payload = buildGeneratePayload(state, null, compose);
  assert.equal(payload.mode, 'compose');
  assert.equal(payload.useProQuality, true);
  assert.equal(payload.compose.topic, 'v4');
  assert.equal(payload.customRules, 'Mention next steps.');
});

test('buildRefinePayload', () => {
  const state = {
    tone: 'professional', length: 'medium', language: 'english',
    useProQuality: false, customRules: '',
  };
  const payload = buildRefinePayload(state, 'Hey Sara, BR,', 'more formal');
  assert.equal(payload.previousDraft, 'Hey Sara, BR,');
  assert.equal(payload.instruction, 'more formal');
  assert.equal(payload.tone, 'professional');
});
