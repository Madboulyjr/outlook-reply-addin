import { test } from 'node:test';
import assert from 'node:assert/strict';
import { buildPrompt, SYSTEM_PROMPT } from '../api/_gemini.js';

test('SYSTEM_PROMPT mentions Ali Madbouly and Nice One', () => {
  assert.match(SYSTEM_PROMPT, /Ali Madbouly/);
  assert.match(SYSTEM_PROMPT, /Nice One/);
  assert.match(SYSTEM_PROMPT, /BR,/);
});

test('buildPrompt for reply mode includes email body and notes', () => {
  const out = buildPrompt({
    mode: 'reply',
    tone: 'professional',
    length: 'medium',
    language: 'auto',
    reply: {
      senderName: 'Sara Ahmed',
      senderEmail: 'sara@example.com',
      subject: 'Launch deck v3',
      body: 'Take a look and let me know.',
      notes: 'tell her v3 looks good but logo too small',
    },
  });
  assert.match(out, /Sara Ahmed/);
  assert.match(out, /Launch deck v3/);
  assert.match(out, /Take a look and let me know\./);
  assert.match(out, /logo too small/);
  assert.match(out, /professional/);
});

test('buildPrompt for compose mode includes recipient and topic', () => {
  const out = buildPrompt({
    mode: 'compose',
    tone: 'casual',
    length: 'short',
    language: 'english',
    compose: {
      to: 'Sara Ahmed',
      topic: 'Launch deck v4 timeline',
      notes: 'need final files Friday',
    },
  });
  assert.match(out, /Sara Ahmed/);
  assert.match(out, /Launch deck v4 timeline/);
  assert.match(out, /final files Friday/);
});

test('buildPrompt for refine includes previous draft and instruction', () => {
  const out = buildPrompt({
    mode: 'refine',
    tone: 'professional',
    length: 'medium',
    language: 'english',
    refine: {
      previousDraft: 'Hey Sara, looks good. BR,',
      instruction: 'make it more formal',
    },
  });
  assert.match(out, /Hey Sara, looks good\. BR,/);
  assert.match(out, /more formal/);
});

test('buildPrompt appends customRules when provided', () => {
  const out = buildPrompt({
    mode: 'reply',
    tone: 'professional',
    length: 'short',
    language: 'english',
    customRules: 'Always mention next steps.',
    reply: { senderName: 'X', senderEmail: 'x@x.com', subject: 's', body: 'b', notes: '' },
  });
  assert.match(out, /Always mention next steps\./);
});
