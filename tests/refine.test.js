import { test, mock } from 'node:test';
import assert from 'node:assert/strict';

mock.module('../api/_gemini.js', {
  namedExports: {
    SYSTEM_PROMPT: 'stub',
    buildPrompt: (p) => `REFINE_PROMPT(${p.refine.instruction})`,
    callGemini: async ({ prompt }) => ({ text: `Refined: ${prompt}`, model: 'gemini-2.5-flash' }),
  },
});

const { default: handler } = await import('../api/refine.js');

function makeReqRes(body, method = 'POST') {
  const req = { method, body, headers: {} };
  let statusCode = 200;
  let payload = null;
  const res = {
    status(c) { statusCode = c; return res; },
    json(p) { payload = p; return res; },
    setHeader() { return res; },
    end() { return res; },
  };
  return { req, res, get statusCode() { return statusCode; }, get payload() { return payload; } };
}

test('refine returns refined draft', async () => {
  const ctx = makeReqRes({
    previousDraft: 'Hey Sara, looks good. BR,',
    instruction: 'make it more formal',
    tone: 'professional',
    length: 'medium',
    language: 'english',
  });
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 200);
  assert.equal(ctx.payload.ok, true);
  assert.match(ctx.payload.draft, /Refined:/);
  assert.match(ctx.payload.draft, /make it more formal/);
});

test('refine rejects empty previousDraft', async () => {
  const ctx = makeReqRes({
    previousDraft: '',
    instruction: 'shorter',
    tone: 'professional', length: 'medium', language: 'english',
  });
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 400);
});

test('refine rejects empty instruction', async () => {
  const ctx = makeReqRes({
    previousDraft: 'something',
    instruction: '',
    tone: 'professional', length: 'medium', language: 'english',
  });
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 400);
});
