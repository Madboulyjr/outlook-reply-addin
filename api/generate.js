import { buildPrompt, callGemini } from './_gemini.js';

const VALID_MODES = new Set(['reply', 'compose']);
const VALID_TONES = new Set(['very_formal', 'professional', 'casual', 'very_casual']);
const VALID_LENGTHS = new Set(['short', 'medium', 'detailed']);
const VALID_LANGUAGES = new Set(['english', 'arabic', 'auto']);

function parseComposeOutput(text) {
  // Try direct JSON parse, then strip code fences if present.
  const trimmed = text.trim();
  const candidates = [trimmed];
  const fenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
  if (fenceMatch) candidates.push(fenceMatch[1]);
  for (const c of candidates) {
    try {
      const obj = JSON.parse(c);
      if (typeof obj.subject === 'string' && typeof obj.body === 'string') return obj;
    } catch { /* try next */ }
  }
  // Fallback — first line as subject, rest as body.
  const lines = trimmed.split('\n');
  const firstNonEmpty = lines.findIndex(l => l.trim());
  return {
    subject: lines[firstNonEmpty] ? lines[firstNonEmpty].replace(/^Subject:\s*/i, '').trim() : '(no subject)',
    body: lines.slice(firstNonEmpty + 1).join('\n').trim() || trimmed,
  };
}

function extractMissingInfo(text) {
  const m = text.match(/⚠️\s*Missing:\s*(.+)$/m);
  return m ? m[1].trim() : null;
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'method_not_allowed' });

  const body = req.body || {};
  if (!VALID_MODES.has(body.mode)) {
    return res.status(400).json({ ok: false, error: 'invalid_mode' });
  }
  if (!VALID_TONES.has(body.tone)) {
    return res.status(400).json({ ok: false, error: 'invalid_tone' });
  }
  if (!VALID_LENGTHS.has(body.length)) {
    return res.status(400).json({ ok: false, error: 'invalid_length' });
  }
  if (!VALID_LANGUAGES.has(body.language)) {
    return res.status(400).json({ ok: false, error: 'invalid_language' });
  }
  if (body.mode === 'reply' && (!body.reply || typeof body.reply !== 'object')) {
    return res.status(400).json({ ok: false, error: 'missing_reply_payload' });
  }
  if (body.mode === 'compose' && (!body.compose || typeof body.compose !== 'object')) {
    return res.status(400).json({ ok: false, error: 'missing_compose_payload' });
  }

  try {
    const prompt = buildPrompt(body);
    const { text, model } = await callGemini({ prompt, useProQuality: body.useProQuality === true });

    if (body.mode === 'compose') {
      const parsed = parseComposeOutput(text);
      return res.status(200).json({
        ok: true,
        draft: parsed.body,
        subject: parsed.subject,
        missingInfo: extractMissingInfo(parsed.body),
        model,
      });
    }

    return res.status(200).json({
      ok: true,
      draft: text.trim(),
      subject: null,
      missingInfo: extractMissingInfo(text),
      model,
    });
  } catch (err) {
    const status = err.status || (err.message?.includes('quota') ? 429 : 500);
    return res.status(status).json({ ok: false, error: err.message || 'unknown_error' });
  }
}
