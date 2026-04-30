import { buildPrompt, callGemini } from './_gemini.js';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'method_not_allowed' });

  const body = req.body || {};
  if (!body.previousDraft || !body.previousDraft.trim()) {
    return res.status(400).json({ ok: false, error: 'missing_previous_draft' });
  }
  if (!body.instruction || !body.instruction.trim()) {
    return res.status(400).json({ ok: false, error: 'missing_instruction' });
  }

  try {
    const prompt = buildPrompt({
      mode: 'refine',
      tone: body.tone || 'professional',
      length: body.length || 'medium',
      language: body.language || 'english',
      customRules: body.customRules || '',
      refine: {
        previousDraft: body.previousDraft,
        instruction: body.instruction,
      },
    });
    const { text, model } = await callGemini({ prompt, useProQuality: body.useProQuality === true });
    return res.status(200).json({
      ok: true,
      draft: text.trim(),
      subject: null,
      missingInfo: null,
      model,
    });
  } catch (err) {
    const status = err.status || (err.message?.includes('quota') ? 429 : 500);
    return res.status(status).json({ ok: false, error: err.message || 'unknown_error' });
  }
}
