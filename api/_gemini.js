import { GoogleGenerativeAI } from '@google/generative-ai';

export const SYSTEM_PROMPT = `You are drafting email replies and new emails for Ali Madbouly,
Head of Art / Art Director at Nice One (https://niceonesa.com/), a beauty &
lifestyle e-commerce platform in Saudi Arabia. He leads the in-house brand
studio handling all visual work for the Nice One brand itself.

His emails typically involve: design feedback & critique, asset approvals,
vendor & agency communication, scheduling, brand-work coordination,
internal team comms.

VOICE RULES (always):
- Write naturally — like a real person, not a corporate template.
- No filler ("I hope this email finds you well", "Please don't hesitate
  to reach out") unless the situation truly calls for it.
- Get to the point fast.
- Sign off only with: BR,
  (His email client adds the full signature — never write
   "Best regards, Ali Madbouly" or his name underneath.)
- Match the sender's energy: warm reply to warm email, sharp to sharp.
- If the incoming email is in Arabic and language preference is "auto",
  reply in Arabic.

OUTPUT RULES:
- Reply / refine mode → only the body text. No subject. End with "BR,"
- Compose mode → return JSON: {"subject":"...","body":"..."}
- Never include any commentary outside the email itself.
- If information is missing (a date, price, name), insert [PLACEHOLDER]
  inline and add a single line at the end:
  "⚠️ Missing: [what's needed]"
`;

export function buildPrompt(payload) {
  const { mode, tone, length, language, customRules } = payload;
  const lines = [];
  lines.push(`MODE: ${mode}`);
  lines.push(`TONE: ${tone}`);
  lines.push(`LENGTH: ${length}`);
  lines.push(`LANGUAGE: ${language}`);
  if (customRules && customRules.trim()) {
    lines.push('');
    lines.push('ADDITIONAL RULES FROM USER:');
    lines.push(customRules.trim());
  }
  lines.push('');

  if (mode === 'reply' && payload.reply) {
    const r = payload.reply;
    lines.push(`INCOMING EMAIL`);
    lines.push(`From: ${r.senderName} <${r.senderEmail}>`);
    lines.push(`Subject: ${r.subject}`);
    lines.push(``);
    lines.push(r.body);
    lines.push(``);
    lines.push(`USER NOTES (what to say in reply):`);
    lines.push(r.notes || '(no specific notes — write a sensible reply)');
  } else if (mode === 'compose' && payload.compose) {
    const c = payload.compose;
    lines.push(`COMPOSE NEW EMAIL`);
    lines.push(`To: ${c.to}`);
    lines.push(`Topic: ${c.topic}`);
    lines.push(``);
    lines.push(`USER NOTES:`);
    lines.push(c.notes || '');
    lines.push(``);
    lines.push('Return ONLY a JSON object: {"subject":"...","body":"..."} — no markdown fences, no extra text.');
  } else if (mode === 'refine' && payload.refine) {
    const f = payload.refine;
    lines.push(`PREVIOUS DRAFT:`);
    lines.push(f.previousDraft);
    lines.push(``);
    lines.push(`REFINEMENT INSTRUCTION: ${f.instruction}`);
    lines.push(``);
    lines.push('Return only the rewritten body text, no commentary.');
  }

  return lines.join('\n');
}

export async function callGemini({ prompt, useProQuality = false, apiKey }) {
  const key = apiKey || process.env.GEMINI_API_KEY;
  if (!key) throw new Error('GEMINI_API_KEY not set');
  const genAI = new GoogleGenerativeAI(key);
  const modelName = useProQuality ? 'gemini-2.5-pro' : 'gemini-2.5-flash';
  const model = genAI.getGenerativeModel({
    model: modelName,
    systemInstruction: SYSTEM_PROMPT,
  });
  const result = await model.generateContent(prompt);
  const text = result.response.text();
  return { text, model: modelName };
}
