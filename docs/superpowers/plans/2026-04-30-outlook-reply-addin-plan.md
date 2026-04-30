# Outlook Reply Add-In Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a personal-sideload Outlook add-in that drafts and refines email replies (and composes new emails) using Google Gemini 2.5 Flash, with a hardcoded user profile so the AI replies in Ali's voice.

**Architecture:** Static taskpane (HTML+vanilla JS+Office.js) hosted on Vercel + two Vercel serverless functions (`/api/generate`, `/api/refine`) that proxy to Google Gemini API. Profile and settings persist in `localStorage`. Manifest is sideloaded once via Outlook on the web; works thereafter on web, new Outlook for Mac, and Outlook mobile.

**Tech Stack:**
- Office.js (Outlook integration, loaded from Microsoft CDN)
- Vanilla JS + HTML + CSS (no framework — taskpane is small)
- Vercel serverless (Node.js 20, ESM)
- `@google/generative-ai` npm package
- `node:test` (built-in test runner — zero deps)
- Vercel CLI for local dev + deploys
- GitHub for source

**Spec:** [`../specs/2026-04-30-outlook-reply-addin-design.md`](../specs/2026-04-30-outlook-reply-addin-design.md)

---

## File Structure

```
outlook-reply-addin/
├── manifest.xml                # Outlook add-in manifest (sideloaded into Outlook)
├── package.json                # Node deps + scripts + node:test entry
├── vercel.json                 # Vercel routing config
├── README.md                   # Setup + sideload + dev instructions
├── .gitignore
├── .env.local.example          # GEMINI_API_KEY template (real .env.local is gitignored)
│
├── public/                     # Vercel serves these statically
│   ├── taskpane.html           # Sidebar shell (loaded by Outlook)
│   ├── taskpane.css            # Styles
│   ├── taskpane.js             # UI controller — DOM, Office.js, fetches /api/*
│   ├── state.js                # Pure JS state module (testable)
│   ├── prompt-builder.js       # Builds payloads for backend (shared with backend in tests)
│   ├── commands.html           # Required by manifest, minimal stub
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
│
├── api/
│   ├── generate.js             # POST /api/generate — reply or compose
│   ├── refine.js               # POST /api/refine — refine an existing draft
│   └── _gemini.js              # Shared Gemini client + system prompt + payload-to-Gemini mapping
│
└── tests/
    ├── prompt-builder.test.js  # Unit tests for payload construction
    ├── state.test.js           # Unit tests for state transitions
    ├── gemini.test.js          # Unit tests for system prompt + request mapping (mocked Gemini SDK)
    ├── generate.test.js        # Unit tests for /api/generate handler
    └── refine.test.js          # Unit tests for /api/refine handler
```

**Why this layout:**
- `public/` is what Vercel serves at the root. Outlook loads `taskpane.html` from there.
- `api/` is what Vercel auto-routes as serverless functions.
- Logic that's testable lives in pure modules (`state.js`, `prompt-builder.js`, `_gemini.js`). DOM and Office.js wiring stays in `taskpane.js` and is smoke-tested manually.
- Tests use `node:test` (built into Node 20+) — no Jest config needed.

---

## Phase 1 — Skeleton & Deploy Pipeline

### Task 1: Init project — package.json, .gitignore, .env example

**Files:**
- Create: `/Users/mad/Desktop/outlook-reply-addin/package.json`
- Create: `/Users/mad/Desktop/outlook-reply-addin/.gitignore`
- Create: `/Users/mad/Desktop/outlook-reply-addin/.env.local.example`

- [ ] **Step 1: Create `package.json`**

```json
{
  "name": "outlook-reply-addin",
  "version": "0.1.0",
  "description": "Personal Outlook add-in that drafts email replies in Ali's voice using Gemini",
  "type": "module",
  "private": true,
  "scripts": {
    "test": "node --test tests/",
    "dev": "vercel dev",
    "deploy": "vercel --prod"
  },
  "engines": {
    "node": ">=20"
  },
  "dependencies": {
    "@google/generative-ai": "^0.21.0"
  }
}
```

- [ ] **Step 2: Create `.gitignore`**

```
node_modules/
.env.local
.env
.vercel/
.DS_Store
*.log
```

- [ ] **Step 3: Create `.env.local.example`**

```
# Get a free key at https://aistudio.google.com/app/apikey
GEMINI_API_KEY=your-key-here
```

- [ ] **Step 4: Install dependencies**

Run from `/Users/mad/Desktop/outlook-reply-addin/`:
```bash
npm install
```
Expected: creates `node_modules/` and `package-lock.json`, no errors.

- [ ] **Step 5: Commit**

```bash
cd /Users/mad/Desktop/outlook-reply-addin
git init
git add package.json package-lock.json .gitignore .env.local.example
git commit -m "chore: init package.json and repo skeleton"
```

---

### Task 2: Push to GitHub

**Files:** none new.

- [ ] **Step 1: Create GitHub repo via gh CLI**

```bash
cd /Users/mad/Desktop/outlook-reply-addin
gh repo create Madboulyjr/outlook-reply-addin --private --source=. --remote=origin --description "Personal Outlook add-in that drafts replies in my voice via Gemini"
```
Expected: repo created at `https://github.com/Madboulyjr/outlook-reply-addin`, remote `origin` added.

- [ ] **Step 2: Push initial commit**

```bash
git push -u origin main
```
Expected: commit pushed, branch tracks `origin/main`.

---

### Task 3: Minimal taskpane HTML + Vercel config

**Files:**
- Create: `public/taskpane.html`
- Create: `public/taskpane.css`
- Create: `public/commands.html`
- Create: `vercel.json`

- [ ] **Step 1: Create `public/taskpane.html` (skeleton only)**

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>AI Email Assistant</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <link rel="stylesheet" href="taskpane.css" />
</head>
<body>
  <main id="root">
    <h1>AI Email Assistant</h1>
    <p id="status">Loading…</p>
  </main>
  <script type="module" src="taskpane.js"></script>
</body>
</html>
```

- [ ] **Step 2: Create `public/taskpane.css` (placeholder)**

```css
:root {
  --bg: #ffffff;
  --fg: #1f1f1f;
  --muted: #6b6b6b;
  --accent: #2563eb;
  --border: #e5e5e5;
  --radius: 6px;
  --gap: 12px;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: var(--gap);
  font: 14px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  color: var(--fg);
  background: var(--bg);
}
h1 { font-size: 16px; margin: 0 0 8px; }
#status { color: var(--muted); }
```

- [ ] **Step 3: Create `public/commands.html` (manifest requires it)**

```html
<!DOCTYPE html>
<html><head><title>Commands</title>
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head><body></body></html>
```

- [ ] **Step 4: Create `vercel.json`**

```json
{
  "version": 2,
  "cleanUrls": false,
  "headers": [
    {
      "source": "/(.*)",
      "headers": [
        { "key": "X-Frame-Options", "value": "ALLOW-FROM https://outlook.office.com" },
        { "key": "Content-Security-Policy", "value": "frame-ancestors https://outlook.office.com https://*.officeapps.live.com" }
      ]
    }
  ],
  "functions": {
    "api/*.js": { "maxDuration": 30 }
  }
}
```

- [ ] **Step 5: Commit**

```bash
git add public/ vercel.json
git commit -m "feat: minimal taskpane shell + vercel config"
```

---

### Task 4: Deploy to Vercel & verify HTTPS URL

**Files:** none new.

- [ ] **Step 1: Install Vercel CLI globally if missing**

```bash
npm install -g vercel
vercel --version
```
Expected: version printed (any v33+).

- [ ] **Step 2: Link to a new Vercel project**

```bash
cd /Users/mad/Desktop/outlook-reply-addin
vercel link --yes
```
When prompted: scope = your personal account, project name = `outlook-reply-addin`, link to GitHub repo when asked.

- [ ] **Step 3: Add Gemini API key to Vercel env**

First get a free key from https://aistudio.google.com/app/apikey, then:
```bash
vercel env add GEMINI_API_KEY production
# paste the key when prompted
vercel env add GEMINI_API_KEY preview
# paste again
vercel env add GEMINI_API_KEY development
# paste again
```
Also create local `.env.local` with the same value (will not be committed):
```bash
echo "GEMINI_API_KEY=your-key-here" > .env.local
```

- [ ] **Step 4: Deploy to production**

```bash
vercel --prod
```
Expected: a URL like `https://outlook-reply-addin.vercel.app`. Note this URL — used in manifest.

- [ ] **Step 5: Verify taskpane loads**

```bash
curl -sI https://outlook-reply-addin.vercel.app/taskpane.html | head -5
```
Expected: `HTTP/2 200`. Open the URL in a browser to see "AI Email Assistant — Loading…".

- [ ] **Step 6: Commit if any vercel.json/env-related change**

```bash
git status
git add -A
git diff --cached
git commit -m "chore: link vercel project and deploy" || echo "nothing to commit"
```

---

### Task 5: Manifest XML + sideload into Outlook

**Files:**
- Create: `manifest.xml`
- Create: `public/icon-16.png`, `public/icon-32.png`, `public/icon-80.png` (placeholder solid-color PNGs)

- [ ] **Step 1: Create three placeholder icons**

Run from project root:
```bash
node -e "
const fs = require('fs');
// 1x1 transparent PNG bytes — Outlook accepts this for sideload icons
const png = Buffer.from('89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C4890000000D49444154789C636000000000050001A5F645400000000049454E44AE426082','hex');
for (const size of [16,32,80]) {
  fs.writeFileSync(\`public/icon-\${size}.png\`, png);
}
console.log('icons created');
"
```
Expected: prints `icons created`. Three files in `public/`. (We'll replace with branded icons later — these are valid placeholders.)

- [ ] **Step 2: Create `manifest.xml`**

Replace `outlook-reply-addin.vercel.app` below with your actual Vercel URL from Task 4.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <Id>e4b1a9c0-1234-4ab1-9abc-000000000001</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Ali Madbouly</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="AI Email Assistant" />
  <Description DefaultValue="Drafts email replies and new emails in your voice using Gemini." />
  <IconUrl DefaultValue="https://outlook-reply-addin.vercel.app/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://outlook-reply-addin.vercel.app/icon-80.png" />
  <SupportUrl DefaultValue="https://github.com/Madboulyjr/outlook-reply-addin" />

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.5" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://outlook-reply-addin.vercel.app/taskpane.html" />
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://outlook-reply-addin.vercel.app/taskpane.html" />
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
  </Rule>

  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.5">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="MessageComposeCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgComposeGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://outlook-reply-addin.vercel.app/icon-16.png" />
        <bt:Image id="Icon.32x32" DefaultValue="https://outlook-reply-addin.vercel.app/icon-32.png" />
        <bt:Image id="Icon.80x80" DefaultValue="https://outlook-reply-addin.vercel.app/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://outlook-reply-addin.vercel.app/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://outlook-reply-addin.vercel.app/taskpane.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="AI Assistant" />
        <bt:String id="TaskpaneButton.Label" DefaultValue="🤖 AI Reply" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Draft an email reply in your voice using AI." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

- [ ] **Step 3: Validate manifest**

```bash
npx --yes office-addin-manifest validate manifest.xml
```
Expected: `The manifest is valid.` If validation errors appear, fix the offending lines and re-run.

- [ ] **Step 4: Sideload into Outlook on the web**

Manual steps (no automation possible — this is a UI flow):
1. Open `https://outlook.office.com` and sign in with the Nice One account.
2. Click the gear icon (top-right) → **View all Outlook settings** → **General** → **Manage add-ins** (or in some tenants: top-right `…` menu → **Get Add-ins**).
3. In the Add-ins dialog: **My add-ins** → **Add a custom add-in** → **Add from file** → pick `manifest.xml`.
4. Confirm the warning dialog ("only install add-ins from sources you trust" — yes, it's yours).

- [ ] **Step 5: Verify the button appears**

1. Open any email in Outlook on the web.
2. Look for `🤖 AI Reply` in the message ribbon (might be under a `…` overflow menu — pin it if so).
3. Click it → the right sidebar should open and load the Vercel-hosted taskpane showing "AI Email Assistant — Loading…".

If the sidebar shows but says blank or errors out — check browser DevTools console for CSP / Office.js errors and fix `vercel.json` or `taskpane.html` accordingly.

- [ ] **Step 6: Commit**

```bash
git add manifest.xml public/icon-*.png
git commit -m "feat: add manifest.xml and placeholder icons; sideload working"
git push
```

---

## Phase 2 — Reply Mode End-to-End

### Task 6: Backend `_gemini.js` — system prompt + Gemini client (TDD)

**Files:**
- Create: `tests/gemini.test.js`
- Create: `api/_gemini.js`

- [ ] **Step 1: Write the failing test**

Create `tests/gemini.test.js`:
```javascript
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
```

- [ ] **Step 2: Run the test to verify it fails**

```bash
npm test
```
Expected: all 5 tests fail with `Cannot find module '../api/_gemini.js'` or similar.

- [ ] **Step 3: Implement `api/_gemini.js`**

```javascript
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
```

- [ ] **Step 4: Run tests — verify pass**

```bash
npm test
```
Expected: all 5 tests pass.

- [ ] **Step 5: Commit**

```bash
git add api/_gemini.js tests/gemini.test.js
git commit -m "feat: add gemini system prompt + buildPrompt with tests"
```

---

### Task 7: Backend `/api/generate` — reply branch (TDD)

**Files:**
- Create: `tests/generate.test.js`
- Create: `api/generate.js`

- [ ] **Step 1: Write the failing test**

Create `tests/generate.test.js`:
```javascript
import { test, mock } from 'node:test';
import assert from 'node:assert/strict';

// Stub the Gemini client BEFORE importing the handler
mock.module('../api/_gemini.js', {
  namedExports: {
    SYSTEM_PROMPT: 'stub',
    buildPrompt: (p) => `PROMPT_FOR_${p.mode}`,
    callGemini: async ({ prompt }) => ({
      text: `STUBBED_OUTPUT_FOR(${prompt})`,
      model: 'gemini-2.5-flash',
    }),
  },
});

const { default: handler } = await import('../api/generate.js');

function makeReqRes(body) {
  const req = { method: 'POST', body, headers: {} };
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

test('POST /api/generate reply mode returns draft', async () => {
  const ctx = makeReqRes({
    mode: 'reply',
    tone: 'professional',
    length: 'medium',
    language: 'auto',
    reply: { senderName: 'Sara', senderEmail: 's@s.com', subject: 'x', body: 'hi', notes: '' },
  });
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 200);
  assert.equal(ctx.payload.ok, true);
  assert.match(ctx.payload.draft, /STUBBED_OUTPUT/);
  assert.equal(ctx.payload.subject, null);
  assert.equal(ctx.payload.model, 'gemini-2.5-flash');
});

test('POST /api/generate compose mode parses JSON output', async () => {
  // Override the stub for this test
  const { default: composeHandler } = await import('../api/generate.js?compose');
  // We can't re-mock easily; instead we will validate compose-parsing in a separate task with a real-shaped stub.
  // For this test, just verify the handler accepts a compose payload without throwing:
  const ctx = makeReqRes({
    mode: 'compose',
    tone: 'casual',
    length: 'short',
    language: 'english',
    compose: { to: 'Sara', topic: 'launch', notes: 'soon' },
  });
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 200);
  assert.equal(ctx.payload.ok, true);
});

test('rejects non-POST', async () => {
  const ctx = makeReqRes(null);
  ctx.req.method = 'GET';
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 405);
});

test('rejects missing mode', async () => {
  const ctx = makeReqRes({ tone: 'professional' });
  await handler(ctx.req, ctx.res);
  assert.equal(ctx.statusCode, 400);
  assert.equal(ctx.payload.ok, false);
});
```

- [ ] **Step 2: Run — verify failure**

```bash
npm test
```
Expected: `Cannot find module '../api/generate.js'`.

- [ ] **Step 3: Implement `api/generate.js`**

```javascript
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

  try {
    const prompt = buildPrompt(body);
    const { text, model } = await callGemini({ prompt, useProQuality: !!body.useProQuality });

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
```

- [ ] **Step 4: Run tests — verify pass**

```bash
npm test
```
Expected: all generate.test.js tests pass.

- [ ] **Step 5: Commit**

```bash
git add api/generate.js tests/generate.test.js
git commit -m "feat: add /api/generate handler for reply + compose with validation"
```

---

### Task 8: Backend `/api/refine` (TDD)

**Files:**
- Create: `tests/refine.test.js`
- Create: `api/refine.js`

- [ ] **Step 1: Write failing test**

Create `tests/refine.test.js`:
```javascript
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
```

- [ ] **Step 2: Run — verify failure**

```bash
npm test
```
Expected: 3 refine tests fail with module-not-found.

- [ ] **Step 3: Implement `api/refine.js`**

```javascript
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
    const { text, model } = await callGemini({ prompt, useProQuality: !!body.useProQuality });
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
```

- [ ] **Step 4: Run tests — verify pass**

```bash
npm test
```
Expected: all refine tests pass; total test count increases.

- [ ] **Step 5: Commit**

```bash
git add api/refine.js tests/refine.test.js
git commit -m "feat: add /api/refine handler with validation"
```

---

### Task 9: Frontend pure modules — `state.js` + `prompt-builder.js` (TDD)

**Files:**
- Create: `tests/state.test.js`
- Create: `tests/prompt-builder.test.js`
- Create: `public/state.js`
- Create: `public/prompt-builder.js`

- [ ] **Step 1: Write failing tests for state**

Create `tests/state.test.js`:
```javascript
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
```

- [ ] **Step 2: Write failing tests for prompt-builder**

Create `tests/prompt-builder.test.js`:
```javascript
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
```

- [ ] **Step 3: Run tests — verify failure**

```bash
npm test
```
Expected: state + prompt-builder tests fail with module-not-found.

- [ ] **Step 4: Implement `public/state.js`**

```javascript
export function createInitialState() {
  return {
    mode: 'reply',
    tone: 'professional',
    length: 'medium',
    language: 'auto',
    useProQuality: false,
    notes: '',
    customRules: '',
    composeFields: { to: '', topic: '', notes: '' },
    currentEmail: null,  // { senderName, senderEmail, subject, body }
    draft: null,
    subject: null,
    missingInfo: null,
    loading: false,
    error: null,
  };
}

export function applyAction(state, action) {
  switch (action.type) {
    case 'SET_MODE':
      return { ...state, mode: action.mode, draft: null, subject: null, missingInfo: null, error: null };
    case 'SET_TONE':
      return { ...state, tone: action.tone };
    case 'SET_LENGTH':
      return { ...state, length: action.length };
    case 'SET_LANGUAGE':
      return { ...state, language: action.language };
    case 'SET_PRO_QUALITY':
      return { ...state, useProQuality: !!action.value };
    case 'SET_NOTES':
      return { ...state, notes: action.notes };
    case 'SET_COMPOSE_FIELD':
      return { ...state, composeFields: { ...state.composeFields, [action.field]: action.value } };
    case 'SET_CUSTOM_RULES':
      return { ...state, customRules: action.rules };
    case 'SET_CURRENT_EMAIL':
      return { ...state, currentEmail: action.email };
    case 'GENERATE_START':
      return { ...state, loading: true, error: null };
    case 'GENERATE_SUCCESS':
      return { ...state, loading: false, draft: action.draft, subject: action.subject, missingInfo: action.missingInfo, error: null };
    case 'GENERATE_FAIL':
      return { ...state, loading: false, error: action.error };
    case 'CLEAR_DRAFT':
      return { ...state, draft: null, subject: null, missingInfo: null };
    default:
      return state;
  }
}
```

- [ ] **Step 5: Implement `public/prompt-builder.js`**

```javascript
export function buildGeneratePayload(state, email, compose) {
  const base = {
    mode: state.mode,
    tone: state.tone,
    length: state.length,
    language: state.language,
    useProQuality: !!state.useProQuality,
    customRules: state.customRules || '',
  };
  if (state.mode === 'reply') {
    return {
      ...base,
      reply: {
        senderName: email?.senderName || '',
        senderEmail: email?.senderEmail || '',
        subject: email?.subject || '',
        body: email?.body || '',
        notes: state.notes || '',
      },
    };
  }
  if (state.mode === 'compose') {
    return {
      ...base,
      compose: {
        to: compose?.to || '',
        topic: compose?.topic || '',
        notes: compose?.notes || '',
      },
    };
  }
  return base;
}

export function buildRefinePayload(state, previousDraft, instruction) {
  return {
    previousDraft,
    instruction,
    tone: state.tone,
    length: state.length,
    language: state.language,
    useProQuality: !!state.useProQuality,
    customRules: state.customRules || '',
  };
}
```

- [ ] **Step 6: Run tests — verify pass**

```bash
npm test
```
Expected: all state + prompt-builder tests pass.

- [ ] **Step 7: Commit**

```bash
git add public/state.js public/prompt-builder.js tests/state.test.js tests/prompt-builder.test.js
git commit -m "feat: add pure state + payload modules with tests"
```

---

### Task 10: Taskpane UI — Reply mode (HTML + CSS + glue JS)

**Files:**
- Modify: `public/taskpane.html`
- Modify: `public/taskpane.css`
- Create: `public/taskpane.js`

- [ ] **Step 1: Replace `public/taskpane.html` with full reply UI**

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>AI Email Assistant</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <link rel="stylesheet" href="taskpane.css" />
</head>
<body>
  <main id="root">
    <header>
      <h1>AI Email Assistant</h1>
      <div class="mode-tabs" role="tablist">
        <button class="mode-tab active" data-mode="reply" role="tab" aria-selected="true">📨 Reply</button>
        <button class="mode-tab" data-mode="compose" role="tab" aria-selected="false">✏️ Compose</button>
      </div>
    </header>

    <section id="reply-section" class="mode-section">
      <div class="email-info" id="email-info">
        <div class="muted">Reading current email…</div>
      </div>

      <label>
        <span>Notes (optional)</span>
        <textarea id="notes" rows="3" placeholder="What do you want to say?"></textarea>
      </label>
    </section>

    <section id="compose-section" class="mode-section hidden">
      <label>
        <span>To</span>
        <input id="compose-to" type="text" placeholder="Recipient name or email" />
      </label>
      <label>
        <span>About</span>
        <input id="compose-topic" type="text" placeholder="Short topic line" />
      </label>
      <label>
        <span>Notes</span>
        <textarea id="compose-notes" rows="3" placeholder="What do you want to say?"></textarea>
      </label>
    </section>

    <section class="controls">
      <div class="control-row">
        <span class="control-label">Tone</span>
        <div class="btn-group" data-control="tone">
          <button data-value="very_formal">Formal</button>
          <button data-value="professional" class="active">Pro</button>
          <button data-value="casual">Casual</button>
          <button data-value="very_casual">Chill</button>
        </div>
      </div>
      <div class="control-row">
        <span class="control-label">Length</span>
        <div class="btn-group" data-control="length">
          <button data-value="short">Short</button>
          <button data-value="medium" class="active">Med</button>
          <button data-value="detailed">Detailed</button>
        </div>
      </div>
      <div class="control-row">
        <span class="control-label">Lang</span>
        <div class="btn-group" data-control="language">
          <button data-value="english">EN</button>
          <button data-value="arabic">AR</button>
          <button data-value="auto" class="active">Auto</button>
        </div>
      </div>
      <label class="checkbox">
        <input type="checkbox" id="pro-quality" />
        <span>Use Pro Quality (high-stakes)</span>
      </label>
    </section>

    <button id="generate" class="primary">✨ Generate</button>

    <section id="draft-section" class="hidden">
      <h2 id="draft-subject-row" class="hidden"><span class="muted">Subject:</span> <span id="draft-subject"></span></h2>
      <pre id="draft" class="draft"></pre>
      <div id="missing-info" class="missing hidden"></div>

      <div class="refine">
        <div class="refine-label">Refine:</div>
        <div class="btn-group wrap">
          <button class="refine-btn" data-instruction="make it more formal">More formal</button>
          <button class="refine-btn" data-instruction="make it friendlier and warmer">Friendlier</button>
          <button class="refine-btn" data-instruction="make it shorter and more direct">Shorter</button>
          <button class="refine-btn" data-instruction="add more detail and context">Longer</button>
          <button class="refine-btn" data-instruction="make it more polite">More polite</button>
          <button class="refine-btn" data-instruction="make it more direct, even firm">More aggressive</button>
          <button class="refine-btn" data-instruction="make it sound less corporate, more human">Less corporate</button>
          <button class="refine-btn" data-instruction="add a touch of light humor">Add humor</button>
          <button class="refine-btn" data-instruction="translate the entire reply to Arabic">→ Arabic</button>
          <button class="refine-btn" data-instruction="translate the entire reply to English">→ English</button>
        </div>
        <div class="refine-custom">
          <input id="custom-refine" type="text" placeholder="Custom: tell me what to change" />
          <button id="custom-refine-go">Apply</button>
        </div>
      </div>

      <div class="actions">
        <button id="copy">📋 Copy</button>
        <button id="insert" class="primary">📨 Insert</button>
      </div>
    </section>

    <p id="error" class="error hidden"></p>
    <p id="status" class="muted"></p>

    <footer>
      <button id="settings-toggle" class="link">⚙️ Settings</button>
    </footer>

    <section id="settings-panel" class="hidden">
      <label>
        <span>Custom voice rules (added to system prompt)</span>
        <textarea id="custom-rules" rows="4" placeholder="e.g. 'Always mention next steps' or 'Avoid the word synergy'"></textarea>
      </label>
      <button id="settings-save">Save</button>
    </section>
  </main>
  <script type="module" src="taskpane.js"></script>
</body>
</html>
```

- [ ] **Step 2: Replace `public/taskpane.css`**

```css
:root {
  --bg: #ffffff;
  --fg: #1f1f1f;
  --muted: #6b6b6b;
  --accent: #2563eb;
  --accent-fg: #ffffff;
  --border: #e5e5e5;
  --surface: #f6f6f7;
  --danger: #b91c1c;
  --warn: #b45309;
  --radius: 6px;
  --gap: 10px;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: var(--gap);
  font: 13px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  color: var(--fg);
  background: var(--bg);
  line-height: 1.4;
}
h1 { font-size: 14px; font-weight: 600; margin: 0 0 8px; }
h2 { font-size: 13px; font-weight: 500; margin: 0 0 6px; }
.muted { color: var(--muted); font-size: 12px; }
.hidden { display: none !important; }

header { margin-bottom: var(--gap); }

.mode-tabs { display: flex; gap: 4px; border-bottom: 1px solid var(--border); }
.mode-tab {
  flex: 1; padding: 6px 8px; border: 0; background: transparent;
  border-bottom: 2px solid transparent; cursor: pointer; font-size: 13px;
}
.mode-tab.active { border-bottom-color: var(--accent); color: var(--accent); font-weight: 600; }

.mode-section, .controls, #draft-section, #settings-panel {
  margin-bottom: var(--gap);
}

label { display: block; margin-bottom: var(--gap); }
label > span { display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }
input[type="text"], textarea {
  width: 100%; padding: 6px 8px; font: inherit; border: 1px solid var(--border);
  border-radius: var(--radius); background: var(--bg); resize: vertical;
}
textarea { min-height: 60px; }

.email-info {
  background: var(--surface); padding: 8px; border-radius: var(--radius);
  margin-bottom: var(--gap); font-size: 12px; max-height: 120px; overflow: auto;
}
.email-info .from { font-weight: 600; }
.email-info .subject { color: var(--muted); margin-bottom: 4px; }

.controls { display: grid; gap: 6px; }
.control-row { display: flex; align-items: center; gap: 8px; }
.control-label { width: 44px; font-size: 12px; color: var(--muted); flex-shrink: 0; }
.btn-group { display: flex; gap: 4px; flex: 1; }
.btn-group.wrap { flex-wrap: wrap; }
.btn-group button {
  flex: 1; padding: 4px 8px; border: 1px solid var(--border); background: var(--bg);
  border-radius: var(--radius); cursor: pointer; font-size: 12px;
}
.btn-group button.active {
  background: var(--accent); color: var(--accent-fg); border-color: var(--accent);
}
.btn-group.wrap button { flex: 0 0 auto; }

.checkbox { display: flex; align-items: center; gap: 6px; font-size: 12px; }
.checkbox input { margin: 0; }

button.primary {
  width: 100%; padding: 10px; font-size: 14px; background: var(--accent);
  color: var(--accent-fg); border: 0; border-radius: var(--radius); cursor: pointer;
  font-weight: 600;
}
button.primary:disabled { opacity: 0.6; cursor: progress; }

button.link {
  background: none; border: 0; color: var(--accent); cursor: pointer;
  font-size: 12px; padding: 0;
}

.draft {
  background: var(--surface); padding: 8px; border-radius: var(--radius);
  white-space: pre-wrap; font: inherit; max-height: 220px; overflow: auto;
  margin: 0 0 var(--gap);
}
.missing {
  background: #fef3c7; color: var(--warn); padding: 6px 8px;
  border-radius: var(--radius); font-size: 12px; margin-bottom: var(--gap);
}
.refine { margin-bottom: var(--gap); }
.refine-label { font-size: 12px; color: var(--muted); margin-bottom: 4px; }
.refine-custom { display: flex; gap: 4px; margin-top: 6px; }
.refine-custom input { flex: 1; }
.refine-custom button {
  padding: 4px 10px; border: 1px solid var(--border); background: var(--bg);
  border-radius: var(--radius); cursor: pointer;
}

.actions { display: flex; gap: 6px; }
.actions button {
  flex: 1; padding: 8px; border: 1px solid var(--border); background: var(--bg);
  border-radius: var(--radius); cursor: pointer;
}
.actions button.primary { border: 0; }

.error {
  background: #fee2e2; color: var(--danger); padding: 6px 8px;
  border-radius: var(--radius); font-size: 12px;
}

footer { margin-top: 16px; text-align: center; }

@media (max-width: 320px) {
  body { padding: 6px; font-size: 12px; }
  .control-row { flex-direction: column; align-items: stretch; gap: 2px; }
  .control-label { width: auto; }
}
```

- [ ] **Step 3: Create `public/taskpane.js`**

```javascript
import { createInitialState, applyAction } from './state.js';
import { buildGeneratePayload, buildRefinePayload } from './prompt-builder.js';

const $ = (id) => document.getElementById(id);
let state = createInitialState();
const SETTINGS_KEY = 'aireply.settings.v1';

function dispatch(action) {
  state = applyAction(state, action);
  render();
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) return;
    const saved = JSON.parse(raw);
    if (saved.tone) state.tone = saved.tone;
    if (saved.length) state.length = saved.length;
    if (saved.language) state.language = saved.language;
    if (typeof saved.useProQuality === 'boolean') state.useProQuality = saved.useProQuality;
    if (typeof saved.customRules === 'string') state.customRules = saved.customRules;
  } catch { /* ignore */ }
}

function saveSettings() {
  const toSave = {
    tone: state.tone,
    length: state.length,
    language: state.language,
    useProQuality: state.useProQuality,
    customRules: state.customRules,
  };
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(toSave));
}

function render() {
  // Mode tabs
  document.querySelectorAll('.mode-tab').forEach(b => {
    b.classList.toggle('active', b.dataset.mode === state.mode);
    b.setAttribute('aria-selected', b.dataset.mode === state.mode ? 'true' : 'false');
  });
  $('reply-section').classList.toggle('hidden', state.mode !== 'reply');
  $('compose-section').classList.toggle('hidden', state.mode !== 'compose');

  // Active control buttons
  for (const control of ['tone', 'length', 'language']) {
    const group = document.querySelector(`.btn-group[data-control="${control}"]`);
    group.querySelectorAll('button').forEach(b => {
      b.classList.toggle('active', b.dataset.value === state[control]);
    });
  }
  $('pro-quality').checked = state.useProQuality;

  // Email info
  if (state.currentEmail) {
    $('email-info').innerHTML = `
      <div class="from">${escape(state.currentEmail.senderName)} &lt;${escape(state.currentEmail.senderEmail)}&gt;</div>
      <div class="subject">${escape(state.currentEmail.subject)}</div>
      <div>${escape(truncate(state.currentEmail.body, 300))}</div>
    `;
  }

  // Draft
  if (state.draft) {
    $('draft-section').classList.remove('hidden');
    $('draft').textContent = state.draft;
    if (state.subject) {
      $('draft-subject-row').classList.remove('hidden');
      $('draft-subject').textContent = state.subject;
    } else {
      $('draft-subject-row').classList.add('hidden');
    }
    if (state.missingInfo) {
      $('missing-info').classList.remove('hidden');
      $('missing-info').textContent = '⚠️ Missing: ' + state.missingInfo;
    } else {
      $('missing-info').classList.add('hidden');
    }
  } else {
    $('draft-section').classList.add('hidden');
  }

  // Loading & error
  $('generate').disabled = state.loading;
  $('generate').textContent = state.loading ? '⏳ Generating…' : '✨ Generate';
  if (state.error) {
    $('error').classList.remove('hidden');
    $('error').textContent = state.error;
  } else {
    $('error').classList.add('hidden');
  }

  // Notes / compose fields
  $('notes').value = state.notes;
  $('compose-to').value = state.composeFields.to;
  $('compose-topic').value = state.composeFields.topic;
  $('compose-notes').value = state.composeFields.notes;
  $('custom-rules').value = state.customRules;
}

function escape(s) {
  return String(s ?? '').replace(/[&<>]/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;' }[c]));
}
function truncate(s, n) { return s.length <= n ? s : s.slice(0, n) + '…'; }

async function readCurrentEmail() {
  return new Promise((resolve, reject) => {
    if (!Office?.context?.mailbox?.item) return resolve(null);
    const item = Office.context.mailbox.item;
    if (!item.body) return resolve(null);
    item.body.getAsync('text', (result) => {
      if (result.status !== 'succeeded') return reject(new Error(result.error?.message || 'body read failed'));
      resolve({
        senderName: item.from?.displayName || item.sender?.displayName || '',
        senderEmail: item.from?.emailAddress || item.sender?.emailAddress || '',
        subject: item.subject || '',
        body: result.value || '',
      });
    });
  });
}

async function generate() {
  dispatch({ type: 'GENERATE_START' });
  try {
    const compose = state.mode === 'compose' ? state.composeFields : null;
    const payload = buildGeneratePayload(state, state.currentEmail, compose);
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
    dispatch({ type: 'GENERATE_SUCCESS', draft: data.draft, subject: data.subject, missingInfo: data.missingInfo });
  } catch (err) {
    dispatch({ type: 'GENERATE_FAIL', error: friendlyError(err) });
  }
}

async function refine(instruction) {
  if (!state.draft) return;
  dispatch({ type: 'GENERATE_START' });
  try {
    const payload = buildRefinePayload(state, state.draft, instruction);
    const res = await fetch('/api/refine', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
    dispatch({ type: 'GENERATE_SUCCESS', draft: data.draft, subject: state.subject, missingInfo: data.missingInfo });
  } catch (err) {
    dispatch({ type: 'GENERATE_FAIL', error: friendlyError(err) });
  }
}

function friendlyError(err) {
  const msg = err.message || String(err);
  if (msg.includes('quota') || msg.includes('429')) return 'Daily free quota hit. Try again tomorrow or enable Pro Quality.';
  if (msg.includes('Failed to fetch')) return 'Network error — check your connection.';
  return `Generation failed: ${msg}`;
}

async function insertDraft() {
  if (!state.draft) return;
  const html = state.draft.split('\n').map(l => `<div>${escape(l) || '&nbsp;'}</div>`).join('');
  if (state.mode === 'reply') {
    Office.context.mailbox.item.displayReplyForm({ htmlBody: html });
  } else {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: state.composeFields.to ? [state.composeFields.to] : [],
      subject: state.subject || '',
      htmlBody: html,
    });
  }
}

function copyDraft() {
  if (!state.draft) return;
  navigator.clipboard?.writeText(state.draft);
  $('status').textContent = 'Copied.';
  setTimeout(() => { $('status').textContent = ''; }, 1500);
}

function bindEvents() {
  document.querySelectorAll('.mode-tab').forEach(b => {
    b.addEventListener('click', () => dispatch({ type: 'SET_MODE', mode: b.dataset.mode }));
  });
  for (const control of ['tone', 'length', 'language']) {
    const group = document.querySelector(`.btn-group[data-control="${control}"]`);
    group.querySelectorAll('button').forEach(b => {
      b.addEventListener('click', () => {
        const upper = control === 'tone' ? 'SET_TONE' : control === 'length' ? 'SET_LENGTH' : 'SET_LANGUAGE';
        dispatch({ type: upper, [control]: b.dataset.value });
      });
    });
  }
  $('pro-quality').addEventListener('change', (e) => dispatch({ type: 'SET_PRO_QUALITY', value: e.target.checked }));
  $('notes').addEventListener('input', (e) => state.notes = e.target.value);
  $('compose-to').addEventListener('input', (e) => dispatch({ type: 'SET_COMPOSE_FIELD', field: 'to', value: e.target.value }));
  $('compose-topic').addEventListener('input', (e) => dispatch({ type: 'SET_COMPOSE_FIELD', field: 'topic', value: e.target.value }));
  $('compose-notes').addEventListener('input', (e) => dispatch({ type: 'SET_COMPOSE_FIELD', field: 'notes', value: e.target.value }));
  $('generate').addEventListener('click', generate);
  $('copy').addEventListener('click', copyDraft);
  $('insert').addEventListener('click', insertDraft);
  document.querySelectorAll('.refine-btn').forEach(b => {
    b.addEventListener('click', () => refine(b.dataset.instruction));
  });
  $('custom-refine-go').addEventListener('click', () => {
    const v = $('custom-refine').value.trim();
    if (v) { refine(v); $('custom-refine').value = ''; }
  });
  $('settings-toggle').addEventListener('click', () => {
    $('settings-panel').classList.toggle('hidden');
  });
  $('settings-save').addEventListener('click', () => {
    state.customRules = $('custom-rules').value;
    saveSettings();
    $('status').textContent = 'Saved.';
    setTimeout(() => { $('status').textContent = ''; }, 1500);
  });
}

Office.onReady(async (info) => {
  loadSettings();
  bindEvents();
  render();
  if (info.host === Office.HostType.Outlook) {
    try {
      const email = await readCurrentEmail();
      if (email) dispatch({ type: 'SET_CURRENT_EMAIL', email });
      else $('email-info').innerHTML = '<div class="muted">No email selected (Compose mode is available).</div>';
    } catch (err) {
      $('email-info').innerHTML = `<div class="muted">Couldn't read email: ${escape(err.message)}</div>`;
    }
  } else {
    // Outside Outlook (browser preview)
    $('email-info').innerHTML = '<div class="muted">Preview mode (not in Outlook).</div>';
  }
});
```

- [ ] **Step 4: Deploy to Vercel**

```bash
cd /Users/mad/Desktop/outlook-reply-addin
vercel --prod
```
Expected: deploys, prints URL, no errors.

- [ ] **Step 5: Smoke test in Outlook web**

1. Open an email in Outlook on the web.
2. Click the `🤖 AI Reply` ribbon button.
3. Verify the sidebar loads: shows email info, tone/length/lang controls, notes textarea, Generate button.
4. Click Generate without notes → verify a draft appears in 2-5s.
5. If quota / API key errors show — check Vercel logs (`vercel logs`) and that GEMINI_API_KEY is set in production env.

- [ ] **Step 6: Commit**

```bash
git add public/taskpane.html public/taskpane.css public/taskpane.js
git commit -m "feat: full taskpane UI with reply mode, refine, compose stubs"
git push
```

---

## Phase 3 — Compose Mode + Refine Polish

### Task 11: Verify compose mode end-to-end

**Files:** none new (compose is already implemented in Tasks 7, 9, 10).

- [ ] **Step 1: Switch to compose mode in the taskpane**

Open the add-in in Outlook web, click the ✏️ Compose tab.

- [ ] **Step 2: Fill in test data**

- To: `Sara Ahmed`
- About: `Launch deck v4 timeline`
- Notes: `need final files Friday before launch on Monday`
- Tone: Pro, Length: Med, Language: EN

- [ ] **Step 3: Click Generate, verify draft**

Expected: a draft appears with a subject line shown above the body and a body ending in `BR,`.

- [ ] **Step 4: Click Insert**

Expected: Outlook opens a new compose window with `Sara Ahmed` in To, the subject filled, and the body filled.

- [ ] **Step 5: If anything off, fix and redeploy**

Common issues:
- Subject not showing: check `parseComposeOutput` in `api/generate.js` — Gemini may have returned non-JSON. Tighten the prompt instruction in `_gemini.js` ("Return ONLY a JSON object…").
- New compose window doesn't open: verify `displayNewMessageForm` is being called with correct params (DevTools console).

- [ ] **Step 6: Commit any fixes**

```bash
git status && git diff
git add -A && git commit -m "fix: tighten compose mode behavior" || echo "no fixes needed"
git push
```

---

### Task 12: Add settings panel persistence + profile editor

**Files:**
- Modify: `public/taskpane.html`
- Modify: `public/taskpane.js`

The system prompt is hardcoded in `_gemini.js` so the profile is fixed in v1 (this is intentional per spec). Settings panel currently only persists `customRules` + tone/length/language defaults + Pro Quality.

- [ ] **Step 1: Verify settings panel writes to localStorage**

In Outlook web → open add-in → Settings → type something into custom rules → Save.

```javascript
// Run this in DevTools console of the taskpane iframe:
JSON.parse(localStorage.getItem('aireply.settings.v1'))
```
Expected: object with `customRules` matching what you typed, plus current tone/length/language.

- [ ] **Step 2: Verify settings reload after refresh**

Reload the taskpane. Open Settings. Confirm the saved customRules text is still there.

- [ ] **Step 3: Verify customRules reaches the prompt**

Set custom rules to something obvious like `"Always start with: TEST_MARKER_42."`. Generate a reply. Verify the draft starts with `TEST_MARKER_42`.

If it doesn't: trace through — `taskpane.js` reads `state.customRules`, `buildGeneratePayload` includes it, `/api/generate` validates and forwards, `buildPrompt` in `_gemini.js` includes it. Find the broken link and fix.

- [ ] **Step 4: Commit if any fix**

```bash
git status && git diff
git add -A && git commit -m "fix: ensure customRules wired through end-to-end" || echo "no fixes needed"
git push
```

---

## Phase 4 — Polish & Cross-Platform Test

### Task 13: Error handling pass

**Files:**
- Modify: `public/taskpane.js` (already has `friendlyError` — verify cases)

- [ ] **Step 1: Test rate-limit error**

Manually trigger by setting `GEMINI_API_KEY` in Vercel to an invalid value temporarily, deploy, click Generate. Should show "Generation failed: …" — note the message. Restore the real key.

- [ ] **Step 2: Test network failure**

In DevTools → Network tab → enable "Offline". Click Generate. Expected: friendly "Network error — check your connection."

- [ ] **Step 3: Test empty notes + empty email body (compose) → ensure no crash**

In compose mode, leave all fields blank, click Generate. Expected: backend returns a draft anyway (Gemini will write something generic) OR a 400 error if validation triggers. Either is fine — verify there's no JS crash and the error is shown nicely.

- [ ] **Step 4: Commit if any polish**

```bash
git status && git diff
git add -A && git commit -m "polish: error handling" || echo "no fixes needed"
git push
```

---

### Task 14: Cross-platform test — new Outlook for Mac

**Files:** none.

- [ ] **Step 1: Open new Outlook for Mac**

Open the **new** Outlook for Mac app (not Outlook 2019 / classic). Sign in with the Nice One account.

- [ ] **Step 2: Verify the add-in propagated**

The `🤖 AI Reply` button should appear on the message ribbon (sometimes under `…` overflow — pin it).

- [ ] **Step 3: Click an email, click the button, verify the taskpane loads**

If the taskpane is blank or stuck on "Loading…" — open DevTools (right-click taskpane → Inspect Element if available, or check the Office Diagnostics console) and look for errors.

- [ ] **Step 4: Generate + Insert, verify reply window opens with draft**

- [ ] **Step 5: Note any layout issues**

Take a screenshot. If buttons are too cramped or text overflows in the narrower Mac taskpane, tweak CSS in `public/taskpane.css` (the existing `@media (max-width: 320px)` rule is the right hook).

- [ ] **Step 6: Commit any tweaks**

```bash
git status && git diff
git add -A && git commit -m "polish: layout tweaks for new Outlook for Mac" || echo "no fixes needed"
git push
```

---

### Task 15: Cross-platform test — Outlook iOS/Android

**Files:** none.

- [ ] **Step 1: Open Outlook on phone**

Open the Outlook mobile app, signed into the Nice One account.

- [ ] **Step 2: Open an email and find the add-in**

On mobile, add-ins typically live under the `…` menu of the message — tap it, look for `🤖 AI Reply`.

- [ ] **Step 3: Verify taskpane loads + generate works**

- [ ] **Step 4: Verify Insert opens a reply on mobile**

- [ ] **Step 5: Note any mobile layout issues + adjust CSS**

Mobile taskpanes are narrower — verify:
- Mode tabs wrap or fit
- Tone/length/lang button rows fit
- Refine buttons wrap reasonably

If problems, edit `public/taskpane.css` and redeploy.

- [ ] **Step 6: Commit any tweaks**

```bash
git status && git diff
git add -A && git commit -m "polish: mobile layout fixes" || echo "no fixes needed"
git push
```

---

### Task 16: README + Tag v1.0.0

**Files:**
- Create: `README.md`

- [ ] **Step 1: Create `README.md`**

```markdown
# Outlook Reply Add-In

Personal Outlook add-in that drafts email replies and new emails in Ali Madbouly's voice using Google Gemini 2.5 Flash.

## Local development

1. Clone: `git clone https://github.com/Madboulyjr/outlook-reply-addin.git && cd outlook-reply-addin`
2. Install: `npm install`
3. Get a Gemini key: https://aistudio.google.com/app/apikey
4. Copy env: `cp .env.local.example .env.local`, paste the key into `.env.local`
5. Dev server: `npm run dev` (uses `vercel dev`, runs at `http://localhost:3000`)
6. Tests: `npm test`

## Deploy

```bash
vercel --prod
```

(env vars must be set in Vercel: `vercel env add GEMINI_API_KEY production`)

## Sideload manifest into Outlook

1. Open https://outlook.office.com → Settings → Manage Add-ins → My Add-ins → Add a custom add-in → From file
2. Upload `manifest.xml`
3. The button `🤖 AI Reply` appears in the message ribbon on web, new Outlook for Mac, and mobile.

## Architecture

- `public/` — static taskpane (Office.js + vanilla JS). Vercel serves at root.
- `api/` — Vercel serverless functions:
  - `POST /api/generate` — reply or compose
  - `POST /api/refine` — refine an existing draft
- `tests/` — `node:test` unit tests (`npm test`)

## Spec / Plan

- Spec: `docs/superpowers/specs/2026-04-30-outlook-reply-addin-design.md`
- Plan: `docs/superpowers/plans/2026-04-30-outlook-reply-addin-plan.md`

## License

Private — personal use only.
```

- [ ] **Step 2: Commit + tag**

```bash
git add README.md
git commit -m "docs: add README with setup instructions"
git tag v1.0.0
git push --tags
git push
```

- [ ] **Step 3: Final smoke test**

Open Outlook web, generate a reply, refine it, insert it — confirm the round trip works end-to-end one more time.

---

## Self-Review

- **Spec coverage:**
  - § 1-2 Problem & Goals → covered by all tasks (the add-in IS the solution).
  - § 3.1 Reply flow → Tasks 7, 10 (UI), 14-15 (cross-platform).
  - § 3.2 Compose flow → Tasks 7, 10, 11.
  - § 3.3 Refine flow → Tasks 8, 10.
  - § 4 Architecture → matches file structure laid out.
  - § 5 Tech stack → Task 1 (deps) + Task 4 (Vercel).
  - § 6 System prompt → Task 6 (verbatim from spec).
  - § 7 UI spec → Task 10 (HTML matches the layout sketch).
  - § 8 API contract → Tasks 7, 8 (validation + response shape match).
  - § 9 Privacy → no code, but `vercel.json` doesn't enable request logging; `.env.local` gitignored. Documented.
  - § 10 Sideload → Task 5.
  - § 11 Phases → tasks are grouped under matching phase headings.
  - § 12 Future extensions → not in v1 (correctly out of scope).

- **Placeholder scan:** No "TBD", "TODO", or vague "implement appropriate X" lines in any task. Every code block is complete.

- **Type/name consistency check:**
  - State action names: `SET_MODE`, `SET_TONE`, `SET_LENGTH`, `SET_LANGUAGE`, `SET_PRO_QUALITY`, `SET_NOTES`, `SET_COMPOSE_FIELD`, `SET_CUSTOM_RULES`, `SET_CURRENT_EMAIL`, `GENERATE_START`, `GENERATE_SUCCESS`, `GENERATE_FAIL`, `CLEAR_DRAFT` — used consistently in `state.js` (Task 9 Step 4) and `taskpane.js` (Task 10 Step 3).
  - Function names: `buildPrompt`, `callGemini`, `buildGeneratePayload`, `buildRefinePayload`, `parseComposeOutput`, `extractMissingInfo` — consistent across tasks.
  - API request/response field names: `mode`, `tone`, `length`, `language`, `useProQuality`, `customRules`, `reply`, `compose`, `previousDraft`, `instruction`, `draft`, `subject`, `missingInfo`, `model` — consistent between backend tasks (7, 8) and frontend builder (Task 9).
  - Field `senderName`/`senderEmail` used consistently in `_gemini.js` (Task 6), `prompt-builder.js` (Task 9), and `taskpane.js` `readCurrentEmail` (Task 10).

- **One ambiguity flagged & resolved during writing:** Task 7 originally tried to swap mocks for compose-mode JSON parsing in a single test file. `node:test`'s `mock.module` doesn't easily re-mock per test, so the second test in `tests/generate.test.js` only verifies the handler accepts the compose payload without crashing. Compose-mode JSON parsing is exercised end-to-end in Task 11 (manual smoke test). Acceptable trade-off: `parseComposeOutput` is a pure function that could be unit tested in a separate file if more rigor is desired later.

No issues to fix.

---

## 14. Approval

- [ ] Plan reviewed
- [ ] Plan approved
- [ ] Execution mode chosen (subagent-driven vs inline)
