# Outlook Reply Add-In — Design Spec

**Date:** 2026-04-30
**Author:** Ali Madbouly (Head of Art / Art Director, Nice One)
**Status:** Draft, awaiting approval

---

## 1. Problem

Ali drafts many work emails per day at Nice One — design feedback, vendor/agency comms, scheduling, brand-work coordination. Drafting in clean English (or Arabic) while keeping a consistent voice eats time. He wants:

- Paste an incoming email → get a draft reply in his voice.
- Optionally write a brief → get a fresh email composed for him.
- Refine drafts via single-click buttons (more formal, friendlier, shorter, etc.).
- All of this **inside Outlook**, working on Mac desktop and the Outlook mobile app.
- **No subscription cost** beyond what he already pays.

## 2. Goals & Non-Goals

### Goals
- Native Outlook integration (taskpane add-in) — no context switch.
- Reply mode + Compose mode + Refine mode in the same UI.
- Persistent profile (name, role, company, tone, signature) hardcoded into every AI call so the model "knows him".
- Cost-free: Gemini 2.5 Flash free tier (1,500 requests/day) + Vercel free hosting tier.
- Personal sideload — no IT involvement, no org deployment.
- Works on: Outlook on the web, new Outlook for Mac, Outlook mobile (iOS/Android).

### Non-Goals (explicitly out of scope)
- Multi-user / org-wide deployment (would need Microsoft 365 admin + IT review).
- Voice-sample learning (paste 3 past emails to mimic exact tone). May be added later if quality isn't enough.
- Long-term memory of past drafts / sent emails (privacy and storage concerns).
- Outlook classic for Mac (Microsoft itself is sunsetting this; new Outlook for Mac is the path).
- Auto-send. The add-in only drafts and inserts; the user always reviews and clicks Send.

## 3. User Flow

### 3.1 Reply Mode (most common)
1. Ali opens an email in Outlook.
2. Clicks `🤖 AI Reply` button in the ribbon (or message surface).
3. Sidebar (taskpane) opens on the right.
4. The add-in auto-reads the current email body & sender.
5. Default tone settings load from his profile (Pro / Medium / Auto-match).
6. Ali optionally types quick notes ("tell her v3 is good but logo is too small, send v4 by Friday").
7. Clicks **Generate Reply**.
8. Draft appears in ~2-4 seconds.
9. Refines via buttons OR types instructions ("make it shorter and warmer").
10. Clicks **Insert into Reply** → Outlook opens compose with the draft pre-filled.
11. Reviews, edits if needed, clicks Send (his existing email signature attaches).

### 3.2 Compose Mode (new email from scratch)
1. Ali clicks `🤖 AI Compose` from the Outlook ribbon (or switches mode in sidebar).
2. Fills in:
   - **To:** recipient name/email
   - **About:** short topic line ("launch deck v4 timeline")
   - **Notes:** what he wants to say in plain language
3. Picks tone settings.
4. Clicks **Generate**.
5. Draft returns with **Subject** + **Body**.
6. Refines or inserts into a new compose window.

### 3.3 Refine
After any draft generates:
- 8 one-click buttons: `More formal`, `Friendlier`, `Shorter`, `Longer`, `More polite`, `More aggressive`, `Less corporate`, `Add humor`
- Plus a free-text box: "Tell me what to change".
- Plus translate buttons: `→ Arabic`, `→ English`.
- Each click sends the previous draft + the refinement instruction to the AI and replaces the draft inline.

## 4. Architecture

```
┌──────────────────────────────────────────────────────┐
│ Outlook (web / Mac / iOS / Android)                  │
│                                                      │
│  ┌─ Taskpane (sidebar) ─────────────────────────┐    │
│  │ - HTML + vanilla JS + CSS                    │    │
│  │ - Office.js (read email, insert into compose)│    │
│  │ - localStorage (profile + settings)          │    │
│  │ - sessionStorage (current draft + refines)   │    │
│  └──────────────────┬───────────────────────────┘    │
│                     │ HTTPS                          │
└─────────────────────┼────────────────────────────────┘
                      │
                      ▼
       ┌─────────────────────────────────┐
       │ Vercel serverless function      │
       │ (Node.js, free tier)            │
       │                                 │
       │ POST /api/generate              │
       │  → Calls Gemini API             │
       │  → Returns draft text           │
       │                                 │
       │ POST /api/refine                │
       │  → Calls Gemini with prev draft │
       │  → Returns refined text         │
       └────────────┬────────────────────┘
                    │
                    ▼
       ┌─────────────────────────────────┐
       │ Google Gemini API                │
       │  - Gemini 2.5 Flash (default)    │
       │  - Gemini 2.5 Pro (high-stakes)  │
       │  - Free tier: 1,500/day Flash    │
       └─────────────────────────────────┘
```

## 5. Tech Stack

| Layer | Choice | Why |
|---|---|---|
| Add-in framework | Office.js (Microsoft's standard) | Native Outlook integration, all platforms |
| Frontend | HTML + vanilla JS + CSS | No framework — taskpane is small, Office.js doesn't pair smoothly with React for taskpanes anyway |
| Backend | Vercel serverless (Node.js) | Free tier covers this 100×, hides API key |
| AI | Google Gemini 2.5 Flash (free tier) | 1,500 req/day free, strong English + Arabic |
| Storage | `localStorage` for profile, `sessionStorage` for live session | No DB needed |
| Hosting | Vercel (taskpane HTML + manifest + serverless API) | Free tier; HTTPS auto-managed |
| Repo | GitHub `Madboulyjr/outlook-reply-addin` (private) | Solo project, his existing GitHub |
| CI/Deploy | Vercel auto-deploy from main branch | Push → live |

## 6. The System Prompt

Hardcoded in the backend; sent with every API call:

```
You are drafting email replies and new emails for Ali Madbouly,
Head of Art / Art Director at Nice One (https://niceonesa.com/),
a beauty & lifestyle e-commerce platform in Saudi Arabia. He leads
the in-house brand studio handling all visual work for the Nice One
brand itself.

His emails typically involve: design feedback & critique, asset
approvals, vendor & agency communication, scheduling, brand-work
coordination, internal team comms.

VOICE RULES (always):
- Write naturally — like a real person, not a corporate template.
- No filler ("I hope this email finds you well", "Please don't
  hesitate to reach out") unless the situation truly calls for it.
- Get to the point fast.
- Sign off only with: BR,
  (His email client adds the full signature — never write
   "Best regards, Ali Madbouly" or his name underneath.)
- Match the sender's energy: warm reply to warm email, sharp to sharp.
- If the incoming email is in Arabic and language preference is
  "auto-match", reply in Arabic.

PER REQUEST you'll receive:
- mode: "reply" | "compose" | "refine"
- tone: very_formal | professional | casual | very_casual
- length: short | medium | detailed
- language: english | arabic | auto
- For reply: the incoming email body + sender name
- For compose: recipient + topic + notes
- For refine: the previous draft + the refinement instruction
- Optional notes from Ali

OUTPUT:
- Reply / refine mode → only the body text. No subject. End with "BR,"
- Compose mode → return JSON: { "subject": "...", "body": "..." }
- Never include any commentary outside the email itself.
- If information is missing (a date, price, name), insert
  [PLACEHOLDER] inline and add a single line at the end:
  "⚠️ Missing: [what's needed]"
```

## 7. UI Spec — Taskpane Layout

```
┌─ AI Email Assistant ─────────────────┐
│                                      │
│  Mode: [📨 Reply] [✏️ Compose]       │
│                                      │
│  ─── Reply Mode ───                  │
│  📩 Reading: "Launch deck v3 review" │
│       From: Sara Ahmed               │
│  [Show email body ▾]                 │
│                                      │
│  💬 Notes (optional):                │
│  ┌────────────────────────────────┐  │
│  │                                │  │
│  └────────────────────────────────┘  │
│                                      │
│  Tone:    [Formal][Pro✓][Casual][Ch] │
│  Length:  [Short][Med✓][Detailed]    │
│  Lang:    [EN][AR][Auto✓]            │
│                                      │
│  □ Use Pro Quality (high-stakes)     │
│                                      │
│       [✨ Generate Reply]            │
│                                      │
│  ─── Draft ───                       │
│  ┌────────────────────────────────┐  │
│  │ Hey Sara,                      │  │
│  │                                │  │
│  │ V3 is solid — only flag is...  │  │
│  │                                │  │
│  │ BR,                            │  │
│  └────────────────────────────────┘  │
│                                      │
│  Refine:                             │
│  [More formal] [Friendlier] [Shorter]│
│  [Longer] [More polite] [Aggressive] │
│  [Less corporate] [Add humor]        │
│  [→ Arabic] [→ English]              │
│                                      │
│  Custom: ┌──────────────┐ [Apply]    │
│          └──────────────┘            │
│                                      │
│  [📋 Copy] [📨 Insert into Reply]    │
│                                      │
│  ─────────────                       │
│  [⚙️ Settings]                       │
└──────────────────────────────────────┘
```

### Settings Panel (⚙️)
- Edit profile (name, title, company, role description) — pre-filled with Ali's info on first load
- Default tone / length / language
- Custom voice rules (free text — appended to system prompt)
- Toggle: auto-show email body
- Toggle: confirm before insert
- API endpoint (advanced — defaults to Vercel deployment)

## 8. API Contract (Backend ↔ Taskpane)

### `POST /api/generate`
```json
{
  "mode": "reply" | "compose",
  "tone": "very_formal" | "professional" | "casual" | "very_casual",
  "length": "short" | "medium" | "detailed",
  "language": "english" | "arabic" | "auto",
  "useProQuality": false,
  "customRules": "",            // appended to system prompt if set
  "reply": {                    // present if mode === "reply"
    "senderName": "Sara Ahmed",
    "senderEmail": "sara@...",
    "subject": "Launch deck v3 review",
    "body": "...",
    "notes": "tell her v3 looks good but logo is too small"
  },
  "compose": {                  // present if mode === "compose"
    "to": "Sara Ahmed",
    "topic": "Launch deck v4 timeline",
    "notes": "need final files Friday"
  }
}
```

**Response:**
```json
{
  "ok": true,
  "draft": "Hey Sara,\n\nV3 is solid...\n\nBR,",
  "subject": null,                 // only set in compose mode
  "missingInfo": null,             // string if model flagged anything
  "model": "gemini-2.5-flash",
  "tokensUsed": 412
}
```

### `POST /api/refine`
```json
{
  "previousDraft": "...",
  "instruction": "more formal" | "friendlier" | "shorter" | ... | "<custom string>",
  "tone": "...",                   // current tone settings carried over
  "length": "...",
  "language": "...",
  "useProQuality": false
}
```

**Response:** same shape as `/api/generate`.

### Errors
- `429` rate-limited (Gemini free tier hit) → taskpane shows: "Daily free quota hit. Try again tomorrow or enable Pro Quality."
- `5xx` → taskpane shows: "Generation failed — try again." with retry button.

## 9. Privacy & Data Handling

- Email content **leaves the user's device** when sent to Vercel → Gemini.
- Vercel serverless functions don't persist request bodies by default; we will explicitly disable any logging that captures the request body.
- `localStorage` (profile, settings) stays in the user's browser — never transmitted except as part of the system prompt.
- The Vercel function only accepts requests from the add-in's manifest origin (CORS lock-down).

### Gemini data-use — decision

Google's free Gemini API (AI Studio key) allows Google to use prompts and responses for product improvement. **Ali has reviewed this and accepted it for v1.**

v1 uses the AI Studio free tier directly. The "High-privacy mode" toggle (paid Vertex AI) is **descoped from v1** to keep the build simple. Can be added later if a sensitive thread comes up.

## 10. Sideload Deployment (Personal-Only)

1. Build & deploy the add-in to Vercel (`outlook-reply-addin.vercel.app`).
2. Generate manifest XML pointing at the deployed taskpane.
3. Sideload steps:
   - Open Outlook on the web (`outlook.office.com`) signed in with Ali's Nice One account
   - Settings → Manage Add-ins → My add-ins → Add a custom add-in → From file
   - Upload `manifest.xml`
   - Done. Add-in appears in:
     - Outlook on the web ✅
     - New Outlook for Mac ✅
     - Outlook iOS / Android ✅
4. To update: redeploy. No re-sideload needed (manifest points at hosted URL).

## 11. Build Phases

### Phase 1 — Skeleton (1-2 hours)
- Repo scaffold (`outlook-reply-addin`).
- Manifest XML generated (validated with `office-addin-manifest`).
- Empty taskpane HTML/JS deployed to Vercel.
- Sideload tested — confirm button shows up in Outlook web.

### Phase 2 — Reply mode end-to-end (3-4 hours)
- Taskpane reads current email via Office.js.
- Tone/length/language UI.
- Backend `/api/generate` with Gemini call + system prompt.
- Draft renders in sidebar.
- "Insert into Reply" wired up.

### Phase 3 — Refine + Compose (2-3 hours)
- 8 refine buttons + custom instruction box.
- Compose mode UI + `/api/generate` compose branch.
- Translate buttons.

### Phase 4 — Settings & polish (1-2 hours)
- Settings panel.
- Profile edit form (pre-filled with Ali's info).
- Custom rules textarea.
- Pro Quality toggle.
- Error handling, rate-limit messaging.

### Phase 5 — Mobile + cross-platform test (1 hour)
- Verify on Outlook iOS.
- Verify on Outlook Mac.
- Tweak responsive layout if needed.

**Total estimated: ~8-12 hours of focused work, fits in 1 working day.**

## 12. Future Extensions (not in v1)

- Voice-sample learning — paste past sent emails, AI mimics exact tone.
- Local Ollama fallback for full privacy.
- Per-contact memory (auto-detect "I always reply to Sara more casually").
- Templates ("New project kickoff", "Vendor cost negotiation").
- Outlook org-wide deployment (after IT/Legal approval).
- Side-by-side draft comparison.
- Auto-sync profile from a cloud source.

## 13. Open Questions

None at this point. All scope decisions resolved during brainstorming.

## 14. Approval

- [x] Ali reviewed this spec (2026-04-30)
- [x] Ali approved (2026-04-30)
- [x] Moving to implementation plan (writing-plans skill)
