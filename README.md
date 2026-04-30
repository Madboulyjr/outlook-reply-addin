# Outlook Reply Add-In

Personal Outlook add-in that drafts email replies and new emails in Ali Madbouly's voice using Google Gemini 2.5 Flash.

Made by [MAD Studio](https://beingmad.co).

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

Private — personal use only. © MAD Studio.
