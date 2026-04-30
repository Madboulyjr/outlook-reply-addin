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
