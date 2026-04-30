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
