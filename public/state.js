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
