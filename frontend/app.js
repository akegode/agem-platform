const STORAGE_KEY = 'agem_platform_state_v1';
const MAX_PROFILE_PHOTO_BYTES = 1_500_000;
const HECTARES_PER_ACRE = 0.40468564224;
const ACRES_PER_HECTARE = 1 / HECTARES_PER_ACRE;
const SQFT_PER_HECTARE = 107639.1041671;
const IMPORT_ONBOARDING_SMS_MAX_LENGTH = 500;
const DEFAULT_LANGUAGE = 'en';
const DEFAULT_IMPORT_ONBOARDING_SMS_TEMPLATE_EN =
  'Agem Portal: Hello {{name}}. You are now registered in the AGEM farmer system. Use USSD {{ussd}} (once active) for farmer services.';
const DEFAULT_IMPORT_ONBOARDING_SMS_TEMPLATE_SW =
  'Agem Portal: Habari {{name}}. Umesajiliwa kwenye mfumo wa wakulima wa AGEM. Tumia USSD {{ussd}} (ukishawashwa) kupata huduma.';
const SMS_OWNER_COST_PER_MESSAGE_KES = 0.25;

const API = {
  enabled: false
};

const state = loadState();
const smsPicker = {
  q: '',
  offset: 0,
  limit: 50,
  total: 0,
  rows: [],
  selectedIds: new Set()
};
const aiState = {
  smsDrafts: [],
  qcInsights: [],
  paymentRiskFlags: [],
  paymentRiskReport: null,
  proposals: [],
  proposalsMeta: null,
  briefs: [],
  alerts: [],
  opsTasks: [],
  opsTasksMeta: null,
  knowledgeDocs: [],
  feedbackSummary: null,
  evalRuns: [],
  lastResponseIds: {
    copilot: '',
    'sms-draft': '',
    'qc-intelligence': '',
    'payment-risk': '',
    'executive-brief': ''
  }
};
const msaidiziState = {
  pane: 'auth',
  role: 'guest',
  contextModules: [],
  status: {
    lastSyncAt: '',
    sourceSignature: '',
    moduleCount: 0
  }
};
const owedPicker = {
  rows: [],
  selectedFarmerIds: new Set(),
  meta: {
    period: 'month',
    from: '',
    to: '',
    count: 0,
    totalBalanceKes: 0
  }
};
const paymentRecommendationPicker = {
  rows: [],
  meta: {
    total: 0,
    counts: {
      pending: 0,
      approved: 0,
      rejected: 0
    }
  }
};
let activePaneId = 'overview';

const elements = {
  authShell: document.getElementById('authShell'),
  appShell: document.getElementById('appShell'),
  dashboardMain: document.querySelector('.dashboard-shell main'),
  scrollTopBtn: document.getElementById('scrollTopBtn'),

  tabs: document.querySelectorAll('#tabs button'),
  panes: document.querySelectorAll('.pane'),
  agentsNavBtn: document.querySelector('#tabs button[data-pane="agents"]'),
  agentsPane: document.getElementById('agents'),
  agentsPaneOption: document.querySelector('#paneSelect option[value="agents"]'),

  roleHint: document.getElementById('roleHint'),
  roleSelect: document.getElementById('roleSelect'),
  paneSelect: document.getElementById('paneSelect'),
  menuPanel: document.getElementById('menuPanel'),
  accountToggleBtn: document.getElementById('accountToggleBtn'),
  accountPanel: document.getElementById('accountPanel'),
  accountCloseBtn: document.getElementById('accountCloseBtn'),

  authState: document.getElementById('authState'),
  dashboardWelcome: document.getElementById('dashboardWelcome'),
  loginForm: document.getElementById('loginForm'),
  loginUsername: document.getElementById('loginUsername'),
  loginPassword: document.getElementById('loginPassword'),
  registrationPanel: document.getElementById('registrationPanel'),
  registerForm: document.getElementById('registerForm'),
  registerName: document.getElementById('registerName'),
  registerPhone: document.getElementById('registerPhone'),
  registerNationalId: document.getElementById('registerNationalId'),
  registerLocation: document.getElementById('registerLocation'),
  registerPreferredLanguage: document.getElementById('registerPreferredLanguage'),
  registerAreaHectares: document.getElementById('registerAreaHectares'),
  registerAreaAcres: document.getElementById('registerAreaAcres'),
  registerAreaSquareFeet: document.getElementById('registerAreaSquareFeet'),
  registerAvocadoAreaHectares: document.getElementById('registerAvocadoAreaHectares'),
  registerAvocadoAreaAcres: document.getElementById('registerAvocadoAreaAcres'),
  registerAvocadoAreaSquareFeet: document.getElementById('registerAvocadoAreaSquareFeet'),
  registerPin: document.getElementById('registerPin'),
  registerConfirmPin: document.getElementById('registerConfirmPin'),
  registerMsg: document.getElementById('registerMsg'),
  changePasswordForm: document.getElementById('changePasswordForm'),
  currentPassword: document.getElementById('currentPassword'),
  newPassword: document.getElementById('newPassword'),
  confirmNewPassword: document.getElementById('confirmNewPassword'),
  changePasswordBtn: document.getElementById('changePasswordBtn'),
  changePasswordMsg: document.getElementById('changePasswordMsg'),
  logoutBtn: document.getElementById('logoutBtn'),
  profilePhoto: document.getElementById('profilePhoto'),
  profilePhotoFallback: document.getElementById('profilePhotoFallback'),
  profilePhotoInput: document.getElementById('profilePhotoInput'),
  clearPhotoBtn: document.getElementById('clearPhotoBtn'),
  recoveryPanel: document.getElementById('recoveryPanel'),
  recoveryForm: document.getElementById('recoveryForm'),
  recoveryRole: document.getElementById('recoveryRole'),
  recoveryCode: document.getElementById('recoveryCode'),
  recoveryHelp: document.getElementById('recoveryHelp'),
  recoverUsernameBtn: document.getElementById('recoverUsernameBtn'),
  recoveryNewPassword: document.getElementById('recoveryNewPassword'),
  recoveryConfirmPassword: document.getElementById('recoveryConfirmPassword'),
  recoverPasswordBtn: document.getElementById('recoverPasswordBtn'),
  recoveryMsg: document.getElementById('recoveryMsg'),
  authMsg: document.getElementById('authMsg'),

  seedBtn: document.getElementById('seedBtn'),
  syncStatus: document.getElementById('syncStatus'),
  kpiGrid: document.getElementById('kpiGrid'),
  launchCheck: document.getElementById('launchCheck'),

  farmerForm: document.getElementById('farmerForm'),
  farmerId: document.getElementById('farmerId'),
  farmerName: document.getElementById('farmerName'),
  farmerPhone: document.getElementById('farmerPhone'),
  farmerNationalId: document.getElementById('farmerNationalId'),
  farmerLocation: document.getElementById('farmerLocation'),
  farmerPreferredLanguage: document.getElementById('farmerPreferredLanguage'),
  treeCount: document.getElementById('treeCount'),
  farmerAreaHectares: document.getElementById('farmerAreaHectares'),
  farmerAreaAcres: document.getElementById('farmerAreaAcres'),
  farmerAreaSquareFeet: document.getElementById('farmerAreaSquareFeet'),
  farmerAvocadoAreaHectares: document.getElementById('farmerAvocadoAreaHectares'),
  farmerAvocadoAreaAcres: document.getElementById('farmerAvocadoAreaAcres'),
  farmerAvocadoAreaSquareFeet: document.getElementById('farmerAvocadoAreaSquareFeet'),
  farmerNotes: document.getElementById('farmerNotes'),
  farmerSubmitBtn: document.getElementById('farmerSubmitBtn'),
  farmerCancelEditBtn: document.getElementById('farmerCancelEditBtn'),
  farmerMsg: document.getElementById('farmerMsg'),
  farmerImportForm: document.getElementById('farmerImportForm'),
  farmerImportCard: document.getElementById('farmerImportCard'),
  farmerImportFile: document.getElementById('farmerImportFile'),
  farmerImportNotifyBySms: document.getElementById('farmerImportNotifyBySms'),
  farmerImportSmsTemplate: document.getElementById('farmerImportSmsTemplate'),
  farmerImportBtn: document.getElementById('farmerImportBtn'),
  farmerImportQaBtn: document.getElementById('farmerImportQaBtn'),
  farmerImportMsg: document.getElementById('farmerImportMsg'),
  farmerImportSummary: document.getElementById('farmerImportSummary'),
  farmerImportQaMsg: document.getElementById('farmerImportQaMsg'),
  farmerImportQaWrap: document.getElementById('farmerImportQaWrap'),
  farmerImportErrors: document.getElementById('farmerImportErrors'),
  farmerPinCard: document.getElementById('farmerPinCard'),
  farmerPinForm: document.getElementById('farmerPinForm'),
  farmerPinSearch: document.getElementById('farmerPinSearch'),
  farmerPinSearchBtn: document.getElementById('farmerPinSearchBtn'),
  farmerPinFarmer: document.getElementById('farmerPinFarmer'),
  farmerPinValue: document.getElementById('farmerPinValue'),
  farmerPinConfirm: document.getElementById('farmerPinConfirm'),
  farmerPinSubmitBtn: document.getElementById('farmerPinSubmitBtn'),
  farmerPinGenerateBtn: document.getElementById('farmerPinGenerateBtn'),
  farmerPinMsg: document.getElementById('farmerPinMsg'),
  farmerPinCredentials: document.getElementById('farmerPinCredentials'),
  farmerTableWrap: document.getElementById('farmerTableWrap'),
  agentForm: document.getElementById('agentForm'),
  agentName: document.getElementById('agentName'),
  agentEmail: document.getElementById('agentEmail'),
  agentPhone: document.getElementById('agentPhone'),
  agentMsg: document.getElementById('agentMsg'),
  agentCredentials: document.getElementById('agentCredentials'),
  agentSearch: document.getElementById('agentSearch'),
  agentRefreshBtn: document.getElementById('agentRefreshBtn'),
  agentTableWrap: document.getElementById('agentTableWrap'),

  produceForm: document.getElementById('produceForm'),
  produceFarmer: document.getElementById('produceFarmer'),
  produceVariety: document.getElementById('produceVariety'),
  produceKgs: document.getElementById('produceKgs'),
  produceSampleSize: document.getElementById('produceSampleSize'),
  produceVisualGrade: document.getElementById('produceVisualGrade'),
  produceQcDecision: document.getElementById('produceQcDecision'),
  produceDryMatterPct: document.getElementById('produceDryMatterPct'),
  produceFirmnessValue: document.getElementById('produceFirmnessValue'),
  produceFirmnessUnit: document.getElementById('produceFirmnessUnit'),
  produceAvgWeightG: document.getElementById('produceAvgWeightG'),
  produceSizeCode: document.getElementById('produceSizeCode'),
  produceInspector: document.getElementById('produceInspector'),
  produceNotes: document.getElementById('produceNotes'),
  produceMsg: document.getElementById('produceMsg'),
  produceTableWrap: document.getElementById('produceTableWrap'),
  purchaseForm: document.getElementById('purchaseForm'),
  purchaseFarmer: document.getElementById('purchaseFarmer'),
  purchaseQcRecord: document.getElementById('purchaseQcRecord'),
  purchaseKgs: document.getElementById('purchaseKgs'),
  purchasePricePerKgKes: document.getElementById('purchasePricePerKgKes'),
  purchaseVariety: document.getElementById('purchaseVariety'),
  purchaseSizeCode: document.getElementById('purchaseSizeCode'),
  purchaseBuyer: document.getElementById('purchaseBuyer'),
  purchaseValueKes: document.getElementById('purchaseValueKes'),
  purchaseNotes: document.getElementById('purchaseNotes'),
  purchaseMsg: document.getElementById('purchaseMsg'),
  purchaseTableWrap: document.getElementById('purchaseTableWrap'),

  paymentForm: document.getElementById('paymentForm'),
  paymentFarmer: document.getElementById('paymentFarmer'),
  amount: document.getElementById('amount'),
  mpesaRef: document.getElementById('mpesaRef'),
  paymentStatus: document.getElementById('paymentStatus'),
  paymentNotes: document.getElementById('paymentNotes'),
  mpesaDisburseBtn: document.getElementById('mpesaDisburseBtn'),
  owedPeriod: document.getElementById('owedPeriod'),
  owedFromDate: document.getElementById('owedFromDate'),
  owedToDate: document.getElementById('owedToDate'),
  owedSearch: document.getElementById('owedSearch'),
  owedRefreshBtn: document.getElementById('owedRefreshBtn'),
  owedSelectVisibleBtn: document.getElementById('owedSelectVisibleBtn'),
  owedClearSelectedBtn: document.getElementById('owedClearSelectedBtn'),
  owedPrepareBtn: document.getElementById('owedPrepareBtn'),
  owedPaySelectedBtn: document.getElementById('owedPaySelectedBtn'),
  owedExportBtn: document.getElementById('owedExportBtn'),
  owedExportFormat: document.getElementById('owedExportFormat'),
  owedSummary: document.getElementById('owedSummary'),
  owedTableWrap: document.getElementById('owedTableWrap'),
  owedMsg: document.getElementById('owedMsg'),
  paymentRecommendationCard: document.getElementById('paymentRecommendationCard'),
  paymentRecommendationStatusFilter: document.getElementById('paymentRecommendationStatusFilter'),
  paymentRecommendationSearch: document.getElementById('paymentRecommendationSearch'),
  paymentRecommendationRefreshBtn: document.getElementById('paymentRecommendationRefreshBtn'),
  paymentRecommendationSummary: document.getElementById('paymentRecommendationSummary'),
  paymentRecommendationForm: document.getElementById('paymentRecommendationForm'),
  paymentRecommendationFarmer: document.getElementById('paymentRecommendationFarmer'),
  paymentRecommendationAmount: document.getElementById('paymentRecommendationAmount'),
  paymentRecommendationReason: document.getElementById('paymentRecommendationReason'),
  paymentRecommendationSubmitBtn: document.getElementById('paymentRecommendationSubmitBtn'),
  paymentRecommendationAgentHint: document.getElementById('paymentRecommendationAgentHint'),
  paymentRecommendationTableWrap: document.getElementById('paymentRecommendationTableWrap'),
  paymentRecommendationMsg: document.getElementById('paymentRecommendationMsg'),
  paymentLogCard: document.getElementById('paymentLogCard'),
  paymentMsg: document.getElementById('paymentMsg'),
  paymentTableWrap: document.getElementById('paymentTableWrap'),

  smsForm: document.getElementById('smsForm'),
  smsTargetMode: document.getElementById('smsTargetMode'),
  smsSingleWrap: document.getElementById('smsSingleWrap'),
  smsRecipientPickerWrap: document.getElementById('smsRecipientPickerWrap'),
  smsRecipientSearch: document.getElementById('smsRecipientSearch'),
  smsRecipientSearchBtn: document.getElementById('smsRecipientSearchBtn'),
  smsSelectVisibleBtn: document.getElementById('smsSelectVisibleBtn'),
  smsClearSelectedBtn: document.getElementById('smsClearSelectedBtn'),
  smsPrevPageBtn: document.getElementById('smsPrevPageBtn'),
  smsNextPageBtn: document.getElementById('smsNextPageBtn'),
  smsSelectionInfo: document.getElementById('smsSelectionInfo'),
  smsRecipientListWrap: document.getElementById('smsRecipientListWrap'),
  smsAllModeNotice: document.getElementById('smsAllModeNotice'),
  smsPhone: document.getElementById('smsPhone'),
  smsMessage: document.getElementById('smsMessage'),
  smsDraftPurpose: document.getElementById('smsDraftPurpose'),
  smsDraftLanguage: document.getElementById('smsDraftLanguage'),
  smsDraftAudience: document.getElementById('smsDraftAudience'),
  smsDraftTone: document.getElementById('smsDraftTone'),
  smsDraftMaxLength: document.getElementById('smsDraftMaxLength'),
  smsDraftBtn: document.getElementById('smsDraftBtn'),
  smsDraftMsg: document.getElementById('smsDraftMsg'),
  smsDraftWrap: document.getElementById('smsDraftWrap'),
  smsMsg: document.getElementById('smsMsg'),
  smsTableWrap: document.getElementById('smsTableWrap'),

  reportMetrics: document.getElementById('reportMetrics'),
  reportNarrative: document.getElementById('reportNarrative'),
  agentStatsWrap: document.getElementById('agentStatsWrap'),
  copilotPrompt: document.getElementById('copilotPrompt'),
  copilotAskBtn: document.getElementById('copilotAskBtn'),
  copilotMsg: document.getElementById('copilotMsg'),
  copilotWrap: document.getElementById('copilotWrap'),
  qcAiRecord: document.getElementById('qcAiRecord'),
  qcAiRunBtn: document.getElementById('qcAiRunBtn'),
  qcAiMsg: document.getElementById('qcAiMsg'),
  qcAiWrap: document.getElementById('qcAiWrap'),
  paymentRiskPeriod: document.getElementById('paymentRiskPeriod'),
  paymentRiskFrom: document.getElementById('paymentRiskFrom'),
  paymentRiskTo: document.getElementById('paymentRiskTo'),
  paymentRiskRunBtn: document.getElementById('paymentRiskRunBtn'),
  paymentRiskMsg: document.getElementById('paymentRiskMsg'),
  paymentRiskWrap: document.getElementById('paymentRiskWrap'),
  briefRunNowBtn: document.getElementById('briefRunNowBtn'),
  briefRefreshBtn: document.getElementById('briefRefreshBtn'),
  briefMsg: document.getElementById('briefMsg'),
  briefWrap: document.getElementById('briefWrap'),
  proposalForm: document.getElementById('proposalForm'),
  proposalActionType: document.getElementById('proposalActionType'),
  proposalConfidence: document.getElementById('proposalConfidence'),
  proposalTitle: document.getElementById('proposalTitle'),
  proposalDescription: document.getElementById('proposalDescription'),
  proposalPayload: document.getElementById('proposalPayload'),
  proposalStatusFilter: document.getElementById('proposalStatusFilter'),
  proposalRefreshBtn: document.getElementById('proposalRefreshBtn'),
  proposalMsg: document.getElementById('proposalMsg'),
  proposalWrap: document.getElementById('proposalWrap'),
  opsFromPaymentRiskBtn: document.getElementById('opsFromPaymentRiskBtn'),
  opsFromQcBtn: document.getElementById('opsFromQcBtn'),
  opsRefreshBtn: document.getElementById('opsRefreshBtn'),
  opsTaskStatusFilter: document.getElementById('opsTaskStatusFilter'),
  opsTaskSeverityFilter: document.getElementById('opsTaskSeverityFilter'),
  opsTaskMsg: document.getElementById('opsTaskMsg'),
  opsTaskWrap: document.getElementById('opsTaskWrap'),
  knowledgeForm: document.getElementById('knowledgeForm'),
  knowledgeTitle: document.getElementById('knowledgeTitle'),
  knowledgeSource: document.getElementById('knowledgeSource'),
  knowledgeTags: document.getElementById('knowledgeTags'),
  knowledgeContent: document.getElementById('knowledgeContent'),
  knowledgeSearch: document.getElementById('knowledgeSearch'),
  knowledgeSearchBtn: document.getElementById('knowledgeSearchBtn'),
  knowledgeMsg: document.getElementById('knowledgeMsg'),
  knowledgeWrap: document.getElementById('knowledgeWrap'),
  aiFeedbackForm: document.getElementById('aiFeedbackForm'),
  aiFeedbackTool: document.getElementById('aiFeedbackTool'),
  aiFeedbackRating: document.getElementById('aiFeedbackRating'),
  aiFeedbackResponseId: document.getElementById('aiFeedbackResponseId'),
  aiFeedbackNote: document.getElementById('aiFeedbackNote'),
  aiFeedbackSummaryBtn: document.getElementById('aiFeedbackSummaryBtn'),
  aiEvalRunBtn: document.getElementById('aiEvalRunBtn'),
  aiEvalRefreshBtn: document.getElementById('aiEvalRefreshBtn'),
  aiFeedbackMsg: document.getElementById('aiFeedbackMsg'),
  aiFeedbackWrap: document.getElementById('aiFeedbackWrap'),

  exportButtons: document.querySelectorAll('.export-btn'),
  exportRange: document.getElementById('exportRange'),
  exportFrom: document.getElementById('exportFrom'),
  exportTo: document.getElementById('exportTo'),
  exportFormat: document.getElementById('exportFormat'),
  backupBtn: document.getElementById('backupBtn'),
  listBackupsBtn: document.getElementById('listBackupsBtn'),
  resetBtn: document.getElementById('resetBtn'),
  exportsMsg: document.getElementById('exportsMsg'),
  backupListWrap: document.getElementById('backupListWrap'),

  msaidiziFab: document.getElementById('msaidiziFab'),
  msaidiziPanel: document.getElementById('msaidiziPanel'),
  msaidiziCloseBtn: document.getElementById('msaidiziCloseBtn'),
  msaidiziMode: document.getElementById('msaidiziMode'),
  msaidiziSyncBtn: document.getElementById('msaidiziSyncBtn'),
  msaidiziQuestion: document.getElementById('msaidiziQuestion'),
  msaidiziAskBtn: document.getElementById('msaidiziAskBtn'),
  msaidiziMsg: document.getElementById('msaidiziMsg'),
  msaidiziSource: document.getElementById('msaidiziSource'),
  msaidiziContextWrap: document.getElementById('msaidiziContextWrap'),
  msaidiziAnswerShort: document.getElementById('msaidiziAnswerShort'),
  msaidiziSteps: document.getElementById('msaidiziSteps'),
  msaidiziTroubleshoot: document.getElementById('msaidiziTroubleshoot'),
  msaidiziDoNow: document.getElementById('msaidiziDoNow'),
  msaidiziAvoid: document.getElementById('msaidiziAvoid'),
  msaidiziWaitFor: document.getElementById('msaidiziWaitFor')
};

void init();

async function init() {
  bindTabs();
  bindRoleSelect();
  bindAccountSettings();
  bindAuth();
  bindAreaConverters();
  bindScrollTools();
  bindFarmers();
  bindAgents();
  bindProduce();
  bindProducePurchases();
  bindPayments();
  bindSms();
  bindExports();
  bindAiTools();
  bindMsaidizi();

  hydrateFarmerSelectors();
  syncCurrentUserPhotoFromState();
  updateAuthUi();
  renderAll();

  await checkApiConnectivity();

  if (API.enabled && state.auth.token) {
    const restored = await restoreSession();
    if (!restored) {
      notifySync('Session expired. Please sign in again.');
    }
  }

  if (API.enabled && isAuthenticated()) {
    await fetchAllData();
  } else if (API.enabled) {
    notifySync('Connected to backend. Sign in to access live data.');
  } else {
    notifySync('Backend offline. Running in local mode.');
  }

  await refreshOwedRows(false);
  await loadPaymentRecommendations();
  await refreshMsaidiziContext();
  updatePermissionUi();
}

function bindAccountSettings() {
  if (!elements.accountToggleBtn || !elements.accountPanel) return;

  elements.accountToggleBtn.addEventListener('click', () => {
    if (elements.accountPanel.hidden) {
      showAccountPanel();
    } else {
      hideAccountPanel();
    }
  });

  if (elements.accountCloseBtn) {
    elements.accountCloseBtn.addEventListener('click', () => {
      hideAccountPanel();
    });
  }
}

function bindAreaConverters() {
  setupAreaConverterGroup({
    hectaresInput: elements.registerAreaHectares,
    acresInput: elements.registerAreaAcres,
    squareFeetInput: elements.registerAreaSquareFeet
  });
  setupAreaConverterGroup({
    hectaresInput: elements.registerAvocadoAreaHectares,
    acresInput: elements.registerAvocadoAreaAcres,
    squareFeetInput: elements.registerAvocadoAreaSquareFeet
  });
  setupAreaConverterGroup({
    hectaresInput: elements.farmerAreaHectares,
    acresInput: elements.farmerAreaAcres,
    squareFeetInput: elements.farmerAreaSquareFeet
  });
  setupAreaConverterGroup({
    hectaresInput: elements.farmerAvocadoAreaHectares,
    acresInput: elements.farmerAvocadoAreaAcres,
    squareFeetInput: elements.farmerAvocadoAreaSquareFeet
  });
}

function setupAreaConverterGroup(group) {
  if (!group.hectaresInput || !group.acresInput || !group.squareFeetInput) return;

  let syncing = false;

  const syncFrom = (source) => {
    if (syncing) return;
    syncing = true;

    const sourceValue =
      source === 'hectares'
        ? group.hectaresInput.value
        : source === 'acres'
          ? group.acresInput.value
          : group.squareFeetInput.value;

    const sourceParsed = parseAreaClient(sourceValue);
    if (!String(sourceValue || '').trim()) {
      group.hectaresInput.value = '';
      group.acresInput.value = '';
      group.squareFeetInput.value = '';
      syncing = false;
      return;
    }

    if (!Number.isFinite(sourceParsed) || sourceParsed <= 0) {
      if (source !== 'hectares') group.hectaresInput.value = '';
      if (source !== 'acres') group.acresInput.value = '';
      if (source !== 'squareFeet') group.squareFeetInput.value = '';
      syncing = false;
      return;
    }

    const hectares =
      source === 'hectares'
        ? sourceParsed
        : source === 'acres'
          ? sourceParsed * HECTARES_PER_ACRE
          : sourceParsed / SQFT_PER_HECTARE;

    const metrics = areaMetricsFromHectares(hectares);

    if (source !== 'hectares') {
      group.hectaresInput.value = formatAreaForInput(metrics.hectares, 3);
    }
    if (source !== 'acres') {
      group.acresInput.value = formatAreaForInput(metrics.acres, 3);
    }
    if (source !== 'squareFeet') {
      group.squareFeetInput.value = formatAreaForInput(metrics.squareFeet, 1);
    }

    syncing = false;
  };

  group.hectaresInput.addEventListener('input', () => syncFrom('hectares'));
  group.acresInput.addEventListener('input', () => syncFrom('acres'));
  group.squareFeetInput.addEventListener('input', () => syncFrom('squareFeet'));

  group.hectaresInput.addEventListener('blur', () => syncFrom('hectares'));
  group.acresInput.addEventListener('blur', () => syncFrom('acres'));
  group.squareFeetInput.addEventListener('blur', () => syncFrom('squareFeet'));
}

function showAccountPanel() {
  if (!elements.accountPanel || !elements.accountToggleBtn) return;
  elements.accountPanel.hidden = false;
  elements.accountToggleBtn.setAttribute('aria-expanded', 'true');
  elements.accountToggleBtn.textContent = 'Hide Account Settings';
}

function hideAccountPanel() {
  if (!elements.accountPanel || !elements.accountToggleBtn) return;
  elements.accountPanel.hidden = true;
  elements.accountToggleBtn.setAttribute('aria-expanded', 'false');
  elements.accountToggleBtn.textContent = 'Account Settings';
}

function bindScrollTools() {
  if (!elements.dashboardMain || !elements.scrollTopBtn) return;

  elements.dashboardMain.addEventListener('scroll', () => {
    elements.scrollTopBtn.hidden = elements.dashboardMain.scrollTop < 260;
  });

  elements.scrollTopBtn.addEventListener('click', () => {
    elements.dashboardMain.scrollTo({ top: 0, behavior: 'smooth' });
  });
}

function msaidiziRole() {
  return isAuthenticated() ? currentRole() : 'guest';
}

function msaidiziPane() {
  return isAuthenticated() ? activePaneId || 'overview' : 'auth';
}

function setMsaidiziOpen(nextOpen) {
  if (!elements.msaidiziPanel || !elements.msaidiziFab) return;
  elements.msaidiziPanel.hidden = !nextOpen;
  elements.msaidiziFab.setAttribute('aria-expanded', nextOpen ? 'true' : 'false');
}

function renderMsaidiziList(target, rows = [], ordered = false) {
  if (!target) return;
  target.innerHTML = '';
  const list = Array.isArray(rows) ? rows.filter(Boolean) : [];
  if (!list.length) {
    const node = document.createElement('li');
    node.textContent = ordered ? 'No steps yet.' : 'No items.';
    target.appendChild(node);
    return;
  }
  list.forEach((row) => {
    const li = document.createElement('li');
    li.textContent = row;
    target.appendChild(li);
  });
}

function renderMsaidiziContext(modules = []) {
  if (!elements.msaidiziContextWrap) return;
  const rows = Array.isArray(modules) ? modules : [];
  if (!rows.length) {
    elements.msaidiziContextWrap.classList.add('empty');
    elements.msaidiziContextWrap.textContent = 'No context tips found for this page yet.';
    return;
  }
  elements.msaidiziContextWrap.classList.remove('empty');
  elements.msaidiziContextWrap.innerHTML = rows
    .map((module) => {
      const steps = Array.isArray(module.steps) ? module.steps.slice(0, 3) : [];
      const stepsHtml = steps.length
        ? `<ol class="msaidizi-list ordered">${steps.map((step) => `<li>${escapeHtml(step)}</li>`).join('')}</ol>`
        : '';
      return `
        <article class="knowledge-card">
          <div class="ai-result-title">${escapeHtml(module.title || 'Instruction module')}</div>
          <p class="meta compact">${escapeHtml(module.short || '')}</p>
          ${stepsHtml}
        </article>
      `;
    })
    .join('');
}

function renderMsaidiziAnswer(answer = {}) {
  if (elements.msaidiziAnswerShort) {
    elements.msaidiziAnswerShort.textContent = String(answer.answerShort || '').trim() || 'No answer generated yet.';
  }
  renderMsaidiziList(elements.msaidiziSteps, answer.steps || [], true);
  renderMsaidiziList(elements.msaidiziTroubleshoot, answer.troubleshoot || []);
  renderMsaidiziList(elements.msaidiziDoNow, answer.doNow || []);
  renderMsaidiziList(elements.msaidiziAvoid, answer.avoid || []);
  renderMsaidiziList(elements.msaidiziWaitFor, answer.waitFor || []);
}

async function refreshMsaidiziStatus() {
  if (!API.enabled || !elements.msaidiziSource) {
    if (elements.msaidiziSource) elements.msaidiziSource.textContent = 'Msaidizi needs backend API connectivity.';
    return;
  }
  try {
    const response = await apiRequest('/api/ai/msaidizi/status', { auth: false });
    const data = response.data || {};
    msaidiziState.status = {
      lastSyncAt: data.lastSyncAt || '',
      sourceSignature: data.sourceSignature || '',
      moduleCount: Number(data.moduleCount) || 0
    };
    const syncLabel = data.lastSyncAt ? formatDate(data.lastSyncAt) : 'not synced yet';
    elements.msaidiziSource.textContent = `Modules: ${msaidiziState.status.moduleCount} | Last sync: ${syncLabel}`;
  } catch (error) {
    elements.msaidiziSource.textContent = error.message;
  }
}

async function refreshMsaidiziContext() {
  msaidiziState.role = msaidiziRole();
  msaidiziState.pane = msaidiziPane();
  if (!API.enabled) {
    renderMsaidiziContext([]);
    if (elements.msaidiziSource) {
      elements.msaidiziSource.textContent = 'Msaidizi needs backend API connectivity.';
    }
    return;
  }

  try {
    const params = new URLSearchParams({
      pane: msaidiziState.pane,
      role: msaidiziState.role,
      limit: '4'
    });
    const response = await apiRequest(`/api/ai/msaidizi/context?${params.toString()}`, { auth: false });
    const data = response.data || {};
    msaidiziState.contextModules = Array.isArray(data.modules) ? data.modules : [];
    renderMsaidiziContext(msaidiziState.contextModules);
    const syncLabel = data.lastSyncAt ? formatDate(data.lastSyncAt) : 'not synced yet';
    if (elements.msaidiziSource) {
      elements.msaidiziSource.textContent = `Context: ${msaidiziState.pane} | Last sync: ${syncLabel}`;
    }
  } catch (error) {
    renderMsaidiziContext([]);
    if (elements.msaidiziSource) elements.msaidiziSource.textContent = error.message;
  }
}

async function askMsaidizi() {
  if (!API.enabled) {
    if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = 'Msaidizi needs backend API connectivity.';
    return;
  }
  const question = String(elements.msaidiziQuestion?.value || '').trim();
  const mode = String(elements.msaidiziMode?.value || 'normal').trim() || 'normal';
  const pane = msaidiziPane();
  if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = 'Generating answer...';

  try {
    const response = await apiRequest('/api/ai/msaidizi/ask', {
      method: 'POST',
      auth: false,
      body: {
        question,
        mode,
        pane,
        role: msaidiziRole()
      }
    });
    const data = response.data || {};
    const answerModules = Array.isArray(data.modules) ? data.modules : [];
    const contextModules = Array.isArray(data.contextModules) ? data.contextModules : msaidiziState.contextModules;
    renderMsaidiziAnswer(data.answer || {});
    renderMsaidiziContext(contextModules || []);
    const sourceLabel = data.source === 'openai' ? `OpenAI (${data.model || 'model'})` : 'Local rules';
    const syncLabel = data.lastSyncAt ? formatDate(data.lastSyncAt) : 'not synced yet';
    const moduleLabel = answerModules.length
      ? ` | Answer modules: ${answerModules.map((module) => module.title).slice(0, 2).join(', ')}`
      : '';
    if (elements.msaidiziSource) {
      elements.msaidiziSource.textContent = `${sourceLabel} | Context page: ${pane} | Last sync: ${syncLabel}${moduleLabel}`;
    }
    if (elements.msaidiziMsg) {
      elements.msaidiziMsg.textContent = data.warning || 'Msaidizi answer ready.';
    }
  } catch (error) {
    if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = error.message;
  }
}

async function syncMsaidiziDocs(force = true) {
  if (!API.enabled) {
    if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = 'Backend API is required for sync.';
    return;
  }
  if (!isAuthenticated() || currentRole() !== 'admin') {
    if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = 'Only admin can sync documentation modules.';
    return;
  }
  if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = 'Syncing documentation...';
  try {
    const response = await apiRequest('/api/ai/msaidizi/sync', {
      method: 'POST',
      body: {
        force,
        reason: 'manual_ui_sync'
      }
    });
    const data = response.data || {};
    if (elements.msaidiziMsg) {
      elements.msaidiziMsg.textContent = data.updated
        ? `Msaidizi synced (${data.moduleCount || 0} modules).`
        : 'No documentation changes detected.';
    }
    await refreshMsaidiziStatus();
    await refreshMsaidiziContext();
  } catch (error) {
    if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = error.message;
  }
}

function bindMsaidizi() {
  if (!elements.msaidiziFab || !elements.msaidiziPanel) return;

  elements.msaidiziFab.addEventListener('click', async () => {
    const willOpen = elements.msaidiziPanel.hidden;
    setMsaidiziOpen(willOpen);
    if (!willOpen) return;
    await refreshMsaidiziStatus();
    await refreshMsaidiziContext();
  });

  if (elements.msaidiziCloseBtn) {
    elements.msaidiziCloseBtn.addEventListener('click', () => {
      setMsaidiziOpen(false);
    });
  }

  if (elements.msaidiziAskBtn) {
    elements.msaidiziAskBtn.addEventListener('click', async () => {
      await askMsaidizi();
    });
  }

  if (elements.msaidiziQuestion) {
    elements.msaidiziQuestion.addEventListener('keydown', async (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      await askMsaidizi();
    });
  }

  if (elements.msaidiziMode) {
    elements.msaidiziMode.addEventListener('change', async () => {
      if (!elements.msaidiziQuestion?.value.trim()) return;
      await askMsaidizi();
    });
  }

  if (elements.msaidiziSyncBtn) {
    elements.msaidiziSyncBtn.addEventListener('click', async () => {
      await syncMsaidiziDocs(true);
    });
  }
}

function bindTabs() {
  elements.tabs.forEach((button) => {
    button.addEventListener('click', () => {
      setActivePane(button.dataset.pane);
    });
  });

  if (elements.paneSelect) {
    elements.paneSelect.addEventListener('change', (event) => {
      setActivePane(event.target.value, true);
    });
  }

  const initiallyActiveButton = [...elements.tabs].find((button) => button.classList.contains('active'));
  setActivePane(initiallyActiveButton?.dataset.pane || 'overview');
}

function setActivePane(paneId, resetScroll = false) {
  if (!paneId) return;

  if (paneId === 'agents' && currentRole() !== 'admin') {
    paneId = 'overview';
  }

  let activeFound = false;
  elements.tabs.forEach((button) => {
    const isActive = button.dataset.pane === paneId;
    button.classList.toggle('active', isActive);
    if (isActive) activeFound = true;
  });

  elements.panes.forEach((pane) => {
    pane.classList.toggle('active', pane.id === paneId);
  });

  if (!activeFound) return;
  activePaneId = paneId;

  if (elements.paneSelect) {
    elements.paneSelect.value = paneId;
  }

  if (elements.menuPanel) {
    elements.menuPanel.open = false;
  }

  if (resetScroll && elements.dashboardMain) {
    elements.dashboardMain.scrollTo({ top: 0, behavior: 'smooth' });
  }

  void refreshMsaidiziContext();
}

function bindRoleSelect() {
  elements.roleSelect.value = state.role;
  elements.roleSelect.addEventListener('change', async (event) => {
    state.role = event.target.value;
    persist();
    renderAll();
    await refreshOwedRows(true);
    await loadPaymentRecommendations();
    updatePermissionUi();
  });
}

function bindAuth() {
  elements.loginForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    if (!API.enabled) {
      elements.authMsg.textContent = 'Backend is offline. Sign-in requires live API.';
      return;
    }

    const identity = elements.loginUsername.value.trim();
    const secret = elements.loginPassword.value;

    if (!identity || !secret) {
      elements.authMsg.textContent = 'Username/phone and password/PIN are required.';
      return;
    }

    try {
      const response = await apiRequest('/api/auth/login', {
        method: 'POST',
        body: { username: identity, password: secret },
        auth: false
      });

      state.auth = {
        token: response.data.token,
        user: response.data.user,
        expiresAt: response.data.expiresAt
      };
      state.role = response.data.user.role;
      syncCurrentUserPhotoFromState();

      elements.loginForm.reset();
      elements.authMsg.textContent = `Signed in as ${response.data.user.username}.`;
      persist();

      updateAuthUi();
      setActivePane('overview', true);
      await fetchAllData();
      updatePermissionUi();
    } catch (error) {
      elements.authMsg.textContent = error.message;
    }
  });

  elements.registerForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    if (!API.enabled) {
      elements.registerMsg.textContent = 'Registration requires backend mode.';
      return;
    }
    if (isAuthenticated()) {
      elements.registerMsg.textContent = 'Sign out first if you want to register a new account.';
      return;
    }

    const name = elements.registerName.value.trim();
    const phone = cleanPhone(elements.registerPhone.value);
    const nationalId = elements.registerNationalId.value.trim();
    const location = elements.registerLocation.value.trim();
    const preferredLanguage = elements.registerPreferredLanguage.value.trim() || DEFAULT_LANGUAGE;
    const areaHectares = elements.registerAreaHectares.value.trim();
    const areaAcres = elements.registerAreaAcres.value.trim();
    const areaSquareFeet = elements.registerAreaSquareFeet.value.trim();
    const avocadoAreaHectares = elements.registerAvocadoAreaHectares.value.trim();
    const avocadoAreaAcres = elements.registerAvocadoAreaAcres.value.trim();
    const avocadoAreaSquareFeet = elements.registerAvocadoAreaSquareFeet.value.trim();
    const pin = elements.registerPin.value.trim();
    const confirmPin = elements.registerConfirmPin.value.trim();

    const totalAreaPayload = buildAreaPayload(areaHectares, areaAcres, areaSquareFeet);
    const avocadoAreaPayload = buildAreaPayload(avocadoAreaHectares, avocadoAreaAcres, avocadoAreaSquareFeet);
    const totalHectares = Number(totalAreaPayload.hectares);
    const avocadoHectares = Number(avocadoAreaPayload.hectares);

    if (
      !name ||
      !phone ||
      !nationalId ||
      !location ||
      (!areaHectares && !areaAcres && !areaSquareFeet) ||
      (!avocadoAreaHectares && !avocadoAreaAcres && !avocadoAreaSquareFeet) ||
      !pin ||
      !confirmPin
    ) {
      elements.registerMsg.textContent = 'All registration fields are required.';
      return;
    }
    if (!/^\d{4}$/.test(pin)) {
      elements.registerMsg.textContent = 'PIN must be exactly 4 digits.';
      return;
    }
    if (pin !== confirmPin) {
      elements.registerMsg.textContent = 'PIN confirmation does not match.';
      return;
    }
    if (!Number.isFinite(totalHectares) || totalHectares <= 0) {
      elements.registerMsg.textContent = 'Total farm area must be greater than 0.';
      return;
    }
    if (!Number.isFinite(avocadoHectares) || avocadoHectares <= 0) {
      elements.registerMsg.textContent = 'Area under avocado must be greater than 0.';
      return;
    }
    if (avocadoHectares > totalHectares) {
      elements.registerMsg.textContent = 'Area under avocado cannot be greater than total farm size.';
      return;
    }

    try {
      const response = await apiRequest('/api/auth/register-farmer', {
        method: 'POST',
        auth: false,
        body: {
          name,
          phone,
          nationalId,
          location,
          preferredLanguage,
          ...totalAreaPayload,
          avocadoHectares: avocadoAreaPayload.hectares,
          avocadoAcres: avocadoAreaPayload.acres,
          avocadoSquareFeet: avocadoAreaPayload.squareFeet,
          pin,
          confirmPin
        }
      });

      state.auth = {
        token: response.data.token,
        user: response.data.user,
        expiresAt: response.data.expiresAt
      };
      state.role = response.data.user.role;
      syncCurrentUserPhotoFromState();
      persist();

      elements.registerForm.reset();
      updateAuthUi();
      setActivePane('overview', true);
      await fetchAllData();
      updatePermissionUi();

      elements.authMsg.textContent = `Welcome ${response.data.user.name}. Your account is active.`;
    } catch (error) {
      elements.registerMsg.textContent = error.message;
    }
  });

  elements.changePasswordForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    if (!API.enabled) {
      elements.changePasswordMsg.textContent = 'Password updates require backend mode.';
      return;
    }
    if (!isAuthenticated()) {
      elements.changePasswordMsg.textContent = 'Sign in first to change your password.';
      return;
    }

    const currentPassword = elements.currentPassword.value;
    const newPassword = elements.newPassword.value;
    const confirmPassword = elements.confirmNewPassword.value;
    const updatingPin = isAuthenticated() && state.auth.user.role === 'farmer';

    if (!currentPassword || !newPassword || !confirmPassword) {
      elements.changePasswordMsg.textContent = updatingPin
        ? 'All PIN fields are required.'
        : 'All password fields are required.';
      return;
    }

    try {
      await apiRequest('/api/auth/change-password', {
        method: 'POST',
        body: { currentPassword, newPassword, confirmPassword }
      });

      elements.changePasswordForm.reset();
      elements.changePasswordMsg.textContent = updatingPin
        ? 'PIN updated successfully.'
        : 'Password updated successfully.';
    } catch (error) {
      elements.changePasswordMsg.textContent = error.message;
    }
  });

  const updateRecoveryUiForRole = () => {
    if (elements.recoveryCode) {
      elements.recoveryCode.placeholder = 'Recovery code';
    }
    if (elements.recoveryHelp) {
      elements.recoveryHelp.textContent = 'Use the recovery code for this account type.';
    }
    if (elements.recoveryNewPassword) {
      elements.recoveryNewPassword.placeholder = 'New password for reset';
    }
    if (elements.recoveryConfirmPassword) {
      elements.recoveryConfirmPassword.placeholder = 'Confirm new password';
    }
    if (elements.recoverPasswordBtn) {
      elements.recoverPasswordBtn.textContent = 'Reset Password';
    }
  };
  updateRecoveryUiForRole();
  elements.recoveryRole?.addEventListener('change', updateRecoveryUiForRole);

  elements.recoverUsernameBtn.addEventListener('click', async () => {
    clearMessages();

    if (!API.enabled) {
      elements.recoveryMsg.textContent = 'Recovery requires backend mode.';
      return;
    }

    const role = elements.recoveryRole.value;
    const recoveryCode = elements.recoveryCode.value.trim();

    if (!recoveryCode) {
      elements.recoveryMsg.textContent = 'Recovery code is required.';
      return;
    }

    try {
      const response = await apiRequest('/api/auth/recover-username', {
        method: 'POST',
        auth: false,
        body: { role, recoveryCode }
      });

      elements.loginUsername.value = response.data.username;
      elements.recoveryMsg.textContent = `Username recovered: ${response.data.username}`;
    } catch (error) {
      elements.recoveryMsg.textContent = error.message;
    }
  });

  elements.recoverPasswordBtn.addEventListener('click', async () => {
    clearMessages();

    if (!API.enabled) {
      elements.recoveryMsg.textContent = 'Recovery requires backend mode.';
      return;
    }

    const role = elements.recoveryRole.value;
    const recoveryCode = elements.recoveryCode.value.trim();
    const newPassword = elements.recoveryNewPassword.value;
    const confirmPassword = elements.recoveryConfirmPassword.value;

    if (!recoveryCode || !newPassword || !confirmPassword) {
      elements.recoveryMsg.textContent = 'Recovery code and both password fields are required.';
      return;
    }

    try {
      const response = await apiRequest('/api/auth/recover-password', {
        method: 'POST',
        auth: false,
        body: { role, recoveryCode, newPassword, confirmPassword }
      });

      elements.recoveryNewPassword.value = '';
      elements.recoveryConfirmPassword.value = '';
      elements.loginUsername.value = response.data.username;
      elements.recoveryMsg.textContent = `Password reset complete for ${response.data.username}.`;
    } catch (error) {
      elements.recoveryMsg.textContent = error.message;
    }
  });

  elements.logoutBtn.addEventListener('click', async () => {
    clearMessages();

    if (API.enabled && state.auth.token) {
      try {
        await apiRequest('/api/auth/logout', { method: 'POST' });
      } catch {
        // Ignore remote logout failures and clear client state.
      }
    }

    state.auth = { token: '', user: null, expiresAt: '' };
    persist();
    elements.changePasswordForm.reset();
    elements.profilePhotoInput.value = '';
    smsPicker.selectedIds.clear();
    smsPicker.rows = [];
    smsPicker.total = 0;
    smsPicker.offset = 0;
    state.agents = [];
    clearAgentCredentialDisplay();
    elements.registrationPanel.open = false;
    elements.recoveryPanel.open = false;

    updateAuthUi();
    updatePermissionUi();

    elements.authMsg.textContent = 'Signed out.';
    if (API.enabled) {
      notifySync('Connected to backend. Sign in to continue.');
    } else {
      notifySync('Signed out. Local mode active.');
    }
  });

  elements.profilePhotoInput.addEventListener('change', async (event) => {
    clearMessages();
    if (!isAuthenticated()) {
      elements.authMsg.textContent = 'Sign in to upload a profile photo.';
      event.target.value = '';
      return;
    }

    const [file] = event.target.files || [];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      elements.changePasswordMsg.textContent = 'Please choose an image file (PNG, JPG, or GIF).';
      event.target.value = '';
      return;
    }

    if (file.size > MAX_PROFILE_PHOTO_BYTES) {
      elements.changePasswordMsg.textContent = 'Image is too large. Please upload a file under 1.5MB.';
      event.target.value = '';
      return;
    }

    try {
      const dataUrl = await readFileAsDataUrl(file);
      setCurrentUserPhoto(dataUrl);
      renderDashboardAccount();
      elements.changePasswordMsg.textContent = 'Profile photo updated.';
    } catch {
      elements.changePasswordMsg.textContent = 'Could not read this image. Please try another file.';
    } finally {
      event.target.value = '';
    }
  });

  elements.clearPhotoBtn.addEventListener('click', () => {
    clearMessages();
    if (!isAuthenticated()) return;
    setCurrentUserPhoto('');
    renderDashboardAccount();
    elements.changePasswordMsg.textContent = 'Profile photo removed.';
  });
}

function bindFarmers() {
  elements.farmerForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    const role = currentRole();
    if (!['admin', 'agent'].includes(role)) {
      elements.farmerMsg.textContent = 'Only admin or agent can manage farmers.';
      return;
    }

    const totalAreaPayload = buildAreaPayload(
      elements.farmerAreaHectares.value.trim(),
      elements.farmerAreaAcres.value.trim(),
      elements.farmerAreaSquareFeet.value.trim()
    );
    const avocadoAreaPayload = buildAreaPayload(
      elements.farmerAvocadoAreaHectares.value.trim(),
      elements.farmerAvocadoAreaAcres.value.trim(),
      elements.farmerAvocadoAreaSquareFeet.value.trim()
    );
    const totalHectares = Number(totalAreaPayload.hectares);
    const avocadoHectares = Number(avocadoAreaPayload.hectares);

    const payload = {
      name: elements.farmerName.value.trim(),
      phone: cleanPhone(elements.farmerPhone.value),
      nationalId: cleanNationalIdClient(elements.farmerNationalId.value),
      location: elements.farmerLocation.value.trim(),
      preferredLanguage: normalizeLanguageClient(elements.farmerPreferredLanguage.value) || DEFAULT_LANGUAGE,
      trees: Number(elements.treeCount.value || 0),
      ...totalAreaPayload,
      avocadoHectares: avocadoAreaPayload.hectares,
      avocadoAcres: avocadoAreaPayload.acres,
      avocadoSquareFeet: avocadoAreaPayload.squareFeet,
      notes: elements.farmerNotes.value.trim()
    };

    if (!Number.isFinite(totalHectares) || totalHectares <= 0) {
      elements.farmerMsg.textContent = 'Total farm area must be greater than 0.';
      return;
    }
    if (!Number.isFinite(avocadoHectares) || avocadoHectares <= 0) {
      elements.farmerMsg.textContent = 'Area under avocado must be greater than 0.';
      return;
    }
    if (avocadoHectares > totalHectares) {
      elements.farmerMsg.textContent = 'Area under avocado cannot be greater than total farm size.';
      return;
    }

    try {
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.farmerMsg.textContent = 'Sign in first to use backend mode.';
          return;
        }

        if (state.editingFarmerId) {
          await apiRequest(`/api/farmers/${encodeURIComponent(state.editingFarmerId)}`, {
            method: 'PATCH',
            body: payload
          });
          elements.farmerMsg.textContent = 'Farmer updated.';
        } else {
          await apiRequest('/api/farmers', {
            method: 'POST',
            body: payload
          });
          elements.farmerMsg.textContent = 'Farmer registered.';
        }

        resetFarmerForm();
        await fetchAllData();
        return;
      }

      const duplicateError = localFarmerDuplicateError(payload, state.editingFarmerId);
      if (duplicateError) {
        elements.farmerMsg.textContent = duplicateError;
        return;
      }

      if (state.editingFarmerId) {
        const existing = state.farmers.find((row) => row.id === state.editingFarmerId);
        if (existing) {
          Object.assign(existing, payload, { updatedAt: new Date().toISOString() });
        }
        elements.farmerMsg.textContent = 'Farmer updated (local mode).';
      } else {
        state.farmers.unshift({
          id: `F-${Date.now()}`,
          ...payload,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          createdBy: 'local'
        });
        elements.farmerMsg.textContent = 'Farmer registered (local mode).';
      }

      resetFarmerForm();
      hydrateFarmerSelectors();
      renderAll();
      persist();
    } catch (error) {
      elements.farmerMsg.textContent = error.message;
    }
  });

  elements.farmerCancelEditBtn.addEventListener('click', () => {
    resetFarmerForm();
    clearMessages();
  });

  if (elements.farmerImportNotifyBySms) {
    elements.farmerImportNotifyBySms.addEventListener('change', () => {
      updateFarmerImportSmsUi();
    });
  }
  updateFarmerImportSmsUi();

  elements.farmerImportForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();
    clearFarmerImportResult();

    const role = currentRole();
    if (role !== 'admin') {
      elements.farmerImportMsg.textContent = 'Only administrators can import farmer files.';
      return;
    }

    const [file] = elements.farmerImportFile.files || [];
    if (!file) {
      elements.farmerImportMsg.textContent = 'Choose a CSV or Excel file first.';
      return;
    }

    try {
      elements.farmerImportMsg.textContent = 'Reading file...';
      const records = await parseFarmerImportFile(file);

      if (!records.length) {
        elements.farmerImportMsg.textContent = 'No data rows found in this file.';
        return;
      }

      const duplicates = findImportDuplicates(records);
      const duplicatePhoneCount = duplicates.phones.length;
      const duplicateNationalIdCount = duplicates.nationalIds.length;
      let duplicateMode = 'skip';
      if (duplicatePhoneCount || duplicateNationalIdCount) {
        const overwrite = confirm(
          `Duplicate matches found: ${duplicatePhoneCount} by phone, ${duplicateNationalIdCount} by National ID.\n\n` +
          'Click OK to overwrite old farmer info with new file info.\n' +
          'Click Cancel to keep old info and skip duplicates.'
        );
        duplicateMode = overwrite ? 'overwrite' : 'skip';
      }
      const sendOnboardingSms = Boolean(elements.farmerImportNotifyBySms?.checked);
      const smsTemplate = String(elements.farmerImportSmsTemplate?.value || '').trim();
      if (sendOnboardingSms && smsTemplate.length > IMPORT_ONBOARDING_SMS_MAX_LENGTH) {
        elements.farmerImportMsg.textContent =
          `SMS template is too long. Keep it under ${IMPORT_ONBOARDING_SMS_MAX_LENGTH} characters.`;
        return;
      }
      const onboardingSmsTemplate = smsTemplate;

      let result = null;
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.farmerImportMsg.textContent = 'Sign in first to import farmers.';
          return;
        }

        const response = await apiRequest('/api/farmers/import', {
          method: 'POST',
          body: {
            records,
            onDuplicate: duplicateMode,
            sendOnboardingSms,
            onboardingSmsTemplate
          }
        });
        result = response.data || {};
        await fetchAllData();
      } else {
        result = importFarmersLocal(records, {
          onDuplicate: duplicateMode,
          sendOnboardingSms,
          onboardingSmsTemplate
        });
        hydrateFarmerSelectors();
        renderAll();
        persist();
      }

      renderFarmerImportResult(result);
      elements.farmerImportFile.value = '';
    } catch (error) {
      elements.farmerImportMsg.textContent = error.message || 'Import failed.';
    }
  });

  if (elements.farmerImportQaBtn) {
    elements.farmerImportQaBtn.addEventListener('click', async () => {
      clearMessages();
      clearFarmerImportResult();
      if (elements.farmerImportQaWrap) {
        elements.farmerImportQaWrap.hidden = true;
        elements.farmerImportQaWrap.innerHTML = '';
      }

      const role = currentRole();
      if (role !== 'admin') {
        elements.farmerImportQaMsg.textContent = 'Only administrators can run Smart Import QA.';
        return;
      }

      const [file] = elements.farmerImportFile.files || [];
      if (!file) {
        elements.farmerImportQaMsg.textContent = 'Choose a CSV or Excel file first.';
        return;
      }

      try {
        elements.farmerImportQaMsg.textContent = 'Running Smart QA...';
        const records = await parseFarmerImportFile(file);
        if (!records.length) {
          elements.farmerImportQaMsg.textContent = 'No data rows found in this file.';
          return;
        }

        let result;
        if (API.enabled) {
          if (!isAuthenticated()) {
            elements.farmerImportQaMsg.textContent = 'Sign in first to run Smart Import QA.';
            return;
          }
          const response = await apiRequest('/api/ai/import-qa', {
            method: 'POST',
            body: { records }
          });
          result = response.data || {};
        } else {
          result = runLocalSmartImportQa(records);
        }

        renderSmartImportQaResult(result);
      } catch (error) {
        elements.farmerImportQaMsg.textContent = error.message || 'Smart QA failed.';
      }
    });
  }

  const setFarmerPinControls = (disabled) => {
    if (!elements.farmerPinForm) return;
    const controls = [
      elements.farmerPinSearch,
      elements.farmerPinSearchBtn,
      elements.farmerPinFarmer,
      elements.farmerPinValue,
      elements.farmerPinConfirm,
      elements.farmerPinSubmitBtn,
      elements.farmerPinGenerateBtn
    ];
    controls.forEach((control) => {
      if (control) control.disabled = disabled;
    });
  };

  const showFarmerPinCredential = (payload, usedGeneratedPin) => {
    if (!elements.farmerPinCredentials) return;
    const username = String(payload?.username || '').trim();
    const phone = String(payload?.phone || '').trim();
    const generatedPin = String(payload?.generatedPin || '').trim();

    if (!username) {
      elements.farmerPinCredentials.hidden = true;
      elements.farmerPinCredentials.textContent = '';
      return;
    }

    const pinLine = usedGeneratedPin && generatedPin
      ? `Generated PIN: <code>${escapeHtml(generatedPin)}</code><br>`
      : '';
    elements.farmerPinCredentials.hidden = false;
    elements.farmerPinCredentials.innerHTML = `
      <strong>Portal access ${payload?.accountCreated ? 'created' : 'updated'}.</strong><br>
      Farmer: <code>${escapeHtml(String(payload?.farmerName || '-'))}</code><br>
      Login phone: <code>${escapeHtml(phone || username)}</code><br>
      Username: <code>${escapeHtml(username)}</code><br>
      ${pinLine}
      <span class="meta compact">Share the phone and PIN with the farmer. They can later change PIN in account settings.</span>
    `;
  };

  const runFarmerPinSearch = async () => {
    if (currentRole() !== 'admin') {
      elements.farmerPinMsg.textContent = 'Only admin can set or reset farmer PIN.';
      return;
    }

    const query = String(elements.farmerPinSearch?.value || '').trim();
    const selectedBefore = elements.farmerPinFarmer?.value || '';

    if (API.enabled) {
      if (!isAuthenticated()) {
        elements.farmerPinMsg.textContent = 'Sign in first to manage farmer PIN.';
        return;
      }

      const params = new URLSearchParams({ limit: '500' });
      if (query) params.set('q', query);
      const response = await apiRequest(`/api/farmers?${params.toString()}`);
      const rows = Array.isArray(response.data) ? response.data : [];
      hydrateFarmerPinOptions(rows, selectedBefore);
      elements.farmerPinMsg.textContent = `Loaded ${rows.length} farmer record(s).`;
      return;
    }

    const rows = state.farmers.filter((row) => {
      if (!query) return true;
      const q = query.toLowerCase();
      return [row.name, row.phone, row.nationalId, row.location]
        .some((field) => String(field || '').toLowerCase().includes(q));
    });
    hydrateFarmerPinOptions(rows, selectedBefore);
    elements.farmerPinMsg.textContent = 'PIN tool works in backend mode only.';
  };

  const submitFarmerPinReset = async (generate = false) => {
    clearMessages();
    if (currentRole() !== 'admin') {
      elements.farmerPinMsg.textContent = 'Only admin can set or reset farmer PIN.';
      return;
    }
    if (!API.enabled) {
      elements.farmerPinMsg.textContent = 'PIN management requires backend mode.';
      return;
    }
    if (!isAuthenticated()) {
      elements.farmerPinMsg.textContent = 'Sign in first to manage farmer PIN.';
      return;
    }

    const farmerId = String(elements.farmerPinFarmer?.value || '').trim();
    const pin = String(elements.farmerPinValue?.value || '').trim();
    const confirmPin = String(elements.farmerPinConfirm?.value || '').trim();

    if (!farmerId) {
      elements.farmerPinMsg.textContent = 'Select a farmer first.';
      return;
    }
    if (!generate) {
      if (!pin || !confirmPin) {
        elements.farmerPinMsg.textContent = 'Enter and confirm a 4-digit PIN.';
        return;
      }
      if (!/^\d{4}$/.test(pin)) {
        elements.farmerPinMsg.textContent = 'PIN must be exactly 4 digits.';
        return;
      }
      if (pin !== confirmPin) {
        elements.farmerPinMsg.textContent = 'PIN confirmation does not match.';
        return;
      }
    }

    setFarmerPinControls(true);
    try {
      const response = await apiRequest(`/api/farmers/${encodeURIComponent(farmerId)}/reset-pin`, {
        method: 'POST',
        body: generate ? { generate: true } : { pin, confirmPin }
      });
      elements.farmerPinValue.value = '';
      elements.farmerPinConfirm.value = '';
      showFarmerPinCredential(response.data || {}, generate);
      elements.farmerPinMsg.textContent = generate
        ? 'PIN generated and saved.'
        : 'PIN updated successfully.';
      await fetchAllData();
    } catch (error) {
      elements.farmerPinMsg.textContent = error.message;
    } finally {
      setFarmerPinControls(false);
    }
  };

  if (elements.farmerPinSearchBtn) {
    elements.farmerPinSearchBtn.addEventListener('click', async () => {
      clearMessages();
      try {
        await runFarmerPinSearch();
      } catch (error) {
        elements.farmerPinMsg.textContent = error.message;
      }
    });
  }

  if (elements.farmerPinSearch) {
    elements.farmerPinSearch.addEventListener('keydown', async (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      clearMessages();
      try {
        await runFarmerPinSearch();
      } catch (error) {
        elements.farmerPinMsg.textContent = error.message;
      }
    });
  }

  if (elements.farmerPinForm) {
    elements.farmerPinForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      await submitFarmerPinReset(false);
    });
  }

  if (elements.farmerPinGenerateBtn) {
    elements.farmerPinGenerateBtn.addEventListener('click', async () => {
      await submitFarmerPinReset(true);
    });
  }

  elements.farmerTableWrap.addEventListener('click', async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;

    const farmerId = button.dataset.id;
    if (!farmerId) return;

    if (button.dataset.action === 'edit-farmer') {
      const farmer = state.farmers.find((row) => row.id === farmerId);
      if (!farmer) return;

      state.editingFarmerId = farmer.id;
      elements.farmerId.value = farmer.id;
      elements.farmerName.value = farmer.name || '';
      elements.farmerPhone.value = farmer.phone || '';
      elements.farmerNationalId.value = farmer.nationalId || '';
      elements.farmerLocation.value = farmer.location || '';
      elements.treeCount.value = farmer.trees || 0;
      elements.farmerPreferredLanguage.value = normalizeLanguageClient(farmer.preferredLanguage) || DEFAULT_LANGUAGE;
      fillAreaInputsFromHectares(
        {
          hectaresInput: elements.farmerAreaHectares,
          acresInput: elements.farmerAreaAcres,
          squareFeetInput: elements.farmerAreaSquareFeet
        },
        farmer.hectares
      );
      fillAreaInputsFromHectares(
        {
          hectaresInput: elements.farmerAvocadoAreaHectares,
          acresInput: elements.farmerAvocadoAreaAcres,
          squareFeetInput: elements.farmerAvocadoAreaSquareFeet
        },
        farmer.avocadoHectares
      );
      elements.farmerNotes.value = farmer.notes || '';

      elements.farmerSubmitBtn.textContent = 'Update Farmer';
      elements.farmerCancelEditBtn.hidden = false;
      elements.farmerMsg.textContent = `Editing ${farmer.name}`;
      return;
    }

    if (button.dataset.action === 'set-pin-farmer') {
      if (currentRole() !== 'admin') {
        elements.farmerMsg.textContent = 'Only admin can set/reset farmer PIN.';
        return;
      }

      if (elements.farmerPinFarmer) {
        elements.farmerPinFarmer.value = farmerId;
      }
      if (elements.farmerPinCard) {
        elements.farmerPinCard.open = true;
      }
      elements.farmerPinValue?.focus();
      elements.farmerPinMsg.textContent = `Ready to set PIN for ${button.dataset.name || 'selected farmer'}.`;
      return;
    }

    if (button.dataset.action === 'delete-farmer') {
      if (!confirm('Delete this farmer? This only works when no produce/payment history exists.')) return;

      try {
        if (API.enabled) {
          if (!isAuthenticated()) {
            elements.farmerMsg.textContent = 'Sign in first to delete farmers.';
            return;
          }

          await apiRequest(`/api/farmers/${encodeURIComponent(farmerId)}`, { method: 'DELETE' });
          elements.farmerMsg.textContent = 'Farmer deleted.';
          await fetchAllData();
          return;
        }

        state.farmers = state.farmers.filter((row) => row.id !== farmerId);
        hydrateFarmerSelectors();
        renderAll();
        persist();
        elements.farmerMsg.textContent = 'Farmer deleted (local mode).';
      } catch (error) {
        elements.farmerMsg.textContent = error.message;
      }
    }
  });
}

function updateFarmerImportSmsUi() {
  if (!elements.farmerImportNotifyBySms || !elements.farmerImportSmsTemplate) return;
  const enabled = !elements.farmerImportNotifyBySms.disabled && elements.farmerImportNotifyBySms.checked;
  elements.farmerImportSmsTemplate.disabled = !enabled;
}

function bindAgents() {
  if (!elements.agentForm) return;

  elements.agentForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();
    clearAgentCredentialDisplay();

    if (currentRole() !== 'admin') {
      elements.agentMsg.textContent = 'Only administrators can create agent accounts.';
      return;
    }
    if (!API.enabled) {
      elements.agentMsg.textContent = 'Agent management requires backend mode.';
      return;
    }
    if (!isAuthenticated()) {
      elements.agentMsg.textContent = 'Sign in first to create agents.';
      return;
    }

    const name = elements.agentName.value.trim();
    const email = elements.agentEmail.value.trim().toLowerCase();
    const phone = cleanPhone(elements.agentPhone.value);

    if (!name || !email || !phone) {
      elements.agentMsg.textContent = 'Name, email, and phone are required.';
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      elements.agentMsg.textContent = 'Enter a valid email address.';
      return;
    }
    if (!/^254\d{9}$/.test(phone)) {
      elements.agentMsg.textContent = 'Enter a valid Kenyan phone number (e.g. 2547XXXXXXXX).';
      return;
    }

    try {
      const response = await apiRequest('/api/agents', {
        method: 'POST',
        body: { name, email, phone }
      });
      elements.agentForm.reset();
      elements.agentMsg.textContent = `Agent account created for ${response.data.name}.`;
      showAgentCredentialDisplay(response.data);
      await loadAgentAccounts();
    } catch (error) {
      elements.agentMsg.textContent = error.message;
    }
  });

  elements.agentRefreshBtn.addEventListener('click', async () => {
    clearMessages();
    try {
      await loadAgentAccounts();
      elements.agentMsg.textContent = `Loaded ${state.agents.length} agent account(s).`;
    } catch (error) {
      elements.agentMsg.textContent = error.message;
    }
  });

  elements.agentSearch.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    clearMessages();
    try {
      await loadAgentAccounts();
    } catch (error) {
      elements.agentMsg.textContent = error.message;
    }
  });
}

function showAgentCredentialDisplay(data) {
  if (!elements.agentCredentials) return;
  const username = String(data?.username || '').trim();
  const temporaryPassword = String(data?.temporaryPassword || '').trim();
  const email = String(data?.email || '').trim();
  const phone = String(data?.phone || '').trim();
  if (!username || !temporaryPassword) {
    elements.agentCredentials.hidden = true;
    elements.agentCredentials.textContent = '';
    return;
  }

  elements.agentCredentials.hidden = false;
  elements.agentCredentials.innerHTML = `
    <strong>Temporary login created.</strong><br>
    Email: <code>${escapeHtml(email || '-')}</code><br>
    Phone: <code>${escapeHtml(phone || '-')}</code><br>
    Username: <code>${escapeHtml(username)}</code><br>
    Password: <code>${escapeHtml(temporaryPassword)}</code><br>
    <span class="meta compact">Share these once, then ask the agent to change password after first sign-in.</span>
  `;
}

function clearAgentCredentialDisplay() {
  if (!elements.agentCredentials) return;
  elements.agentCredentials.hidden = true;
  elements.agentCredentials.textContent = '';
}

function bindProduce() {
  elements.produceForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    const role = currentRole();
    if (!['admin', 'agent'].includes(role)) {
      elements.produceMsg.textContent = 'Only admin or agent can record produce.';
      return;
    }

    if (!state.farmers.length) {
      elements.produceMsg.textContent = 'Register a farmer first.';
      return;
    }

    const payload = {
      farmerId: elements.produceFarmer.value,
      kgs: Number(elements.produceKgs.value || 0),
      lotWeightKgs: Number(elements.produceKgs.value || 0),
      variety: elements.produceVariety.value.trim(),
      sampleSize: Number(elements.produceSampleSize.value || 0),
      visualGrade: elements.produceVisualGrade.value.trim(),
      qcDecision: elements.produceQcDecision.value.trim(),
      dryMatterPct: elements.produceDryMatterPct.value.trim(),
      firmnessValue: Number(elements.produceFirmnessValue.value || 0),
      firmnessUnit: elements.produceFirmnessUnit.value.trim(),
      avgFruitWeightG: Number(elements.produceAvgWeightG.value || 0),
      sizeCode: elements.produceSizeCode.value.trim(),
      inspector: elements.produceInspector.value.trim(),
      notes: elements.produceNotes.value.trim()
    };

    try {
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.produceMsg.textContent = 'Sign in first to use backend mode.';
          return;
        }

        await apiRequest('/api/produce', {
          method: 'POST',
          body: payload
        });
        elements.produceForm.reset();
        elements.produceMsg.textContent = 'Produce recorded.';
        await fetchAllData();
        return;
      }

      const farmerName = farmerNameById(payload.farmerId);
      state.produce.unshift({
        id: `P-${Date.now()}`,
        ...payload,
        farmerName,
        createdBy: 'local',
        createdAt: new Date().toISOString()
      });

      elements.produceForm.reset();
      elements.produceMsg.textContent = 'Produce recorded (local mode).';
      renderAll();
      persist();
    } catch (error) {
      elements.produceMsg.textContent = error.message;
    }
  });

  elements.produceTableWrap.addEventListener('click', async (event) => {
    const button = event.target.closest('button[data-action="delete-produce"]');
    if (!button) return;

    const produceId = button.dataset.id;
    if (!produceId) return;

    if (!confirm('Delete this produce record?')) return;

    try {
      if (API.enabled) {
        await apiRequest(`/api/produce/${encodeURIComponent(produceId)}`, { method: 'DELETE' });
        elements.produceMsg.textContent = 'Produce record deleted.';
        await fetchAllData();
        return;
      }

      state.produce = state.produce.filter((row) => row.id !== produceId);
      renderAll();
      persist();
      elements.produceMsg.textContent = 'Produce record deleted (local mode).';
    } catch (error) {
      elements.produceMsg.textContent = error.message;
    }
  });
}

function bindProducePurchases() {
  elements.purchaseFarmer.addEventListener('change', () => {
    hydratePurchaseQcOptions();
  });

  elements.purchaseForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    const role = currentRole();
    if (!['admin', 'agent'].includes(role)) {
      elements.purchaseMsg.textContent = 'Only admin or agent can record purchased produce.';
      return;
    }

    if (!state.farmers.length) {
      elements.purchaseMsg.textContent = 'Register a farmer first.';
      return;
    }

    const payload = {
      farmerId: elements.purchaseFarmer.value,
      qcRecordId: elements.purchaseQcRecord.value.trim(),
      purchasedKgs: Number(elements.purchaseKgs.value || 0),
      pricePerKgKes: elements.purchasePricePerKgKes.value.trim(),
      variety: elements.purchaseVariety.value.trim(),
      sizeCode: elements.purchaseSizeCode.value.trim(),
      buyer: elements.purchaseBuyer.value.trim(),
      purchaseValueKes: elements.purchaseValueKes.value.trim(),
      notes: elements.purchaseNotes.value.trim()
    };

    try {
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.purchaseMsg.textContent = 'Sign in first to use backend mode.';
          return;
        }

        await apiRequest('/api/produce-purchases', {
          method: 'POST',
          body: payload
        });
        elements.purchaseForm.reset();
        elements.purchaseMsg.textContent = 'Purchased produce recorded.';
        await fetchAllData();
        await refreshOwedRows(false);
        return;
      }

      const linkedQc = state.produce.find((row) => row.id === payload.qcRecordId) || null;
      const variety = payload.variety || linkedQc?.variety || '';
      if (!variety) {
        elements.purchaseMsg.textContent = 'Select variety or link a QC lot.';
        return;
      }
      const purchasedKgs = Number(payload.purchasedKgs || 0);
      if (!(purchasedKgs > 0)) {
        elements.purchaseMsg.textContent = 'Purchased weight must be greater than 0.';
        return;
      }

      if (!(Number(payload.pricePerKgKes) > 0) && !(Number(payload.purchaseValueKes) > 0)) {
        elements.purchaseMsg.textContent = 'Enter agreed price per kg or total purchase value.';
        return;
      }

      const pricePerKgKes = payload.pricePerKgKes ? Number(payload.pricePerKgKes) : null;
      const computedValue = pricePerKgKes ? purchasedKgs * pricePerKgKes : null;
      const purchaseValueKes = payload.purchaseValueKes
        ? Number(payload.purchaseValueKes)
        : computedValue;

      state.producePurchases.unshift({
        id: `PR-${Date.now()}`,
        farmerId: payload.farmerId,
        farmerName: farmerNameById(payload.farmerId),
        qcRecordId: payload.qcRecordId,
        variety,
        sizeCode: payload.sizeCode || linkedQc?.sizeCode || '',
        purchasedKgs: Number(purchasedKgs.toFixed(2)),
        pricePerKgKes: pricePerKgKes ? Number(pricePerKgKes.toFixed(2)) : null,
        purchaseValueKes: purchaseValueKes ? Number(purchaseValueKes.toFixed(2)) : null,
        buyer: payload.buyer,
        notes: payload.notes,
        createdBy: 'local',
        createdAt: new Date().toISOString()
      });

      elements.purchaseForm.reset();
      elements.purchaseMsg.textContent = 'Purchased produce recorded (local mode).';
      renderAll();
      await refreshOwedRows(false);
      persist();
    } catch (error) {
      elements.purchaseMsg.textContent = error.message;
    }
  });

  elements.purchaseTableWrap.addEventListener('click', async (event) => {
    const button = event.target.closest('button[data-action="delete-purchase"]');
    if (!button) return;

    const purchaseId = button.dataset.id;
    if (!purchaseId) return;

    if (!confirm('Delete this purchased produce record?')) return;

    try {
      if (API.enabled) {
        await apiRequest(`/api/produce-purchases/${encodeURIComponent(purchaseId)}`, { method: 'DELETE' });
        elements.purchaseMsg.textContent = 'Purchased produce record deleted.';
        await fetchAllData();
        await refreshOwedRows(false);
        return;
      }

      state.producePurchases = state.producePurchases.filter((row) => row.id !== purchaseId);
      renderAll();
      await refreshOwedRows(false);
      persist();
      elements.purchaseMsg.textContent = 'Purchased produce record deleted (local mode).';
    } catch (error) {
      elements.purchaseMsg.textContent = error.message;
    }
  });
}

function bindPayments() {
  elements.paymentFarmer.addEventListener('change', () => {
    const owed = owedPicker.rows.find((row) => row.farmerId === elements.paymentFarmer.value);
    if (owed && owed.balanceKes > 0) {
      elements.amount.value = String(owed.balanceKes);
    }
  });

  if (elements.paymentRecommendationStatusFilter) {
    elements.paymentRecommendationStatusFilter.addEventListener('change', async () => {
      clearMessages();
      await loadPaymentRecommendations();
    });
  }

  if (elements.paymentRecommendationSearch) {
    elements.paymentRecommendationSearch.addEventListener('keydown', async (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      clearMessages();
      await loadPaymentRecommendations();
    });
  }

  if (elements.paymentRecommendationRefreshBtn) {
    elements.paymentRecommendationRefreshBtn.addEventListener('click', async () => {
      clearMessages();
      await loadPaymentRecommendations();
    });
  }

  if (elements.paymentRecommendationForm) {
    elements.paymentRecommendationForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      clearMessages();

      if (currentRole() !== 'agent') {
        elements.paymentRecommendationMsg.textContent = 'Only agents can submit payment recommendations.';
        return;
      }

      const payload = {
        farmerId: String(elements.paymentRecommendationFarmer.value || '').trim(),
        amount: Number(elements.paymentRecommendationAmount.value || 0),
        reason: String(elements.paymentRecommendationReason.value || '').trim()
      };

      if (!payload.farmerId || !(payload.amount > 0) || !payload.reason) {
        elements.paymentRecommendationMsg.textContent = 'Farmer, amount, and reason are required.';
        return;
      }

      try {
        if (API.enabled) {
          if (!isAuthenticated()) {
            elements.paymentRecommendationMsg.textContent = 'Sign in first to submit recommendation.';
            return;
          }
          await apiRequest('/api/payments/recommendations', {
            method: 'POST',
            body: payload
          });
        elements.paymentRecommendationForm.reset();
        elements.paymentRecommendationMsg.textContent = 'Recommendation sent to admin for approval.';
        await fetchAllData();
        await refreshOwedRows(false);
        await loadPaymentRecommendations();
        return;
      }

        const farmer = state.farmers.find((row) => row.id === payload.farmerId);
        const now = new Date().toISOString();
        state.paymentRecommendations = Array.isArray(state.paymentRecommendations) ? state.paymentRecommendations : [];
        state.paymentRecommendations.unshift({
          id: `PREQ-${Date.now()}`,
          farmerId: payload.farmerId,
          farmerName: farmer?.name || 'Unknown',
          farmerPhone: farmer?.phone || '',
          amount: Number(payload.amount.toFixed(2)),
          requestedOwedKes: Number(payload.amount.toFixed(2)),
          status: 'pending',
          reason: payload.reason,
          decisionNote: '',
          createdBy: state.auth?.user?.username || 'agent-local',
          createdAt: now,
          updatedAt: now,
          approvedBy: '',
          approvedAt: '',
          rejectedBy: '',
          rejectedAt: '',
          rejectionReason: '',
          paymentId: ''
        });
        elements.paymentRecommendationForm.reset();
        elements.paymentRecommendationMsg.textContent = 'Recommendation queued (local mode).';
        await loadPaymentRecommendations();
        persist();
      } catch (error) {
        elements.paymentRecommendationMsg.textContent = error.message;
      }
    });
  }

  elements.owedPeriod.addEventListener('change', async () => {
    clearMessages();
    await refreshOwedRows(true);
  });

  elements.owedFromDate.addEventListener('change', async () => {
    clearMessages();
    await refreshOwedRows(true);
  });

  elements.owedToDate.addEventListener('change', async () => {
    clearMessages();
    await refreshOwedRows(true);
  });

  elements.owedSearch.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    clearMessages();
    await refreshOwedRows(true);
  });

  elements.owedRefreshBtn.addEventListener('click', async () => {
    clearMessages();
    await refreshOwedRows(false);
  });

  elements.owedSelectVisibleBtn.addEventListener('click', () => {
    for (const row of owedPicker.rows) {
      owedPicker.selectedFarmerIds.add(row.farmerId);
    }
    renderOwedTable();
  });

  elements.owedClearSelectedBtn.addEventListener('click', () => {
    owedPicker.selectedFarmerIds.clear();
    renderOwedTable();
  });

  elements.owedTableWrap.addEventListener('change', (event) => {
    const checkbox = event.target.closest('input[data-action="toggle-owed-farmer"]');
    if (!checkbox) return;
    const farmerId = String(checkbox.dataset.id || '');
    if (!farmerId) return;

    if (checkbox.checked) {
      owedPicker.selectedFarmerIds.add(farmerId);
    } else {
      owedPicker.selectedFarmerIds.delete(farmerId);
    }
    renderOwedTable();
  });

  elements.owedPrepareBtn.addEventListener('click', async () => {
    clearMessages();
    await settleSelectedOwed('Pending');
  });

  elements.owedPaySelectedBtn.addEventListener('click', async () => {
    clearMessages();
    await settleSelectedOwed('Received');
  });

  elements.owedExportBtn.addEventListener('click', async () => {
    clearMessages();
    await exportOwedCsv();
  });

  elements.paymentForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    if (currentRole() !== 'admin') {
      elements.paymentMsg.textContent = 'Only admin can log payments.';
      return;
    }

    if (!state.farmers.length) {
      elements.paymentMsg.textContent = 'Register a farmer first.';
      return;
    }

    const payload = {
      farmerId: elements.paymentFarmer.value,
      amount: Number(elements.amount.value || 0),
      ref: elements.mpesaRef.value.trim(),
      status: elements.paymentStatus.value,
      notes: elements.paymentNotes.value.trim(),
      method: 'M-PESA'
    };

    try {
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.paymentMsg.textContent = 'Sign in first to use backend mode.';
          return;
        }

        await apiRequest('/api/payments', {
          method: 'POST',
          body: payload
        });

        elements.paymentForm.reset();
        elements.paymentMsg.textContent = 'Payment logged.';
        await fetchAllData();
        await refreshOwedRows(false);
        return;
      }

      state.payments.unshift({
        id: `TX-${Date.now()}`,
        farmerId: payload.farmerId,
        farmerName: farmerNameById(payload.farmerId),
        amount: payload.amount,
        ref: payload.ref,
        status: payload.status,
        notes: payload.notes,
        method: payload.method,
        createdBy: 'local',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });

      elements.paymentForm.reset();
      elements.paymentMsg.textContent = 'Payment logged (local mode).';
      renderAll();
      await refreshOwedRows(false);
      persist();
    } catch (error) {
      elements.paymentMsg.textContent = error.message;
    }
  });

  elements.mpesaDisburseBtn.addEventListener('click', async () => {
    clearMessages();

    if (currentRole() !== 'admin') {
      elements.paymentMsg.textContent = 'Only admin can simulate disbursements.';
      return;
    }

    if (!elements.paymentFarmer.value || !(Number(elements.amount.value) > 0)) {
      elements.paymentMsg.textContent = 'Select farmer and amount first.';
      return;
    }

    if (!API.enabled) {
      elements.paymentMsg.textContent = 'M-PESA simulation requires backend mode.';
      return;
    }

    try {
      await apiRequest('/api/integrations/mpesa/disburse', {
        method: 'POST',
        body: {
          farmerId: elements.paymentFarmer.value,
          amount: Number(elements.amount.value),
          narration: elements.paymentNotes.value.trim()
        }
      });

      elements.paymentMsg.textContent = 'M-PESA disbursement simulated successfully.';
      await fetchAllData();
      await refreshOwedRows(false);
    } catch (error) {
      elements.paymentMsg.textContent = error.message;
    }
  });

  elements.paymentTableWrap.addEventListener('click', async (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;

    const paymentId = button.dataset.id;
    const nextStatus = button.dataset.status;

    if (button.dataset.action !== 'set-payment-status' || !paymentId || !nextStatus) return;

    if (currentRole() !== 'admin') {
      elements.paymentMsg.textContent = 'Only admin can change payment status.';
      return;
    }

    try {
      if (API.enabled) {
        await apiRequest(`/api/payments/${encodeURIComponent(paymentId)}/status`, {
          method: 'PATCH',
          body: { status: nextStatus }
        });
        elements.paymentMsg.textContent = `Payment status changed to ${nextStatus}.`;
        await fetchAllData();
        await refreshOwedRows(false);
        return;
      }

      const payment = state.payments.find((row) => row.id === paymentId);
      if (payment) {
        payment.status = nextStatus;
        payment.updatedAt = new Date().toISOString();
      }
      renderAll();
      await refreshOwedRows(false);
      persist();
      elements.paymentMsg.textContent = `Payment status changed to ${nextStatus} (local mode).`;
    } catch (error) {
      elements.paymentMsg.textContent = error.message;
    }
  });

  if (elements.paymentRecommendationTableWrap) {
    elements.paymentRecommendationTableWrap.addEventListener('click', async (event) => {
      const button = event.target.closest('button[data-action]');
      if (!button) return;

      const recId = String(button.dataset.id || '').trim();
      const action = String(button.dataset.action || '').trim();
      if (!recId || !action) return;

      if (!['approve-payment-recommendation', 'reject-payment-recommendation'].includes(action)) return;
      if (currentRole() !== 'admin') {
        elements.paymentRecommendationMsg.textContent = 'Only admin can approve or reject recommendations.';
        return;
      }

      try {
        if (API.enabled) {
          if (!isAuthenticated()) {
            elements.paymentRecommendationMsg.textContent = 'Sign in first.';
            return;
          }

          if (action === 'approve-payment-recommendation') {
            const decisionNote = prompt('Optional approval note (press Cancel to skip):', '') || '';
            const response = await apiRequest(
              `/api/payments/recommendations/${encodeURIComponent(recId)}/approve`,
              {
                method: 'POST',
                body: { decisionNote }
              }
            );
            const smsFarmer = response?.meta?.smsNotified?.farmer ? 'farmer SMS sent' : 'farmer SMS skipped';
            const smsAgent = response?.meta?.smsNotified?.agent ? 'agent SMS sent' : 'agent SMS skipped';
            elements.paymentRecommendationMsg.textContent = `Recommendation approved and payment posted (${smsFarmer}, ${smsAgent}).`;
          } else {
            const reason = prompt('Reason for rejection (optional):', '') || '';
            const response = await apiRequest(
              `/api/payments/recommendations/${encodeURIComponent(recId)}/reject`,
              {
                method: 'POST',
                body: { reason }
              }
            );
            const smsFarmer = response?.meta?.smsNotified?.farmer ? 'farmer SMS sent' : 'farmer SMS skipped';
            const smsAgent = response?.meta?.smsNotified?.agent ? 'agent SMS sent' : 'agent SMS skipped';
            elements.paymentRecommendationMsg.textContent = `Recommendation rejected (${smsFarmer}, ${smsAgent}).`;
          }
          await fetchAllData();
          await refreshOwedRows(false);
          await loadPaymentRecommendations();
          return;
        }

        state.paymentRecommendations = Array.isArray(state.paymentRecommendations) ? state.paymentRecommendations : [];
        const rec = state.paymentRecommendations.find((row) => row.id === recId);
        if (!rec || String(rec.status).toLowerCase() !== 'pending') return;
        const now = new Date().toISOString();
        if (action === 'approve-payment-recommendation') {
          rec.status = 'approved';
          rec.approvedBy = state.auth?.user?.username || 'admin-local';
          rec.approvedAt = now;
          rec.updatedAt = now;
          const ref = `APR${Date.now().toString().slice(-8)}`;
          const farmer = state.farmers.find((row) => row.id === rec.farmerId);
          const paymentRecord = {
            id: `TX-${Date.now()}`,
            farmerId: rec.farmerId,
            farmerName: rec.farmerName || farmer?.name || 'Unknown',
            amount: Number(rec.amount || 0),
            ref,
            status: 'Received',
            method: 'M-PESA(Mock)',
            notes: `Approved from recommendation ${rec.id}`,
            source: 'agent-recommendation',
            recommendationId: rec.id,
            createdBy: rec.approvedBy,
            createdAt: now,
            updatedAt: now
          };
          rec.paymentId = paymentRecord.id;
          state.payments.unshift(paymentRecord);
          elements.paymentRecommendationMsg.textContent = 'Recommendation approved and payment posted (local mode).';
        } else {
          const reason = prompt('Reason for rejection (optional):', '') || '';
          rec.status = 'rejected';
          rec.rejectedBy = state.auth?.user?.username || 'admin-local';
          rec.rejectedAt = now;
          rec.updatedAt = now;
          rec.rejectionReason = reason;
          elements.paymentRecommendationMsg.textContent = 'Recommendation rejected (local mode).';
        }
        await loadPaymentRecommendations();
        renderAll();
        persist();
      } catch (error) {
        elements.paymentRecommendationMsg.textContent = error.message;
      }
    });
  }
}

function moneyValue(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Number(num.toFixed(2));
}

function normalizePaymentStatusValue(status) {
  const lower = String(status || '').trim().toLowerCase();
  if (lower === 'received') return 'Received';
  if (lower === 'pending') return 'Pending';
  return 'Failed';
}

function parseDateBound(value, isEnd) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  const stamp = /^\d{4}-\d{2}-\d{2}$/.test(raw)
    ? `${raw}T${isEnd ? '23:59:59.999' : '00:00:00.000'}`
    : raw;
  const ms = new Date(stamp).getTime();
  return Number.isFinite(ms) ? ms : null;
}

function formatDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function syncExportFilterUi() {
  if (!elements.exportRange || !elements.exportFrom || !elements.exportTo) return;
  const range = String(elements.exportRange.value || 'today').trim().toLowerCase();
  const custom = range === 'custom';
  elements.exportFrom.hidden = !custom;
  elements.exportTo.hidden = !custom;
  elements.exportFrom.disabled = !custom;
  elements.exportTo.disabled = !custom;
}

function buildExportQueryParams() {
  const params = new URLSearchParams();
  if (!elements.exportRange) return params;

  const range = String(elements.exportRange.value || 'today').trim().toLowerCase();
  params.set('period', range);

  if (range === 'today') {
    const today = formatDateInputValue(new Date());
    params.set('from', today);
    params.set('to', today);
    return params;
  }

  if (range === 'last7') {
    const toDate = new Date();
    const fromDate = new Date(toDate);
    fromDate.setDate(fromDate.getDate() - 6);
    params.set('from', formatDateInputValue(fromDate));
    params.set('to', formatDateInputValue(toDate));
    return params;
  }

  if (range === 'thismonth' || range === 'month') {
    const toDate = new Date();
    const fromDate = new Date(toDate.getFullYear(), toDate.getMonth(), 1);
    params.set('from', formatDateInputValue(fromDate));
    params.set('to', formatDateInputValue(toDate));
    return params;
  }

  if (range === 'last30') {
    const toDate = new Date();
    const fromDate = new Date(toDate);
    fromDate.setDate(fromDate.getDate() - 29);
    params.set('from', formatDateInputValue(fromDate));
    params.set('to', formatDateInputValue(toDate));
    return params;
  }

  if (range === 'all') {
    return params;
  }

  if (range === 'custom') {
    let from = String(elements.exportFrom?.value || '').trim();
    let to = String(elements.exportTo?.value || '').trim();
    if (!from || !to) {
      throw new Error('For custom export, choose both start and end date.');
    }
    if (from > to) {
      const swap = from;
      from = to;
      to = swap;
    }
    params.set('from', from);
    params.set('to', to);
  }

  return params;
}

function exportRangeLabel(periodRaw, fromRaw, toRaw) {
  const period = String(periodRaw || '').trim().toLowerCase();
  if (period === 'today') return 'Today';
  if (period === 'last7') return 'Last 7 Days';
  if (period === 'thismonth' || period === 'month') return 'This Month';
  if (period === 'last30') return 'Last 30 Days';
  if (period === 'all') return 'All Dates';
  if (period === 'custom') {
    const from = clean(fromRaw);
    const to = clean(toRaw);
    if (from && to) return `Custom (${from} to ${to})`;
    return 'Custom Date Range';
  }
  return 'Selected Range';
}

let exportLogoDataUrlPromise = null;
async function getExportLogoDataUrl() {
  if (!exportLogoDataUrlPromise) {
    exportLogoDataUrlPromise = fetch('/assets/agem-logo-transparent.png', { cache: 'force-cache' })
      .then(async (response) => {
        if (!response.ok) return '';
        const blob = await response.blob();
        return await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result || ''));
          reader.onerror = () => resolve('');
          reader.readAsDataURL(blob);
        });
      })
      .catch(() => '');
  }
  return exportLogoDataUrlPromise;
}

function parseCsvTextRows(csvText) {
  if (!window.XLSX) {
    throw new Error('Spreadsheet engine is not available yet. Refresh the page and retry.');
  }
  const workbook = window.XLSX.read(csvText, { type: 'string' });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const headerRows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const headers = Array.isArray(headerRows[0]) ? headerRows[0].map((value) => String(value || '').trim()) : [];
  const records = window.XLSX.utils.sheet_to_json(sheet, { defval: '' });
  return { headers, records };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function escapeExportHtml(value) {
  return String(value == null ? '' : value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function csvCellEscaped(value) {
  const text = String(value == null ? '' : value);
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function buildBrandedCsv(csvText, type, rangeLabel) {
  const lines = [
    [ 'Agem Portal Export', type ],
    [ 'Brand', 'Agem Handshake' ],
    [ 'Range', rangeLabel ],
    [ 'Generated', new Date().toISOString() ],
    []
  ].map((row) => row.map((cell) => csvCellEscaped(cell)).join(','));

  const base = String(csvText || '');
  return `${lines.join('\n')}\n${base}`;
}

async function exportAsExcelHtml(type, csvText, rangeLabel) {
  const parsed = parseCsvTextRows(csvText);
  const logoDataUrl = await getExportLogoDataUrl();
  const headers = parsed.headers;
  const rows = parsed.records;
  const generatedAt = new Date().toISOString();

  const tableHead = headers.map((header) => `<th>${escapeExportHtml(header)}</th>`).join('');
  const tableBody = rows
    .map((record) => {
      const cells = headers.map((header) => `<td>${escapeExportHtml(record?.[header])}</td>`).join('');
      return `<tr>${cells}</tr>`;
    })
    .join('');

  const html = `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      body { font-family: Arial, sans-serif; font-size: 12px; color: #123f2a; }
      .header { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }
      .header img { height: 46px; width: auto; }
      .title { font-size: 20px; font-weight: 700; margin: 0; }
      .meta { margin: 2px 0; color: #2f5c45; }
      table { border-collapse: collapse; width: 100%; margin-top: 12px; }
      th, td { border: 1px solid #b8cfc0; padding: 6px 8px; text-align: left; vertical-align: top; }
      th { background: #e8f3ea; color: #123f2a; font-weight: 700; }
      tr:nth-child(even) td { background: #f7fbf8; }
    </style>
  </head>
  <body>
    <div class="header">
      ${logoDataUrl ? `<img src="${logoDataUrl}" alt="Agem Handshake Logo" />` : ''}
      <div>
        <p class="title">Agem Portal Export: ${escapeExportHtml(type)}</p>
        <p class="meta">Range: ${escapeExportHtml(rangeLabel)}</p>
        <p class="meta">Generated: ${escapeExportHtml(generatedAt)}</p>
      </div>
    </div>
    <table>
      <thead>
        <tr>${tableHead}</tr>
      </thead>
      <tbody>
        ${tableBody}
      </tbody>
    </table>
  </body>
</html>`;

  return new Blob([html], {
    type: 'application/vnd.ms-excel;charset=utf-8'
  });
}

async function exportAsPdf(type, csvText, rangeLabel) {
  const parsed = parseCsvTextRows(csvText);
  const headers = parsed.headers;
  const rows = parsed.records;

  if (!window.jspdf || !window.jspdf.jsPDF) {
    throw new Error('PDF engine is not available yet. Refresh the page and retry.');
  }
  const orientation = headers.length > 8 ? 'landscape' : 'portrait';
  const doc = new window.jspdf.jsPDF({ orientation, unit: 'pt', format: 'a4' });
  if (typeof doc.autoTable !== 'function') {
    throw new Error('PDF table engine is not available yet. Refresh the page and retry.');
  }

  const body = rows.map((record) => headers.map((header) => String(record?.[header] ?? '')));
  const logoDataUrl = await getExportLogoDataUrl();

  if (logoDataUrl) {
    doc.addImage(logoDataUrl, 'PNG', 40, 26, 120, 38);
  }
  doc.setFontSize(16);
  doc.setTextColor(18, 63, 42);
  doc.text(`Agem Portal Export: ${type}`, 40, 86);
  doc.setFontSize(10);
  doc.setTextColor(62, 94, 77);
  doc.text(`Range: ${rangeLabel}`, 40, 104);
  doc.text(`Generated: ${new Date().toISOString()}`, 40, 118);

  doc.autoTable({
    startY: 132,
    head: [headers],
    body,
    styles: { fontSize: 8, cellPadding: 4, overflow: 'linebreak' },
    headStyles: { fillColor: [22, 93, 58], textColor: 255, fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [246, 250, 247] },
    margin: { left: 24, right: 24 }
  });

  return doc.output('blob');
}

async function exportDatasetFile({ type, csvText, format, rangeLabel, fileSuffix = '' }) {
  const dateStamp = new Date().toISOString().slice(0, 10);
  const safeType = String(type || 'export').toLowerCase();
  const suffix = fileSuffix ? `-${fileSuffix}` : '';

  if (format === 'excel') {
    const excelBlob = await exportAsExcelHtml(type, csvText, rangeLabel);
    downloadBlob(excelBlob, `${safeType}${suffix}-${dateStamp}.xls`);
    return 'Excel';
  }

  if (format === 'pdf') {
    const pdfBlob = await exportAsPdf(type, csvText, rangeLabel);
    downloadBlob(pdfBlob, `${safeType}${suffix}-${dateStamp}.pdf`);
    return 'PDF';
  }

  const brandedCsv = buildBrandedCsv(csvText, type, rangeLabel);
  const csvBlob = new Blob([brandedCsv], { type: 'text/csv;charset=utf-8' });
  downloadBlob(csvBlob, `${safeType}${suffix}-${dateStamp}.csv`);
  return 'CSV';
}

function resolveOwedFilters() {
  const period = String(elements.owedPeriod.value || 'month').toLowerCase();
  let from = parseDateBound(elements.owedFromDate.value, false);
  let to = parseDateBound(elements.owedToDate.value, true);

  if (from !== null && to !== null && from > to) {
    const swap = from;
    from = to;
    to = swap;
  }

  if (from === null && to === null) {
    const now = new Date();
    if (period === 'today') {
      from = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0).getTime();
      to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999).getTime();
    } else if (period === 'week') {
      to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999).getTime();
      from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 6, 0, 0, 0, 0).getTime();
    } else if (period === 'month') {
      from = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0).getTime();
      to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999).getTime();
    }
  }

  return {
    period,
    q: String(elements.owedSearch.value || '').trim(),
    from,
    to,
    fromRaw: String(elements.owedFromDate.value || '').trim(),
    toRaw: String(elements.owedToDate.value || '').trim()
  };
}

function purchaseTotalValue(row) {
  const explicit = Number(row.purchaseValueKes);
  if (Number.isFinite(explicit) && explicit > 0) return moneyValue(explicit);
  const unit = Number(row.pricePerKgKes);
  const qty = Number(row.purchasedKgs);
  if (Number.isFinite(unit) && unit > 0 && Number.isFinite(qty) && qty > 0) {
    return moneyValue(unit * qty);
  }
  return 0;
}

function dateValue(input) {
  const ms = new Date(input).getTime();
  return Number.isFinite(ms) ? ms : 0;
}

function reconcileLocalPurchases() {
  const receivedByFarmer = {};
  for (const payment of state.payments) {
    if (normalizePaymentStatusValue(payment.status) !== 'Received') continue;
    const farmerId = String(payment.farmerId || '').trim();
    if (!farmerId) continue;
    receivedByFarmer[farmerId] = moneyValue((receivedByFarmer[farmerId] || 0) + moneyValue(payment.amount));
  }

  const purchases = [...state.producePurchases].sort((a, b) => dateValue(a.createdAt) - dateValue(b.createdAt));
  for (const purchase of purchases) {
    const farmerId = String(purchase.farmerId || '').trim();
    const totalValue = purchaseTotalValue(purchase);
    const available = moneyValue(receivedByFarmer[farmerId] || 0);
    const paidAmountKes = totalValue > 0 ? moneyValue(Math.min(totalValue, available)) : 0;
    const balanceKes = totalValue > 0 ? moneyValue(totalValue - paidAmountKes) : 0;

    purchase.purchasedKgs = moneyValue(purchase.purchasedKgs);
    purchase.pricePerKgKes = Number.isFinite(Number(purchase.pricePerKgKes)) && Number(purchase.pricePerKgKes) > 0
      ? moneyValue(purchase.pricePerKgKes)
      : null;
    purchase.purchaseValueKes = totalValue > 0 ? totalValue : null;
    purchase.paidAmountKes = paidAmountKes;
    purchase.balanceKes = balanceKes;
    purchase.settlementStatus = totalValue <= 0
      ? 'Unpriced'
      : balanceKes <= 0
        ? 'Paid'
        : paidAmountKes > 0
          ? 'Partially Paid'
          : 'Unpaid';

    receivedByFarmer[farmerId] = moneyValue(Math.max(0, available - paidAmountKes));
  }
}

function buildLocalOwedRows(filters) {
  reconcileLocalPurchases();
  const farmersById = new Map((state.farmers || []).map((row) => [row.id, row]));
  const grouped = new Map();

  for (const purchase of state.producePurchases) {
    const createdMs = dateValue(purchase.createdAt);
    if (filters.from !== null && createdMs < filters.from) continue;
    if (filters.to !== null && createdMs > filters.to) continue;

    const farmerId = String(purchase.farmerId || '').trim();
    const balanceKes = moneyValue(purchase.balanceKes);
    if (!farmerId || !(balanceKes > 0)) continue;

    if (!grouped.has(farmerId)) {
      const farmer = farmersById.get(farmerId) || {};
      grouped.set(farmerId, {
        farmerId,
        farmerName: purchase.farmerName || farmer.name || '-',
        phone: farmer.phone || '',
        nationalId: farmer.nationalId || '',
        location: farmer.location || '',
        purchaseCount: 0,
        purchasedKgs: 0,
        totalValueKes: 0,
        paidKes: 0,
        balanceKes: 0,
        lastPurchaseAt: ''
      });
    }

    const row = grouped.get(farmerId);
    row.purchaseCount += 1;
    row.purchasedKgs = moneyValue(row.purchasedKgs + moneyValue(purchase.purchasedKgs));
    row.totalValueKes = moneyValue(row.totalValueKes + moneyValue(purchase.purchaseValueKes));
    row.paidKes = moneyValue(row.paidKes + moneyValue(purchase.paidAmountKes));
    row.balanceKes = moneyValue(row.balanceKes + balanceKes);
    if (!row.lastPurchaseAt || dateValue(purchase.createdAt) > dateValue(row.lastPurchaseAt)) {
      row.lastPurchaseAt = purchase.createdAt;
    }
  }

  let rows = Array.from(grouped.values());
  const q = String(filters.q || '').toLowerCase();
  if (q) {
    rows = rows.filter((row) =>
      [row.farmerName, row.phone, row.nationalId, row.location].some((value) => String(value || '').toLowerCase().includes(q))
    );
  }

  rows.sort((a, b) => {
    if (b.balanceKes !== a.balanceKes) return b.balanceKes - a.balanceKes;
    return dateValue(b.lastPurchaseAt) - dateValue(a.lastPurchaseAt);
  });

  return rows;
}

function renderOwedTable() {
  const selectedCount = owedPicker.rows.filter((row) => owedPicker.selectedFarmerIds.has(row.farmerId)).length;
  const totalBalance = moneyValue(owedPicker.rows.reduce((sum, row) => sum + moneyValue(row.balanceKes), 0));

  elements.owedSummary.textContent =
    `${owedPicker.rows.length} farmer(s) owed, KES ${formatCurrency(totalBalance)} total. ${selectedCount} selected.`;

  if (!owedPicker.rows.length) {
    elements.owedTableWrap.innerHTML = '<div class="empty">No owed farmers in this period.</div>';
    return;
  }

  const rows = owedPicker.rows
    .slice(0, 500)
    .map((row) => `
      <tr>
        <td><input class="table-check" type="checkbox" data-action="toggle-owed-farmer" data-id="${escapeHtml(row.farmerId)}" ${owedPicker.selectedFarmerIds.has(row.farmerId) ? 'checked' : ''}></td>
        <td>${escapeHtml(row.farmerName)}</td>
        <td>${escapeHtml(row.phone || '-')}</td>
        <td>${escapeHtml(row.nationalId || '-')}</td>
        <td>${escapeHtml(String(row.purchaseCount))}</td>
        <td>${escapeHtml(String(row.purchasedKgs))}</td>
        <td>${escapeHtml(formatCurrency(row.totalValueKes))}</td>
        <td>${escapeHtml(formatCurrency(row.paidKes))}</td>
        <td>${escapeHtml(formatCurrency(row.balanceKes))}</td>
        <td>${escapeHtml(dateShort(row.lastPurchaseAt))}</td>
      </tr>
    `)
    .join('');

  elements.owedTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Select</th>
          <th>Farmer</th>
          <th>Phone</th>
          <th>National ID</th>
          <th>Lots</th>
          <th>Purchased Kg</th>
          <th>Total Value</th>
          <th>Paid</th>
          <th>Balance</th>
          <th>Last Purchase</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function normalizeRecommendationStatus(status) {
  const value = String(status || '').trim().toLowerCase();
  if (value === 'approved') return 'approved';
  if (value === 'rejected') return 'rejected';
  return 'pending';
}

function recommendationCounts(rows) {
  const counts = { pending: 0, approved: 0, rejected: 0 };
  for (const row of rows || []) {
    const key = normalizeRecommendationStatus(row.status);
    counts[key] += 1;
  }
  return counts;
}

function filteredLocalPaymentRecommendations(statusFilter, q) {
  let rows = Array.isArray(state.paymentRecommendations) ? [...state.paymentRecommendations] : [];
  if (currentRole() === 'agent') {
    const me = String(state.auth?.user?.username || '').trim();
    rows = rows.filter((row) => String(row.createdBy || '').trim() === me);
  }

  if (statusFilter && statusFilter !== 'all') {
    rows = rows.filter((row) => normalizeRecommendationStatus(row.status) === statusFilter);
  }

  const query = String(q || '').trim().toLowerCase();
  if (query) {
    rows = rows.filter((row) =>
      [row.id, row.farmerName, row.farmerPhone, row.reason, row.createdBy, row.status, row.rejectionReason]
        .some((field) => String(field || '').toLowerCase().includes(query))
    );
  }

  rows.sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());
  return rows;
}

async function loadPaymentRecommendations() {
  if (!['admin', 'agent'].includes(currentRole())) {
    paymentRecommendationPicker.rows = [];
    paymentRecommendationPicker.meta = {
      total: 0,
      counts: { pending: 0, approved: 0, rejected: 0 }
    };
    renderPaymentRecommendations();
    return;
  }

  const status = String(elements.paymentRecommendationStatusFilter?.value || 'all').trim().toLowerCase();
  const q = String(elements.paymentRecommendationSearch?.value || '').trim();

  if (API.enabled) {
    if (!isAuthenticated()) {
      paymentRecommendationPicker.rows = [];
      renderPaymentRecommendations();
      return;
    }
    const query = new URLSearchParams();
    if (status) query.set('status', status);
    if (q) query.set('q', q);
    query.set('limit', '500');

    try {
      const response = await apiRequest(`/api/payments/recommendations?${query.toString()}`);
      const rows = Array.isArray(response.data) ? response.data : [];
      state.paymentRecommendations = rows;
      paymentRecommendationPicker.rows = rows;
      paymentRecommendationPicker.meta = response.meta || {
        total: rows.length,
        counts: recommendationCounts(rows)
      };
      renderPaymentRecommendations();
      persist();
    } catch (error) {
      elements.paymentRecommendationMsg.textContent = error.message;
    }
    return;
  }

  const rows = filteredLocalPaymentRecommendations(status, q);
  paymentRecommendationPicker.rows = rows;
  paymentRecommendationPicker.meta = {
    total: rows.length,
    counts: recommendationCounts(rows)
  };
  renderPaymentRecommendations();
}

function renderPaymentRecommendations() {
  if (!elements.paymentRecommendationTableWrap) return;

  const role = currentRole();
  const isAgent = role === 'agent';
  const isAdmin = role === 'admin';

  if (elements.paymentRecommendationForm) {
    elements.paymentRecommendationForm.hidden = !isAgent;
  }
  if (elements.paymentRecommendationAgentHint) {
    elements.paymentRecommendationAgentHint.hidden = !isAgent;
  }

  if (!isAdmin && !isAgent) {
    if (elements.paymentRecommendationSummary) {
      elements.paymentRecommendationSummary.textContent = 'Only admin and agents can view payment recommendations.';
    }
    elements.paymentRecommendationTableWrap.innerHTML =
      '<div class="empty">Sign in as admin or agent to access this queue.</div>';
    return;
  }

  const rows =
    Array.isArray(paymentRecommendationPicker.rows) && paymentRecommendationPicker.rows.length
      ? paymentRecommendationPicker.rows
      : Array.isArray(state.paymentRecommendations)
        ? state.paymentRecommendations
        : [];
  const counts = paymentRecommendationPicker.meta?.counts || recommendationCounts(rows);
  const total = Number(paymentRecommendationPicker.meta?.total ?? rows.length);
  if (elements.paymentRecommendationSummary) {
    elements.paymentRecommendationSummary.textContent =
      `Total ${total} | Pending ${counts.pending || 0} | Approved ${counts.approved || 0} | Rejected ${counts.rejected || 0}`;
  }

  if (!rows.length) {
    elements.paymentRecommendationTableWrap.innerHTML = '<div class="empty">No payment recommendations in this filter.</div>';
    return;
  }

  const rowsHtml = rows
    .slice(0, 300)
    .map((row) => {
      const status = normalizeRecommendationStatus(row.status);
      const statusLabel = status.charAt(0).toUpperCase() + status.slice(1);
      const canAdminAction = isAdmin && status === 'pending';
      const actionButtons = canAdminAction
        ? `
          <button class="table-btn" data-action="approve-payment-recommendation" data-id="${escapeHtml(row.id)}">Approve</button>
          <button class="table-btn danger" data-action="reject-payment-recommendation" data-id="${escapeHtml(row.id)}">Reject</button>
        `
        : '-';

      return `
        <tr>
          <td>${escapeHtml(row.id || '-')}</td>
          <td>${escapeHtml(row.farmerName || '-')}</td>
          <td>${escapeHtml(row.farmerPhone || '-')}</td>
          <td>${escapeHtml(formatCurrency(row.amount || 0))}</td>
          <td>${escapeHtml(row.reason || '-')}</td>
          <td>${escapeHtml(row.createdBy || '-')}</td>
          <td><span class="severity-badge ${status === 'rejected' ? 'high' : status === 'approved' ? '' : 'medium'}">${escapeHtml(statusLabel)}</span></td>
          <td>${escapeHtml(row.decisionNote || row.rejectionReason || '-')}</td>
          <td>${escapeHtml(dateShort(row.updatedAt || row.createdAt))}</td>
          <td class="actions">${actionButtons}</td>
        </tr>
      `;
    })
    .join('');

  elements.paymentRecommendationTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Farmer</th>
          <th>Phone</th>
          <th>Amount (KES)</th>
          <th>Reason</th>
          <th>Requested By</th>
          <th>Status</th>
          <th>Decision</th>
          <th>Updated</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;
}

async function refreshOwedRows(resetSelection = false) {
  if (currentRole() !== 'admin') {
    owedPicker.rows = [];
    owedPicker.selectedFarmerIds.clear();
    elements.owedSummary.textContent = 'Owed list is visible to admin only.';
    elements.owedTableWrap.innerHTML =
      '<div class="empty">Sign in as admin to view owed balances. Agents can still submit payment recommendations.</div>';
    return;
  }

  const filters = resolveOwedFilters();
  if (resetSelection) {
    owedPicker.selectedFarmerIds.clear();
  }

  if (API.enabled) {
    if (!isAuthenticated()) {
      owedPicker.rows = [];
      elements.owedSummary.textContent = 'Sign in to view owed balances.';
      elements.owedTableWrap.innerHTML = '<div class="empty">Sign in to load owed balances from backend.</div>';
      return;
    }

    const query = new URLSearchParams();
    if (filters.period) query.set('period', filters.period);
    if (filters.fromRaw) query.set('from', filters.fromRaw);
    if (filters.toRaw) query.set('to', filters.toRaw);
    if (filters.q) query.set('q', filters.q);
    query.set('limit', '2000');

    try {
      const response = await apiRequest(`/api/payments/owed?${query.toString()}`);
      owedPicker.rows = response.data || [];
      owedPicker.meta = response.meta || owedPicker.meta;
    } catch (error) {
      elements.owedMsg.textContent = error.message;
      return;
    }
  } else {
    owedPicker.rows = buildLocalOwedRows(filters);
    owedPicker.meta = {
      period: filters.period,
      from: filters.from ? new Date(filters.from).toISOString() : '',
      to: filters.to ? new Date(filters.to).toISOString() : '',
      count: owedPicker.rows.length,
      totalBalanceKes: moneyValue(owedPicker.rows.reduce((sum, row) => sum + moneyValue(row.balanceKes), 0))
    };
  }

  for (const farmerId of Array.from(owedPicker.selectedFarmerIds)) {
    if (!owedPicker.rows.some((row) => row.farmerId === farmerId)) {
      owedPicker.selectedFarmerIds.delete(farmerId);
    }
  }

  renderOwedTable();
}

async function settleSelectedOwed(status) {
  if (currentRole() !== 'admin') {
    elements.owedMsg.textContent = 'Only admin can run owed settlements.';
    return;
  }

  const selected = owedPicker.rows.filter((row) => owedPicker.selectedFarmerIds.has(row.farmerId) && row.balanceKes > 0);
  if (!selected.length) {
    elements.owedMsg.textContent = 'Select at least one owed farmer.';
    return;
  }

  const filters = resolveOwedFilters();
  const farmerIds = selected.map((row) => row.farmerId);

  try {
    if (API.enabled) {
      const response = await apiRequest('/api/payments/settle', {
        method: 'POST',
        body: {
          farmerIds,
          status,
          method: 'M-PESA(Mock)',
          period: filters.period,
          from: filters.fromRaw,
          to: filters.toRaw
        }
      });
      elements.owedMsg.textContent =
        `${response.data.createdCount} payment(s) created, total KES ${formatCurrency(response.data.totalAmount)} (${status}).`;
      await fetchAllData();
      await refreshOwedRows(false);
      return;
    }

    for (let index = 0; index < selected.length; index += 1) {
      const row = selected[index];
      const record = {
        id: `TX-${Date.now()}-${index + 1}`,
        farmerId: row.farmerId,
        farmerName: row.farmerName,
        amount: moneyValue(row.balanceKes),
        ref: `${status === 'Received' ? 'MPB' : 'SET'}${Date.now().toString().slice(-8)}${String(index + 1).padStart(2, '0')}`,
        status,
        method: status === 'Received' ? 'M-PESA(Mock)' : 'M-PESA',
        notes: 'Settlement from owed panel',
        source: 'purchase-owed',
        createdBy: 'local',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      state.payments.unshift(record);
    }

    elements.owedMsg.textContent = `${selected.length} payment(s) created (${status}) in local mode.`;
    persist();
    renderAll();
    await refreshOwedRows(false);
  } catch (error) {
    elements.owedMsg.textContent = error.message;
  }
}

function csvCell(value) {
  const raw = value == null ? '' : String(value);
  if (/[",\n]/.test(raw)) {
    return `"${raw.replace(/"/g, '""')}"`;
  }
  return raw;
}

function rowsToCsv(rows, headers) {
  const head = headers.map((header) => header.label).join(',');
  const body = rows.map((row) => headers.map((header) => csvCell(row[header.key])).join(',')).join('\n');
  return `${head}\n${body}`;
}

async function exportOwedCsv() {
  if (currentRole() !== 'admin') {
    elements.owedMsg.textContent = 'Only admin can export owed balances.';
    return;
  }

  const filters = resolveOwedFilters();
  const query = new URLSearchParams();
  if (filters.period) query.set('period', filters.period);
  if (filters.fromRaw) query.set('from', filters.fromRaw);
  if (filters.toRaw) query.set('to', filters.toRaw);
  if (filters.q) query.set('q', filters.q);

  try {
    let content = '';
    if (API.enabled) {
      content = await apiRequest(`/api/exports/payments-owed.csv?${query.toString()}`, {
        response: 'text'
      });
    } else {
      const rows = buildLocalOwedRows(filters);
      content = rowsToCsv(rows, [
        { key: 'farmerId', label: 'farmerId' },
        { key: 'farmerName', label: 'farmerName' },
        { key: 'phone', label: 'phone' },
        { key: 'nationalId', label: 'nationalId' },
        { key: 'location', label: 'location' },
        { key: 'purchaseCount', label: 'purchaseCount' },
        { key: 'purchasedKgs', label: 'purchasedKgs' },
        { key: 'totalValueKes', label: 'totalValueKes' },
        { key: 'paidKes', label: 'paidKes' },
        { key: 'balanceKes', label: 'balanceKes' },
        { key: 'lastPurchaseAt', label: 'lastPurchaseAt' }
      ]);
    }

    const period = filters.period || 'all';
    const rangeLabel = exportRangeLabel(period, filters.fromRaw, filters.toRaw);
    const format = String(elements.owedExportFormat?.value || 'csv').toLowerCase();
    const formatLabel = await exportDatasetFile({
      type: 'payments-owed',
      csvText: content,
      format,
      rangeLabel,
      fileSuffix: period
    });

    elements.owedMsg.textContent = `Owed balances ${formatLabel} exported.`;
  } catch (error) {
    elements.owedMsg.textContent = error.message;
  }
}

function bindSms() {
  elements.smsTargetMode.addEventListener('change', async () => {
    clearMessages();
    updateSmsComposerUi();
    if (currentSmsMode() === 'selected') {
      await loadSmsRecipients(true);
    }
  });

  elements.smsRecipientSearchBtn.addEventListener('click', async () => {
    clearMessages();
    await loadSmsRecipients(true);
  });

  elements.smsRecipientSearch.addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    clearMessages();
    await loadSmsRecipients(true);
  });

  elements.smsPrevPageBtn.addEventListener('click', async () => {
    if (smsPicker.offset === 0) return;
    smsPicker.offset = Math.max(0, smsPicker.offset - smsPicker.limit);
    await loadSmsRecipients(false);
  });

  elements.smsNextPageBtn.addEventListener('click', async () => {
    if (smsPicker.offset + smsPicker.rows.length >= smsPicker.total) return;
    smsPicker.offset += smsPicker.limit;
    await loadSmsRecipients(false);
  });

  elements.smsSelectVisibleBtn.addEventListener('click', () => {
    for (const row of smsPicker.rows) {
      smsPicker.selectedIds.add(row.id);
    }
    renderSmsRecipientList();
  });

  elements.smsClearSelectedBtn.addEventListener('click', () => {
    smsPicker.selectedIds.clear();
    renderSmsRecipientList();
  });

  elements.smsRecipientListWrap.addEventListener('change', (event) => {
    const checkbox = event.target.closest('input[data-farmer-id]');
    if (!checkbox) return;
    const farmerId = checkbox.dataset.farmerId;
    if (!farmerId) return;

    if (checkbox.checked) {
      smsPicker.selectedIds.add(farmerId);
    } else {
      smsPicker.selectedIds.delete(farmerId);
    }
    renderSmsRecipientList();
  });

  elements.smsForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    if (!['admin', 'agent'].includes(currentRole())) {
      elements.smsMsg.textContent = 'Only admin or agent can send SMS.';
      return;
    }

    const mode = currentSmsMode();
    const message = elements.smsMessage.value.trim();

    if (!message) {
      elements.smsMsg.textContent = 'Message is required.';
      return;
    }

    if (mode !== 'single' && currentRole() !== 'admin') {
      elements.smsMsg.textContent = 'Bulk SMS is admin-only. Switch mode to Single Mobile Number.';
      return;
    }

    try {
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.smsMsg.textContent = 'Sign in first to use backend mode.';
          return;
        }

        if (mode === 'single') {
          const phone = elements.smsPhone.value.trim();
          if (!phone) {
            elements.smsMsg.textContent = 'Phone number is required for single send mode.';
            return;
          }

          await apiRequest('/api/sms/send', {
            method: 'POST',
            body: { phone, message }
          });
          elements.smsMsg.textContent = 'SMS sent.';
        } else if (mode === 'all') {
          const response = await apiRequest('/api/sms/send-bulk', {
            method: 'POST',
            body: { mode: 'all', message }
          });
          elements.smsMsg.textContent = `Bulk SMS sent to ${response.data.sentCount} recipients.`;
        } else {
          if (smsPicker.selectedIds.size === 0) {
            elements.smsMsg.textContent = 'Select at least one recipient first.';
            return;
          }

          const response = await apiRequest('/api/sms/send-bulk', {
            method: 'POST',
            body: {
              mode: 'selected',
              farmerIds: [...smsPicker.selectedIds],
              message
            }
          });
          elements.smsMsg.textContent = `Bulk SMS sent to ${response.data.sentCount} recipients.`;
        }

        elements.smsMessage.value = '';
        await fetchAllData();
        return;
      }

      if (mode === 'single') {
        const phone = elements.smsPhone.value.trim();
        if (!phone) {
          elements.smsMsg.textContent = 'Phone number is required for single send mode.';
          return;
        }

        state.smsLogs.unshift({
          id: `SMS-${Date.now()}`,
          farmerId: '',
          farmerName: '',
          phone,
          message,
          provider: 'Local Mock',
          ownerCostKes: SMS_OWNER_COST_PER_MESSAGE_KES,
          status: 'Sent',
          createdBy: 'local',
          createdAt: new Date().toISOString()
        });
        elements.smsMsg.textContent = 'SMS sent (local mode).';
      } else {
        if (mode === 'selected' && smsPicker.selectedIds.size === 0) {
          elements.smsMsg.textContent = 'Select at least one recipient first.';
          return;
        }

        let targets = [];
        if (mode === 'all') {
          targets = state.farmers.filter((row) => String(row.phone || '').trim());
        } else {
          targets = state.farmers.filter((row) => smsPicker.selectedIds.has(row.id) && String(row.phone || '').trim());
        }

        const byPhone = new Map();
        for (const farmer of targets) {
          const phone = String(farmer.phone || '').trim();
          if (!phone || byPhone.has(phone)) continue;
          byPhone.set(phone, farmer);
        }

        const now = new Date().toISOString();
        const logs = [];
        for (const [phone, farmer] of byPhone.entries()) {
          logs.push({
            id: `SMS-${Date.now()}-${logs.length + 1}`,
            farmerId: farmer.id || '',
            farmerName: farmer.name || '',
            phone,
            message,
            provider: 'Local Mock',
            ownerCostKes: SMS_OWNER_COST_PER_MESSAGE_KES,
            status: 'Sent',
            createdBy: 'local',
            createdAt: now
          });
        }

        state.smsLogs = logs.reverse().concat(state.smsLogs);
        elements.smsMsg.textContent = `Bulk SMS sent to ${logs.length} recipients (local mode).`;
      }

      elements.smsMessage.value = '';
      renderAll();
      persist();
    } catch (error) {
      elements.smsMsg.textContent = error.message;
    }
  });

  updateSmsComposerUi();
  void loadSmsRecipients(true);
}

function bindAiTools() {
  if (elements.smsDraftWrap) {
    elements.smsDraftWrap.hidden = true;
  }
  if (elements.farmerImportQaWrap) {
    elements.farmerImportQaWrap.hidden = true;
  }
  if (elements.qcAiWrap) {
    elements.qcAiWrap.hidden = false;
  }
  if (elements.paymentRiskWrap) {
    elements.paymentRiskWrap.hidden = false;
  }
  syncPaymentRiskFilterUi();

  if (elements.smsDraftBtn) {
    elements.smsDraftBtn.addEventListener('click', async () => {
      clearMessages();

      if (!API.enabled) {
        elements.smsDraftMsg.textContent = 'AI drafting requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.smsDraftMsg.textContent = 'Sign in first to use AI drafting.';
        return;
      }
      if (!['admin', 'agent'].includes(currentRole())) {
        elements.smsDraftMsg.textContent = 'Only admin or agent can draft SMS.';
        return;
      }

      const purpose = String(elements.smsDraftPurpose?.value || '').trim();
      if (!purpose) {
        elements.smsDraftMsg.textContent = 'Enter a message goal first.';
        return;
      }

      const payload = {
        purpose,
        audience: String(elements.smsDraftAudience?.value || '').trim(),
        tone: String(elements.smsDraftTone?.value || '').trim() || 'professional',
        language: String(elements.smsDraftLanguage?.value || '').trim() || 'bilingual',
        maxLength: String(elements.smsDraftMaxLength?.value || '').trim()
      };

      try {
        const response = await apiRequest('/api/ai/sms-draft', {
          method: 'POST',
          body: payload
        });
        renderSmsDraftResult(response.data || {});
      } catch (error) {
        elements.smsDraftMsg.textContent = error.message;
      }
    });
  }

  if (elements.smsDraftWrap) {
    elements.smsDraftWrap.addEventListener('click', (event) => {
      const applyBtn = event.target.closest('[data-apply-sms-draft]');
      if (!applyBtn) return;
      const index = Number(applyBtn.getAttribute('data-apply-sms-draft'));
      const message = String(aiState.smsDrafts[index]?.message || '').trim();
      if (!message) return;
      elements.smsMessage.value = message;
      elements.smsDraftMsg.textContent = 'Draft applied to message box.';
    });
  }

  if (elements.copilotAskBtn) {
    elements.copilotAskBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.copilotMsg.textContent = 'Admin Copilot requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.copilotMsg.textContent = 'Sign in first to use Admin Copilot.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.copilotMsg.textContent = 'Admin Copilot is admin-only.';
        return;
      }

      const question = String(elements.copilotPrompt?.value || '').trim();
      if (!question) {
        elements.copilotMsg.textContent = 'Ask a question first.';
        return;
      }

      try {
        const response = await apiRequest('/api/ai/copilot', {
          method: 'POST',
          body: { question }
        });
        renderCopilotResult(response.data || {});
      } catch (error) {
        elements.copilotMsg.textContent = error.message;
      }
    });
  }

  if (elements.copilotPrompt) {
    elements.copilotPrompt.addEventListener('keydown', async (event) => {
      if (event.key !== 'Enter') return;
      event.preventDefault();
      elements.copilotAskBtn?.click();
    });
  }

  if (elements.qcAiRunBtn) {
    elements.qcAiRunBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.qcAiMsg.textContent = 'QC intelligence requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.qcAiMsg.textContent = 'Sign in first to run QC intelligence.';
        return;
      }
      if (!['admin', 'agent'].includes(currentRole())) {
        elements.qcAiMsg.textContent = 'Only admin or agent can run QC intelligence.';
        return;
      }

      try {
        const qcRecordId = String(elements.qcAiRecord?.value || '').trim();
        const response = await apiRequest('/api/ai/qc-intelligence', {
          method: 'POST',
          body: {
            qcRecordId,
            limit: 30
          }
        });
        renderQcAiResult(response.data || {});
      } catch (error) {
        elements.qcAiMsg.textContent = error.message;
      }
    });
  }

  if (elements.paymentRiskPeriod) {
    elements.paymentRiskPeriod.addEventListener('change', () => {
      syncPaymentRiskFilterUi();
    });
  }

  if (elements.paymentRiskRunBtn) {
    elements.paymentRiskRunBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.paymentRiskMsg.textContent = 'Payment risk check requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.paymentRiskMsg.textContent = 'Sign in first to run payment risk check.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.paymentRiskMsg.textContent = 'Payment risk check is admin-only.';
        return;
      }

      try {
        const period = String(elements.paymentRiskPeriod?.value || 'week').trim().toLowerCase();
        const payload = { period };
        if (period === 'custom') {
          const from = String(elements.paymentRiskFrom?.value || '').trim();
          const to = String(elements.paymentRiskTo?.value || '').trim();
          if (!from || !to) {
            elements.paymentRiskMsg.textContent = 'Choose start and end date for custom range.';
            return;
          }
          payload.from = from;
          payload.to = to;
        }

        const response = await apiRequest('/api/ai/payment-risk', {
          method: 'POST',
          body: payload
        });
        renderPaymentRiskResult(response.data || {});
      } catch (error) {
        elements.paymentRiskMsg.textContent = error.message;
      }
    });
  }

  if (elements.briefRunNowBtn) {
    elements.briefRunNowBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.briefMsg.textContent = 'Executive briefs require backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.briefMsg.textContent = 'Sign in first to run an executive brief.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.briefMsg.textContent = 'Run-Now brief is admin-only.';
        return;
      }

      try {
        const response = await apiRequest('/api/ai/briefs/run-now', {
          method: 'POST',
          body: {}
        });
        const brief = response.data || null;
        if (brief) {
          aiState.briefs = [brief, ...aiState.briefs.filter((row) => row.id !== brief.id)].slice(0, 60);
          if (brief.alert) {
            aiState.alerts = [brief.alert, ...aiState.alerts.filter((row) => row.id !== brief.alert.id)].slice(0, 60);
          }
          aiState.lastResponseIds['executive-brief'] = String(brief.id || '');
          syncAiFeedbackResponseIdFromLatest(true);
          renderExecutiveBriefs();
        }
        elements.briefMsg.textContent = brief?.warning || 'Executive brief generated.';
      } catch (error) {
        elements.briefMsg.textContent = error.message;
      }
    });
  }

  if (elements.briefRefreshBtn) {
    elements.briefRefreshBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.briefMsg.textContent = 'Executive briefs require backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.briefMsg.textContent = 'Sign in first to refresh executive briefs.';
        return;
      }
      if (!['admin', 'agent'].includes(currentRole())) {
        elements.briefMsg.textContent = 'Only admin or agent can view executive briefs.';
        return;
      }
      try {
        await fetchAiBriefs();
        renderExecutiveBriefs();
        elements.briefMsg.textContent = 'Executive briefs refreshed.';
      } catch (error) {
        elements.briefMsg.textContent = error.message;
      }
    });
  }

  if (elements.proposalRefreshBtn) {
    elements.proposalRefreshBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.proposalMsg.textContent = 'Proposal queue requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.proposalMsg.textContent = 'Sign in first to refresh proposal queue.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.proposalMsg.textContent = 'Proposal queue is admin-only.';
        return;
      }
      try {
        await fetchAiProposals();
        renderProposalQueue();
        elements.proposalMsg.textContent = 'Proposal queue refreshed.';
      } catch (error) {
        elements.proposalMsg.textContent = error.message;
      }
    });
  }

  if (elements.proposalStatusFilter) {
    elements.proposalStatusFilter.addEventListener('change', async () => {
      if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
        renderProposalQueue();
        return;
      }
      try {
        await fetchAiProposals();
      } catch (error) {
        elements.proposalMsg.textContent = error.message;
      }
      renderProposalQueue();
    });
  }

  if (elements.proposalForm) {
    elements.proposalForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      clearMessages();
      if (!API.enabled) {
        elements.proposalMsg.textContent = 'Proposal queue requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.proposalMsg.textContent = 'Sign in first to create a proposal.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.proposalMsg.textContent = 'Only admin can create proposals.';
        return;
      }

      const actionType = String(elements.proposalActionType?.value || '').trim();
      const confidenceRaw = String(elements.proposalConfidence?.value || '').trim();
      const confidence = confidenceRaw ? Number(confidenceRaw) : undefined;
      const title = String(elements.proposalTitle?.value || '').trim();
      const description = String(elements.proposalDescription?.value || '').trim();
      const payloadText = String(elements.proposalPayload?.value || '').trim();

      let proposalPayload = {};
      if (payloadText) {
        try {
          const parsed = JSON.parse(payloadText);
          if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
            elements.proposalMsg.textContent = 'Payload JSON must be an object.';
            return;
          }
          proposalPayload = parsed;
        } catch {
          elements.proposalMsg.textContent = 'Payload JSON is invalid.';
          return;
        }
      }

      try {
        const response = await apiRequest('/api/ai/proposals', {
          method: 'POST',
          body: {
            actionType,
            confidence,
            title,
            description,
            payload: proposalPayload
          }
        });
        const proposal = response.data || null;
        if (proposal) {
          aiState.proposals = [proposal, ...aiState.proposals.filter((row) => row.id !== proposal.id)];
          elements.proposalForm.reset();
          if (elements.proposalActionType) {
            elements.proposalActionType.value = proposal.actionType || 'payment_risk_to_tasks';
          }
          renderProposalQueue();
        }
        elements.proposalMsg.textContent = 'Proposal created.';
      } catch (error) {
        elements.proposalMsg.textContent = error.message;
      }
    });
  }

  if (elements.proposalWrap) {
    elements.proposalWrap.addEventListener('click', async (event) => {
      const button = event.target.closest('[data-proposal-action]');
      if (!button) return;
      clearMessages();
      if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
        elements.proposalMsg.textContent = 'Sign in as admin to manage proposals.';
        return;
      }

      const proposalId = String(button.getAttribute('data-proposal-id') || '').trim();
      const action = String(button.getAttribute('data-proposal-action') || '').trim().toLowerCase();
      if (!proposalId || !action) return;

      button.disabled = true;
      try {
        if (action === 'approve') {
          await apiRequest(`/api/ai/proposals/${encodeURIComponent(proposalId)}/approve`, {
            method: 'POST',
            body: {}
          });
          elements.proposalMsg.textContent = 'Proposal approved.';
        } else if (action === 'reject') {
          const reason = window.prompt('Optional rejection reason:', '') || '';
          await apiRequest(`/api/ai/proposals/${encodeURIComponent(proposalId)}/reject`, {
            method: 'POST',
            body: { reason }
          });
          elements.proposalMsg.textContent = 'Proposal rejected.';
        } else if (action === 'execute') {
          const response = await apiRequest(`/api/ai/proposals/${encodeURIComponent(proposalId)}/execute`, {
            method: 'POST',
            body: {}
          });
          const resultSummary = String(response?.result?.summary || '').trim();
          elements.proposalMsg.textContent = resultSummary || 'Proposal executed.';
          await fetchAiOpsTasks();
          await fetchAiBriefs();
          renderOpsTasks();
          renderExecutiveBriefs();
        }

        await fetchAiProposals();
        renderProposalQueue();
      } catch (error) {
        elements.proposalMsg.textContent = error.message;
      } finally {
        button.disabled = false;
      }
    });
  }

  if (elements.opsFromPaymentRiskBtn) {
    elements.opsFromPaymentRiskBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.opsTaskMsg.textContent = 'Ops task creation requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.opsTaskMsg.textContent = 'Sign in first to create ops tasks.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.opsTaskMsg.textContent = 'Create-from-AI task actions are admin-only.';
        return;
      }

      const period = String(elements.paymentRiskPeriod?.value || 'week').trim().toLowerCase();
      const body = {
        source: 'payment-risk',
        period
      };
      if (period === 'custom') {
        const from = String(elements.paymentRiskFrom?.value || '').trim();
        const to = String(elements.paymentRiskTo?.value || '').trim();
        if (!from || !to) {
          elements.opsTaskMsg.textContent = 'Choose start and end date first for custom range.';
          return;
        }
        body.from = from;
        body.to = to;
      }

      try {
        const response = await apiRequest('/api/ops/tasks/from-ai', {
          method: 'POST',
          body
        });
        const createdCount = Number(response.meta?.createdCount || (Array.isArray(response.data) ? response.data.length : 0));
        elements.opsTaskMsg.textContent = `Created ${createdCount} ops task(s) from payment risk.`;
        await fetchAiOpsTasks();
        renderOpsTasks();
      } catch (error) {
        elements.opsTaskMsg.textContent = error.message;
      }
    });
  }

  if (elements.opsFromQcBtn) {
    elements.opsFromQcBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.opsTaskMsg.textContent = 'Ops task creation requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.opsTaskMsg.textContent = 'Sign in first to create ops tasks.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.opsTaskMsg.textContent = 'Create-from-AI task actions are admin-only.';
        return;
      }

      try {
        const qcRecordId = String(elements.qcAiRecord?.value || '').trim();
        const response = await apiRequest('/api/ops/tasks/from-ai', {
          method: 'POST',
          body: {
            source: 'qc-intelligence',
            qcRecordId,
            limit: 30
          }
        });
        const createdCount = Number(response.meta?.createdCount || (Array.isArray(response.data) ? response.data.length : 0));
        elements.opsTaskMsg.textContent = `Created ${createdCount} ops task(s) from QC intelligence.`;
        await fetchAiOpsTasks();
        renderOpsTasks();
      } catch (error) {
        elements.opsTaskMsg.textContent = error.message;
      }
    });
  }

  if (elements.opsRefreshBtn) {
    elements.opsRefreshBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.opsTaskMsg.textContent = 'Ops tasks require backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.opsTaskMsg.textContent = 'Sign in first to refresh ops tasks.';
        return;
      }
      if (!['admin', 'agent'].includes(currentRole())) {
        elements.opsTaskMsg.textContent = 'Only admin or agent can view ops tasks.';
        return;
      }
      try {
        await fetchAiOpsTasks();
        renderOpsTasks();
        elements.opsTaskMsg.textContent = 'Ops tasks refreshed.';
      } catch (error) {
        elements.opsTaskMsg.textContent = error.message;
      }
    });
  }

  if (elements.opsTaskStatusFilter) {
    elements.opsTaskStatusFilter.addEventListener('change', async () => {
      if (!API.enabled || !isAuthenticated() || !['admin', 'agent'].includes(currentRole())) {
        renderOpsTasks();
        return;
      }
      try {
        await fetchAiOpsTasks();
      } catch (error) {
        elements.opsTaskMsg.textContent = error.message;
      }
      renderOpsTasks();
    });
  }

  if (elements.opsTaskSeverityFilter) {
    elements.opsTaskSeverityFilter.addEventListener('change', async () => {
      if (!API.enabled || !isAuthenticated() || !['admin', 'agent'].includes(currentRole())) {
        renderOpsTasks();
        return;
      }
      try {
        await fetchAiOpsTasks();
      } catch (error) {
        elements.opsTaskMsg.textContent = error.message;
      }
      renderOpsTasks();
    });
  }

  if (elements.opsTaskWrap) {
    elements.opsTaskWrap.addEventListener('click', async (event) => {
      const saveBtn = event.target.closest('[data-task-save]');
      if (!saveBtn) return;
      clearMessages();
      if (!API.enabled || !isAuthenticated() || !['admin', 'agent'].includes(currentRole())) {
        elements.opsTaskMsg.textContent = 'Sign in as admin or agent to update tasks.';
        return;
      }

      const taskId = String(saveBtn.getAttribute('data-task-save') || '').trim();
      const card = saveBtn.closest('[data-task-id]');
      if (!taskId || !card) return;

      const status = String(card.querySelector('[data-task-field="status"]')?.value || '').trim();
      const assignedTo = String(card.querySelector('[data-task-field="assignedTo"]')?.value || '').trim();
      const notes = String(card.querySelector('[data-task-field="notes"]')?.value || '').trim();

      saveBtn.disabled = true;
      try {
        await apiRequest(`/api/ops/tasks/${encodeURIComponent(taskId)}`, {
          method: 'PATCH',
          body: { status, assignedTo, notes }
        });
        await fetchAiOpsTasks();
        renderOpsTasks();
        elements.opsTaskMsg.textContent = 'Task updated.';
      } catch (error) {
        elements.opsTaskMsg.textContent = error.message;
      } finally {
        saveBtn.disabled = false;
      }
    });
  }

  if (elements.knowledgeForm) {
    elements.knowledgeForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      clearMessages();
      if (!API.enabled) {
        elements.knowledgeMsg.textContent = 'Knowledge base requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.knowledgeMsg.textContent = 'Sign in first to add knowledge.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.knowledgeMsg.textContent = 'Only admin can add knowledge documents.';
        return;
      }

      const title = String(elements.knowledgeTitle?.value || '').trim();
      const source = String(elements.knowledgeSource?.value || '').trim();
      const tags = String(elements.knowledgeTags?.value || '').trim();
      const content = String(elements.knowledgeContent?.value || '').trim();
      if (!title || !content) {
        elements.knowledgeMsg.textContent = 'Title and content are required.';
        return;
      }

      try {
        await apiRequest('/api/ai/knowledge', {
          method: 'POST',
          body: { title, source, tags, content }
        });
        elements.knowledgeForm.reset();
        await fetchAiKnowledgeDocs();
        renderKnowledgeDocs();
        elements.knowledgeMsg.textContent = 'Knowledge document added.';
      } catch (error) {
        elements.knowledgeMsg.textContent = error.message;
      }
    });
  }

  if (elements.knowledgeSearchBtn) {
    elements.knowledgeSearchBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.knowledgeMsg.textContent = 'Knowledge base requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.knowledgeMsg.textContent = 'Sign in first to search knowledge.';
        return;
      }
      if (!['admin', 'agent'].includes(currentRole())) {
        elements.knowledgeMsg.textContent = 'Only admin or agent can view knowledge.';
        return;
      }

      try {
        await fetchAiKnowledgeDocs();
        renderKnowledgeDocs();
        elements.knowledgeMsg.textContent = 'Knowledge list refreshed.';
      } catch (error) {
        elements.knowledgeMsg.textContent = error.message;
      }
    });
  }

  if (elements.knowledgeWrap) {
    elements.knowledgeWrap.addEventListener('click', async (event) => {
      const deleteBtn = event.target.closest('[data-knowledge-delete]');
      if (!deleteBtn) return;
      clearMessages();
      if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
        elements.knowledgeMsg.textContent = 'Sign in as admin to delete knowledge documents.';
        return;
      }
      const docId = String(deleteBtn.getAttribute('data-knowledge-delete') || '').trim();
      if (!docId) return;
      if (!window.confirm('Delete this knowledge document?')) return;

      deleteBtn.disabled = true;
      try {
        await apiRequest(`/api/ai/knowledge/${encodeURIComponent(docId)}`, {
          method: 'DELETE',
          body: {}
        });
        await fetchAiKnowledgeDocs();
        renderKnowledgeDocs();
        elements.knowledgeMsg.textContent = 'Knowledge document deleted.';
      } catch (error) {
        elements.knowledgeMsg.textContent = error.message;
      } finally {
        deleteBtn.disabled = false;
      }
    });
  }

  if (elements.aiFeedbackTool) {
    elements.aiFeedbackTool.addEventListener('change', () => {
      syncAiFeedbackResponseIdFromLatest(true);
    });
  }

  if (elements.aiFeedbackForm) {
    elements.aiFeedbackForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      clearMessages();
      if (!API.enabled) {
        elements.aiFeedbackMsg.textContent = 'AI feedback requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.aiFeedbackMsg.textContent = 'Sign in first to submit feedback.';
        return;
      }
      if (!['admin', 'agent'].includes(currentRole())) {
        elements.aiFeedbackMsg.textContent = 'Only admin or agent can submit AI feedback.';
        return;
      }

      try {
        const response = await apiRequest('/api/ai/feedback', {
          method: 'POST',
          body: {
            tool: selectedAiFeedbackTool() || 'other',
            rating: String(elements.aiFeedbackRating?.value || 'neutral').trim(),
            responseId: String(elements.aiFeedbackResponseId?.value || '').trim(),
            note: String(elements.aiFeedbackNote?.value || '').trim(),
            prompt: String(elements.copilotPrompt?.value || '').trim()
          }
        });
        const entry = response.data || null;
        if (entry && currentRole() === 'admin') {
          await fetchAiFeedbackSummary();
        }
        renderAiFeedbackAndEvals();
        elements.aiFeedbackMsg.textContent = 'Feedback submitted.';
      } catch (error) {
        elements.aiFeedbackMsg.textContent = error.message;
      }
    });
  }

  if (elements.aiFeedbackSummaryBtn) {
    elements.aiFeedbackSummaryBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.aiFeedbackMsg.textContent = 'Feedback summary requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.aiFeedbackMsg.textContent = 'Sign in first to refresh feedback summary.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.aiFeedbackMsg.textContent = 'Feedback summary is admin-only.';
        return;
      }

      try {
        await fetchAiFeedbackSummary();
        renderAiFeedbackAndEvals();
        elements.aiFeedbackMsg.textContent = 'Feedback summary refreshed.';
      } catch (error) {
        elements.aiFeedbackMsg.textContent = error.message;
      }
    });
  }

  if (elements.aiEvalRunBtn) {
    elements.aiEvalRunBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.aiFeedbackMsg.textContent = 'AI eval suite requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.aiFeedbackMsg.textContent = 'Sign in first to run AI evals.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.aiFeedbackMsg.textContent = 'AI eval suite is admin-only.';
        return;
      }

      try {
        await apiRequest('/api/ai/evals/run', {
          method: 'POST',
          body: {}
        });
        await fetchAiEvalRuns();
        renderAiFeedbackAndEvals();
        elements.aiFeedbackMsg.textContent = 'AI eval run completed.';
      } catch (error) {
        elements.aiFeedbackMsg.textContent = error.message;
      }
    });
  }

  if (elements.aiEvalRefreshBtn) {
    elements.aiEvalRefreshBtn.addEventListener('click', async () => {
      clearMessages();
      if (!API.enabled) {
        elements.aiFeedbackMsg.textContent = 'AI eval history requires backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.aiFeedbackMsg.textContent = 'Sign in first to refresh AI eval runs.';
        return;
      }
      if (currentRole() !== 'admin') {
        elements.aiFeedbackMsg.textContent = 'AI eval history is admin-only.';
        return;
      }

      try {
        await fetchAiEvalRuns();
        renderAiFeedbackAndEvals();
        elements.aiFeedbackMsg.textContent = 'AI eval history refreshed.';
      } catch (error) {
        elements.aiFeedbackMsg.textContent = error.message;
      }
    });
  }
}

function renderSmsDraftResult(data) {
  const responseId = String(data?.responseId || '').trim();
  if (responseId) {
    aiState.lastResponseIds['sms-draft'] = responseId;
    syncAiFeedbackResponseIdFromLatest();
  }
  const drafts = Array.isArray(data?.drafts) ? data.drafts : [];
  const source = String(data?.source || 'local-rules');
  const warning = String(data?.warning || '').trim();
  const sourceText = source === 'openai' ? `OpenAI (${escapeHtml(data.model || '')})` : 'Local fallback';
  aiState.smsDrafts = drafts.map((item) => ({
    language: String(item?.language || 'Message'),
    message: String(item?.message || '')
  }));

  if (!drafts.length) {
    elements.smsDraftMsg.textContent = warning || 'No drafts returned.';
    if (elements.smsDraftWrap) {
      elements.smsDraftWrap.hidden = true;
      elements.smsDraftWrap.innerHTML = '';
    }
    return;
  }

  const list = drafts
    .map((draft, index) => {
      const language = escapeHtml(String(draft.language || 'Message'));
      const message = escapeHtml(String(draft.message || ''));
      return `
        <li>
          <strong>${language}</strong>: ${message}
          <button class="table-btn" type="button" data-apply-sms-draft="${index}">Use This Draft</button>
        </li>
      `;
    })
    .join('');

  elements.smsDraftMsg.textContent = warning || `Drafts generated via ${sourceText}.`;
  if (!elements.smsDraftWrap) return;
  elements.smsDraftWrap.hidden = false;
  elements.smsDraftWrap.innerHTML = `
    <div class="ai-result-title">SMS Draft Suggestions</div>
    <ul class="ai-result-list">${list}</ul>
  `;
}

function renderCopilotResult(data) {
  const responseId = String(data?.responseId || '').trim();
  if (responseId) {
    aiState.lastResponseIds.copilot = responseId;
    syncAiFeedbackResponseIdFromLatest();
  }
  const answer = String(data?.answer || '').trim();
  const insights = Array.isArray(data?.insights) ? data.insights.map((item) => String(item || '').trim()).filter(Boolean) : [];
  const actions = Array.isArray(data?.actions) ? data.actions.map((item) => String(item || '').trim()).filter(Boolean) : [];
  const source = String(data?.source || 'local-rules');
  const warning = String(data?.warning || '').trim();
  const sourceText = source === 'openai' ? `OpenAI (${escapeHtml(data.model || '')})` : 'Local fallback';

  const insightsHtml = insights.length
    ? `<div class="ai-result-title">Insights</div><ul class="ai-result-list">${insights.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
    : '';
  const actionsHtml = actions.length
    ? `<div class="ai-result-title">Suggested Actions</div><ul class="ai-result-list">${actions.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
    : '';

  elements.copilotMsg.textContent = warning || `Copilot response generated via ${sourceText}.`;
  if (!elements.copilotWrap) return;
  elements.copilotWrap.classList.remove('empty');
  elements.copilotWrap.innerHTML = `
    <div class="ai-result-title">Answer</div>
    <p class="meta compact">${escapeHtml(answer || 'No answer returned.')}</p>
    ${insightsHtml}
    ${actionsHtml}
  `;
}

function renderQcAiResult(data) {
  const responseId = String(data?.responseId || '').trim();
  if (responseId) {
    aiState.lastResponseIds['qc-intelligence'] = responseId;
    syncAiFeedbackResponseIdFromLatest();
  }
  const summary = data?.summary || {};
  const items = Array.isArray(data?.items) ? data.items : [];
  const warning = String(data?.warning || '').trim();
  const source = String(data?.source || 'local-rules');
  const sourceText = source === 'openai' ? `OpenAI (${escapeHtml(data.model || '')})` : 'Local fallback';

  aiState.qcInsights = items;
  if (!items.length) {
    elements.qcAiMsg.textContent = warning || 'No QC intelligence results returned.';
    if (elements.qcAiWrap) {
      elements.qcAiWrap.classList.add('empty');
      elements.qcAiWrap.textContent = 'No QC intelligence output returned for this request.';
    }
    return;
  }

  const rows = items
    .slice(0, 30)
    .map((item) => {
      const reasons = Array.isArray(item?.reasons) ? item.reasons : [];
      const actions = Array.isArray(item?.actions) ? item.actions : [];
      return `
        <li>
          <strong>${escapeHtml(String(item.farmerName || '-'))}</strong> | Lot ${escapeHtml(String(item.qcRecordId || '-'))}
          <br>Risk: <strong>${escapeHtml(String(item.riskLevel || '-').toUpperCase())}</strong>
          | Score: ${escapeHtml(String(item.riskScore ?? '-'))}
          | Current: ${escapeHtml(String(item.currentDecision || '-'))}
          | Suggested: <strong>${escapeHtml(String(item.recommendedDecision || '-'))}</strong>
          <br>${escapeHtml(String(item.summary || ''))}
          ${reasons.length ? `<br><em>Reasons:</em> ${escapeHtml(reasons.join(' | '))}` : ''}
          ${actions.length ? `<br><em>Actions:</em> ${escapeHtml(actions.join(' | '))}` : ''}
        </li>
      `;
    })
    .join('');

  elements.qcAiMsg.textContent =
    warning ||
    `QC intelligence generated via ${sourceText}. High risk: ${summary.highRisk || 0}, medium: ${summary.mediumRisk || 0}, low: ${summary.lowRisk || 0}.`;

  if (!elements.qcAiWrap) return;
  elements.qcAiWrap.classList.remove('empty');
  elements.qcAiWrap.innerHTML = `
    <div class="ai-result-title">QC Intelligence Results</div>
    <p class="meta compact">
      Analyzed: ${escapeHtml(String(summary.totalAnalyzed || items.length))},
      Flagged: ${escapeHtml(String(summary.flagged || 0))}
    </p>
    <ul class="ai-result-list">${rows}</ul>
  `;
}

function renderPaymentRiskResult(data) {
  const responseId = String(data?.responseId || '').trim();
  if (responseId) {
    aiState.lastResponseIds['payment-risk'] = responseId;
    syncAiFeedbackResponseIdFromLatest();
  }
  const summary = data?.summary || {};
  const flags = Array.isArray(data?.flags) ? data.flags : [];
  const actions = Array.isArray(data?.actions) ? data.actions : [];
  const warning = String(data?.warning || '').trim();
  const source = String(data?.source || 'local-rules');
  const sourceText = source === 'openai' ? `OpenAI (${escapeHtml(data.model || '')})` : 'Local fallback';
  const narrative = String(data?.narrative || '').trim();

  aiState.paymentRiskFlags = flags;

  const actionsHtml = actions.length
    ? `<div class="ai-result-title">Suggested Actions</div><ul class="ai-result-list">${actions.map((row) => `<li>${escapeHtml(String(row))}</li>`).join('')}</ul>`
    : '';
  const flagsHtml = flags.length
    ? `<div class="ai-result-title">Risk Flags</div><ul class="ai-result-list">${flags
      .slice(0, 40)
      .map(
        (flag) => `<li><strong>${escapeHtml(String(flag.severity || 'medium').toUpperCase())}</strong> - ${escapeHtml(
          String(flag.title || flag.code || 'Flag')
        )}: ${escapeHtml(String(flag.detail || ''))}</li>`
      )
      .join('')}</ul>`
    : '<div class="empty">No risk flags detected for the selected period.</div>';

  elements.paymentRiskMsg.textContent =
    warning ||
    `Payment risk report generated via ${sourceText}. Overall risk: ${String(summary.overallRisk || 'low').toUpperCase()}.`;

  if (!elements.paymentRiskWrap) return;
  elements.paymentRiskWrap.classList.remove('empty');
  elements.paymentRiskWrap.innerHTML = `
    <div class="ai-result-title">Payment Risk Summary</div>
    <p class="meta compact">
      Payments analyzed: ${escapeHtml(String(data.paymentCount || 0))} |
      Overall risk: <strong>${escapeHtml(String(summary.overallRisk || 'low').toUpperCase())}</strong> |
      High: ${escapeHtml(String(summary.highFlags || 0))},
      Medium: ${escapeHtml(String(summary.mediumFlags || 0))},
      Low: ${escapeHtml(String(summary.lowFlags || 0))}
    </p>
    <p class="meta compact">${escapeHtml(narrative || 'No narrative returned.')}</p>
    ${actionsHtml}
    ${flagsHtml}
  `;
}

function syncPaymentRiskFilterUi() {
  if (!elements.paymentRiskPeriod || !elements.paymentRiskFrom || !elements.paymentRiskTo) return;
  const period = String(elements.paymentRiskPeriod.value || 'week').trim().toLowerCase();
  const custom = period === 'custom';
  elements.paymentRiskFrom.hidden = !custom;
  elements.paymentRiskTo.hidden = !custom;
  elements.paymentRiskFrom.disabled = !custom;
  elements.paymentRiskTo.disabled = !custom;
}

function selectedAiFeedbackTool() {
  return String(elements.aiFeedbackTool?.value || '').trim().toLowerCase();
}

function syncAiFeedbackResponseIdFromLatest(force = false) {
  if (!elements.aiFeedbackResponseId || !elements.aiFeedbackTool) return;
  const tool = selectedAiFeedbackTool();
  const latest = String(aiState.lastResponseIds?.[tool] || '').trim();
  if (!latest) return;
  if (!force && String(elements.aiFeedbackResponseId.value || '').trim()) return;
  elements.aiFeedbackResponseId.value = latest;
}

function normalizeProposalStatusClient(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (['pending', 'approved', 'rejected', 'executed'].includes(normalized)) return normalized;
  return 'pending';
}

function normalizeOpsTaskStatusClient(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (['open', 'in_progress', 'resolved', 'closed'].includes(normalized)) return normalized;
  return 'open';
}

function normalizeSeverityClient(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (['critical', 'high', 'medium', 'low'].includes(normalized)) return normalized;
  return 'low';
}

function textOrDash(value) {
  const text = String(value || '').trim();
  return text || '-';
}

async function fetchAiProposals() {
  if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
    aiState.proposals = [];
    aiState.proposalsMeta = null;
    return [];
  }

  const params = new URLSearchParams({ limit: '100' });
  const status = String(elements.proposalStatusFilter?.value || '').trim().toLowerCase();
  if (status) params.set('status', status);
  const response = await apiRequest(`/api/ai/proposals?${params.toString()}`);
  aiState.proposals = Array.isArray(response.data) ? response.data : [];
  aiState.proposalsMeta = response.meta || null;
  return aiState.proposals;
}

async function fetchAiBriefs() {
  if (!API.enabled || !isAuthenticated() || !['admin', 'agent'].includes(currentRole())) {
    aiState.briefs = [];
    aiState.alerts = [];
    return [];
  }

  const response = await apiRequest('/api/ai/briefs?limit=20&includeAlerts=true');
  aiState.briefs = Array.isArray(response.data?.briefs) ? response.data.briefs : [];
  aiState.alerts = Array.isArray(response.data?.alerts) ? response.data.alerts : [];
  if (aiState.briefs.length) {
    aiState.lastResponseIds['executive-brief'] = String(aiState.briefs[0].id || '').trim();
    syncAiFeedbackResponseIdFromLatest();
  }
  return aiState.briefs;
}

async function fetchAiOpsTasks() {
  if (!API.enabled || !isAuthenticated() || !['admin', 'agent'].includes(currentRole())) {
    aiState.opsTasks = [];
    aiState.opsTasksMeta = null;
    return [];
  }

  const params = new URLSearchParams({ limit: '200' });
  const status = String(elements.opsTaskStatusFilter?.value || '').trim().toLowerCase();
  const severity = String(elements.opsTaskSeverityFilter?.value || '').trim().toLowerCase();
  if (status) params.set('status', status);
  if (severity) params.set('severity', severity);
  const response = await apiRequest(`/api/ops/tasks?${params.toString()}`);
  aiState.opsTasks = Array.isArray(response.data) ? response.data : [];
  aiState.opsTasksMeta = response.meta || null;
  return aiState.opsTasks;
}

async function fetchAiKnowledgeDocs() {
  if (!API.enabled || !isAuthenticated() || !['admin', 'agent'].includes(currentRole())) {
    aiState.knowledgeDocs = [];
    return [];
  }

  const params = new URLSearchParams({ limit: '80' });
  const q = String(elements.knowledgeSearch?.value || '').trim();
  if (q) params.set('q', q);
  const response = await apiRequest(`/api/ai/knowledge?${params.toString()}`);
  aiState.knowledgeDocs = Array.isArray(response.data) ? response.data : [];
  return aiState.knowledgeDocs;
}

async function fetchAiFeedbackSummary() {
  if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
    aiState.feedbackSummary = null;
    return null;
  }
  const response = await apiRequest('/api/ai/feedback/summary?limit=30');
  aiState.feedbackSummary = response.data || null;
  return aiState.feedbackSummary;
}

async function fetchAiEvalRuns() {
  if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
    aiState.evalRuns = [];
    return [];
  }
  const response = await apiRequest('/api/ai/evals?limit=20');
  aiState.evalRuns = Array.isArray(response.data) ? response.data : [];
  return aiState.evalRuns;
}

function renderExecutiveBriefs() {
  if (!elements.briefWrap) return;

  if (!API.enabled) {
    elements.briefWrap.classList.add('empty');
    elements.briefWrap.textContent = 'Executive briefs are available in backend mode only.';
    return;
  }
  if (!isAuthenticated()) {
    elements.briefWrap.classList.add('empty');
    elements.briefWrap.textContent = 'Sign in to view executive briefs.';
    return;
  }
  if (!['admin', 'agent'].includes(currentRole())) {
    elements.briefWrap.classList.add('empty');
    elements.briefWrap.textContent = 'Executive briefs are available for admin and agent roles.';
    return;
  }

  const briefs = [...(Array.isArray(aiState.briefs) ? aiState.briefs : [])]
    .sort((a, b) => new Date(b.generatedAt || b.createdAt || 0).getTime() - new Date(a.generatedAt || a.createdAt || 0).getTime())
    .slice(0, 20);
  const alerts = [...(Array.isArray(aiState.alerts) ? aiState.alerts : [])]
    .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime())
    .slice(0, 20);

  if (!briefs.length && !alerts.length) {
    elements.briefWrap.classList.add('empty');
    elements.briefWrap.textContent = 'No executive briefs yet. Use "Run Brief Now" to generate one.';
    return;
  }

  const briefCards = briefs
    .map((brief) => {
      const risk = normalizeSeverityClient(brief?.overallRisk);
      const findings = Array.isArray(brief?.keyFindings) ? brief.keyFindings : [];
      const actions = Array.isArray(brief?.actions) ? brief.actions : [];
      return `
        <article class="brief-card">
          <div class="inline-actions wrap">
            <strong>${escapeHtml(textOrDash(brief.title || `Brief ${brief.id}`))}</strong>
            <span class="severity-badge ${escapeHtml(risk)}">${escapeHtml(risk)}</span>
          </div>
          <p class="meta compact">
            ID: ${escapeHtml(textOrDash(brief.id))} | Trigger: ${escapeHtml(textOrDash(brief.trigger))} |
            Time: ${escapeHtml(dateShort(brief.generatedAt || brief.createdAt))}
          </p>
          <p class="meta compact">${escapeHtml(textOrDash(brief.summary || brief.narrative))}</p>
          ${
            findings.length
              ? `<div class="ai-result-title">Key Findings</div><ul class="ai-result-list">${findings
                .slice(0, 4)
                .map((row) => `<li>${escapeHtml(String(row))}</li>`)
                .join('')}</ul>`
              : ''
          }
          ${
            actions.length
              ? `<div class="ai-result-title">Recommended Actions</div><ul class="ai-result-list">${actions
                .slice(0, 4)
                .map((row) => `<li>${escapeHtml(String(row))}</li>`)
                .join('')}</ul>`
              : ''
          }
          ${brief.warning ? `<p class="dim">${escapeHtml(String(brief.warning))}</p>` : ''}
        </article>
      `;
    })
    .join('');

  const alertCards = alerts
    .map((alert) => {
      const severity = normalizeSeverityClient(alert?.severity);
      return `
        <article class="brief-card">
          <div class="inline-actions wrap">
            <strong>${escapeHtml(textOrDash(alert.title || 'Alert'))}</strong>
            <span class="severity-badge ${escapeHtml(severity)}">${escapeHtml(severity)}</span>
          </div>
          <p class="meta compact">${escapeHtml(textOrDash(alert.message || alert.summary))}</p>
          <p class="dim">${escapeHtml(dateShort(alert.createdAt))}</p>
        </article>
      `;
    })
    .join('');

  elements.briefWrap.classList.remove('empty');
  elements.briefWrap.innerHTML = `
    <p class="meta compact">Briefs: ${escapeHtml(String(briefs.length))} | Alerts: ${escapeHtml(String(alerts.length))}</p>
    ${briefCards}
    ${alertCards ? `<div class="ai-result-title">Alerts</div>${alertCards}` : ''}
  `;
}

function renderProposalQueue() {
  if (!elements.proposalWrap) return;

  if (!API.enabled) {
    elements.proposalWrap.classList.add('empty');
    elements.proposalWrap.textContent = 'Proposal queue requires backend mode.';
    return;
  }
  if (!isAuthenticated()) {
    elements.proposalWrap.classList.add('empty');
    elements.proposalWrap.textContent = 'Sign in to access the proposal queue.';
    return;
  }
  if (currentRole() !== 'admin') {
    elements.proposalWrap.classList.add('empty');
    elements.proposalWrap.textContent = 'Proposal queue is visible to admin only.';
    return;
  }

  const filter = String(elements.proposalStatusFilter?.value || '').trim().toLowerCase();
  const rows = (Array.isArray(aiState.proposals) ? aiState.proposals : [])
    .filter((row) => !filter || normalizeProposalStatusClient(row.status) === filter)
    .sort((a, b) => new Date(b.updatedAt || b.createdAt || 0).getTime() - new Date(a.updatedAt || a.createdAt || 0).getTime());

  if (!rows.length) {
    elements.proposalWrap.classList.add('empty');
    elements.proposalWrap.textContent = filter ? `No proposals with status "${filter}".` : 'No proposals available.';
    return;
  }

  const counts = aiState.proposalsMeta?.counts || {};
  const summaryLine = ['pending', 'approved', 'rejected', 'executed']
    .map((status) => `${status}: ${Number(counts?.[status] || 0)}`)
    .join(' | ');

  const html = rows
    .map((proposal) => {
      const status = normalizeProposalStatusClient(proposal.status);
      const canApprove = status === 'pending';
      const canReject = status === 'pending' || status === 'approved';
      const canExecute = status === 'approved';
      const payloadText = proposal?.payload && Object.keys(proposal.payload).length
        ? JSON.stringify(proposal.payload, null, 2)
        : '';

      return `
        <article class="proposal-card">
          <div class="inline-actions wrap">
            <strong>${escapeHtml(textOrDash(proposal.title || proposal.actionType))}</strong>
            <span class="severity-badge ${status === 'rejected' ? 'high' : status === 'executed' ? 'low' : 'medium'}">${escapeHtml(status)}</span>
          </div>
          <p class="meta compact">
            ID: ${escapeHtml(textOrDash(proposal.id))} |
            Action: ${escapeHtml(textOrDash(proposal.actionType))} |
            Confidence: ${escapeHtml(proposal.confidence === '' || proposal.confidence === undefined ? '-' : String(proposal.confidence))}
          </p>
          <p class="meta compact">${escapeHtml(textOrDash(proposal.description))}</p>
          ${payloadText ? `<pre class="meta compact">${escapeHtml(payloadText)}</pre>` : ''}
          <p class="dim">
            Created by ${escapeHtml(textOrDash(proposal.createdBy))} on ${escapeHtml(dateShort(proposal.createdAt))}
            | Updated ${escapeHtml(dateShort(proposal.updatedAt || proposal.createdAt))}
          </p>
          ${
            proposal.executionSummary
              ? `<p class="meta compact"><strong>Execution:</strong> ${escapeHtml(String(proposal.executionSummary))}</p>`
              : ''
          }
          ${
            proposal.rejectionReason
              ? `<p class="meta compact"><strong>Rejection reason:</strong> ${escapeHtml(String(proposal.rejectionReason))}</p>`
              : ''
          }
          <div class="proposal-actions">
            ${canApprove ? `<button class="table-btn" type="button" data-proposal-id="${escapeHtml(proposal.id)}" data-proposal-action="approve">Approve</button>` : ''}
            ${canReject ? `<button class="table-btn" type="button" data-proposal-id="${escapeHtml(proposal.id)}" data-proposal-action="reject">Reject</button>` : ''}
            ${canExecute ? `<button class="table-btn" type="button" data-proposal-id="${escapeHtml(proposal.id)}" data-proposal-action="execute">Execute</button>` : ''}
          </div>
        </article>
      `;
    })
    .join('');

  elements.proposalWrap.classList.remove('empty');
  elements.proposalWrap.innerHTML = `
    <p class="meta compact">Queue counts: ${escapeHtml(summaryLine)}</p>
    ${html}
  `;
}

function renderOpsTasks() {
  if (!elements.opsTaskWrap) return;

  if (!API.enabled) {
    elements.opsTaskWrap.classList.add('empty');
    elements.opsTaskWrap.textContent = 'Operations tasks require backend mode.';
    return;
  }
  if (!isAuthenticated()) {
    elements.opsTaskWrap.classList.add('empty');
    elements.opsTaskWrap.textContent = 'Sign in to view operations tasks.';
    return;
  }
  if (!['admin', 'agent'].includes(currentRole())) {
    elements.opsTaskWrap.classList.add('empty');
    elements.opsTaskWrap.textContent = 'Operations tasks are available for admin and agent roles.';
    return;
  }

  const statusFilter = String(elements.opsTaskStatusFilter?.value || '').trim().toLowerCase();
  const severityFilter = String(elements.opsTaskSeverityFilter?.value || '').trim().toLowerCase();
  const rows = (Array.isArray(aiState.opsTasks) ? aiState.opsTasks : [])
    .filter((task) => {
      const statusOk = !statusFilter || normalizeOpsTaskStatusClient(task.status) === statusFilter;
      const severityOk = !severityFilter || normalizeSeverityClient(task.severity) === severityFilter;
      return statusOk && severityOk;
    })
    .sort((a, b) => new Date(b.updatedAt || b.createdAt || 0).getTime() - new Date(a.updatedAt || a.createdAt || 0).getTime());

  if (!rows.length) {
    elements.opsTaskWrap.classList.add('empty');
    elements.opsTaskWrap.textContent = 'No operations tasks for the selected filters.';
    return;
  }

  const canEdit = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(currentRole());
  const statusOptions = ['open', 'in_progress', 'resolved', 'closed'];
  const counts = aiState.opsTasksMeta?.counts || {};
  const countLine = statusOptions.map((status) => `${status}: ${Number(counts[status] || 0)}`).join(' | ');

  const html = rows
    .map((task) => {
      const severity = normalizeSeverityClient(task.severity);
      const status = normalizeOpsTaskStatusClient(task.status);
      const statusSelect = statusOptions
        .map((option) => `<option value="${option}" ${status === option ? 'selected' : ''}>${escapeHtml(option.replace('_', ' '))}</option>`)
        .join('');
      return `
        <article class="task-card" data-task-id="${escapeHtml(task.id)}">
          <div class="inline-actions wrap">
            <strong>${escapeHtml(textOrDash(task.title))}</strong>
            <span class="severity-badge ${escapeHtml(severity)}">${escapeHtml(severity)}</span>
          </div>
          <p class="meta compact">${escapeHtml(textOrDash(task.description))}</p>
          <p class="meta compact">
            Source: ${escapeHtml(textOrDash(task.sourceType))} (${escapeHtml(textOrDash(task.sourceRef))}) |
            Created by ${escapeHtml(textOrDash(task.createdBy))}
          </p>
          <div class="row">
            <select data-task-field="status" ${canEdit ? '' : 'disabled'}>${statusSelect}</select>
            <input data-task-field="assignedTo" placeholder="Assigned to" value="${escapeHtml(String(task.assignedTo || ''))}" ${canEdit ? '' : 'disabled'}>
          </div>
          <textarea data-task-field="notes" rows="2" placeholder="Notes" ${canEdit ? '' : 'disabled'}>${escapeHtml(String(task.notes || ''))}</textarea>
          <div class="task-actions">
            <button class="table-btn" type="button" data-task-save="${escapeHtml(task.id)}" ${canEdit ? '' : 'disabled'}>Save Task</button>
            <span class="dim">Updated: ${escapeHtml(dateShort(task.updatedAt || task.createdAt))}</span>
          </div>
        </article>
      `;
    })
    .join('');

  elements.opsTaskWrap.classList.remove('empty');
  elements.opsTaskWrap.innerHTML = `
    <p class="meta compact">Task counts: ${escapeHtml(countLine)}</p>
    ${html}
  `;
}

function renderKnowledgeDocs() {
  if (!elements.knowledgeWrap) return;

  if (!API.enabled) {
    elements.knowledgeWrap.classList.add('empty');
    elements.knowledgeWrap.textContent = 'Knowledge base requires backend mode.';
    return;
  }
  if (!isAuthenticated()) {
    elements.knowledgeWrap.classList.add('empty');
    elements.knowledgeWrap.textContent = 'Sign in to view the knowledge base.';
    return;
  }
  if (!['admin', 'agent'].includes(currentRole())) {
    elements.knowledgeWrap.classList.add('empty');
    elements.knowledgeWrap.textContent = 'Knowledge base is available for admin and agent roles.';
    return;
  }

  const query = String(elements.knowledgeSearch?.value || '').trim().toLowerCase();
  const tokens = query ? query.split(/\s+/).filter(Boolean) : [];
  const docs = (Array.isArray(aiState.knowledgeDocs) ? aiState.knowledgeDocs : [])
    .filter((doc) => {
      if (!tokens.length) return true;
      const haystack = [
        doc?.title,
        doc?.source,
        Array.isArray(doc?.tags) ? doc.tags.join(' ') : '',
        doc?.content,
        doc?.snippet
      ]
        .map((row) => String(row || '').toLowerCase())
        .join(' ');
      return tokens.every((token) => haystack.includes(token));
    })
    .sort((a, b) => new Date(b.updatedAt || b.createdAt || 0).getTime() - new Date(a.updatedAt || a.createdAt || 0).getTime());

  if (!docs.length) {
    elements.knowledgeWrap.classList.add('empty');
    elements.knowledgeWrap.textContent = query ? `No knowledge docs matched "${query}".` : 'No knowledge documents available.';
    return;
  }

  const canDelete = currentRole() === 'admin';
  const html = docs
    .map((doc) => {
      const tags = Array.isArray(doc?.tags) ? doc.tags : [];
      const snippet = String(doc.snippet || doc.content || '').trim();
      return `
        <article class="knowledge-card">
          <div class="inline-actions wrap">
            <strong>${escapeHtml(textOrDash(doc.title))}</strong>
            ${canDelete ? `<button class="table-btn danger" type="button" data-knowledge-delete="${escapeHtml(doc.id)}">Delete</button>` : ''}
          </div>
          <p class="meta compact">Source: ${escapeHtml(textOrDash(doc.source))} | Updated: ${escapeHtml(dateShort(doc.updatedAt || doc.createdAt))}</p>
          ${tags.length ? `<p class="meta compact">Tags: ${escapeHtml(tags.join(', '))}</p>` : ''}
          <p class="meta compact">${escapeHtml(snippet || 'No snippet available.')}</p>
        </article>
      `;
    })
    .join('');

  elements.knowledgeWrap.classList.remove('empty');
  elements.knowledgeWrap.innerHTML = html;
}

function renderAiFeedbackAndEvals() {
  if (!elements.aiFeedbackWrap) return;

  if (!API.enabled) {
    elements.aiFeedbackWrap.classList.add('empty');
    elements.aiFeedbackWrap.textContent = 'AI feedback and evals require backend mode.';
    return;
  }
  if (!isAuthenticated()) {
    elements.aiFeedbackWrap.classList.add('empty');
    elements.aiFeedbackWrap.textContent = 'Sign in to access AI feedback and evals.';
    return;
  }
  if (!['admin', 'agent'].includes(currentRole())) {
    elements.aiFeedbackWrap.classList.add('empty');
    elements.aiFeedbackWrap.textContent = 'AI feedback tools are available for admin and agent roles.';
    return;
  }

  const role = currentRole();
  const summary = aiState.feedbackSummary || null;
  const evalRuns = Array.isArray(aiState.evalRuns) ? aiState.evalRuns : [];
  const cards = [];

  if (role === 'admin' && summary) {
    const totals = summary.totals || {};
    cards.push(`
      <article class="feedback-card">
        <strong>Feedback Totals</strong>
        <p class="meta compact">
          Total: ${escapeHtml(String(totals.total || 0))} |
          Positive: ${escapeHtml(String(totals.positive || 0))} |
          Neutral: ${escapeHtml(String(totals.neutral || 0))} |
          Negative: ${escapeHtml(String(totals.negative || 0))} |
          Positive Rate: ${escapeHtml(String(totals.positiveRatePct || 0))}%
        </p>
      </article>
    `);

    const byTool = Array.isArray(summary.byTool) ? summary.byTool : [];
    if (byTool.length) {
      const toolRows = byTool
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(textOrDash(row.tool))}</td>
              <td>${escapeHtml(String(row.total || 0))}</td>
              <td>${escapeHtml(String(row.positive || 0))}</td>
              <td>${escapeHtml(String(row.neutral || 0))}</td>
              <td>${escapeHtml(String(row.negative || 0))}</td>
              <td>${escapeHtml(String(row.positiveRatePct || 0))}%</td>
            </tr>
          `
        )
        .join('');
      cards.push(`
        <article class="feedback-card">
          <strong>Feedback by Tool</strong>
          <table>
            <thead>
              <tr>
                <th>Tool</th>
                <th>Total</th>
                <th>Up</th>
                <th>Neutral</th>
                <th>Down</th>
                <th>Up %</th>
              </tr>
            </thead>
            <tbody>${toolRows}</tbody>
          </table>
        </article>
      `);
    }

    const recent = Array.isArray(summary.recent) ? summary.recent.slice(0, 10) : [];
    if (recent.length) {
      cards.push(`
        <article class="feedback-card">
          <strong>Recent Feedback</strong>
          <ul class="ai-result-list">
            ${recent
              .map((row) => `<li>${escapeHtml(dateShort(row.createdAt))} | ${escapeHtml(textOrDash(row.tool))} | ${escapeHtml(textOrDash(row.rating))} | ${escapeHtml(textOrDash(row.actor))}</li>`)
              .join('')}
          </ul>
        </article>
      `);
    }
  } else if (role === 'agent') {
    cards.push(`
      <article class="feedback-card">
        <strong>Feedback Saved</strong>
        <p class="meta compact">Agents can submit feedback. Summary and eval dashboards are visible to admin.</p>
      </article>
    `);
  } else if (role === 'admin') {
    cards.push(`
      <article class="feedback-card">
        <strong>Feedback Summary</strong>
        <p class="meta compact">No feedback summary loaded yet. Click "Refresh Feedback Summary".</p>
      </article>
    `);
  }

  if (role === 'admin') {
    if (evalRuns.length) {
      const evalHtml = evalRuns
        .slice(0, 10)
        .map((run) => {
          const failedChecks = Array.isArray(run?.checks) ? run.checks.filter((row) => !row.pass).slice(0, 3) : [];
          const failList = failedChecks.length
            ? `<ul class="ai-result-list">${failedChecks.map((row) => `<li>${escapeHtml(textOrDash(row.label || row.id))}: ${escapeHtml(textOrDash(row.details))}</li>`).join('')}</ul>`
            : '<p class="meta compact">No failed checks.</p>';
          return `
            <article class="feedback-card">
              <div class="inline-actions wrap">
                <strong>${escapeHtml(textOrDash(run.id))}</strong>
                <span class="severity-badge ${run.failCount > 0 ? 'high' : 'low'}">${escapeHtml(String(run.scorePct || 0))}%</span>
              </div>
              <p class="meta compact">
                ${escapeHtml(dateShort(run.createdAt))} |
                Checks: ${escapeHtml(String(run.totalChecks || 0))} |
                Pass: ${escapeHtml(String(run.passCount || 0))} |
                Fail: ${escapeHtml(String(run.failCount || 0))}
              </p>
              ${failList}
            </article>
          `;
        })
        .join('');
      cards.push(`<div class="ai-result-title">Eval Runs</div>${evalHtml}`);
    } else {
      cards.push(`
        <article class="feedback-card">
          <strong>Eval Runs</strong>
          <p class="meta compact">No eval runs yet. Click "Run AI Eval Suite".</p>
        </article>
      `);
    }
  }

  elements.aiFeedbackWrap.classList.remove('empty');
  elements.aiFeedbackWrap.innerHTML = cards.join('');
}

function currentSmsMode() {
  return String(elements.smsTargetMode.value || 'selected').trim().toLowerCase();
}

function updateSmsComposerUi() {
  let mode = currentSmsMode();

  if (currentRole() !== 'admin' && mode !== 'single') {
    mode = 'single';
    elements.smsTargetMode.value = 'single';
  }

  elements.smsSingleWrap.hidden = mode !== 'single';
  elements.smsRecipientPickerWrap.hidden = mode !== 'selected';
  elements.smsAllModeNotice.hidden = mode !== 'all';

  const submitBtn = elements.smsForm.querySelector('button[type="submit"]');
  if (mode === 'all') {
    submitBtn.textContent = 'Send SMS to All Mobile Numbers';
  } else if (mode === 'selected') {
    submitBtn.textContent = 'Send SMS to Selected Recipients';
  } else {
    submitBtn.textContent = 'Send SMS';
  }

  const allowBulk = currentRole() === 'admin';
  elements.smsTargetMode.disabled = !allowBulk;
  elements.smsSelectVisibleBtn.disabled = !allowBulk || mode !== 'selected' || smsPicker.rows.length === 0;
  elements.smsClearSelectedBtn.disabled = !allowBulk || smsPicker.selectedIds.size === 0;
  elements.smsPrevPageBtn.disabled = !allowBulk || mode !== 'selected' || smsPicker.offset === 0;
  elements.smsNextPageBtn.disabled = !allowBulk || mode !== 'selected' || smsPicker.offset + smsPicker.rows.length >= smsPicker.total;
  elements.smsRecipientSearchBtn.disabled = !allowBulk || mode !== 'selected';
  elements.smsRecipientSearch.disabled = !allowBulk || mode !== 'selected';
}

async function loadSmsRecipients(reset = false) {
  if (currentSmsMode() !== 'selected') {
    renderSmsRecipientList();
    return;
  }

  if (reset) {
    smsPicker.offset = 0;
  }

  const q = elements.smsRecipientSearch.value.trim();
  smsPicker.q = q;

  try {
    if (API.enabled) {
      if (!isAuthenticated()) {
        smsPicker.rows = [];
        smsPicker.total = 0;
        renderSmsRecipientList();
        return;
      }

      const params = new URLSearchParams();
      if (q) params.set('q', q);
      params.set('limit', String(smsPicker.limit));
      params.set('offset', String(smsPicker.offset));

      const response = await apiRequest(`/api/sms/recipients?${params.toString()}`);
      smsPicker.rows = Array.isArray(response.data) ? response.data : [];
      smsPicker.total = Number(response.meta?.total || smsPicker.rows.length);

      if (smsPicker.offset > 0 && smsPicker.rows.length === 0) {
        smsPicker.offset = Math.max(0, smsPicker.offset - smsPicker.limit);
        return loadSmsRecipients(false);
      }

      renderSmsRecipientList();
      return;
    }

    let rows = state.farmers.filter((row) => String(row.phone || '').trim());
    if (q) {
      const needle = q.toLowerCase();
      rows = rows.filter((row) => {
        const text =
          `${row.id || ''} ${row.name || ''} ${row.phone || ''} ${row.nationalId || ''} ` +
          `${row.location || ''} ${row.notes || ''} ${row.preferredLanguage || ''}`.toLowerCase();
        return text.includes(needle);
      });
    }

    smsPicker.total = rows.length;
    smsPicker.rows = rows.slice(smsPicker.offset, smsPicker.offset + smsPicker.limit).map((row) => ({
      id: row.id,
      name: row.name,
      phone: row.phone,
      nationalId: row.nationalId,
      location: row.location,
      preferredLanguage: languageOrDefaultClient(row.preferredLanguage)
    }));

    if (smsPicker.offset > 0 && smsPicker.rows.length === 0) {
      smsPicker.offset = Math.max(0, smsPicker.offset - smsPicker.limit);
      return loadSmsRecipients(false);
    }

    renderSmsRecipientList();
  } catch (error) {
    smsPicker.rows = [];
    smsPicker.total = 0;
    renderSmsRecipientList();
    elements.smsMsg.textContent = error.message;
  }
}

function renderSmsRecipientList() {
  if (currentSmsMode() !== 'selected') {
    elements.smsRecipientListWrap.innerHTML = '';
    elements.smsSelectionInfo.textContent = '';
    updateSmsComposerUi();
    return;
  }

  const selectedCount = smsPicker.selectedIds.size;
  const start = smsPicker.total ? smsPicker.offset + 1 : 0;
  const end = smsPicker.offset + smsPicker.rows.length;

  elements.smsSelectionInfo.textContent = `${selectedCount} selected. Showing ${start}-${end} of ${smsPicker.total}${
    smsPicker.q ? ` (query: "${smsPicker.q}")` : ''
  }.`;

  if (!smsPicker.rows.length) {
    elements.smsRecipientListWrap.innerHTML = '<div class="empty">No recipients found for this search.</div>';
    updateSmsComposerUi();
    return;
  }

  const rows = smsPicker.rows
    .map(
      (row) => `
        <tr>
          <td>
            <input type="checkbox" data-farmer-id="${escapeHtml(row.id)}" ${smsPicker.selectedIds.has(row.id) ? 'checked' : ''}>
          </td>
          <td>${escapeHtml(row.name || '-')}</td>
          <td>${escapeHtml(row.phone || '-')}</td>
          <td>${escapeHtml(row.nationalId || '-')}</td>
          <td>${escapeHtml(row.location || '-')}</td>
          <td>${escapeHtml(languageOrDefaultClient(row.preferredLanguage) === 'sw' ? 'Kiswahili' : 'English')}</td>
        </tr>
      `
    )
    .join('');

  elements.smsRecipientListWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Select</th>
          <th>Name</th>
          <th>Phone</th>
          <th>National ID</th>
          <th>Location</th>
          <th>Language</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  updateSmsComposerUi();
}

function bindExports() {
  if (elements.exportRange) {
    const today = formatDateInputValue(new Date());
    const monthStart = new Date();
    monthStart.setDate(1);

    if (elements.exportFrom && !elements.exportFrom.value) {
      elements.exportFrom.value = formatDateInputValue(monthStart);
    }
    if (elements.exportTo && !elements.exportTo.value) {
      elements.exportTo.value = today;
    }

    elements.exportRange.addEventListener('change', () => {
      syncExportFilterUi();
    });

    syncExportFilterUi();
  }

  elements.exportButtons.forEach((button) => {
    button.addEventListener('click', async () => {
      clearMessages();

      const type = button.dataset.export;
      if (!type) return;

      if (!API.enabled) {
        elements.exportsMsg.textContent = 'Exports require backend mode.';
        return;
      }
      if (!isAuthenticated()) {
        elements.exportsMsg.textContent = 'Sign in first to export data.';
        return;
      }

      try {
        const params = buildExportQueryParams();
        const period = String(params.get('period') || 'all').toLowerCase();
        const format = String(elements.exportFormat?.value || 'csv').toLowerCase();
        const rangeLabel = exportRangeLabel(period, params.get('from'), params.get('to'));

        let endpoint = `/api/exports/${encodeURIComponent(type)}.csv`;
        const query = params.toString();
        if (query) endpoint = `${endpoint}?${query}`;

        const content = await apiRequest(endpoint, {
          method: 'GET',
          response: 'text'
        });

        const formatLabel = await exportDatasetFile({
          type,
          csvText: content,
          format,
          rangeLabel,
          fileSuffix: period
        });

        elements.exportsMsg.textContent = `Exported ${type} (${formatLabel}).`;
      } catch (error) {
        elements.exportsMsg.textContent = error.message;
      }
    });
  });

  elements.backupBtn.addEventListener('click', async () => {
    clearMessages();

    if (!API.enabled) {
      elements.exportsMsg.textContent = 'Backup requires backend mode.';
      return;
    }

    try {
      const response = await apiRequest('/api/admin/backup', {
        method: 'POST',
        body: {}
      });
      elements.exportsMsg.textContent = `Backup created: ${response.data.filename}`;
      await listBackups();
    } catch (error) {
      elements.exportsMsg.textContent = error.message;
    }
  });

  elements.listBackupsBtn.addEventListener('click', async () => {
    clearMessages();
    await listBackups();
  });

  elements.backupListWrap.addEventListener('click', async (event) => {
    const restoreBtn = event.target.closest('[data-restore-backup]');
    if (!restoreBtn) return;
    if (!API.enabled || !isAuthenticated() || currentRole() !== 'admin') {
      elements.exportsMsg.textContent = 'Sign in as admin to restore backups.';
      return;
    }

    const encoded = restoreBtn.getAttribute('data-restore-backup') || '';
    const filename = decodeURIComponent(encoded);
    if (!filename) return;

    const allow = confirm(
      `Restore backup ${filename}? This replaces current farmers, QC, purchases, payments, and SMS logs.`
    );
    if (!allow) return;

    restoreBtn.disabled = true;
    try {
      const response = await apiRequest('/api/admin/restore', {
        method: 'POST',
        body: { filename, confirm: true }
      });
      elements.exportsMsg.textContent = `Restore completed: ${response.data.fromBackup}`;
      await fetchAllData();
      await listBackups();
    } catch (error) {
      elements.exportsMsg.textContent = error.message;
    } finally {
      restoreBtn.disabled = false;
    }
  });

  elements.resetBtn.addEventListener('click', async () => {
    clearMessages();

    if (!confirm('This will clear all farmers, QC records, produce purchases, payments, and SMS data. Continue?')) return;

    if (!API.enabled) {
      state.farmers = [];
      state.produce = [];
      state.producePurchases = [];
      state.payments = [];
      state.smsLogs = [];
      state.summary = null;
      state.agentStats = [];
      state.backups = [];
      resetFarmerForm();
      hydrateFarmerSelectors();
      renderAll();
      await refreshOwedRows(true);
      persist();
      elements.exportsMsg.textContent = 'Local dataset reset.';
      return;
    }

    try {
      const response = await apiRequest('/api/admin/reset', {
        method: 'POST',
        body: { confirm: true }
      });
      elements.exportsMsg.textContent = `Platform data reset. Backup: ${response.data.backup}`;
      resetFarmerForm();
      await fetchAllData();
      await listBackups();
    } catch (error) {
      elements.exportsMsg.textContent = error.message;
    }
  });

  elements.seedBtn.addEventListener('click', async () => {
    clearMessages();

    if (state.farmers.length || state.produce.length || state.producePurchases.length || state.payments.length) {
      notifySync('Seed skipped: data already exists. Use reset if you need a clean slate.');
      return;
    }

    if (API.enabled) {
      if (!isAuthenticated() || currentRole() !== 'admin') {
        notifySync('Sign in as admin to seed backend data.');
        return;
      }

      try {
        await apiRequest('/api/seed', { method: 'POST', body: {} });
        notifySync('Sample data loaded from backend.');
        await fetchAllData();
      } catch (error) {
        notifySync(error.message);
      }
      return;
    }

    seedLocalData();
    notifySync('Sample data loaded in local mode.');
  });
}

async function listBackups() {
  if (!API.enabled) {
    elements.backupListWrap.innerHTML = '<div class="empty">Backups are available in backend mode only.</div>';
    return;
  }

  try {
    const response = await apiRequest('/api/admin/backups');
    state.backups = response.data || [];
    persist();
    renderBackups();
  } catch (error) {
    elements.exportsMsg.textContent = error.message;
  }
}

async function checkApiConnectivity() {
  try {
    await apiRequest('/api/health', { auth: false });
    API.enabled = true;
  } catch {
    API.enabled = false;
  }
}

async function restoreSession() {
  try {
    const response = await apiRequest('/api/auth/me');
    state.auth.user = response.data.user;
    state.auth.expiresAt = response.data.expiresAt;
    state.role = response.data.user.role;
    syncCurrentUserPhotoFromState();
    persist();
    updateAuthUi();
    return true;
  } catch {
    state.auth = { token: '', user: null, expiresAt: '' };
    persist();
    updateAuthUi();
    return false;
  }
}

async function fetchPhase2bSnapshots(role) {
  const fallback = {
    proposals: [],
    proposalsMeta: null,
    briefs: [],
    alerts: [],
    opsTasks: [],
    opsTasksMeta: null,
    knowledgeDocs: [],
    feedbackSummary: null,
    evalRuns: []
  };

  if (!API.enabled || !isAuthenticated()) return fallback;

  const calls = [];
  if (role === 'admin') {
    calls.push(['proposals', apiRequest('/api/ai/proposals?limit=100')]);
    calls.push(['briefs', apiRequest('/api/ai/briefs?limit=20&includeAlerts=true')]);
    calls.push(['opsTasks', apiRequest('/api/ops/tasks?limit=200')]);
    calls.push(['knowledgeDocs', apiRequest('/api/ai/knowledge?limit=80')]);
    calls.push(['feedbackSummary', apiRequest('/api/ai/feedback/summary?limit=30')]);
    calls.push(['evalRuns', apiRequest('/api/ai/evals?limit=20')]);
  } else if (role === 'agent') {
    calls.push(['briefs', apiRequest('/api/ai/briefs?limit=20&includeAlerts=true')]);
    calls.push(['opsTasks', apiRequest('/api/ops/tasks?limit=200')]);
    calls.push(['knowledgeDocs', apiRequest('/api/ai/knowledge?limit=80')]);
  } else {
    return fallback;
  }

  const settled = await Promise.allSettled(calls.map((entry) => entry[1]));
  settled.forEach((result, idx) => {
    const key = calls[idx][0];
    if (result.status !== 'fulfilled') return;
    const payload = result.value || {};
    if (key === 'proposals') {
      fallback.proposals = Array.isArray(payload.data) ? payload.data : [];
      fallback.proposalsMeta = payload.meta || null;
    } else if (key === 'briefs') {
      fallback.briefs = Array.isArray(payload.data?.briefs) ? payload.data.briefs : [];
      fallback.alerts = Array.isArray(payload.data?.alerts) ? payload.data.alerts : [];
    } else if (key === 'opsTasks') {
      fallback.opsTasks = Array.isArray(payload.data) ? payload.data : [];
      fallback.opsTasksMeta = payload.meta || null;
    } else if (key === 'knowledgeDocs') {
      fallback.knowledgeDocs = Array.isArray(payload.data) ? payload.data : [];
    } else if (key === 'feedbackSummary') {
      fallback.feedbackSummary = payload.data || null;
    } else if (key === 'evalRuns') {
      fallback.evalRuns = Array.isArray(payload.data) ? payload.data : [];
    }
  });

  return fallback;
}

async function fetchAllData() {
  if (!API.enabled || !isAuthenticated()) {
    renderAll();
    updatePermissionUi();
    await loadPaymentRecommendations();
    return;
  }

  try {
    const [farmers, produce, producePurchases, payments, sms, summary, agentStats] = await Promise.all([
      apiRequest('/api/farmers'),
      apiRequest('/api/produce'),
      apiRequest('/api/produce-purchases'),
      apiRequest('/api/payments'),
      apiRequest('/api/sms'),
      apiRequest('/api/reports/summary'),
      apiRequest('/api/reports/agents')
    ]);

    let paymentRecommendations = { data: [], meta: null };
    if (['admin', 'agent'].includes(currentRole())) {
      paymentRecommendations = await apiRequest('/api/payments/recommendations?status=all&limit=500');
    }

    let managedAgents = [];
    if (currentRole() === 'admin') {
      const agentAccounts = await apiRequest('/api/agents?includeDisabled=true&limit=500');
      managedAgents = Array.isArray(agentAccounts.data) ? agentAccounts.data : [];
    }

    const phase2b = await fetchPhase2bSnapshots(currentRole());

    state.farmers = farmers.data || [];
    state.produce = produce.data || [];
    state.producePurchases = producePurchases.data || [];
    state.payments = payments.data || [];
    state.paymentRecommendations = paymentRecommendations.data || [];
    state.smsLogs = sms.data || [];
    state.summary = summary.data || null;
    state.agentStats = agentStats.data || [];
    state.agents = managedAgents;
    paymentRecommendationPicker.meta = paymentRecommendations.meta || paymentRecommendationPicker.meta;

    aiState.proposals = phase2b.proposals || [];
    aiState.proposalsMeta = phase2b.proposalsMeta || null;
    aiState.briefs = phase2b.briefs || [];
    aiState.alerts = phase2b.alerts || [];
    aiState.opsTasks = phase2b.opsTasks || [];
    aiState.opsTasksMeta = phase2b.opsTasksMeta || null;
    aiState.knowledgeDocs = phase2b.knowledgeDocs || [];
    aiState.feedbackSummary = phase2b.feedbackSummary || null;
    aiState.evalRuns = phase2b.evalRuns || [];

    hydrateFarmerSelectors();
    renderAll();
    updatePermissionUi();
    persist();

    notifySync(`Connected to backend as ${currentRole()}.`);
    await refreshOwedRows(false);
    await loadPaymentRecommendations();
    if (currentSmsMode() === 'selected') {
      await loadSmsRecipients(false);
    }
  } catch (error) {
    notifySync(error.message);
  }
}

async function loadAgentAccounts() {
  if (currentRole() !== 'admin') {
    state.agents = [];
    renderAgents();
    return;
  }

  if (!API.enabled || !isAuthenticated()) {
    state.agents = [];
    renderAgents();
    return;
  }

  const params = new URLSearchParams({ includeDisabled: 'true', limit: '500' });
  const query = String(elements.agentSearch?.value || '').trim();
  if (query) params.set('q', query);

  const response = await apiRequest(`/api/agents?${params.toString()}`);
  state.agents = Array.isArray(response.data) ? response.data : [];
  renderAgents();
  persist();
}

async function apiRequest(path, options = {}) {
  const method = options.method || 'GET';
  const requireAuth = options.auth !== false;
  const responseType = options.response || 'json';

  const headers = {};
  if (options.body !== undefined) {
    headers['Content-Type'] = 'application/json';
  }
  if (requireAuth) {
    if (!state.auth.token) {
      throw new Error('Sign in first.');
    }
    headers.Authorization = `Bearer ${state.auth.token}`;
  }

  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), 6000);

  try {
    const response = await fetch(path, {
      method,
      headers,
      body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
      signal: controller.signal
    });

    if (responseType === 'text') {
      if (!response.ok) {
        let message = `Request failed (${response.status})`;
        try {
          const payload = await response.json();
          if (payload.error) message = payload.error;
        } catch {
          // no-op
        }
        throw new Error(message);
      }
      return response.text();
    }

    let payload = {};
    try {
      payload = await response.json();
    } catch {
      payload = {};
    }

    if (!response.ok) {
      if (response.status === 401 && requireAuth) {
        state.auth = { token: '', user: null, expiresAt: '' };
        persist();
        updateAuthUi();
        updatePermissionUi();
      }
      throw new Error(payload.error || `Request failed (${response.status})`);
    }

    return payload;
  } catch (error) {
    if (error.name === 'AbortError') {
      throw new Error('Request timed out.');
    }
    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }
}

function updateAuthUi() {
  if (isAuthenticated()) {
    elements.authShell.hidden = true;
    elements.appShell.hidden = false;
    document.body.classList.add('dashboard-open');
    if (elements.accountToggleBtn) {
      elements.accountToggleBtn.hidden = false;
    }
    elements.authState.textContent = `Signed in: ${state.auth.user.username} (${state.auth.user.role})`;
    elements.loginForm.hidden = false;
    elements.loginForm.style.display = 'grid';
    elements.registrationPanel.hidden = false;
    elements.changePasswordForm.hidden = false;
    elements.changePasswordForm.style.display = 'grid';
    elements.logoutBtn.hidden = false;
    elements.logoutBtn.style.display = 'block';
    elements.recoveryPanel.hidden = false;
    elements.roleSelect.disabled = true;
    elements.roleSelect.value = state.auth.user.role;
    state.role = state.auth.user.role;
    elements.registrationPanel.open = false;
    elements.recoveryPanel.open = false;
  } else {
    elements.authShell.hidden = false;
    elements.appShell.hidden = true;
    elements.scrollTopBtn.hidden = true;
    if (elements.dashboardMain) elements.dashboardMain.scrollTop = 0;
    document.body.classList.remove('dashboard-open');
    hideAccountPanel();
    if (elements.accountToggleBtn) {
      elements.accountToggleBtn.hidden = true;
    }
    if (elements.menuPanel) {
      elements.menuPanel.open = false;
    }
    elements.authState.textContent = 'Not signed in';
    elements.loginForm.hidden = false;
    elements.loginForm.style.display = 'grid';
    elements.registrationPanel.hidden = false;
    elements.changePasswordForm.hidden = true;
    elements.changePasswordForm.style.display = 'none';
    elements.logoutBtn.hidden = true;
    elements.logoutBtn.style.display = 'none';
    elements.recoveryPanel.hidden = false;
    elements.roleSelect.disabled = false;
    elements.roleSelect.value = state.role;
  }

  updateCredentialUiForRole();
  renderDashboardAccount();
  updateRoleHint();
  msaidiziState.role = isAuthenticated() ? currentRole() : 'guest';
  if (!isAuthenticated()) {
    activePaneId = 'auth';
  }
  void refreshMsaidiziContext();
}

function updatePermissionUi() {
  const role = currentRole();
  const backendAuthReady = !API.enabled || isAuthenticated();

  const farmersAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const produceAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const producePurchaseAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const owedAllowed = backendAuthReady && role === 'admin';
  const paymentRecommendationViewAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const paymentRecommendationCreateAllowed = backendAuthReady && role === 'agent';
  const paymentRecommendationDecisionAllowed = backendAuthReady && role === 'admin';
  const paymentAllowed = backendAuthReady && role === 'admin';
  const smsAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const aiSmsAllowed = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(role);
  const copilotAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const qcAiAllowed = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(role);
  const paymentRiskAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const executiveBriefViewAllowed = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(role);
  const executiveBriefRunAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const proposalViewAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const proposalManageAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const opsTaskViewAllowed = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(role);
  const opsTaskCreateFromAiAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const knowledgeViewAllowed = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(role);
  const knowledgeManageAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const feedbackAllowed = API.enabled && isAuthenticated() && ['admin', 'agent'].includes(role);
  const feedbackSummaryAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const evalAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const msaidiziAvailable = API.enabled;
  const msaidiziSyncAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const smartQaAllowed = role === 'admin';
  const adminAllowed = backendAuthReady && role === 'admin';
  const importAllowed = backendAuthReady && role === 'admin';
  const farmerPinManageAllowed = API.enabled && isAuthenticated() && role === 'admin';
  const agentManageAllowed = API.enabled && isAuthenticated() && role === 'admin';

  elements.farmerSubmitBtn.disabled = !farmersAllowed;
  elements.farmerCancelEditBtn.disabled = !farmersAllowed;
  elements.farmerImportBtn.disabled = !importAllowed;
  if (elements.farmerImportQaBtn) elements.farmerImportQaBtn.disabled = !smartQaAllowed;
  elements.farmerImportFile.disabled = !importAllowed;
  elements.farmerImportNotifyBySms.disabled = !importAllowed;
  if (elements.farmerImportCard) {
    elements.farmerImportCard.hidden = !importAllowed;
  }
  if (elements.farmerPinSearch) elements.farmerPinSearch.disabled = !farmerPinManageAllowed;
  if (elements.farmerPinSearchBtn) elements.farmerPinSearchBtn.disabled = !farmerPinManageAllowed;
  if (elements.farmerPinFarmer) elements.farmerPinFarmer.disabled = !farmerPinManageAllowed;
  if (elements.farmerPinValue) elements.farmerPinValue.disabled = !farmerPinManageAllowed;
  if (elements.farmerPinConfirm) elements.farmerPinConfirm.disabled = !farmerPinManageAllowed;
  if (elements.farmerPinSubmitBtn) elements.farmerPinSubmitBtn.disabled = !farmerPinManageAllowed;
  if (elements.farmerPinGenerateBtn) elements.farmerPinGenerateBtn.disabled = !farmerPinManageAllowed;
  updateFarmerImportSmsUi();
  elements.agentName.disabled = !agentManageAllowed;
  elements.agentEmail.disabled = !agentManageAllowed;
  elements.agentPhone.disabled = !agentManageAllowed;
  elements.agentRefreshBtn.disabled = !agentManageAllowed;
  elements.agentSearch.disabled = !agentManageAllowed;
  elements.agentForm.querySelector('button[type="submit"]').disabled = !agentManageAllowed;

  if (elements.agentsNavBtn) {
    elements.agentsNavBtn.hidden = !agentManageAllowed;
    elements.agentsNavBtn.disabled = !agentManageAllowed;
  }
  if (elements.agentsPaneOption) {
    const inSelect = elements.agentsPaneOption.parentElement === elements.paneSelect;
    if (!agentManageAllowed && inSelect) {
      elements.paneSelect.removeChild(elements.agentsPaneOption);
    } else if (agentManageAllowed && !inSelect) {
      const produceOption = elements.paneSelect?.querySelector('option[value="produce"]');
      if (produceOption) {
        elements.paneSelect.insertBefore(elements.agentsPaneOption, produceOption);
      } else {
        elements.paneSelect.appendChild(elements.agentsPaneOption);
      }
    }
    elements.agentsPaneOption.hidden = !agentManageAllowed;
    elements.agentsPaneOption.disabled = !agentManageAllowed;
  }
  if (elements.agentsPane) {
    elements.agentsPane.hidden = !agentManageAllowed;
  }
  if (!agentManageAllowed && elements.agentsPane?.classList.contains('active')) {
    setActivePane('overview');
  }
  if (!agentManageAllowed && elements.paneSelect?.value === 'agents') {
    elements.paneSelect.value = 'overview';
  }
  if (!agentManageAllowed) {
    clearAgentCredentialDisplay();
  }

  if (elements.msaidiziFab) {
    elements.msaidiziFab.disabled = !msaidiziAvailable;
    elements.msaidiziFab.title = msaidiziAvailable
      ? 'Ask Msaidizi'
      : 'Msaidizi needs backend API connectivity.';
  }
  if (elements.msaidiziAskBtn) {
    elements.msaidiziAskBtn.disabled = !msaidiziAvailable;
  }
  if (elements.msaidiziSyncBtn) {
    elements.msaidiziSyncBtn.hidden = !msaidiziSyncAllowed;
    elements.msaidiziSyncBtn.disabled = !msaidiziSyncAllowed;
  }

  elements.produceForm.querySelector('button[type="submit"]').disabled = !produceAllowed;
  elements.purchaseForm.querySelector('button[type="submit"]').disabled = !producePurchaseAllowed;

  elements.owedPeriod.disabled = !owedAllowed;
  elements.owedFromDate.disabled = !owedAllowed;
  elements.owedToDate.disabled = !owedAllowed;
  elements.owedSearch.disabled = !owedAllowed;
  elements.owedRefreshBtn.disabled = !owedAllowed;
  elements.owedSelectVisibleBtn.disabled = !owedAllowed;
  elements.owedClearSelectedBtn.disabled = !owedAllowed;
  elements.owedPrepareBtn.disabled = !owedAllowed;
  elements.owedPaySelectedBtn.disabled = !owedAllowed;
  elements.owedExportBtn.disabled = !owedAllowed;
  elements.owedPrepareBtn.hidden = !owedAllowed;
  elements.owedPaySelectedBtn.hidden = !owedAllowed;

  if (elements.paymentRecommendationCard) {
    elements.paymentRecommendationCard.hidden = !paymentRecommendationViewAllowed;
  }
  if (elements.paymentRecommendationStatusFilter) {
    elements.paymentRecommendationStatusFilter.disabled = !paymentRecommendationViewAllowed;
  }
  if (elements.paymentRecommendationSearch) {
    elements.paymentRecommendationSearch.disabled = !paymentRecommendationViewAllowed;
  }
  if (elements.paymentRecommendationRefreshBtn) {
    elements.paymentRecommendationRefreshBtn.disabled = !paymentRecommendationViewAllowed;
  }
  if (elements.paymentRecommendationFarmer) {
    elements.paymentRecommendationFarmer.disabled = !paymentRecommendationCreateAllowed;
  }
  if (elements.paymentRecommendationAmount) {
    elements.paymentRecommendationAmount.disabled = !paymentRecommendationCreateAllowed;
  }
  if (elements.paymentRecommendationReason) {
    elements.paymentRecommendationReason.disabled = !paymentRecommendationCreateAllowed;
  }
  if (elements.paymentRecommendationSubmitBtn) {
    elements.paymentRecommendationSubmitBtn.disabled = !paymentRecommendationCreateAllowed;
  }
  if (elements.paymentRecommendationAgentHint) {
    elements.paymentRecommendationAgentHint.hidden = !paymentRecommendationCreateAllowed;
  }
  if (elements.paymentRecommendationTableWrap) {
    elements.paymentRecommendationTableWrap.classList.toggle('admin-review-mode', paymentRecommendationDecisionAllowed);
  }

  elements.paymentForm.querySelector('button[type="submit"]').disabled = !paymentAllowed;
  elements.mpesaDisburseBtn.disabled = !paymentAllowed;
  if (elements.paymentLogCard) {
    elements.paymentLogCard.hidden = !paymentAllowed;
  }

  elements.smsForm.querySelector('button[type="submit"]').disabled = !smsAllowed;
  if (elements.smsDraftPurpose) elements.smsDraftPurpose.disabled = !aiSmsAllowed;
  if (elements.smsDraftLanguage) elements.smsDraftLanguage.disabled = !aiSmsAllowed;
  if (elements.smsDraftAudience) elements.smsDraftAudience.disabled = !aiSmsAllowed;
  if (elements.smsDraftTone) elements.smsDraftTone.disabled = !aiSmsAllowed;
  if (elements.smsDraftMaxLength) elements.smsDraftMaxLength.disabled = !aiSmsAllowed;
  if (elements.smsDraftBtn) elements.smsDraftBtn.disabled = !aiSmsAllowed;

  if (elements.copilotPrompt) elements.copilotPrompt.disabled = !copilotAllowed;
  if (elements.copilotAskBtn) elements.copilotAskBtn.disabled = !copilotAllowed;
  if (elements.qcAiRecord) elements.qcAiRecord.disabled = !qcAiAllowed;
  if (elements.qcAiRunBtn) elements.qcAiRunBtn.disabled = !qcAiAllowed;
  if (elements.paymentRiskPeriod) elements.paymentRiskPeriod.disabled = !paymentRiskAllowed;
  if (elements.paymentRiskRunBtn) elements.paymentRiskRunBtn.disabled = !paymentRiskAllowed;
  if (elements.paymentRiskFrom) elements.paymentRiskFrom.disabled = !paymentRiskAllowed || elements.paymentRiskFrom.hidden;
  if (elements.paymentRiskTo) elements.paymentRiskTo.disabled = !paymentRiskAllowed || elements.paymentRiskTo.hidden;
  if (elements.briefRunNowBtn) elements.briefRunNowBtn.disabled = !executiveBriefRunAllowed;
  if (elements.briefRefreshBtn) elements.briefRefreshBtn.disabled = !executiveBriefViewAllowed;

  if (elements.proposalStatusFilter) elements.proposalStatusFilter.disabled = !proposalViewAllowed;
  if (elements.proposalRefreshBtn) elements.proposalRefreshBtn.disabled = !proposalViewAllowed;
  if (elements.proposalActionType) elements.proposalActionType.disabled = !proposalManageAllowed;
  if (elements.proposalConfidence) elements.proposalConfidence.disabled = !proposalManageAllowed;
  if (elements.proposalTitle) elements.proposalTitle.disabled = !proposalManageAllowed;
  if (elements.proposalDescription) elements.proposalDescription.disabled = !proposalManageAllowed;
  if (elements.proposalPayload) elements.proposalPayload.disabled = !proposalManageAllowed;
  if (elements.proposalForm) {
    const proposalSubmit = elements.proposalForm.querySelector('button[type="submit"]');
    if (proposalSubmit) proposalSubmit.disabled = !proposalManageAllowed;
  }

  if (elements.opsFromPaymentRiskBtn) elements.opsFromPaymentRiskBtn.disabled = !opsTaskCreateFromAiAllowed;
  if (elements.opsFromQcBtn) elements.opsFromQcBtn.disabled = !opsTaskCreateFromAiAllowed;
  if (elements.opsRefreshBtn) elements.opsRefreshBtn.disabled = !opsTaskViewAllowed;
  if (elements.opsTaskStatusFilter) elements.opsTaskStatusFilter.disabled = !opsTaskViewAllowed;
  if (elements.opsTaskSeverityFilter) elements.opsTaskSeverityFilter.disabled = !opsTaskViewAllowed;

  if (elements.knowledgeSearch) elements.knowledgeSearch.disabled = !knowledgeViewAllowed;
  if (elements.knowledgeSearchBtn) elements.knowledgeSearchBtn.disabled = !knowledgeViewAllowed;
  if (elements.knowledgeTitle) elements.knowledgeTitle.disabled = !knowledgeManageAllowed;
  if (elements.knowledgeSource) elements.knowledgeSource.disabled = !knowledgeManageAllowed;
  if (elements.knowledgeTags) elements.knowledgeTags.disabled = !knowledgeManageAllowed;
  if (elements.knowledgeContent) elements.knowledgeContent.disabled = !knowledgeManageAllowed;
  if (elements.knowledgeForm) {
    const knowledgeSubmit = elements.knowledgeForm.querySelector('button[type="submit"]');
    if (knowledgeSubmit) knowledgeSubmit.disabled = !knowledgeManageAllowed;
  }

  if (elements.aiFeedbackTool) elements.aiFeedbackTool.disabled = !feedbackAllowed;
  if (elements.aiFeedbackRating) elements.aiFeedbackRating.disabled = !feedbackAllowed;
  if (elements.aiFeedbackResponseId) elements.aiFeedbackResponseId.disabled = !feedbackAllowed;
  if (elements.aiFeedbackNote) elements.aiFeedbackNote.disabled = !feedbackAllowed;
  if (elements.aiFeedbackForm) {
    const feedbackSubmit = elements.aiFeedbackForm.querySelector('button[type="submit"]');
    if (feedbackSubmit) feedbackSubmit.disabled = !feedbackAllowed;
  }
  if (elements.aiFeedbackSummaryBtn) elements.aiFeedbackSummaryBtn.disabled = !feedbackSummaryAllowed;
  if (elements.aiEvalRunBtn) elements.aiEvalRunBtn.disabled = !evalAllowed;
  if (elements.aiEvalRefreshBtn) elements.aiEvalRefreshBtn.disabled = !evalAllowed;

  elements.seedBtn.disabled = !(adminAllowed || !API.enabled);
  elements.backupBtn.disabled = !adminAllowed;
  elements.listBackupsBtn.disabled = !adminAllowed;
  elements.resetBtn.disabled = !(adminAllowed || !API.enabled);

  const baseExportAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  if (elements.exportRange) elements.exportRange.disabled = !baseExportAllowed;
  if (elements.exportFormat) elements.exportFormat.disabled = !baseExportAllowed;
  if (elements.exportFrom) elements.exportFrom.disabled = !baseExportAllowed || elements.exportFrom.hidden;
  if (elements.exportTo) elements.exportTo.disabled = !baseExportAllowed || elements.exportTo.hidden;
  if (elements.owedExportFormat) elements.owedExportFormat.disabled = !owedAllowed;

  elements.exportButtons.forEach((button) => {
    const type = button.dataset.export;
    if (type === 'payments' || type === 'activity') {
      button.disabled = !adminAllowed;
    } else {
      button.disabled = !baseExportAllowed;
    }
  });

  updateSmsComposerUi();
}

function updateRoleHint() {
  const role = currentRole();
  if (role === 'admin') {
    elements.roleHint.textContent = 'Admin mode: full control over users, payments, reports, and backups.';
  } else if (role === 'agent') {
    elements.roleHint.textContent = 'Agent mode: farmer onboarding, produce capture, and SMS engagement.';
  } else {
    elements.roleHint.textContent = 'Farmer mode: read-only visibility into operations and payments.';
  }
}

function renderAll() {
  renderDashboardAccount();
  renderOverview();
  renderFarmers();
  renderAgents();
  renderProduce();
  renderProducePurchases();
  renderPayments();
  renderPaymentRecommendations();
  renderSms();
  renderReports();
  renderBackups();
  updateRoleHint();
}

function renderOverview() {
  const summary = state.summary || deriveSummaryFromState();

  const cards = [
    { label: 'Registered Farmers', value: summary.farmers || 0 },
    { label: 'QC Records', value: summary.qcRecords ?? summary.produceRecords ?? 0 },
    { label: 'Purchased Kg', value: Number(summary.totalPurchasedKg || 0).toFixed(1) },
    { label: 'Farmers Owed', value: summary.owedFarmers || 0 },
    { label: 'Amount Owed (KES)', value: formatCurrency(summary.totalOwedKes || 0) },
    { label: 'Payments (KES)', value: formatCurrency(summary.paymentsReceived || 0) },
    { label: 'SMS Sent', value: summary.smsSent || 0 },
    { label: 'SMS Sent (24h)', value: summary.smsSentLast24h || 0 },
    { label: 'SMS Spend (24h KES)', value: formatKesWithCents(summary.smsSpentLast24hKes || 0) },
    { label: 'Payment Success Rate', value: `${summary.paymentSuccessRate || 0}%` }
  ];

  elements.kpiGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="kpi">
          <div class="label">${escapeHtml(card.label)}</div>
          <div class="value">${escapeHtml(String(card.value))}</div>
        </article>
      `
    )
    .join('');

  elements.launchCheck.innerHTML = summary.launchReady
    ? '<span class="ok">Soft launch readiness: ON TRACK</span>'
    : '<span class="danger">Soft launch readiness: Missing one or more core modules.</span>';
}

function renderFarmers() {
  if (!state.farmers.length) {
    elements.farmerTableWrap.innerHTML = '<div class="empty">No farmers yet.</div>';
    return;
  }

  const canEdit = ['admin', 'agent'].includes(currentRole());
  const canDelete = currentRole() === 'admin';
  const canSetPin = currentRole() === 'admin';

  const rows = state.farmers
    .slice(0, 200)
    .map((farmer) => {
      const actions = canEdit
        ? `
          <button class="table-btn" data-action="edit-farmer" data-id="${escapeHtml(farmer.id)}">Edit</button>
          ${canSetPin ? `<button class="table-btn" data-action="set-pin-farmer" data-id="${escapeHtml(farmer.id)}" data-name="${escapeHtml(farmer.name)}">Set PIN</button>` : ''}
          ${canDelete ? `<button class="table-btn danger" data-action="delete-farmer" data-id="${escapeHtml(farmer.id)}">Delete</button>` : ''}
        `
        : '-';
      const portalAccess = farmer.hasPortalAccess ? 'Enabled' : 'Not Set';

      return `
        <tr>
          <td>${escapeHtml(farmer.id)}</td>
          <td>${escapeHtml(farmer.name)}</td>
          <td>${escapeHtml(farmer.phone)}</td>
          <td>${escapeHtml(farmer.nationalId || '-')}</td>
          <td>${escapeHtml(farmer.location)}</td>
          <td>${escapeHtml(languageOrDefaultClient(farmer.preferredLanguage) === 'sw' ? 'Kiswahili' : 'English')}</td>
          <td>${escapeHtml(formatHectares(farmer.hectares))}</td>
          <td>${escapeHtml(formatAcres(farmer.hectares))}</td>
          <td>${escapeHtml(formatSquareFeet(farmer.hectares))}</td>
          <td>${escapeHtml(formatHectares(farmer.avocadoHectares))}</td>
          <td>${escapeHtml(formatAcres(farmer.avocadoHectares))}</td>
          <td>${escapeHtml(formatSquareFeet(farmer.avocadoHectares))}</td>
          <td>${escapeHtml(String(farmer.trees ?? '0'))}</td>
          <td>${escapeHtml(portalAccess)}</td>
          <td>${escapeHtml(dateShort(farmer.updatedAt || farmer.createdAt))}</td>
          <td class="actions">${actions}</td>
        </tr>
      `;
    })
    .join('');

  elements.farmerTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Name</th>
          <th>Phone</th>
          <th>National ID</th>
          <th>Location</th>
          <th>Language</th>
          <th>Hectares</th>
          <th>Acres</th>
          <th>Square Feet</th>
          <th>Avocado Hectares</th>
          <th>Avocado Acres</th>
          <th>Avocado Sq Ft</th>
          <th>Trees</th>
          <th>Portal Access</th>
          <th>Updated</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderAgents() {
  if (currentRole() !== 'admin') {
    elements.agentTableWrap.innerHTML = '<div class="empty">Sign in as admin to manage agent accounts.</div>';
    return;
  }

  if (!API.enabled) {
    elements.agentTableWrap.innerHTML = '<div class="empty">Agent management is available in backend mode only.</div>';
    return;
  }

  if (!state.agents.length) {
    elements.agentTableWrap.innerHTML = '<div class="empty">No agent accounts found. Create one above.</div>';
    return;
  }

  const rows = state.agents
    .slice(0, 500)
    .map((agent) => `
      <tr>
        <td>${escapeHtml(agent.name || '-')}</td>
        <td>${escapeHtml(agent.username || '-')}</td>
        <td>${escapeHtml(agent.email || '-')}</td>
        <td>${escapeHtml(agent.phone || '-')}</td>
        <td>${escapeHtml(agent.status || '-')}</td>
        <td>${escapeHtml(agent.provisioning === 'environment' ? 'Environment' : 'Admin')}</td>
        <td>${escapeHtml(dateShort(agent.updatedAt || agent.createdAt))}</td>
      </tr>
    `)
    .join('');

  elements.agentTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Name</th>
          <th>Username</th>
          <th>Email</th>
          <th>Phone</th>
          <th>Status</th>
          <th>Source</th>
          <th>Updated</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderProduce() {
  if (!state.produce.length) {
    elements.produceTableWrap.innerHTML = '<div class="empty">No produce entries yet.</div>';
    hydratePurchaseQcOptions();
    hydrateQcAiOptions();
    return;
  }

  const canDelete = currentRole() === 'admin';

  const rows = state.produce
    .slice(0, 250)
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.id)}</td>
        <td>${escapeHtml(row.farmerName)}</td>
        <td>${escapeHtml(String(row.variety || '-'))}</td>
        <td>${escapeHtml(String(row.sizeCode || '-'))}</td>
        <td>${escapeHtml(String(row.lotWeightKgs ?? row.kgs ?? '-'))}</td>
        <td>${escapeHtml(String(row.avgFruitWeightG ?? '-'))}</td>
        <td>${escapeHtml(String(row.sampleSize ?? '-'))}</td>
        <td>${escapeHtml(String(row.visualGrade || row.quality || '-'))}</td>
        <td>${escapeHtml(row.dryMatterPct == null || row.dryMatterPct === '' ? '-' : String(row.dryMatterPct))}</td>
        <td>${escapeHtml(`${row.firmnessValue ?? '-'} ${row.firmnessUnit || ''}`.trim() || '-')}</td>
        <td>${escapeHtml(String(row.qcDecision || '-'))}</td>
        <td>${escapeHtml(String(row.inspector || row.agent || '-'))}</td>
        <td>${escapeHtml(dateShort(row.createdAt))}</td>
        <td class="actions">
          ${
            canDelete
              ? `<button class="table-btn danger" data-action="delete-produce" data-id="${escapeHtml(row.id)}">Delete</button>`
              : '-'
          }
        </td>
      </tr>
    `)
    .join('');

  elements.produceTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Farmer</th>
          <th>Variety</th>
          <th>Size</th>
          <th>Lot Kg</th>
          <th>Avg g</th>
          <th>Sample</th>
          <th>Visual</th>
          <th>Dry %</th>
          <th>Firmness</th>
          <th>Decision</th>
          <th>Inspector</th>
          <th>Date</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
  hydratePurchaseQcOptions();
  hydrateQcAiOptions();
}

function renderProducePurchases() {
  if (!API.enabled) {
    reconcileLocalPurchases();
  }

  if (!state.producePurchases.length) {
    elements.purchaseTableWrap.innerHTML = '<div class="empty">No purchased produce entries yet.</div>';
    return;
  }

  const canDelete = currentRole() === 'admin';
  const rows = state.producePurchases
    .slice(0, 250)
    .map((row) => {
      const actions = canDelete
        ? `<button class="table-btn danger" data-action="delete-purchase" data-id="${escapeHtml(row.id)}">Delete</button>`
        : '-';

      return `
        <tr>
          <td>${escapeHtml(row.id)}</td>
          <td>${escapeHtml(row.farmerName || '-')}</td>
          <td>${escapeHtml(String(row.qcRecordId || '-'))}</td>
          <td>${escapeHtml(String(row.variety || '-'))}</td>
          <td>${escapeHtml(String(row.sizeCode || '-'))}</td>
          <td>${escapeHtml(String(row.purchasedKgs ?? '-'))}</td>
          <td>${escapeHtml(row.pricePerKgKes == null || row.pricePerKgKes === '' ? '-' : formatCurrency(row.pricePerKgKes))}</td>
          <td>${escapeHtml(row.purchaseValueKes == null || row.purchaseValueKes === '' ? '-' : formatCurrency(row.purchaseValueKes))}</td>
          <td>${escapeHtml(row.paidAmountKes == null || row.paidAmountKes === '' ? '-' : formatCurrency(row.paidAmountKes))}</td>
          <td>${escapeHtml(row.balanceKes == null || row.balanceKes === '' ? '-' : formatCurrency(row.balanceKes))}</td>
          <td>${escapeHtml(String(row.settlementStatus || '-'))}</td>
          <td>${escapeHtml(String(row.buyer || '-'))}</td>
          <td>${escapeHtml(dateShort(row.createdAt))}</td>
          <td class="actions">${actions}</td>
        </tr>
      `;
    })
    .join('');

  elements.purchaseTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Farmer</th>
          <th>QC Lot</th>
          <th>Variety</th>
          <th>Size</th>
          <th>Purchased Kg</th>
          <th>Price/Kg (KES)</th>
          <th>Value (KES)</th>
          <th>Paid (KES)</th>
          <th>Balance (KES)</th>
          <th>Status</th>
          <th>Buyer</th>
          <th>Date</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderPayments() {
  if (!state.payments.length) {
    elements.paymentTableWrap.innerHTML = '<div class="empty">No payments yet.</div>';
    return;
  }

  const canChangeStatus = currentRole() === 'admin';

  const rows = state.payments
    .slice(0, 250)
    .map((row) => {
      const statusActions = canChangeStatus
        ? `
          <button class="table-btn" data-action="set-payment-status" data-id="${escapeHtml(row.id)}" data-status="Received">Received</button>
          <button class="table-btn warn" data-action="set-payment-status" data-id="${escapeHtml(row.id)}" data-status="Pending">Pending</button>
          <button class="table-btn danger" data-action="set-payment-status" data-id="${escapeHtml(row.id)}" data-status="Failed">Failed</button>
        `
        : '-';

      return `
        <tr>
          <td>${escapeHtml(row.id)}</td>
          <td>${escapeHtml(row.farmerName)}</td>
          <td>${escapeHtml(formatCurrency(row.amount))}</td>
          <td>${escapeHtml(row.ref)}</td>
          <td>${escapeHtml(row.status)}</td>
          <td>${escapeHtml(row.method || 'M-PESA')}</td>
          <td>${escapeHtml(dateShort(row.createdAt))}</td>
          <td class="actions">${statusActions}</td>
        </tr>
      `;
    })
    .join('');

  elements.paymentTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Farmer</th>
          <th>Amount</th>
          <th>Reference</th>
          <th>Status</th>
          <th>Method</th>
          <th>Date</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderSms() {
  const summary = state.summary || deriveSummaryFromState();
  const smsCostSummary = `
    <p class="meta">
      Owner SMS spend: <strong>KES ${escapeHtml(formatKesWithCents(summary.smsSpentLast24hKes || 0))}</strong> in the last 24 hours
      (${escapeHtml(String(summary.smsSentLast24h || 0))} billable SMS), total
      <strong>KES ${escapeHtml(formatKesWithCents(summary.smsSpentKes || 0))}</strong>.
    </p>
  `;

  if (!state.smsLogs.length) {
    elements.smsTableWrap.innerHTML = `${smsCostSummary}<div class="empty">No SMS messages logged yet.</div>`;
    return;
  }

  const rows = state.smsLogs
    .slice(0, 250)
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.id)}</td>
          <td>${escapeHtml(row.phone)}</td>
          <td>${escapeHtml(row.farmerName || '-')}</td>
          <td>${escapeHtml(row.message)}</td>
          <td>${escapeHtml(row.status)}</td>
          <td>${escapeHtml(formatKesWithCents(smsOwnerCostKesClient(row, summary.smsOwnerCostPerMessageKes || SMS_OWNER_COST_PER_MESSAGE_KES)))}</td>
          <td>${escapeHtml(row.createdBy || '-')}</td>
          <td>${escapeHtml(dateShort(row.createdAt))}</td>
        </tr>
      `
    )
    .join('');

  elements.smsTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th>
          <th>Phone</th>
          <th>Farmer</th>
          <th>Message</th>
          <th>Status</th>
          <th>Owner Cost (KES)</th>
          <th>By</th>
          <th>Date</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  elements.smsTableWrap.innerHTML = `${smsCostSummary}${elements.smsTableWrap.innerHTML}`;
}

function renderReports() {
  const summary = state.summary || deriveSummaryFromState();
  const cards = [
    { label: 'Farmers', value: summary.farmers || 0 },
    { label: 'QC Lot Weight (kg)', value: Number(summary.totalProduceKg || 0).toFixed(1) },
    { label: 'Purchased Produce (kg)', value: Number(summary.totalPurchasedKg || 0).toFixed(1) },
    { label: 'Farmers Owed', value: summary.owedFarmers || 0 },
    { label: 'Amount Owed (KES)', value: formatCurrency(summary.totalOwedKes || 0) },
    { label: 'Payments Received (KES)', value: formatCurrency(summary.paymentsReceived || 0) },
    { label: 'SMS Sent (24h)', value: summary.smsSentLast24h || 0 },
    { label: 'SMS Spend (24h KES)', value: formatKesWithCents(summary.smsSpentLast24hKes || 0) },
    { label: 'SMS Spend Total (KES)', value: formatKesWithCents(summary.smsSpentKes || 0) },
    { label: 'Payment Success', value: `${summary.paymentSuccessRate || 0}%` }
  ];

  elements.reportMetrics.innerHTML = cards
    .map(
      (card) => `
        <article class="kpi">
          <div class="label">${escapeHtml(card.label)}</div>
          <div class="value">${escapeHtml(String(card.value))}</div>
        </article>
      `
    )
    .join('');

  const dataSource = API.enabled ? 'backend API' : 'local browser storage';
  const user = isAuthenticated() ? `${state.auth.user.username} (${state.auth.user.role})` : `offline role: ${state.role}`;

  elements.reportNarrative.textContent =
    `Current operations show ${summary.farmers || 0} farmers, ${summary.qcRecords ?? summary.produceRecords ?? 0} QC entries, ` +
    `${summary.purchasedRecords || 0} purchase entries, ${summary.owedFarmers || 0} farmers currently owed, ` +
    `${summary.paymentRecords || 0} payment records, and ${summary.smsSent || 0} SMS messages. ` +
    `SMS owner spend: KES ${formatKesWithCents(summary.smsSpentLast24hKes || 0)} in the last 24 hours ` +
    `(total KES ${formatKesWithCents(summary.smsSpentKes || 0)}). ` +
    `Data source: ${dataSource}. Session: ${user}.`;

  if (!state.agentStats.length) {
    elements.agentStatsWrap.innerHTML = '<div class="empty">No agent stats yet.</div>';
  } else {
    const rows = state.agentStats
      .map(
        (row) => `
          <tr>
            <td>${escapeHtml(row.actor)}</td>
            <td>${escapeHtml(String(row.farmers))}</td>
            <td>${escapeHtml(String(row.produceKg))}</td>
            <td>${escapeHtml(String(row.purchasedKg || 0))}</td>
            <td>${escapeHtml(String(row.sms))}</td>
          </tr>
        `
      )
      .join('');

    elements.agentStatsWrap.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>Actor</th>
            <th>Farmers</th>
            <th>QC Kg</th>
            <th>Purchased Kg</th>
            <th>SMS Sent</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
  }

  renderExecutiveBriefs();
  renderProposalQueue();
  renderOpsTasks();
  renderKnowledgeDocs();
  renderAiFeedbackAndEvals();
}

function renderBackups() {
  if (!state.backups || !state.backups.length) {
    elements.backupListWrap.innerHTML = '<div class="empty">No backup snapshots listed.</div>';
    return;
  }

  const canRestore = API.enabled && isAuthenticated() && currentRole() === 'admin';
  const rows = state.backups
    .slice(0, 50)
    .map(
      (item) => `
        <tr>
          <td>${escapeHtml(item.filename)}</td>
          <td>${escapeHtml(String(item.sizeBytes))}</td>
          <td>${escapeHtml(dateShort(item.createdAt))}</td>
          <td>
            ${canRestore
              ? `<button type="button" class="secondary mini-btn" data-restore-backup="${encodeURIComponent(item.filename)}">Restore</button>`
              : '<span class="dim">Admin only</span>'}
          </td>
        </tr>
      `
    )
    .join('');

  elements.backupListWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Filename</th>
          <th>Size (bytes)</th>
          <th>Created</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function clearFarmerImportResult() {
  elements.farmerImportSummary.textContent = '';
  elements.farmerImportErrors.innerHTML = '';
  elements.farmerImportErrors.classList.remove('import-errors');
}

function renderFarmerImportResult(result) {
  const totalRows = Number(result?.totalRows || 0);
  const imported = Number(result?.imported || 0);
  const updated = Number(result?.updated || 0);
  const skipped = Number(result?.skipped || 0);
  const smsSent = Number(result?.smsSent || 0);
  const anomalyCount = Number(result?.anomalyCount || 0);
  const duplicateMode = result?.duplicateMode || 'skip';
  const errors = Array.isArray(result?.errors) ? result.errors : [];

  if (imported || updated) {
    elements.farmerImportMsg.textContent =
      `Import complete. Added ${imported}, updated ${updated}.` +
      (smsSent ? ` Onboarding SMS sent: ${smsSent}.` : '');
  } else {
    elements.farmerImportMsg.textContent = 'Import completed with no new farmers added.';
  }
  elements.farmerImportSummary.textContent =
    `Rows processed: ${totalRows}. Imported: ${imported}. Updated: ${updated}. Skipped: ${skipped}. ` +
    `Duplicate mode: ${duplicateMode}. Onboarding SMS sent: ${smsSent}. Anomalies flagged: ${anomalyCount}.`;

  if (!errors.length) {
    elements.farmerImportErrors.innerHTML = '';
    elements.farmerImportErrors.classList.remove('import-errors');
    return;
  }

  const maxShown = 25;
  const listItems = errors
    .slice(0, maxShown)
    .map((entry) => {
      const row = escapeHtml(String(entry.row || '?'));
      const error = escapeHtml(entry.error || 'Invalid row');
      const suggestion = String(entry?.suggestion || '').trim();
      const observed = entry?.observed && typeof entry.observed === 'object'
        ? Object.entries(entry.observed)
          .map(([key, value]) => `${key}: ${value}`)
          .join(', ')
        : '';
      const suggestionHtml = suggestion
        ? `<br><span class="meta compact"><strong>Suggested fix:</strong> ${escapeHtml(suggestion)}</span>`
        : '';
      const observedHtml = observed
        ? `<br><span class="meta compact"><strong>Observed:</strong> ${escapeHtml(observed)}</span>`
        : '';
      return `<li>Row ${row}: ${error}${suggestionHtml}${observedHtml}</li>`;
    })
    .join('');
  const remaining = errors.length - maxShown;
  const more = remaining > 0 ? `<li>...and ${escapeHtml(String(remaining))} more row errors.</li>` : '';

  elements.farmerImportErrors.classList.add('import-errors');
  elements.farmerImportErrors.innerHTML = `<ul>${listItems}${more}</ul>`;
}

function toAreaBundleClient(hectaresValue) {
  const metrics = areaMetricsFromHectares(hectaresValue);
  if (!Number.isFinite(metrics.hectares)) {
    return { hectares: '', acres: '', squareFeet: '' };
  }
  return {
    hectares: Number(metrics.hectares.toFixed(3)),
    acres: Number(metrics.acres.toFixed(3)),
    squareFeet: Number(metrics.squareFeet.toFixed(1))
  };
}

function smartImportAutoFixClient(mapped, invalid) {
  const patch = {};
  let reason = '';
  if (invalid === 'area under avocado cannot be greater than total farm size') {
    if (Number.isFinite(mapped.hectares) && mapped.hectares > 0) {
      patch.avocadoHectares = Number(mapped.hectares.toFixed(3));
      reason = 'Set avocado area equal to total farm size.';
    }
  } else if (invalid === 'avocadoHectares/avocadoAcres/avocadoSquareFeet is required') {
    if (Number.isFinite(mapped.hectares) && mapped.hectares > 0) {
      patch.avocadoHectares = Number(mapped.hectares.toFixed(3));
      reason = 'Missing avocado area. Suggested temporary value equals total farm size.';
    }
  } else if (invalid === 'hectares/acres/square feet is required') {
    if (Number.isFinite(mapped.avocadoHectares) && mapped.avocadoHectares > 0) {
      patch.hectares = Number(mapped.avocadoHectares.toFixed(3));
      reason = 'Missing total area. Suggested temporary value equals avocado area.';
    }
  } else if (invalid === 'trees must be a number') {
    patch.trees = 0;
    reason = 'Tree count converted to 0 (requires review).';
  }

  if (!Object.keys(patch).length) {
    return null;
  }

  const corrected = { ...mapped, ...patch };
  return {
    reason,
    corrected: {
      ...corrected,
      totalArea: toAreaBundleClient(corrected.hectares),
      avocadoArea: toAreaBundleClient(corrected.avocadoHectares)
    }
  };
}

function runLocalSmartImportQa(records) {
  const existingByPhone = new Map(
    state.farmers
      .map((row) => [cleanPhone(row.phone), row])
      .filter(([phone]) => Boolean(phone))
  );
  const existingByNationalId = new Map(
    state.farmers
      .map((row) => [normalizeNationalIdClient(row.nationalId), row])
      .filter(([nationalId]) => Boolean(nationalId))
  );

  const rows = records.map((raw, index) => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      return {
        row: index + 1,
        status: 'blocked',
        confidence: 0.2,
        mapped: {},
        issues: [
          {
            code: 'invalid_row_shape',
            severity: 'error',
            message: 'Row is not a valid object.',
            suggestion: 'Check spreadsheet headers and row structure.'
          }
        ],
        proposedFix: null,
        duplicate: null
      };
    }

    const mapped = mapFarmerImportRecordClient(raw);
    const invalid = validateImportedFarmer(mapped);
    const issues = [];
    let confidence = 0.9;
    let proposedFix = null;
    let duplicate = null;

    const phoneKey = cleanPhone(mapped.phone);
    const nationalIdKey = normalizeNationalIdClient(mapped.nationalId);
    const existingPhone = phoneKey ? existingByPhone.get(phoneKey) : null;
    const existingNational = nationalIdKey ? existingByNationalId.get(nationalIdKey) : null;

    if (existingPhone || existingNational) {
      const found = existingNational || existingPhone;
      duplicate = {
        farmerId: found.id,
        farmerName: found.name || '-',
        phone: found.phone || '',
        nationalId: found.nationalId || '',
        match: existingPhone && existingNational ? 'phone+nationalId' : existingPhone ? 'phone' : 'nationalId'
      };
      issues.push({
        code: 'duplicate_farmer',
        severity: 'warning',
        message: 'Row matches an existing farmer.',
        suggestion: 'Choose overwrite mode to replace existing farmer details.'
      });
      confidence = Math.min(confidence, 0.72);
    }

    if (invalid) {
      const issue = buildFarmerImportIssueClient(mapped, invalid);
      issues.unshift({
        code: issue.code || 'invalid_import_row',
        severity: 'error',
        message: issue.error || invalid,
        suggestion: issue.suggestion || 'Fix the row and retry.'
      });

      const autoFix = smartImportAutoFixClient(mapped, invalid);
      if (autoFix) {
        proposedFix = autoFix.corrected;
        issues.push({
          code: 'auto_fix_proposed',
          severity: 'info',
          message: autoFix.reason,
          suggestion: 'Review this suggested correction before import.'
        });
        confidence = 0.6;
      } else {
        confidence = 0.4;
      }
    }

    const status = issues.some((item) => item.severity === 'error')
      ? 'blocked'
      : issues.length
        ? 'review'
        : 'ready';

    return {
      row: index + 1,
      status,
      confidence: Number(confidence.toFixed(2)),
      mapped: {
        ...mapped,
        totalArea: toAreaBundleClient(mapped.hectares),
        avocadoArea: toAreaBundleClient(mapped.avocadoHectares)
      },
      issues,
      proposedFix,
      duplicate
    };
  });

  const counts = rows.reduce(
    (acc, row) => {
      acc[row.status] = (acc[row.status] || 0) + 1;
      return acc;
    },
    { ready: 0, review: 0, blocked: 0 }
  );
  const issueSummaryMap = {};
  rows.forEach((row) => {
    row.issues.forEach((issue) => {
      issueSummaryMap[issue.code] = (issueSummaryMap[issue.code] || 0) + 1;
    });
  });

  return {
    totalRows: records.length,
    readyRows: counts.ready || 0,
    reviewRows: counts.review || 0,
    blockedRows: counts.blocked || 0,
    autoFixableRows: rows.filter((row) => Boolean(row.proposedFix)).length,
    issueSummary: Object.entries(issueSummaryMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([code, count]) => ({ code, count })),
    rows: rows.slice(0, 500)
  };
}

function renderSmartImportQaResult(result) {
  const totalRows = Number(result?.totalRows || 0);
  const readyRows = Number(result?.readyRows || 0);
  const reviewRows = Number(result?.reviewRows || 0);
  const blockedRows = Number(result?.blockedRows || 0);
  const autoFixableRows = Number(result?.autoFixableRows || 0);
  const rows = Array.isArray(result?.rows) ? result.rows : [];

  elements.farmerImportQaMsg.textContent =
    `Smart QA complete. Ready: ${readyRows}, Review: ${reviewRows}, Blocked: ${blockedRows}, Auto-fixable: ${autoFixableRows}.`;

  if (!elements.farmerImportQaWrap) return;
  if (!rows.length) {
    elements.farmerImportQaWrap.hidden = false;
    elements.farmerImportQaWrap.innerHTML = '<div class="empty">No row diagnostics returned.</div>';
    return;
  }

  const issueSummary = Array.isArray(result?.issueSummary) ? result.issueSummary : [];
  const summaryHtml = issueSummary.length
    ? `<p class="meta compact"><strong>Top issue codes:</strong> ${issueSummary
      .map((item) => `${escapeHtml(item.code)} (${escapeHtml(String(item.count))})`)
      .join(', ')}</p>`
    : '';

  const rowsHtml = rows
    .slice(0, 30)
    .map((row) => {
      const issues = (row.issues || [])
        .map((issue) => `${issue.code}: ${issue.message}`)
        .join(' | ');
      const fix = row.proposedFix
        ? `Proposed fix -> total ha: ${row.proposedFix.totalArea?.hectares || '-'}, avocado ha: ${row.proposedFix.avocadoArea?.hectares || '-'}`
        : 'No auto-fix';
      return `<li>Row ${escapeHtml(String(row.row))} [${escapeHtml(row.status)}] - ${escapeHtml(issues || 'No issues')} (${escapeHtml(fix)})</li>`;
    })
    .join('');
  const overflow = rows.length > 30 ? `<li>...and ${escapeHtml(String(rows.length - 30))} more rows.</li>` : '';

  elements.farmerImportQaWrap.hidden = false;
  elements.farmerImportQaWrap.innerHTML = `
    <div class="ai-result-title">Smart Import QA Preview (${escapeHtml(String(totalRows))} rows)</div>
    ${summaryHtml}
    <ul class="ai-result-list">${rowsHtml}${overflow}</ul>
  `;
}

function normalizeImportHeaderClient(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function firstNonEmptyClient(values) {
  for (const value of values) {
    const text = String(value ?? '').trim();
    if (text) return text;
  }
  return '';
}

function joinNonEmptyClient(values, separator = ' ') {
  return values
    .map((value) => String(value ?? '').trim())
    .filter(Boolean)
    .join(separator);
}

function parseTreeCountClient(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return 0;
  const direct = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(direct)) return direct;
  const match = raw.match(/-?\d[\d,]*(?:\.\d+)?/);
  if (!match) return NaN;
  const num = Number(match[0].replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function parseAreaClient(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return NaN;
  const direct = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(direct)) return direct;
  const match = raw.match(/-?\d[\d,]*(?:\.\d+)?/);
  if (!match) return NaN;
  const num = Number(match[0].replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function parseAreaToHectaresClient(value, defaultUnit = 'hectares') {
  const raw = String(value ?? '').trim();
  if (!raw) return NaN;
  const num = parseAreaClient(raw);
  if (!Number.isFinite(num)) return NaN;
  const lower = raw.toLowerCase();
  if (/(hectare|hectares|\bha\b)/.test(lower)) return num;
  if (/(acre|acres|acreage)/.test(lower)) return num * HECTARES_PER_ACRE;
  if (/(square\s*feet|square\s*foot|sq\s*ft|sqft|ft2|ft²)/.test(lower)) return num / SQFT_PER_HECTARE;
  if (defaultUnit === 'acres') return num * HECTARES_PER_ACRE;
  if (defaultUnit === 'squareFeet') return num / SQFT_PER_HECTARE;
  return num;
}

function areaMetricsFromHectares(hectaresValue) {
  const hectares = Number(hectaresValue);
  if (!Number.isFinite(hectares) || hectares <= 0) {
    return { hectares: NaN, acres: NaN, squareFeet: NaN };
  }

  return {
    hectares,
    acres: hectares * ACRES_PER_HECTARE,
    squareFeet: hectares * SQFT_PER_HECTARE
  };
}

function toHectaresFromInputsClient(hectaresInput, acresInput, squareFeetInput) {
  const hectaresText = String(hectaresInput ?? '').trim();
  if (hectaresText) return parseAreaToHectaresClient(hectaresText, 'hectares');

  const acresText = String(acresInput ?? '').trim();
  if (acresText) return parseAreaToHectaresClient(acresText, 'acres');

  const squareFeetText = String(squareFeetInput ?? '').trim();
  if (squareFeetText) return parseAreaToHectaresClient(squareFeetText, 'squareFeet');

  return NaN;
}

function formatAreaForInput(value, precision) {
  const num = Number(value);
  if (!Number.isFinite(num) || num <= 0) return '';
  return num.toFixed(precision).replace(/\.?0+$/, '');
}

function fillAreaInputsFromHectares(group, hectaresValue) {
  if (!group?.hectaresInput || !group?.acresInput || !group?.squareFeetInput) return;
  const metrics = areaMetricsFromHectares(hectaresValue);
  if (!Number.isFinite(metrics.hectares)) {
    group.hectaresInput.value = '';
    group.acresInput.value = '';
    group.squareFeetInput.value = '';
    return;
  }

  group.hectaresInput.value = formatAreaForInput(metrics.hectares, 3);
  group.acresInput.value = formatAreaForInput(metrics.acres, 3);
  group.squareFeetInput.value = formatAreaForInput(metrics.squareFeet, 1);
}

function buildAreaPayload(hectaresInput, acresInput, squareFeetInput) {
  const hectares = toHectaresFromInputsClient(hectaresInput, acresInput, squareFeetInput);
  if (!Number.isFinite(hectares) || hectares <= 0) {
    return {
      hectares: String(hectaresInput ?? '').trim(),
      acres: String(acresInput ?? '').trim(),
      squareFeet: String(squareFeetInput ?? '').trim()
    };
  }

  const metrics = areaMetricsFromHectares(hectares);
  return {
    hectares: Number(metrics.hectares.toFixed(3)),
    acres: Number(metrics.acres.toFixed(3)),
    squareFeet: Number(metrics.squareFeet.toFixed(1))
  };
}

function mapFarmerImportRecordClient(raw) {
  const flat = {};
  Object.entries(raw || {}).forEach(([key, value]) => {
    flat[normalizeImportHeaderClient(key)] = value;
  });

  const firstName = firstNonEmptyClient([flat.firstname, flat.firstnameseller, flat.sellerfirstname]);
  const middleName = firstNonEmptyClient([flat.middlename, flat.sellermiddlename]);
  const lastName = firstNonEmptyClient([flat.lastname, flat.lastnameseller, flat.sellerlastname, flat.surname]);
  const composedName = joinNonEmptyClient([firstName, middleName, lastName]);
  const name = firstNonEmptyClient([
    flat.name,
    flat.farmername,
    flat.fullname,
    flat.farmerfullname,
    flat.growername,
    flat.sellername,
    composedName
  ]);
  const phone = firstNonEmptyClient([
    flat.phone,
    flat.phonenumber,
    flat.mobile,
    flat.mobilenumber,
    flat.msisdn,
    flat.contactnumber
  ]);
  const nationalId = firstNonEmptyClient([
    flat.nationalid,
    flat.nationalidnumber,
    flat.idnumber,
    flat.idno,
    flat.nationalidno,
    flat.governmentid,
    flat.governmentidnumber,
    flat.identificationnumber
  ]);
  const explicitLocation = firstNonEmptyClient([
    flat.location,
    flat.area,
    flat.village,
    flat.farmlocation,
    flat.nearestcollectionpointdepot
  ]);
  const ward = firstNonEmptyClient([flat.ward, flat.locationward, flat.wardname, flat['2']]);
  const county = firstNonEmptyClient([flat.county, flat.region, flat.locationcounty, flat.countyname, flat['3']]);
  const wardCountyLocation = joinNonEmptyClient([ward, county], ', ');
  const location = firstNonEmptyClient([explicitLocation, wardCountyLocation, ward, county, flat.locationofsocietycooperative]);
  const notes = firstNonEmptyClient([
    flat.notes,
    flat.note,
    flat.comments,
    flat.remarks,
    flat.description,
    flat.additionalinformationcommentsnotes
  ]);
  const treesRaw = firstNonEmptyClient([
    flat.trees,
    flat.treecount,
    flat.numberoftrees,
    flat.treequantity,
    flat.treenumber,
    flat.numberapproximateofhasstrees,
    flat.numberofhasstrees
  ]);
  const hectaresRaw = firstNonEmptyClient([
    flat.hectares,
    flat.hectare,
    flat.farmsizeha,
    flat.farmsizehectares,
    flat.landsizehectares,
    flat.totalfarmsizehectares
  ]);
  const acresRaw = firstNonEmptyClient([
    flat.acres,
    flat.acre,
    flat.acreage,
    flat.farmsizeacres,
    flat.landsizeacres,
    flat.farmsize,
    flat.totalfarmsize
  ]);
  const squareFeetRaw = firstNonEmptyClient([
    flat.squarefeet,
    flat.squarefoot,
    flat.squareft,
    flat.sqft,
    flat.ft2,
    flat.squarefeetarea,
    flat.farmsizesquarefeet,
    flat.landsizesquarefeet
  ]);
  const avocadoHectaresRaw = firstNonEmptyClient([
    flat.avocadohectares,
    flat.areaunderavocadohectares,
    flat.avocadoareahectares,
    flat.avocadoplothectares
  ]);
  const avocadoAcresRaw = firstNonEmptyClient([
    flat.avocadoacres,
    flat.avocadoacreage,
    flat.areaunderavocadoacres,
    flat.avocadoareaacres,
    flat.avocadoplotacres,
    flat.areaunderhassavocado,
    flat.areaunderavocado
  ]);
  const avocadoSquareFeetRaw = firstNonEmptyClient([
    flat.avocadosquarefeet,
    flat.avocadosquarefoot,
    flat.avocadosqft,
    flat.areaunderavocadosquarefeet,
    flat.areaunderavocadosqft
  ]);
  const preferredLanguageRaw = firstNonEmptyClient([
    flat.preferredlanguage,
    flat.language,
    flat.farmerlanguage,
    flat.smslanguage,
    flat.ussdlanguage,
    flat.lugha
  ]);

  return {
    name,
    phone: cleanPhone(phone),
    nationalId: cleanNationalIdClient(nationalId),
    location,
    preferredLanguage: languageOrDefaultClient(preferredLanguageRaw),
    hectares: toHectaresFromInputsClient(hectaresRaw, acresRaw, squareFeetRaw),
    avocadoHectares: toHectaresFromInputsClient(avocadoHectaresRaw, avocadoAcresRaw, avocadoSquareFeetRaw),
    trees: parseTreeCountClient(treesRaw),
    notes
  };
}

function validateImportedFarmer(mapped) {
  if (!mapped.name) return 'name is required';
  if (!mapped.phone) return 'phone is required';
  if (!mapped.nationalId) return 'nationalId is required';
  if (!mapped.location) return 'location is required';
  if (!Number.isFinite(mapped.hectares)) return 'hectares/acres/square feet is required';
  if (!Number.isFinite(mapped.avocadoHectares)) return 'avocadoHectares/avocadoAcres/avocadoSquareFeet is required';
  if (mapped.hectares <= 0) return 'hectares must be greater than 0';
  if (mapped.avocadoHectares <= 0) return 'area under avocado must be greater than 0';
  if (mapped.avocadoHectares > mapped.hectares) return 'area under avocado cannot be greater than total farm size';
  if (Number.isNaN(mapped.trees)) return 'trees must be a number';
  return '';
}

function formatAreaHintClient(hectares) {
  const num = Number(hectares);
  if (!Number.isFinite(num) || num <= 0) return '';
  const acres = num * ACRES_PER_HECTARE;
  return `${num.toFixed(3)} ha (~${acres.toFixed(2)} acres)`;
}

function buildFarmerImportIssueClient(mapped, invalid) {
  const totalHectares = Number(mapped?.hectares);
  const avocadoHectares = Number(mapped?.avocadoHectares);
  const issue = {
    error: invalid,
    code: 'invalid_import_row',
    suggestion: 'Check the row fields and retry import.'
  };

  if (invalid === 'name is required') {
    issue.code = 'missing_name';
    issue.suggestion = 'Fill farmer name (or first and last seller names) in this row.';
    return issue;
  }
  if (invalid === 'phone is required') {
    issue.code = 'missing_phone';
    issue.suggestion = 'Fill a valid farmer phone number (Kenya format supported: 07..., 7..., or 254...).';
    return issue;
  }
  if (invalid === 'nationalId is required') {
    issue.code = 'missing_national_id';
    issue.suggestion = 'Fill the National ID / Identification Number column for this farmer.';
    return issue;
  }
  if (invalid === 'location is required') {
    issue.code = 'missing_location';
    issue.suggestion = 'Fill location (ward/county/area) for this farmer.';
    return issue;
  }
  if (invalid === 'hectares/acres/square feet is required') {
    issue.code = 'missing_total_area';
    issue.suggestion = 'Fill total farm size with a number (e.g., 2.5, 5 acres, or 1 hectare).';
    return issue;
  }
  if (invalid === 'avocadoHectares/avocadoAcres/avocadoSquareFeet is required') {
    issue.code = 'missing_avocado_area';
    issue.suggestion = 'Fill area under avocado with a number (e.g., 1.2, 3 acres, or 0.8 hectare).';
    return issue;
  }
  if (invalid === 'hectares must be greater than 0') {
    issue.code = 'invalid_total_area';
    issue.suggestion = 'Set total farm size to a value greater than zero.';
    return issue;
  }
  if (invalid === 'area under avocado must be greater than 0') {
    issue.code = 'invalid_avocado_area';
    issue.suggestion = 'Set area under avocado to a value greater than zero.';
    return issue;
  }
  if (invalid === 'area under avocado cannot be greater than total farm size') {
    issue.code = 'avocado_area_exceeds_total';
    issue.suggestion =
      Number.isFinite(totalHectares) && totalHectares > 0
        ? `Set area under avocado to <= total farm size (${formatAreaHintClient(totalHectares)}), or fix units in the source row.`
        : 'Set area under avocado to a value less than or equal to total farm size.';
    issue.observed = {
      totalHectares: Number.isFinite(totalHectares) ? Number(totalHectares.toFixed(3)) : '',
      avocadoHectares: Number.isFinite(avocadoHectares) ? Number(avocadoHectares.toFixed(3)) : ''
    };
    return issue;
  }
  if (invalid === 'trees must be a number') {
    issue.code = 'invalid_tree_count';
    issue.suggestion = 'Use a numeric value for tree count (e.g., 400).';
    return issue;
  }

  return issue;
}

function cleanPhone(value) {
  const raw = String(value ?? '').trim();
  const digits = raw.replace(/[^\d]/g, '');
  if (!digits) return '';
  if (digits.startsWith('254') && digits.length === 12) return digits;
  if (digits.startsWith('0') && digits.length === 10) return `254${digits.slice(1)}`;
  if (digits.length === 9 && digits.startsWith('7')) return `254${digits}`;
  return digits || raw;
}

function normalizeNationalIdClient(value) {
  return String(value ?? '')
    .trim()
    .toUpperCase()
    .replace(/[\s-]/g, '');
}

function cleanNationalIdClient(value) {
  return String(value ?? '').trim().toUpperCase();
}

function isTruthyClient(value) {
  if (typeof value === 'boolean') return value;
  const lower = String(value ?? '').trim().toLowerCase();
  return lower === 'true' || lower === '1' || lower === 'yes' || lower === 'on';
}

function normalizeLanguageClient(value) {
  const lower = String(value ?? '').trim().toLowerCase();
  if (['2', 'sw', 'kiswahili', 'swahili'].includes(lower)) return 'sw';
  if (['1', 'en', 'english'].includes(lower)) return 'en';
  return '';
}

function languageOrDefaultClient(value) {
  return normalizeLanguageClient(value) || DEFAULT_LANGUAGE;
}

function defaultImportOnboardingSmsTemplateClient(language) {
  return languageOrDefaultClient(language) === 'sw'
    ? DEFAULT_IMPORT_ONBOARDING_SMS_TEMPLATE_SW
    : DEFAULT_IMPORT_ONBOARDING_SMS_TEMPLATE_EN;
}

function renderImportOnboardingSmsTemplateClient(template, farmer) {
  const lang = languageOrDefaultClient(farmer?.preferredLanguage);
  const rawTemplate = String(template || '').trim() || defaultImportOnboardingSmsTemplateClient(lang);
  const values = {
    name: String(farmer?.name || '').trim() || 'farmer',
    phone: String(farmer?.phone || '').trim(),
    nationalId: String(farmer?.nationalId || '').trim(),
    location: String(farmer?.location || '').trim(),
    ussd: '*483#',
    portal: 'Agem Portal'
  };
  const rendered = rawTemplate.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => values[key] || '');
  return rendered.trim();
}

function hasFarmerPhoneConflict(phone, excludeId = '') {
  const target = cleanPhone(phone);
  if (!target) return false;
  return state.farmers.some((row) => cleanPhone(row.phone) === target && row.id !== excludeId);
}

function hasFarmerNationalIdConflict(nationalId, excludeId = '') {
  const target = normalizeNationalIdClient(nationalId);
  if (!target) return false;
  return state.farmers.some(
    (row) => normalizeNationalIdClient(row.nationalId) === target && row.id !== excludeId
  );
}

function localFarmerDuplicateError(payload, excludeId = '') {
  if (hasFarmerPhoneConflict(payload.phone, excludeId)) {
    return 'A farmer with this phone number already exists.';
  }
  if (hasFarmerNationalIdConflict(payload.nationalId, excludeId)) {
    return 'A farmer with this National ID already exists.';
  }
  return '';
}

function findImportDuplicates(records) {
  const existingPhones = new Set(state.farmers.map((row) => cleanPhone(row.phone)).filter(Boolean));
  const existingNationalIds = new Set(
    state.farmers.map((row) => normalizeNationalIdClient(row.nationalId)).filter(Boolean)
  );
  const duplicatePhones = new Set();
  const duplicateNationalIds = new Set();

  records.forEach((raw) => {
    const mapped = mapFarmerImportRecordClient(raw);
    const phone = cleanPhone(mapped.phone);
    const nationalId = normalizeNationalIdClient(mapped.nationalId);
    if (phone && existingPhones.has(phone)) {
      duplicatePhones.add(phone);
    }
    if (nationalId && existingNationalIds.has(nationalId)) {
      duplicateNationalIds.add(nationalId);
    }
  });

  return {
    phones: [...duplicatePhones],
    nationalIds: [...duplicateNationalIds]
  };
}

function importFarmersLocal(records, options = {}) {
  const onDuplicate = options.onDuplicate === 'overwrite' ? 'overwrite' : 'skip';
  const sendOnboardingSms = isTruthyClient(options.sendOnboardingSms);
  const onboardingSmsTemplate = String(options.onboardingSmsTemplate || '').trim();
  const farmersByPhone = new Map(
    state.farmers
      .map((row) => [cleanPhone(row.phone), row])
      .filter(([phone]) => Boolean(phone))
  );
  const farmersByNationalId = new Map(
    state.farmers
      .map((row) => [normalizeNationalIdClient(row.nationalId), row])
      .filter(([nationalId]) => Boolean(nationalId))
  );
  const created = [];
  const updated = [];
  const errors = [];
  const now = Date.now();

  records.forEach((raw, idx) => {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      errors.push({ row: idx + 1, error: 'Row is not a valid object' });
      return;
    }

    const mapped = mapFarmerImportRecordClient(raw);
    const invalid = validateImportedFarmer(mapped);
    if (invalid) {
      const issue = buildFarmerImportIssueClient(mapped, invalid);
      errors.push({ row: idx + 1, ...issue });
      return;
    }

    const phone = cleanPhone(mapped.phone);
    const nationalId = cleanNationalIdClient(mapped.nationalId);
    const nationalIdKey = normalizeNationalIdClient(nationalId);
    const existingByNationalId = farmersByNationalId.get(nationalIdKey);
    const existingByPhone = farmersByPhone.get(phone);
    const existing = existingByNationalId || existingByPhone;
    if (existing) {
      if (onDuplicate === 'overwrite') {
        if (
          existingByNationalId &&
          existingByPhone &&
          existingByNationalId.id !== existingByPhone.id
        ) {
          errors.push({
            row: idx + 1,
            error: 'Conflicting duplicate match: phone and National ID belong to different farmers',
            code: 'duplicate_conflict',
            suggestion: 'Review this row manually: phone and National ID currently belong to different existing farmers.'
          });
          return;
        }

        const oldPhoneKey = cleanPhone(existing.phone);
        const oldNationalIdKey = normalizeNationalIdClient(existing.nationalId);
        if (oldPhoneKey && farmersByPhone.get(oldPhoneKey)?.id === existing.id) {
          farmersByPhone.delete(oldPhoneKey);
        }
        if (oldNationalIdKey && farmersByNationalId.get(oldNationalIdKey)?.id === existing.id) {
          farmersByNationalId.delete(oldNationalIdKey);
        }

        existing.name = mapped.name;
        existing.phone = phone;
        existing.nationalId = nationalId;
        existing.location = mapped.location;
        existing.hectares = Number(mapped.hectares.toFixed(3));
        existing.avocadoHectares = Number(mapped.avocadoHectares.toFixed(3));
        existing.trees = Number(mapped.trees.toFixed(2));
        existing.preferredLanguage = languageOrDefaultClient(mapped.preferredLanguage);
        existing.notes = mapped.notes;
        existing.updatedAt = new Date().toISOString();

        farmersByPhone.set(phone, existing);
        farmersByNationalId.set(nationalIdKey, existing);
        updated.push(existing);
      } else {
        errors.push({
          row: idx + 1,
          error: 'Duplicate farmer (matching phone or National ID)',
          code: 'duplicate_farmer',
          suggestion: 'Use overwrite mode to update existing farmer details, or keep skip mode to retain old data.'
        });
      }
      return;
    }

    const record = {
      id: `F-${now}-${idx + 1}`,
      name: mapped.name,
      phone,
      nationalId,
      location: mapped.location,
      preferredLanguage: languageOrDefaultClient(mapped.preferredLanguage),
      hectares: Number(mapped.hectares.toFixed(3)),
      avocadoHectares: Number(mapped.avocadoHectares.toFixed(3)),
      trees: Number(mapped.trees.toFixed(2)),
      notes: mapped.notes,
      createdBy: 'local',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    farmersByPhone.set(phone, record);
    farmersByNationalId.set(nationalIdKey, record);
    created.push(record);
  });

  if (created.length) {
    state.farmers = created.reverse().concat(state.farmers);
  }

  let smsSent = 0;
  if (sendOnboardingSms && (created.length || updated.length)) {
    const byPhone = new Map();
    for (const farmer of created.concat(updated)) {
      const phone = cleanPhone(farmer.phone);
      if (!phone || byPhone.has(phone)) continue;
      byPhone.set(phone, farmer);
    }

    const nowIso = new Date().toISOString();
    const logs = [];
    for (const [phone, farmer] of byPhone.entries()) {
      const message = renderImportOnboardingSmsTemplateClient(onboardingSmsTemplate, farmer);
      if (!message) continue;
      logs.push({
        id: `SMS-${Date.now()}-${logs.length + 1}`,
        farmerId: farmer.id || '',
        farmerName: farmer.name || '',
        phone,
        message,
        provider: 'Local Mock',
        ownerCostKes: SMS_OWNER_COST_PER_MESSAGE_KES,
        status: 'Sent',
        createdBy: 'local',
        createdAt: nowIso
      });
    }
    if (logs.length) {
      state.smsLogs = logs.reverse().concat(state.smsLogs);
      smsSent = logs.length;
    }
  }

  return {
    totalRows: records.length,
    imported: created.length,
    updated: updated.length,
    skipped: records.length - created.length - updated.length,
    duplicateMode: onDuplicate,
    sendOnboardingSms,
    smsSent,
    anomalyCount: errors.filter((entry) => String(entry?.suggestion || '').trim()).length,
    errors: errors.slice(0, 250)
  };
}

async function parseFarmerImportFile(file) {
  const ext = String(file.name || '')
    .split('.')
    .pop()
    .toLowerCase();
  const type = String(file.type || '').toLowerCase();
  const isCsv = ext === 'csv' || type.includes('csv') || type === 'text/plain';
  const isExcel =
    ext === 'xlsx' ||
    ext === 'xls' ||
    type.includes('spreadsheetml') ||
    type.includes('ms-excel') ||
    type.includes('excel');

  if (isCsv) {
    const content = await readFileAsText(file);
    return parseCsvRecords(content);
  }

  if (isExcel) {
    if (!window.XLSX || typeof window.XLSX.read !== 'function') {
      throw new Error('Excel support is still loading. Wait a few seconds and try again, or use CSV.');
    }

    const buffer = await readFileAsArrayBuffer(file);
    const workbook = window.XLSX.read(buffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames?.[0];
    if (!firstSheetName) return [];

    const sheet = workbook.Sheets[firstSheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, {
      defval: '',
      raw: false,
      blankrows: false
    });

    return rows.filter(
      (row) =>
        row &&
        typeof row === 'object' &&
        !Array.isArray(row) &&
        Object.values(row).some((value) => String(value ?? '').trim())
    );
  }

  throw new Error('Unsupported file type. Use .csv, .xlsx, or .xls.');
}

function parseCsvRecords(content) {
  const rows = parseCsvRows(String(content || ''));
  if (!rows.length) return [];

  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell ?? '').trim()));
  if (headerIndex < 0) return [];

  const headerRow = rows[headerIndex];
  const headerCounts = {};
  const headers = headerRow.map((cell, idx) => {
    const base = String(cell ?? '').trim() || `Column_${idx + 1}`;
    const used = headerCounts[base] || 0;
    headerCounts[base] = used + 1;
    return used ? `${base}_${used + 1}` : base;
  });

  const records = [];
  for (let i = headerIndex + 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row || !row.some((cell) => String(cell ?? '').trim())) continue;

    const record = {};
    headers.forEach((header, idx) => {
      record[header] = row[idx] ?? '';
    });
    records.push(record);
  }

  return records;
}

function parseCsvRows(content) {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < content.length; i += 1) {
    const char = content[i];

    if (inQuotes) {
      if (char === '"') {
        if (content[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        cell += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      continue;
    }

    if (char === ',') {
      row.push(cell);
      cell = '';
      continue;
    }

    if (char === '\n') {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      continue;
    }

    if (char === '\r') {
      if (content[i + 1] === '\n') i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      continue;
    }

    cell += char;
  }

  row.push(cell);
  if (row.length > 1 || String(row[0] ?? '').trim()) {
    rows.push(row);
  }

  return rows;
}

function resetFarmerForm() {
  state.editingFarmerId = '';
  elements.farmerId.value = '';
  elements.farmerForm.reset();
  if (elements.farmerPreferredLanguage) {
    elements.farmerPreferredLanguage.value = DEFAULT_LANGUAGE;
  }
  elements.farmerSubmitBtn.textContent = 'Register Farmer';
  elements.farmerCancelEditBtn.hidden = true;
}

function hydrateFarmerPinOptions(rows = state.farmers, preferredId = '') {
  if (!elements.farmerPinFarmer) return;

  const options = rows
    .slice(0, 500)
    .map((farmer) => {
      const label = `${farmer.name || '-'} | ${farmer.phone || '-'} | ID ${farmer.nationalId || '-'}`;
      return `<option value="${escapeHtml(farmer.id)}">${escapeHtml(label)}</option>`;
    })
    .join('');

  elements.farmerPinFarmer.innerHTML = '<option value="">Select farmer</option>' + options;
  if (preferredId) {
    elements.farmerPinFarmer.value = preferredId;
  }
}

function hydrateFarmerSelectors() {
  const selectedPinFarmerId = elements.farmerPinFarmer?.value || '';
  const options = state.farmers
    .map((farmer) => `<option value="${escapeHtml(farmer.id)}">${escapeHtml(farmer.name)} (${escapeHtml(farmer.location)})</option>`)
    .join('');

  const fallback = '<option value="">No farmers yet</option>';
  elements.produceFarmer.innerHTML = options || fallback;
  elements.purchaseFarmer.innerHTML = options || fallback;
  elements.paymentFarmer.innerHTML = options || fallback;
  if (elements.paymentRecommendationFarmer) {
    elements.paymentRecommendationFarmer.innerHTML = options || fallback;
  }
  hydrateFarmerPinOptions(state.farmers, selectedPinFarmerId);
  hydratePurchaseQcOptions();
  hydrateQcAiOptions();

  if (!API.enabled && currentSmsMode() === 'selected') {
    void loadSmsRecipients(false);
  }
}

function hydratePurchaseQcOptions() {
  if (!elements.purchaseQcRecord) return;
  const selectedFarmerId = elements.purchaseFarmer?.value || '';
  const qcOptions = state.produce
    .filter((row) => !selectedFarmerId || row.farmerId === selectedFarmerId)
    .slice(0, 400)
    .map((row) => {
      const label = `${row.id} | ${row.farmerName || '-'} | ${row.variety || '-'} | ${row.lotWeightKgs ?? row.kgs ?? '-'}kg`;
      return `<option value="${escapeHtml(row.id)}">${escapeHtml(label)}</option>`;
    })
    .join('');

  elements.purchaseQcRecord.innerHTML = '<option value="">Link QC lot (optional)</option>' + qcOptions;
}

function hydrateQcAiOptions() {
  if (!elements.qcAiRecord) return;
  const selectedId = String(elements.qcAiRecord.value || '').trim();
  const options = state.produce
    .slice(0, 300)
    .map((row) => {
      const label = `${row.id} | ${row.farmerName || '-'} | ${row.variety || '-'} | ${row.qcDecision || '-'} | ${dateShort(row.createdAt)}`;
      return `<option value="${escapeHtml(row.id)}">${escapeHtml(label)}</option>`;
    })
    .join('');

  elements.qcAiRecord.innerHTML = '<option value="">Analyze latest QC lots (up to 30)</option>' + options;
  if (selectedId && state.produce.some((row) => row.id === selectedId)) {
    elements.qcAiRecord.value = selectedId;
  }
}

function seedLocalData() {
  const now = new Date().toISOString();

  const farmerA = {
    id: `F-${Date.now()}-01`,
    name: 'Mercy Achieng',
    phone: '254712330001',
    nationalId: '28643197',
    location: 'Muranga',
    preferredLanguage: 'en',
    hectares: 1.9,
    avocadoHectares: 1.2,
    trees: 48,
    notes: 'Group A',
    createdBy: 'local',
    createdAt: now,
    updatedAt: now
  };

  const farmerB = {
    id: `F-${Date.now()}-02`,
    name: 'David Mwangi',
    phone: '254712330002',
    nationalId: '30984522',
    location: 'Nyeri',
    preferredLanguage: 'sw',
    hectares: 2.4,
    avocadoHectares: 1.6,
    trees: 62,
    notes: 'Organic',
    createdBy: 'local',
    createdAt: now,
    updatedAt: now
  };

  state.farmers = [farmerA, farmerB];
  state.produce = [
    {
      id: `P-${Date.now()}-01`,
      farmerId: farmerA.id,
      farmerName: farmerA.name,
      kgs: 320.4,
      lotWeightKgs: 320.4,
      variety: 'Hass',
      sampleSize: 18,
      visualGrade: 'Pass',
      dryMatterPct: 24.1,
      firmnessValue: 78.5,
      firmnessUnit: 'N',
      avgFruitWeightG: 248.6,
      sizeCode: 'C20',
      qcDecision: 'Accept',
      inspector: 'Agent Njoroge',
      quality: 'Pass',
      agent: 'Agent Njoroge',
      notes: 'Seed dataset',
      createdBy: 'local',
      createdAt: now
    }
  ];
  state.producePurchases = [
    {
      id: `PR-${Date.now()}-01`,
      farmerId: farmerA.id,
      farmerName: farmerA.name,
      qcRecordId: state.produce[0].id,
      variety: 'Hass',
      sizeCode: 'C20',
      purchasedKgs: 302.5,
      pricePerKgKes: 152.5,
      purchaseValueKes: Number((302.5 * 152.5).toFixed(2)),
      buyer: 'Agent Njoroge',
      notes: 'Accepted lot from farm-gate QC',
      createdBy: 'local',
      createdAt: now
    }
  ];
  state.payments = [
    {
      id: `TX-${Date.now()}-01`,
      farmerId: farmerA.id,
      farmerName: farmerA.name,
      amount: 45800,
      ref: 'MPSEED001',
      status: 'Received',
      method: 'M-PESA(Mock)',
      notes: 'Seed payment',
      createdBy: 'local',
      createdAt: now,
      updatedAt: now
    }
  ];
  state.paymentRecommendations = [];
  state.smsLogs = [];
  state.summary = deriveSummaryFromState();
  state.agentStats = [];

  hydrateFarmerSelectors();
  renderAll();
  void refreshOwedRows(true);
  persist();
}

function deriveSummaryFromState() {
  if (!API.enabled) {
    reconcileLocalPurchases();
  }

  const totalProduceKg = state.produce.reduce((sum, row) => sum + Number(row.kgs || 0), 0);
  const totalPurchasedKg = state.producePurchases.reduce((sum, row) => sum + Number(row.purchasedKgs || 0), 0);
  const purchasedValueKes = state.producePurchases.reduce((sum, row) => sum + Number(row.purchaseValueKes || 0), 0);
  const totalOwedKes = state.producePurchases.reduce((sum, row) => sum + Number(row.balanceKes || 0), 0);
  const owedFarmers = new Set(
    state.producePurchases
      .filter((row) => Number(row.balanceKes || 0) > 0)
      .map((row) => row.farmerId)
  ).size;
  const paymentsReceived = state.payments
    .filter((row) => row.status === 'Received')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);

  const paymentSuccessRate = state.payments.length
    ? Math.round((state.payments.filter((row) => row.status === 'Received').length / state.payments.length) * 100)
    : 0;
  const smsCostSummary = smsCostSummaryFromLogs(state.smsLogs, SMS_OWNER_COST_PER_MESSAGE_KES);

  return {
    farmers: state.farmers.length,
    qcRecords: state.produce.length,
    produceRecords: state.produce.length,
    purchasedRecords: state.producePurchases.length,
    paymentRecords: state.payments.length,
    smsSent: state.smsLogs.length,
    totalProduceKg,
    totalPurchasedKg,
    purchasedValueKes,
    totalOwedKes,
    owedFarmers,
    paymentsReceived,
    paymentSuccessRate,
    smsOwnerCostPerMessageKes: SMS_OWNER_COST_PER_MESSAGE_KES,
    smsSentLast24h: smsCostSummary.smsSentLast24h,
    smsSpentKes: smsCostSummary.smsSpentKes,
    smsSpentLast24hKes: smsCostSummary.smsSpentLast24hKes,
    launchReady:
      state.farmers.length > 0 &&
      state.produce.length > 0 &&
      state.payments.some((row) => row.status === 'Received')
  };
}

function farmerNameById(id) {
  const farmer = state.farmers.find((row) => row.id === id);
  return farmer ? farmer.name : 'Unknown';
}

function currentRole() {
  return isAuthenticated() ? state.auth.user.role : state.role;
}

function isAuthenticated() {
  return Boolean(state.auth && state.auth.token && state.auth.user);
}

function updateCredentialUiForRole() {
  const isFarmer = isAuthenticated() && state.auth.user.role === 'farmer';

  if (isFarmer) {
    elements.currentPassword.placeholder = 'Current PIN';
    elements.newPassword.placeholder = 'New 4-digit PIN';
    elements.confirmNewPassword.placeholder = 'Confirm new 4-digit PIN';
    if (elements.changePasswordBtn) {
      elements.changePasswordBtn.textContent = 'Change PIN';
    }
    return;
  }

  elements.currentPassword.placeholder = 'Current password';
  elements.newPassword.placeholder = 'New password (min 10 chars)';
  elements.confirmNewPassword.placeholder = 'Confirm new password';
  if (elements.changePasswordBtn) {
    elements.changePasswordBtn.textContent = 'Change Password';
  }
}

function currentUserPhotoKey() {
  if (!isAuthenticated()) return '';
  return String(state.auth.user.username || '')
    .trim()
    .toLowerCase();
}

function syncCurrentUserPhotoFromState() {
  if (!isAuthenticated()) return;
  const key = currentUserPhotoKey();
  if (!key) return;
  const stored = state.profilePhotos?.[key] || '';
  state.auth.user.photo = stored;
}

function setCurrentUserPhoto(dataUrl) {
  if (!isAuthenticated()) return;
  const key = currentUserPhotoKey();
  if (!key) return;
  if (!state.profilePhotos || typeof state.profilePhotos !== 'object') {
    state.profilePhotos = {};
  }

  if (dataUrl) {
    state.profilePhotos[key] = dataUrl;
    state.auth.user.photo = dataUrl;
  } else {
    delete state.profilePhotos[key];
    state.auth.user.photo = '';
  }

  persist();
}

function renderDashboardAccount() {
  if (!isAuthenticated()) {
    elements.dashboardWelcome.textContent = 'Welcome back.';
    elements.profilePhoto.hidden = true;
    elements.profilePhoto.removeAttribute('src');
    elements.profilePhotoFallback.hidden = false;
    elements.profilePhotoFallback.textContent = 'A';
    elements.clearPhotoBtn.disabled = true;
    return;
  }

  const name = state.auth.user.name || state.auth.user.username || 'User';
  elements.dashboardWelcome.textContent = `Welcome back, ${name}.`;

  const firstLetter = String(name).trim().charAt(0).toUpperCase() || 'A';
  const photo = state.auth.user.photo || '';
  if (photo) {
    elements.profilePhoto.src = photo;
    elements.profilePhoto.hidden = false;
    elements.profilePhotoFallback.hidden = true;
    elements.clearPhotoBtn.disabled = false;
  } else {
    elements.profilePhoto.hidden = true;
    elements.profilePhoto.removeAttribute('src');
    elements.profilePhotoFallback.hidden = false;
    elements.profilePhotoFallback.textContent = firstLetter;
    elements.clearPhotoBtn.disabled = true;
  }
}

async function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function notifySync(message) {
  elements.syncStatus.textContent = message;
}

function clearMessages() {
  if (elements.msaidiziMsg) elements.msaidiziMsg.textContent = '';
  elements.authMsg.textContent = '';
  elements.registerMsg.textContent = '';
  elements.changePasswordMsg.textContent = '';
  elements.recoveryMsg.textContent = '';
  elements.farmerMsg.textContent = '';
  elements.farmerImportMsg.textContent = '';
  elements.farmerImportSummary.textContent = '';
  if (elements.farmerImportQaMsg) elements.farmerImportQaMsg.textContent = '';
  if (elements.farmerImportQaWrap) {
    elements.farmerImportQaWrap.hidden = true;
    elements.farmerImportQaWrap.innerHTML = '';
  }
  elements.farmerImportErrors.innerHTML = '';
  elements.farmerImportErrors.classList.remove('import-errors');
  elements.farmerPinMsg.textContent = '';
  if (elements.farmerPinCredentials) {
    elements.farmerPinCredentials.hidden = true;
    elements.farmerPinCredentials.textContent = '';
  }
  elements.agentMsg.textContent = '';
  elements.produceMsg.textContent = '';
  elements.purchaseMsg.textContent = '';
  elements.owedMsg.textContent = '';
  if (elements.paymentRecommendationMsg) elements.paymentRecommendationMsg.textContent = '';
  elements.paymentMsg.textContent = '';
  elements.smsMsg.textContent = '';
  if (elements.smsDraftMsg) elements.smsDraftMsg.textContent = '';
  if (elements.smsDraftWrap) {
    elements.smsDraftWrap.hidden = true;
    elements.smsDraftWrap.innerHTML = '';
  }
  aiState.smsDrafts = [];
  if (elements.copilotMsg) elements.copilotMsg.textContent = '';
  if (elements.copilotWrap) {
    elements.copilotWrap.classList.add('empty');
    elements.copilotWrap.textContent = 'Ask a question to get AI-guided operational insights.';
  }
  if (elements.qcAiMsg) elements.qcAiMsg.textContent = '';
  if (elements.qcAiWrap) {
    elements.qcAiWrap.classList.add('empty');
    elements.qcAiWrap.textContent = 'Run analysis to get lot risk levels and pass/hold/reject guidance.';
  }
  if (elements.paymentRiskMsg) elements.paymentRiskMsg.textContent = '';
  if (elements.paymentRiskWrap) {
    elements.paymentRiskWrap.classList.add('empty');
    elements.paymentRiskWrap.textContent = 'Run analysis to detect duplicate refs, outliers, and payout risks.';
  }
  if (elements.briefMsg) elements.briefMsg.textContent = '';
  if (elements.proposalMsg) elements.proposalMsg.textContent = '';
  if (elements.opsTaskMsg) elements.opsTaskMsg.textContent = '';
  if (elements.knowledgeMsg) elements.knowledgeMsg.textContent = '';
  if (elements.aiFeedbackMsg) elements.aiFeedbackMsg.textContent = '';
  if (elements.briefWrap) {
    elements.briefWrap.classList.add('empty');
    elements.briefWrap.textContent = 'Run or refresh executive briefs to view strategic summaries and alerts.';
  }
  if (elements.proposalWrap) {
    elements.proposalWrap.classList.add('empty');
    elements.proposalWrap.textContent = 'No proposals loaded yet.';
  }
  if (elements.opsTaskWrap) {
    elements.opsTaskWrap.classList.add('empty');
    elements.opsTaskWrap.textContent = 'No ops tasks loaded yet.';
  }
  if (elements.knowledgeWrap) {
    elements.knowledgeWrap.classList.add('empty');
    elements.knowledgeWrap.textContent = 'No knowledge documents loaded yet.';
  }
  if (elements.aiFeedbackWrap) {
    elements.aiFeedbackWrap.classList.add('empty');
    elements.aiFeedbackWrap.textContent = 'Feedback summary and eval runs will appear here.';
  }
  aiState.qcInsights = [];
  aiState.paymentRiskFlags = [];
  elements.exportsMsg.textContent = '';
}

function formatCurrency(value) {
  return Number(value || 0).toLocaleString('en-US');
}

function formatDate(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString('en-US');
}

function formatKesWithCents(value) {
  const num = Number(value || 0);
  if (!Number.isFinite(num)) return '0.00';
  return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function isBillableSmsStatusClient(status) {
  const normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return true;
  return !['failed', 'rejected', 'cancelled', 'undelivered'].includes(normalized);
}

function smsOwnerCostKesClient(log, fallbackCostPerMessageKes = SMS_OWNER_COST_PER_MESSAGE_KES) {
  if (!isBillableSmsStatusClient(log?.status)) return 0;

  const explicitOwnerCost = Number(log?.ownerCostKes);
  if (Number.isFinite(explicitOwnerCost) && explicitOwnerCost >= 0) return Number(explicitOwnerCost.toFixed(2));

  const legacyCost = Number(log?.costKes);
  if (Number.isFinite(legacyCost) && legacyCost >= 0) return Number(legacyCost.toFixed(2));

  return Number(Number(fallbackCostPerMessageKes || 0).toFixed(2));
}

function smsCostSummaryFromLogs(logs, fallbackCostPerMessageKes = SMS_OWNER_COST_PER_MESSAGE_KES) {
  const rows = Array.isArray(logs) ? logs : [];
  const cutoff = Date.now() - 24 * 60 * 60 * 1000;
  let sentLast24h = 0;
  let smsSpentKes = 0;
  let smsSpentLast24hKes = 0;

  for (const row of rows) {
    const ownerCostKes = smsOwnerCostKesClient(row, fallbackCostPerMessageKes);
    smsSpentKes += ownerCostKes;

    const createdAtMs = new Date(row?.createdAt || '').getTime();
    if (Number.isFinite(createdAtMs) && createdAtMs >= cutoff && ownerCostKes > 0) {
      sentLast24h += 1;
      smsSpentLast24hKes += ownerCostKes;
    }
  }

  return {
    smsSentLast24h: sentLast24h,
    smsSpentKes: Number(smsSpentKes.toFixed(2)),
    smsSpentLast24hKes: Number(smsSpentLast24hKes.toFixed(2))
  };
}

function formatHectares(value) {
  const metrics = areaMetricsFromHectares(value);
  if (!Number.isFinite(metrics.hectares)) return '-';
  return metrics.hectares.toFixed(2);
}

function formatAcres(value) {
  const metrics = areaMetricsFromHectares(value);
  if (!Number.isFinite(metrics.acres)) return '-';
  return metrics.acres.toFixed(2);
}

function formatSquareFeet(value) {
  const metrics = areaMetricsFromHectares(value);
  if (!Number.isFinite(metrics.squareFeet)) return '-';
  return Math.round(metrics.squareFeet).toLocaleString('en-US');
}

function dateShort(value) {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '-';
  return date.toLocaleString();
}

function normalizeLocalFarmerRow(row) {
  if (!row || typeof row !== 'object') return row;
  const normalized = { ...row };
  const hectares = Number(normalized.hectares);
  if (Number.isFinite(hectares) && hectares > 0) {
    normalized.hectares = Number(hectares.toFixed(3));
    const avocadoHectares = Number(normalized.avocadoHectares);
    if (!Number.isFinite(avocadoHectares) || avocadoHectares <= 0 || avocadoHectares > hectares) {
      normalized.avocadoHectares = Number(hectares.toFixed(3));
    } else {
      normalized.avocadoHectares = Number(avocadoHectares.toFixed(3));
    }
  }
  return normalized;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return freshState();

    const parsed = JSON.parse(raw);
    return {
      role: parsed.role || 'admin',
      auth: {
        token: parsed.auth?.token || '',
        user: parsed.auth?.user || null,
        expiresAt: parsed.auth?.expiresAt || ''
      },
      farmers: Array.isArray(parsed.farmers) ? parsed.farmers.map(normalizeLocalFarmerRow) : [],
      agents: Array.isArray(parsed.agents) ? parsed.agents : [],
      produce: Array.isArray(parsed.produce) ? parsed.produce : [],
      producePurchases: Array.isArray(parsed.producePurchases) ? parsed.producePurchases : [],
      payments: Array.isArray(parsed.payments) ? parsed.payments : [],
      paymentRecommendations: Array.isArray(parsed.paymentRecommendations) ? parsed.paymentRecommendations : [],
      smsLogs: Array.isArray(parsed.smsLogs) ? parsed.smsLogs : [],
      summary: parsed.summary || null,
      agentStats: Array.isArray(parsed.agentStats) ? parsed.agentStats : [],
      backups: Array.isArray(parsed.backups) ? parsed.backups : [],
      profilePhotos: parsed.profilePhotos && typeof parsed.profilePhotos === 'object' ? parsed.profilePhotos : {},
      editingFarmerId: ''
    };
  } catch {
    return freshState();
  }
}

function freshState() {
  return {
    role: 'admin',
    auth: { token: '', user: null, expiresAt: '' },
    farmers: [],
    agents: [],
    produce: [],
    producePurchases: [],
    payments: [],
    paymentRecommendations: [],
    smsLogs: [],
    summary: null,
    agentStats: [],
    backups: [],
    profilePhotos: {},
    editingFarmerId: ''
  };
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
