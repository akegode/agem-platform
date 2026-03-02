const STORAGE_KEY = 'agem_platform_state_v1';
const MAX_PROFILE_PHOTO_BYTES = 1_500_000;
const HECTARES_PER_ACRE = 0.40468564224;
const ACRES_PER_HECTARE = 1 / HECTARES_PER_ACRE;
const SQFT_PER_HECTARE = 107639.1041671;

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
  registerAreaHectares: document.getElementById('registerAreaHectares'),
  registerAreaAcres: document.getElementById('registerAreaAcres'),
  registerAreaSquareFeet: document.getElementById('registerAreaSquareFeet'),
  registerAvocadoAreaHectares: document.getElementById('registerAvocadoAreaHectares'),
  registerAvocadoAreaAcres: document.getElementById('registerAvocadoAreaAcres'),
  registerAvocadoAreaSquareFeet: document.getElementById('registerAvocadoAreaSquareFeet'),
  registerUsername: document.getElementById('registerUsername'),
  registerPassword: document.getElementById('registerPassword'),
  registerConfirmPassword: document.getElementById('registerConfirmPassword'),
  registerMsg: document.getElementById('registerMsg'),
  changePasswordForm: document.getElementById('changePasswordForm'),
  currentPassword: document.getElementById('currentPassword'),
  newPassword: document.getElementById('newPassword'),
  confirmNewPassword: document.getElementById('confirmNewPassword'),
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
  farmerImportFile: document.getElementById('farmerImportFile'),
  farmerImportBtn: document.getElementById('farmerImportBtn'),
  farmerImportMsg: document.getElementById('farmerImportMsg'),
  farmerImportSummary: document.getElementById('farmerImportSummary'),
  farmerImportErrors: document.getElementById('farmerImportErrors'),
  farmerTableWrap: document.getElementById('farmerTableWrap'),
  agentForm: document.getElementById('agentForm'),
  agentName: document.getElementById('agentName'),
  agentEmail: document.getElementById('agentEmail'),
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
  owedSummary: document.getElementById('owedSummary'),
  owedTableWrap: document.getElementById('owedTableWrap'),
  owedMsg: document.getElementById('owedMsg'),
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
  smsMsg: document.getElementById('smsMsg'),
  smsTableWrap: document.getElementById('smsTableWrap'),

  reportMetrics: document.getElementById('reportMetrics'),
  reportNarrative: document.getElementById('reportNarrative'),
  agentStatsWrap: document.getElementById('agentStatsWrap'),

  exportButtons: document.querySelectorAll('.export-btn'),
  backupBtn: document.getElementById('backupBtn'),
  listBackupsBtn: document.getElementById('listBackupsBtn'),
  resetBtn: document.getElementById('resetBtn'),
  exportsMsg: document.getElementById('exportsMsg'),
  backupListWrap: document.getElementById('backupListWrap')
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

  if (elements.paneSelect) {
    elements.paneSelect.value = paneId;
  }

  if (elements.menuPanel) {
    elements.menuPanel.open = false;
  }

  if (resetScroll && elements.dashboardMain) {
    elements.dashboardMain.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

function bindRoleSelect() {
  elements.roleSelect.value = state.role;
  elements.roleSelect.addEventListener('change', async (event) => {
    state.role = event.target.value;
    persist();
    renderAll();
    await refreshOwedRows(true);
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

    const username = elements.loginUsername.value.trim();
    const password = elements.loginPassword.value;

    if (!username || !password) {
      elements.authMsg.textContent = 'Username and password are required.';
      return;
    }

    try {
      const response = await apiRequest('/api/auth/login', {
        method: 'POST',
        body: { username, password },
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
    const phone = elements.registerPhone.value.trim();
    const nationalId = elements.registerNationalId.value.trim();
    const location = elements.registerLocation.value.trim();
    const areaHectares = elements.registerAreaHectares.value.trim();
    const areaAcres = elements.registerAreaAcres.value.trim();
    const areaSquareFeet = elements.registerAreaSquareFeet.value.trim();
    const avocadoAreaHectares = elements.registerAvocadoAreaHectares.value.trim();
    const avocadoAreaAcres = elements.registerAvocadoAreaAcres.value.trim();
    const avocadoAreaSquareFeet = elements.registerAvocadoAreaSquareFeet.value.trim();
    const username = elements.registerUsername.value.trim();
    const password = elements.registerPassword.value;
    const confirmPassword = elements.registerConfirmPassword.value;

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
      !username ||
      !password ||
      !confirmPassword
    ) {
      elements.registerMsg.textContent = 'All registration fields are required.';
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
          ...totalAreaPayload,
          avocadoHectares: avocadoAreaPayload.hectares,
          avocadoAcres: avocadoAreaPayload.acres,
          avocadoSquareFeet: avocadoAreaPayload.squareFeet,
          username,
          password,
          confirmPassword
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

    if (!currentPassword || !newPassword || !confirmPassword) {
      elements.changePasswordMsg.textContent = 'All password fields are required.';
      return;
    }

    try {
      await apiRequest('/api/auth/change-password', {
        method: 'POST',
        body: { currentPassword, newPassword, confirmPassword }
      });

      elements.changePasswordForm.reset();
      elements.changePasswordMsg.textContent = 'Password updated successfully.';
    } catch (error) {
      elements.changePasswordMsg.textContent = error.message;
    }
  });

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
      phone: elements.farmerPhone.value.trim(),
      nationalId: cleanNationalIdClient(elements.farmerNationalId.value),
      location: elements.farmerLocation.value.trim(),
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

      let result = null;
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.farmerImportMsg.textContent = 'Sign in first to import farmers.';
          return;
        }

        const response = await apiRequest('/api/farmers/import', {
          method: 'POST',
          body: { records, onDuplicate: duplicateMode }
        });
        result = response.data || {};
        await fetchAllData();
      } else {
        result = importFarmersLocal(records, { onDuplicate: duplicateMode });
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

    if (!name || !email) {
      elements.agentMsg.textContent = 'Name and email are required.';
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      elements.agentMsg.textContent = 'Enter a valid email address.';
      return;
    }

    try {
      const response = await apiRequest('/api/agents', {
        method: 'POST',
        body: { name, email }
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
  if (!username || !temporaryPassword) {
    elements.agentCredentials.hidden = true;
    elements.agentCredentials.textContent = '';
    return;
  }

  elements.agentCredentials.hidden = false;
  elements.agentCredentials.innerHTML = `
    <strong>Temporary login created.</strong><br>
    Email: <code>${escapeHtml(email || '-')}</code><br>
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

async function refreshOwedRows(resetSelection = false) {
  if (currentRole() !== 'admin') {
    owedPicker.rows = [];
    owedPicker.selectedFarmerIds.clear();
    elements.owedSummary.textContent = 'Owed list is visible to admin only.';
    elements.owedTableWrap.innerHTML = '<div class="empty">Sign in as admin to view owed balances.</div>';
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

    const blob = new Blob([content], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `payments-owed-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    elements.owedMsg.textContent = 'Owed balances CSV exported.';
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
        const text = `${row.id || ''} ${row.name || ''} ${row.phone || ''} ${row.nationalId || ''} ${row.location || ''} ${row.notes || ''}`.toLowerCase();
        return text.includes(needle);
      });
    }

    smsPicker.total = rows.length;
    smsPicker.rows = rows.slice(smsPicker.offset, smsPicker.offset + smsPicker.limit).map((row) => ({
      id: row.id,
      name: row.name,
      phone: row.phone,
      nationalId: row.nationalId,
      location: row.location
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
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  updateSmsComposerUi();
}

function bindExports() {
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
        const content = await apiRequest(`/api/exports/${encodeURIComponent(type)}.csv`, {
          method: 'GET',
          response: 'text'
        });

        const blob = new Blob([content], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${type}-${new Date().toISOString().slice(0, 10)}.csv`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);

        elements.exportsMsg.textContent = `Exported ${type}.csv`;
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

async function fetchAllData() {
  if (!API.enabled || !isAuthenticated()) {
    renderAll();
    updatePermissionUi();
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

    let managedAgents = [];
    if (currentRole() === 'admin') {
      const agentAccounts = await apiRequest('/api/agents?includeDisabled=true&limit=500');
      managedAgents = Array.isArray(agentAccounts.data) ? agentAccounts.data : [];
    }

    state.farmers = farmers.data || [];
    state.produce = produce.data || [];
    state.producePurchases = producePurchases.data || [];
    state.payments = payments.data || [];
    state.smsLogs = sms.data || [];
    state.summary = summary.data || null;
    state.agentStats = agentStats.data || [];
    state.agents = managedAgents;

    hydrateFarmerSelectors();
    renderAll();
    updatePermissionUi();
    persist();

    notifySync(`Connected to backend as ${currentRole()}.`);
    await refreshOwedRows(false);
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

  renderDashboardAccount();
  updateRoleHint();
}

function updatePermissionUi() {
  const role = currentRole();
  const backendAuthReady = !API.enabled || isAuthenticated();

  const farmersAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const produceAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const producePurchaseAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const owedAllowed = backendAuthReady && role === 'admin';
  const paymentAllowed = backendAuthReady && role === 'admin';
  const smsAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const adminAllowed = backendAuthReady && role === 'admin';
  const importAllowed = backendAuthReady && role === 'admin';
  const agentManageAllowed = API.enabled && isAuthenticated() && role === 'admin';

  elements.farmerSubmitBtn.disabled = !farmersAllowed;
  elements.farmerCancelEditBtn.disabled = !farmersAllowed;
  elements.farmerImportBtn.disabled = !importAllowed;
  elements.farmerImportFile.disabled = !importAllowed;
  elements.agentName.disabled = !agentManageAllowed;
  elements.agentEmail.disabled = !agentManageAllowed;
  elements.agentRefreshBtn.disabled = !agentManageAllowed;
  elements.agentSearch.disabled = !agentManageAllowed;
  elements.agentForm.querySelector('button[type="submit"]').disabled = !agentManageAllowed;

  if (elements.agentsNavBtn) {
    elements.agentsNavBtn.hidden = !agentManageAllowed;
  }
  if (elements.agentsPaneOption) {
    elements.agentsPaneOption.hidden = !agentManageAllowed;
  }
  if (!agentManageAllowed && elements.agentsPane?.classList.contains('active')) {
    setActivePane('overview');
  }
  if (!agentManageAllowed) {
    clearAgentCredentialDisplay();
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

  elements.paymentForm.querySelector('button[type="submit"]').disabled = !paymentAllowed;
  elements.mpesaDisburseBtn.disabled = !paymentAllowed;

  elements.smsForm.querySelector('button[type="submit"]').disabled = !smsAllowed;

  elements.seedBtn.disabled = !(adminAllowed || !API.enabled);
  elements.backupBtn.disabled = !adminAllowed;
  elements.listBackupsBtn.disabled = !adminAllowed;
  elements.resetBtn.disabled = !(adminAllowed || !API.enabled);

  elements.exportButtons.forEach((button) => {
    const type = button.dataset.export;
    if (type === 'payments' || type === 'activity') {
      button.disabled = !adminAllowed;
    } else {
      button.disabled = !(backendAuthReady && ['admin', 'agent'].includes(role));
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

  const rows = state.farmers
    .slice(0, 200)
    .map((farmer) => {
      const actions = canEdit
        ? `
          <button class="table-btn" data-action="edit-farmer" data-id="${escapeHtml(farmer.id)}">Edit</button>
          ${canDelete ? `<button class="table-btn danger" data-action="delete-farmer" data-id="${escapeHtml(farmer.id)}">Delete</button>` : ''}
        `
        : '-';

      return `
        <tr>
          <td>${escapeHtml(farmer.id)}</td>
          <td>${escapeHtml(farmer.name)}</td>
          <td>${escapeHtml(farmer.phone)}</td>
          <td>${escapeHtml(farmer.nationalId || '-')}</td>
          <td>${escapeHtml(farmer.location)}</td>
          <td>${escapeHtml(formatHectares(farmer.hectares))}</td>
          <td>${escapeHtml(formatAcres(farmer.hectares))}</td>
          <td>${escapeHtml(formatSquareFeet(farmer.hectares))}</td>
          <td>${escapeHtml(formatHectares(farmer.avocadoHectares))}</td>
          <td>${escapeHtml(formatAcres(farmer.avocadoHectares))}</td>
          <td>${escapeHtml(formatSquareFeet(farmer.avocadoHectares))}</td>
          <td>${escapeHtml(String(farmer.trees ?? '0'))}</td>
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
          <th>Hectares</th>
          <th>Acres</th>
          <th>Square Feet</th>
          <th>Avocado Hectares</th>
          <th>Avocado Acres</th>
          <th>Avocado Sq Ft</th>
          <th>Trees</th>
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
  if (!state.smsLogs.length) {
    elements.smsTableWrap.innerHTML = '<div class="empty">No SMS messages logged yet.</div>';
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
          <th>By</th>
          <th>Date</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
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
    `Data source: ${dataSource}. Session: ${user}.`;

  if (!state.agentStats.length) {
    elements.agentStatsWrap.innerHTML = '<div class="empty">No agent stats yet.</div>';
    return;
  }

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

function renderBackups() {
  if (!state.backups || !state.backups.length) {
    elements.backupListWrap.innerHTML = '<div class="empty">No backup snapshots listed.</div>';
    return;
  }

  const rows = state.backups
    .slice(0, 50)
    .map(
      (item) => `
        <tr>
          <td>${escapeHtml(item.filename)}</td>
          <td>${escapeHtml(String(item.sizeBytes))}</td>
          <td>${escapeHtml(dateShort(item.createdAt))}</td>
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
  const duplicateMode = result?.duplicateMode || 'skip';
  const errors = Array.isArray(result?.errors) ? result.errors : [];

  if (imported || updated) {
    elements.farmerImportMsg.textContent = `Import complete. Added ${imported}, updated ${updated}.`;
  } else {
    elements.farmerImportMsg.textContent = 'Import completed with no new farmers added.';
  }
  elements.farmerImportSummary.textContent =
    `Rows processed: ${totalRows}. Imported: ${imported}. Updated: ${updated}. Skipped: ${skipped}. ` +
    `Duplicate mode: ${duplicateMode}.`;

  if (!errors.length) {
    elements.farmerImportErrors.innerHTML = '';
    elements.farmerImportErrors.classList.remove('import-errors');
    return;
  }

  const maxShown = 25;
  const listItems = errors
    .slice(0, maxShown)
    .map((entry) => `<li>Row ${escapeHtml(String(entry.row || '?'))}: ${escapeHtml(entry.error || 'Invalid row')}</li>`)
    .join('');
  const remaining = errors.length - maxShown;
  const more = remaining > 0 ? `<li>...and ${escapeHtml(String(remaining))} more row errors.</li>` : '';

  elements.farmerImportErrors.classList.add('import-errors');
  elements.farmerImportErrors.innerHTML = `<ul>${listItems}${more}</ul>`;
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

function parseTreeCountClient(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return 0;
  const num = Number(raw.replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function parseAreaClient(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return NaN;
  const num = Number(raw.replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
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
  if (hectaresText) return parseAreaClient(hectaresText);

  const acresText = String(acresInput ?? '').trim();
  if (acresText) {
    const acres = parseAreaClient(acresText);
    if (!Number.isFinite(acres)) return NaN;
    return acres * HECTARES_PER_ACRE;
  }

  const squareFeetText = String(squareFeetInput ?? '').trim();
  if (squareFeetText) {
    const squareFeet = parseAreaClient(squareFeetText);
    if (!Number.isFinite(squareFeet)) return NaN;
    return squareFeet / SQFT_PER_HECTARE;
  }

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

  const name = firstNonEmptyClient([flat.name, flat.farmername, flat.fullname, flat.farmerfullname, flat.growername]);
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
    flat.governmentidnumber
  ]);
  const location = firstNonEmptyClient([flat.location, flat.area, flat.ward, flat.county, flat.village, flat.region]);
  const notes = firstNonEmptyClient([flat.notes, flat.note, flat.comments, flat.remarks, flat.description]);
  const treesRaw = firstNonEmptyClient([flat.trees, flat.treecount, flat.numberoftrees, flat.treequantity, flat.treenumber]);
  const hectaresRaw = firstNonEmptyClient([flat.hectares, flat.hectare, flat.farmsizeha, flat.farmsizehectares, flat.landsizehectares]);
  const acresRaw = firstNonEmptyClient([flat.acres, flat.acre, flat.acreage, flat.farmsizeacres, flat.landsizeacres]);
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
    flat.avocadoplotacres
  ]);
  const avocadoSquareFeetRaw = firstNonEmptyClient([
    flat.avocadosquarefeet,
    flat.avocadosquarefoot,
    flat.avocadosqft,
    flat.areaunderavocadosquarefeet,
    flat.areaunderavocadosqft
  ]);

  return {
    name,
    phone,
    nationalId: cleanNationalIdClient(nationalId),
    location,
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

function cleanPhone(value) {
  return String(value ?? '').trim();
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
      errors.push({ row: idx + 1, error: invalid });
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
            error: 'Conflicting duplicate match: phone and National ID belong to different farmers'
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
        existing.notes = mapped.notes;
        existing.updatedAt = new Date().toISOString();

        farmersByPhone.set(phone, existing);
        farmersByNationalId.set(nationalIdKey, existing);
        updated.push(existing.id);
      } else {
        errors.push({ row: idx + 1, error: 'Duplicate farmer (matching phone or National ID)' });
      }
      return;
    }

    const record = {
      id: `F-${now}-${idx + 1}`,
      name: mapped.name,
      phone,
      nationalId,
      location: mapped.location,
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

  return {
    totalRows: records.length,
    imported: created.length,
    updated: updated.length,
    skipped: records.length - created.length - updated.length,
    duplicateMode: onDuplicate,
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
  elements.farmerSubmitBtn.textContent = 'Register Farmer';
  elements.farmerCancelEditBtn.hidden = true;
}

function hydrateFarmerSelectors() {
  const options = state.farmers
    .map((farmer) => `<option value="${escapeHtml(farmer.id)}">${escapeHtml(farmer.name)} (${escapeHtml(farmer.location)})</option>`)
    .join('');

  const fallback = '<option value="">No farmers yet</option>';
  elements.produceFarmer.innerHTML = options || fallback;
  elements.purchaseFarmer.innerHTML = options || fallback;
  elements.paymentFarmer.innerHTML = options || fallback;
  hydratePurchaseQcOptions();

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

function seedLocalData() {
  const now = new Date().toISOString();

  const farmerA = {
    id: `F-${Date.now()}-01`,
    name: 'Mercy Achieng',
    phone: '254712330001',
    nationalId: '28643197',
    location: 'Muranga',
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
  elements.authMsg.textContent = '';
  elements.registerMsg.textContent = '';
  elements.changePasswordMsg.textContent = '';
  elements.recoveryMsg.textContent = '';
  elements.farmerMsg.textContent = '';
  elements.farmerImportMsg.textContent = '';
  elements.farmerImportSummary.textContent = '';
  elements.farmerImportErrors.innerHTML = '';
  elements.farmerImportErrors.classList.remove('import-errors');
  elements.agentMsg.textContent = '';
  elements.produceMsg.textContent = '';
  elements.purchaseMsg.textContent = '';
  elements.owedMsg.textContent = '';
  elements.paymentMsg.textContent = '';
  elements.smsMsg.textContent = '';
  elements.exportsMsg.textContent = '';
}

function formatCurrency(value) {
  return Number(value || 0).toLocaleString('en-US');
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
