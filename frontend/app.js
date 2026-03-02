const STORAGE_KEY = 'agem_platform_state_v1';

const API = {
  enabled: false
};

const state = loadState();

const elements = {
  tabs: document.querySelectorAll('#tabs button'),
  panes: document.querySelectorAll('.pane'),

  roleHint: document.getElementById('roleHint'),
  roleSelect: document.getElementById('roleSelect'),

  authState: document.getElementById('authState'),
  loginForm: document.getElementById('loginForm'),
  loginUsername: document.getElementById('loginUsername'),
  loginPassword: document.getElementById('loginPassword'),
  registrationPanel: document.getElementById('registrationPanel'),
  registerForm: document.getElementById('registerForm'),
  registerName: document.getElementById('registerName'),
  registerPhone: document.getElementById('registerPhone'),
  registerLocation: document.getElementById('registerLocation'),
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
  farmerLocation: document.getElementById('farmerLocation'),
  treeCount: document.getElementById('treeCount'),
  farmerNotes: document.getElementById('farmerNotes'),
  farmerSubmitBtn: document.getElementById('farmerSubmitBtn'),
  farmerCancelEditBtn: document.getElementById('farmerCancelEditBtn'),
  farmerMsg: document.getElementById('farmerMsg'),
  farmerTableWrap: document.getElementById('farmerTableWrap'),

  produceForm: document.getElementById('produceForm'),
  produceFarmer: document.getElementById('produceFarmer'),
  produceKgs: document.getElementById('produceKgs'),
  quality: document.getElementById('quality'),
  collector: document.getElementById('collector'),
  produceNotes: document.getElementById('produceNotes'),
  produceMsg: document.getElementById('produceMsg'),
  produceTableWrap: document.getElementById('produceTableWrap'),

  paymentForm: document.getElementById('paymentForm'),
  paymentFarmer: document.getElementById('paymentFarmer'),
  amount: document.getElementById('amount'),
  mpesaRef: document.getElementById('mpesaRef'),
  paymentStatus: document.getElementById('paymentStatus'),
  paymentNotes: document.getElementById('paymentNotes'),
  mpesaDisburseBtn: document.getElementById('mpesaDisburseBtn'),
  paymentMsg: document.getElementById('paymentMsg'),
  paymentTableWrap: document.getElementById('paymentTableWrap'),

  smsForm: document.getElementById('smsForm'),
  smsFarmer: document.getElementById('smsFarmer'),
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
  bindAuth();
  bindFarmers();
  bindProduce();
  bindPayments();
  bindSms();
  bindExports();

  hydrateFarmerSelectors();
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

  updatePermissionUi();
}

function bindTabs() {
  elements.tabs.forEach((button) => {
    button.addEventListener('click', () => {
      elements.tabs.forEach((item) => item.classList.remove('active'));
      elements.panes.forEach((pane) => pane.classList.remove('active'));
      button.classList.add('active');
      const pane = document.getElementById(button.dataset.pane);
      if (pane) pane.classList.add('active');
    });
  });
}

function bindRoleSelect() {
  elements.roleSelect.value = state.role;
  elements.roleSelect.addEventListener('change', (event) => {
    state.role = event.target.value;
    persist();
    renderAll();
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

      elements.loginForm.reset();
      elements.authMsg.textContent = `Signed in as ${response.data.user.username}.`;
      persist();

      updateAuthUi();
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
    const location = elements.registerLocation.value.trim();
    const username = elements.registerUsername.value.trim();
    const password = elements.registerPassword.value;
    const confirmPassword = elements.registerConfirmPassword.value;

    if (!name || !phone || !location || !username || !password || !confirmPassword) {
      elements.registerMsg.textContent = 'All registration fields are required.';
      return;
    }

    try {
      const response = await apiRequest('/api/auth/register-farmer', {
        method: 'POST',
        auth: false,
        body: {
          name,
          phone,
          location,
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
      persist();

      elements.registerForm.reset();
      updateAuthUi();
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

    const payload = {
      name: elements.farmerName.value.trim(),
      phone: elements.farmerPhone.value.trim(),
      location: elements.farmerLocation.value.trim(),
      trees: Number(elements.treeCount.value || 0),
      notes: elements.farmerNotes.value.trim()
    };

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
      elements.farmerLocation.value = farmer.location || '';
      elements.treeCount.value = farmer.trees || 0;
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
      quality: elements.quality.value.trim(),
      agent: elements.collector.value.trim(),
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

function bindPayments() {
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
        return;
      }

      const payment = state.payments.find((row) => row.id === paymentId);
      if (payment) {
        payment.status = nextStatus;
        payment.updatedAt = new Date().toISOString();
      }
      renderAll();
      persist();
      elements.paymentMsg.textContent = `Payment status changed to ${nextStatus} (local mode).`;
    } catch (error) {
      elements.paymentMsg.textContent = error.message;
    }
  });
}

function bindSms() {
  elements.smsFarmer.addEventListener('change', () => {
    if (!elements.smsFarmer.value) return;
    const farmer = state.farmers.find((row) => row.id === elements.smsFarmer.value);
    if (farmer) {
      elements.smsPhone.value = farmer.phone || '';
    }
  });

  elements.smsForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearMessages();

    if (!['admin', 'agent'].includes(currentRole())) {
      elements.smsMsg.textContent = 'Only admin or agent can send SMS.';
      return;
    }

    const payload = {
      farmerId: elements.smsFarmer.value || '',
      phone: elements.smsPhone.value.trim(),
      message: elements.smsMessage.value.trim()
    };

    try {
      if (API.enabled) {
        if (!isAuthenticated()) {
          elements.smsMsg.textContent = 'Sign in first to use backend mode.';
          return;
        }

        await apiRequest('/api/sms/send', {
          method: 'POST',
          body: payload
        });

        elements.smsForm.reset();
        elements.smsMsg.textContent = 'SMS sent.';
        await fetchAllData();
        return;
      }

      if (!payload.message) {
        elements.smsMsg.textContent = 'Message is required.';
        return;
      }

      let phone = payload.phone;
      let farmerName = '';
      if (payload.farmerId) {
        const farmer = state.farmers.find((row) => row.id === payload.farmerId);
        if (farmer) {
          phone = farmer.phone;
          farmerName = farmer.name;
        }
      }

      if (!phone) {
        elements.smsMsg.textContent = 'Provide phone or select farmer.';
        return;
      }

      state.smsLogs.unshift({
        id: `SMS-${Date.now()}`,
        farmerId: payload.farmerId,
        farmerName,
        phone,
        message: payload.message,
        provider: 'Local Mock',
        status: 'Sent',
        createdBy: 'local',
        createdAt: new Date().toISOString()
      });

      elements.smsForm.reset();
      elements.smsMsg.textContent = 'SMS sent (local mode).';
      renderAll();
      persist();
    } catch (error) {
      elements.smsMsg.textContent = error.message;
    }
  });
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

    if (!confirm('This will clear all farmers, produce, payments, and SMS data. Continue?')) return;

    if (!API.enabled) {
      state.farmers = [];
      state.produce = [];
      state.payments = [];
      state.smsLogs = [];
      state.summary = null;
      state.agentStats = [];
      state.backups = [];
      resetFarmerForm();
      hydrateFarmerSelectors();
      renderAll();
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

    if (state.farmers.length || state.produce.length || state.payments.length) {
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
    const [farmers, produce, payments, sms, summary, agents] = await Promise.all([
      apiRequest('/api/farmers'),
      apiRequest('/api/produce'),
      apiRequest('/api/payments'),
      apiRequest('/api/sms'),
      apiRequest('/api/reports/summary'),
      apiRequest('/api/reports/agents')
    ]);

    state.farmers = farmers.data || [];
    state.produce = produce.data || [];
    state.payments = payments.data || [];
    state.smsLogs = sms.data || [];
    state.summary = summary.data || null;
    state.agentStats = agents.data || [];

    hydrateFarmerSelectors();
    renderAll();
    updatePermissionUi();
    persist();

    notifySync(`Connected to backend as ${currentRole()}.`);
  } catch (error) {
    notifySync(error.message);
  }
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
    elements.authState.textContent = `Signed in: ${state.auth.user.username} (${state.auth.user.role})`;
    elements.loginForm.hidden = true;
    elements.loginForm.style.display = 'none';
    elements.registrationPanel.hidden = true;
    elements.changePasswordForm.hidden = false;
    elements.changePasswordForm.style.display = 'grid';
    elements.logoutBtn.hidden = false;
    elements.logoutBtn.style.display = 'block';
    elements.recoveryPanel.hidden = true;
    elements.roleSelect.disabled = true;
    elements.roleSelect.value = state.auth.user.role;
    state.role = state.auth.user.role;
  } else {
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

  updateRoleHint();
}

function updatePermissionUi() {
  const role = currentRole();
  const backendAuthReady = !API.enabled || isAuthenticated();

  const farmersAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const produceAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const paymentAllowed = backendAuthReady && role === 'admin';
  const smsAllowed = backendAuthReady && ['admin', 'agent'].includes(role);
  const adminAllowed = backendAuthReady && role === 'admin';

  elements.farmerSubmitBtn.disabled = !farmersAllowed;
  elements.farmerCancelEditBtn.disabled = !farmersAllowed;

  elements.produceForm.querySelector('button[type="submit"]').disabled = !produceAllowed;

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
  renderOverview();
  renderFarmers();
  renderProduce();
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
    { label: 'Produce Records', value: summary.produceRecords || 0 },
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
          <td>${escapeHtml(farmer.location)}</td>
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
          <th>Location</th>
          <th>Trees</th>
          <th>Updated</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderProduce() {
  if (!state.produce.length) {
    elements.produceTableWrap.innerHTML = '<div class="empty">No produce entries yet.</div>';
    return;
  }

  const canDelete = currentRole() === 'admin';

  const rows = state.produce
    .slice(0, 250)
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.id)}</td>
        <td>${escapeHtml(row.farmerName)}</td>
        <td>${escapeHtml(String(row.kgs))}</td>
        <td>${escapeHtml(row.quality)}</td>
        <td>${escapeHtml(row.agent)}</td>
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
          <th>Kg</th>
          <th>Quality</th>
          <th>Agent</th>
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
    { label: 'Produce (kg)', value: Number(summary.totalProduceKg || 0).toFixed(1) },
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
    `Current operations show ${summary.farmers || 0} farmers, ${summary.produceRecords || 0} produce entries, ` +
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
          <th>Produce Kg</th>
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
  elements.paymentFarmer.innerHTML = options || fallback;
  elements.smsFarmer.innerHTML = `<option value="">Select farmer (optional)</option>${options}`;
}

function seedLocalData() {
  const now = new Date().toISOString();

  const farmerA = {
    id: `F-${Date.now()}-01`,
    name: 'Mercy Achieng',
    phone: '254712330001',
    location: 'Muranga',
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
    location: 'Nyeri',
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
      quality: 'A',
      agent: 'Agent Njoroge',
      notes: 'Seed dataset',
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
  persist();
}

function deriveSummaryFromState() {
  const totalProduceKg = state.produce.reduce((sum, row) => sum + Number(row.kgs || 0), 0);
  const paymentsReceived = state.payments
    .filter((row) => row.status === 'Received')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);

  const paymentSuccessRate = state.payments.length
    ? Math.round((state.payments.filter((row) => row.status === 'Received').length / state.payments.length) * 100)
    : 0;

  return {
    farmers: state.farmers.length,
    produceRecords: state.produce.length,
    paymentRecords: state.payments.length,
    smsSent: state.smsLogs.length,
    totalProduceKg,
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

function notifySync(message) {
  elements.syncStatus.textContent = message;
}

function clearMessages() {
  elements.authMsg.textContent = '';
  elements.registerMsg.textContent = '';
  elements.changePasswordMsg.textContent = '';
  elements.recoveryMsg.textContent = '';
  elements.farmerMsg.textContent = '';
  elements.produceMsg.textContent = '';
  elements.paymentMsg.textContent = '';
  elements.smsMsg.textContent = '';
  elements.exportsMsg.textContent = '';
}

function formatCurrency(value) {
  return Number(value || 0).toLocaleString('en-US');
}

function dateShort(value) {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '-';
  return date.toLocaleString();
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
      farmers: Array.isArray(parsed.farmers) ? parsed.farmers : [],
      produce: Array.isArray(parsed.produce) ? parsed.produce : [],
      payments: Array.isArray(parsed.payments) ? parsed.payments : [],
      smsLogs: Array.isArray(parsed.smsLogs) ? parsed.smsLogs : [],
      summary: parsed.summary || null,
      agentStats: Array.isArray(parsed.agentStats) ? parsed.agentStats : [],
      backups: Array.isArray(parsed.backups) ? parsed.backups : [],
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
    produce: [],
    payments: [],
    smsLogs: [],
    summary: null,
    agentStats: [],
    backups: [],
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
