const assert = require('assert');
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServer(baseUrl, timeoutMs = 12000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(`${baseUrl}/api/health`);
      if (res.ok) return;
    } catch {
      // keep waiting
    }
    await sleep(200);
  }
  throw new Error('Server did not become ready in time');
}

async function request(baseUrl, route, options = {}) {
  const headers = {
    ...(options.headers || {})
  };
  if (options.token) headers.Authorization = `Bearer ${options.token}`;
  if (options.body !== undefined && !headers['Content-Type']) headers['Content-Type'] = 'application/json';

  const response = await fetch(`${baseUrl}${route}`, {
    method: options.method || 'GET',
    headers,
    body:
      options.bodyRaw !== undefined
        ? options.bodyRaw
        : options.body !== undefined
          ? JSON.stringify(options.body)
          : undefined
  });

  if (options.expectStatus !== undefined) {
    assert.strictEqual(
      response.status,
      options.expectStatus,
      `Expected ${options.expectStatus} for ${route}, got ${response.status}`
    );
  }

  if (options.responseType === 'text') {
    return response.text();
  }

  let payload = {};
  try {
    payload = await response.json();
  } catch {
    payload = {};
  }
  return payload;
}

async function run() {
  const root = path.resolve(__dirname, '..');
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'agem-test-'));
  const storePath = path.join(tempDir, 'store.json');
  const port = 19000 + Math.floor(Math.random() * 1000);
  const baseUrl = `http://127.0.0.1:${port}`;
  const adminUsername = 'opsadmin';
  const adminPassword = 'OpsAdmin#2026!';
  const adminPasswordUpdated = 'OpsAdmin#2026!X2';
  const adminRecoveryCode = 'AdminRecovery#2026';
  const agentUsername = 'fieldagent';
  const agentPassword = 'FieldAgent#2026!';
  const agentPasswordUpdated = 'FieldAgent#2026!X2';
  const additionalAgentEmail = 'mary.wanjiku@agemlimited.com';
  const additionalAgentPhone = '254711223344';
  const agentRecoveryCode = 'AgentRecovery#2026';
  const farmerSelfPhone = '254700111222';
  const farmerSelfPin = '1728';
  const farmerSelfPinUpdated = '1839';
  const ussdSecret = 'UssdSecret#2026';
  const ussdCode = '*483#';

  const server = spawn('node', ['backend/server.js'], {
    cwd: root,
    env: {
      ...process.env,
      HOST: '127.0.0.1',
      PORT: String(port),
      STORE_PATH: storePath,
      ALLOW_DEMO_USERS: 'false',
      ADMIN_USERNAME: adminUsername,
      ADMIN_PASSWORD: adminPassword,
      AGENT_USERNAME: agentUsername,
      AGENT_PASSWORD: agentPassword,
      ADMIN_RECOVERY_CODE: adminRecoveryCode,
      AGENT_RECOVERY_CODE: agentRecoveryCode,
      USSD_ENABLED: 'true',
      USSD_CODE: ussdCode,
      USSD_HELP_PHONE: '+254700111999',
      USSD_SHARED_SECRET: ussdSecret
    },
    stdio: ['ignore', 'pipe', 'pipe']
  });

  let stderr = '';
  server.stderr.on('data', (chunk) => {
    stderr += chunk.toString();
  });

  try {
    await waitForServer(baseUrl);

    await request(baseUrl, '/api/farmers', { expectStatus: 401 });

    const adminLogin = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: adminUsername, password: adminPassword },
      expectStatus: 200
    });
    let adminToken = adminLogin.data.token;
    assert.ok(adminToken, 'Admin token missing');

    const agentLogin = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: agentUsername, password: agentPassword },
      expectStatus: 200
    });
    let agentToken = agentLogin.data.token;

    const recoveredAdminUsername = await request(baseUrl, '/api/auth/recover-username', {
      method: 'POST',
      body: { role: 'admin', recoveryCode: adminRecoveryCode },
      expectStatus: 200
    });
    assert.strictEqual(recoveredAdminUsername.data.username, adminUsername);

    await request(baseUrl, '/api/auth/change-password', {
      method: 'POST',
      token: adminToken,
      body: {
        currentPassword: adminPassword,
        newPassword: adminPasswordUpdated,
        confirmPassword: adminPasswordUpdated
      },
      expectStatus: 200
    });

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: adminUsername, password: adminPassword },
      expectStatus: 401
    });

    const adminLoginUpdated = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: adminUsername, password: adminPasswordUpdated },
      expectStatus: 200
    });
    adminToken = adminLoginUpdated.data.token;

    const recoverAgentPassword = await request(baseUrl, '/api/auth/recover-password', {
      method: 'POST',
      body: {
        role: 'agent',
        recoveryCode: agentRecoveryCode,
        newPassword: agentPasswordUpdated,
        confirmPassword: agentPasswordUpdated
      },
      expectStatus: 200
    });
    assert.strictEqual(recoverAgentPassword.data.username, agentUsername);

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: agentUsername, password: agentPassword },
      expectStatus: 401
    });

    const agentLoginUpdated = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: agentUsername, password: agentPasswordUpdated },
      expectStatus: 200
    });
    agentToken = agentLoginUpdated.data.token;

    await request(baseUrl, '/api/agents', {
      token: agentToken,
      expectStatus: 403
    });

    const createAgent = await request(baseUrl, '/api/agents', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Mary Wanjiku',
        email: additionalAgentEmail,
        phone: additionalAgentPhone
      },
      expectStatus: 201
    });
    assert.strictEqual(createAgent.data.role, 'agent');
    assert.strictEqual(createAgent.data.email, additionalAgentEmail);
    assert.strictEqual(createAgent.data.phone, additionalAgentPhone);
    assert.ok(/^[a-z0-9._-]{3,32}$/.test(createAgent.data.username), 'Expected generated username format');
    assert.ok(
      typeof createAgent.data.temporaryPassword === 'string' && createAgent.data.temporaryPassword.length >= 10,
      'Expected generated temporary password'
    );

    await request(baseUrl, '/api/agents', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Duplicate Agent',
        email: additionalAgentEmail,
        phone: '254712000000'
      },
      expectStatus: 409
    });

    await request(baseUrl, '/api/agents', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Duplicate Phone Agent',
        email: 'duplicate.phone@agemlimited.com',
        phone: additionalAgentPhone
      },
      expectStatus: 409
    });

    const agentsList = await request(baseUrl, '/api/agents?includeDisabled=true&limit=20', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(agentsList.data), 'Expected list of agents');
    const agentUsernames = new Set((agentsList.data || []).map((row) => row.username));
    const agentEmails = new Set((agentsList.data || []).map((row) => row.email).filter(Boolean));
    const agentPhones = new Set((agentsList.data || []).map((row) => row.phone).filter(Boolean));
    assert.ok(agentUsernames.has(agentUsername), 'Expected env agent in list');
    assert.ok(agentEmails.has(additionalAgentEmail), 'Expected created agent email in list');
    assert.ok(agentPhones.has(additionalAgentPhone), 'Expected created agent phone in list');

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: createAgent.data.username, password: createAgent.data.temporaryPassword },
      expectStatus: 200
    });

    const selfRegister = await request(baseUrl, '/api/auth/register-farmer', {
      method: 'POST',
      body: {
        name: 'Self Grower',
        phone: farmerSelfPhone,
        nationalId: '11223344',
        location: 'Muranga',
        preferredLanguage: 'sw',
        hectares: 1.2,
        avocadoHectares: 0.8,
        pin: farmerSelfPin,
        confirmPin: farmerSelfPin
      },
      expectStatus: 201
    });
    assert.strictEqual(selfRegister.data.user.role, 'farmer');
    assert.strictEqual(selfRegister.data.user.username, farmerSelfPhone);
    assert.ok(selfRegister.data.farmerId, 'Farmer ID missing after self-registration');

    await request(baseUrl, '/api/auth/register-farmer', {
      method: 'POST',
      body: {
        name: 'Duplicate Grower',
        phone: farmerSelfPhone,
        nationalId: '55443322',
        location: 'Muranga',
        hectares: 1.1,
        avocadoHectares: 0.6,
        pin: '1944',
        confirmPin: '1944'
      },
      expectStatus: 409
    });

    const farmerLogin = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: `+${farmerSelfPhone}`, password: farmerSelfPin },
      expectStatus: 200
    });
    let farmerToken = farmerLogin.data.token;
    const farmerMe = await request(baseUrl, '/api/auth/me', {
      token: farmerToken,
      expectStatus: 200
    });
    assert.strictEqual(farmerMe.data.user.username, farmerSelfPhone);

    await request(baseUrl, '/api/auth/change-password', {
      method: 'POST',
      token: farmerToken,
      body: {
        currentPassword: farmerSelfPin,
        newPassword: farmerSelfPinUpdated,
        confirmPassword: farmerSelfPinUpdated
      },
      expectStatus: 200
    });

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: farmerSelfPhone, password: farmerSelfPin },
      expectStatus: 401
    });

    const farmerLoginUpdated = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { phone: farmerSelfPhone, pin: farmerSelfPinUpdated },
      expectStatus: 200
    });
    farmerToken = farmerLoginUpdated.data.token;

    await request(baseUrl, '/api/auth/recover-username', {
      method: 'POST',
      body: { role: 'farmer', phone: farmerSelfPhone, recoveryCode: farmerSelfPinUpdated },
      expectStatus: 422
    });

    await request(baseUrl, '/api/auth/recover-password', {
      method: 'POST',
      body: {
        role: 'farmer',
        phone: farmerSelfPhone,
        recoveryCode: farmerSelfPinUpdated,
        newPassword: '1940',
        confirmPassword: '1940'
      },
      expectStatus: 422
    });

    const farmer = await request(baseUrl, '/api/farmers', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Grace Njoki',
        phone: '254711223344',
        nationalId: '87654321',
        location: 'Kiambu',
        hectares: 2.35,
        avocadoHectares: 1.5,
        trees: 44,
        notes: 'Integration test farmer'
      },
      expectStatus: 201
    });

    const farmerId = farmer.data.id;
    assert.ok(farmerId, 'Farmer ID missing after create');

    await request(baseUrl, '/api/farmers', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Grace Duplicate',
        phone: '254711223345',
        nationalId: '87654321',
        location: 'Kiambu',
        hectares: 1.2,
        avocadoHectares: 0.9,
        trees: 10
      },
      expectStatus: 409
    });

    const farmerPatchedArea = await request(baseUrl, `/api/farmers/${encodeURIComponent(farmerId)}`, {
      method: 'PATCH',
      token: adminToken,
      body: {
        squareFeet: 107639.1,
        avocadoSquareFeet: 53819.55
      },
      expectStatus: 200
    });
    assert.ok(Math.abs(Number(farmerPatchedArea.data.hectares) - 1.0) < 0.02, 'Square feet update should convert to hectares');
    assert.ok(
      Math.abs(Number(farmerPatchedArea.data.avocadoHectares) - 0.5) < 0.02,
      'Avocado square feet update should convert to hectares'
    );

    await request(baseUrl, '/api/farmers/import', {
      method: 'POST',
      token: agentToken,
      body: {
        records: [
          {
            name: 'Agent Attempt',
            phone: '254700000001',
            nationalId: '10000111',
            location: 'Nakuru',
            hectares: 1.4,
            avocadoHectares: 1.0
          }
        ]
      },
      expectStatus: 403
    });

    const importResponse = await request(baseUrl, '/api/farmers/import', {
      method: 'POST',
      token: adminToken,
      body: {
        sendOnboardingSms: true,
        onboardingSmsTemplate: 'Hello {{name}}, you are now on Agem Portal. Dial {{ussd}} when active.',
        records: [
          {
            'Farmer Name': 'Peter Mwangi',
            'Mobile Number': '254700123456',
            'National ID Number': '23597146',
            Ward: 'Kirinyaga',
            acreage: '4.5',
            'Area Under Avocado (Acres)': '2.0',
            'Tree Count': '30',
            'Random Extra Column': 'ignore this'
          },
          {
            name: 'Duplicate Grace',
            phone: '254711223344',
            nationalId: '87654321',
            location: 'Kiambu',
            hectares: 1.1,
            avocadoHectares: 0.5,
            trees: 55
          },
          {
            Name: '',
            Phone: '254722000001',
            'ID No': '29988777',
            Location: 'Nyandarua'
          },
          {
            'Full Name': 'Alice Otieno',
            MSISDN: '254733445566',
            'ID No': '30011223',
            County: 'Nyeri',
            'Square Feet': '107639.1',
            'Area Under Avocado (Square Feet)': '53819.55',
            'Number of Trees': '14',
            Remarks: 'Excel alias columns'
          }
        ]
      },
      expectStatus: 200
    });
    assert.strictEqual(importResponse.data.totalRows, 4);
    assert.strictEqual(importResponse.data.imported, 2);
    assert.strictEqual(importResponse.data.updated || 0, 0);
    assert.strictEqual(importResponse.data.skipped, 2);
    assert.strictEqual(importResponse.data.smsSent, 2);
    assert.ok(Array.isArray(importResponse.data.errors), 'Expected row errors array');
    assert.ok(importResponse.data.errors.length >= 2, 'Expected duplicate and validation row errors');

    const smsAfterImport = await request(baseUrl, '/api/sms?limit=20', {
      token: adminToken,
      expectStatus: 200
    });
    assert.strictEqual((smsAfterImport.data || []).length, 2, 'Expected two onboarding SMS logs from import');
    assert.ok(
      (smsAfterImport.data || []).every((row) => String(row.message || '').includes('Dial *483# when active.')),
      'Expected onboarding SMS template to be applied with USSD placeholder'
    );

    const farmersAfterImport = await request(baseUrl, '/api/farmers?limit=20', {
      token: adminToken,
      expectStatus: 200
    });
    const farmersByPhone = new Map((farmersAfterImport.data || []).map((row) => [row.phone, row]));
    const farmerPhones = new Set(farmersByPhone.keys());
    assert.ok(farmerPhones.has('254700123456'), 'Imported CSV row missing');
    assert.ok(farmerPhones.has('254733445566'), 'Imported alias row missing');

    const peter = farmersByPhone.get('254700123456');
    const alice = farmersByPhone.get('254733445566');
    assert.ok(peter && Number.isFinite(Number(peter.hectares)), 'Expected hectares for Peter');
    assert.ok(alice && Number.isFinite(Number(alice.hectares)), 'Expected hectares for Alice');
    assert.ok(peter && Number.isFinite(Number(peter.avocadoHectares)), 'Expected avocado hectares for Peter');
    assert.ok(alice && Number.isFinite(Number(alice.avocadoHectares)), 'Expected avocado hectares for Alice');
    assert.ok(Math.abs(Number(peter.hectares) - 1.821) < 0.02, 'Acreage to hectares conversion should be applied');
    assert.ok(Math.abs(Number(alice.hectares) - 1.0) < 0.02, 'Square-feet to hectares conversion should be applied');
    assert.ok(Math.abs(Number(peter.avocadoHectares) - 0.809) < 0.02, 'Avocado acreage to hectares conversion should be applied');
    assert.ok(Math.abs(Number(alice.avocadoHectares) - 0.5) < 0.02, 'Avocado square-feet to hectares conversion should be applied');
    assert.strictEqual(Boolean(peter.hasPortalAccess), false, 'Imported farmer should not have portal access before PIN set');

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { phone: peter.phone, pin: '3728' },
      expectStatus: 401
    });

    await request(baseUrl, `/api/farmers/${encodeURIComponent(peter.id)}/reset-pin`, {
      method: 'POST',
      token: agentToken,
      body: { pin: '3728', confirmPin: '3728' },
      expectStatus: 403
    });

    const pinProvision = await request(baseUrl, `/api/farmers/${encodeURIComponent(peter.id)}/reset-pin`, {
      method: 'POST',
      token: adminToken,
      body: { pin: '3728', confirmPin: '3728' },
      expectStatus: 200
    });
    assert.strictEqual(pinProvision.data.farmerId, peter.id);
    assert.strictEqual(pinProvision.data.accountCreated, true);
    assert.strictEqual(pinProvision.data.username, peter.phone);
    assert.strictEqual(pinProvision.data.generatedPin, '');

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { phone: peter.phone, pin: '3728' },
      expectStatus: 200
    });

    const overwriteImport = await request(baseUrl, '/api/farmers/import', {
      method: 'POST',
      token: adminToken,
      body: {
        onDuplicate: 'overwrite',
        records: [
          {
            name: 'Peter Mwangi Updated',
            phone: '254700999999',
            nationalId: '23597146',
            location: 'Embu',
            hectares: 2.2,
            avocadoHectares: 1.7,
            trees: 33,
            notes: 'Updated from import'
          }
        ]
      },
      expectStatus: 200
    });
    assert.strictEqual(overwriteImport.data.imported, 0);
    assert.strictEqual(overwriteImport.data.updated, 1);
    assert.strictEqual(overwriteImport.data.skipped, 0);
    assert.strictEqual(overwriteImport.data.duplicateMode, 'overwrite');

    const farmersAfterOverwrite = await request(baseUrl, '/api/farmers?limit=20', {
      token: adminToken,
      expectStatus: 200
    });
    const updatedPeter = (farmersAfterOverwrite.data || []).find((row) => row.nationalId === '23597146');
    assert.ok(updatedPeter, 'Expected updated farmer by National ID');
    assert.strictEqual(updatedPeter.phone, '254700999999');
    assert.strictEqual(updatedPeter.location, 'Embu');
    assert.ok(Math.abs(Number(updatedPeter.hectares) - 2.2) < 0.01);
    assert.ok(Math.abs(Number(updatedPeter.avocadoHectares) - 1.7) < 0.01);
    assert.strictEqual(Boolean(updatedPeter.hasPortalAccess), true, 'Expected imported farmer portal access to persist');

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { phone: '254700123456', pin: '3728' },
      expectStatus: 401
    });

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { phone: '254700999999', pin: '3728' },
      expectStatus: 200
    });

    const generatedPinReset = await request(baseUrl, `/api/farmers/${encodeURIComponent(updatedPeter.id)}/reset-pin`, {
      method: 'POST',
      token: adminToken,
      body: { generate: true },
      expectStatus: 200
    });
    assert.strictEqual(generatedPinReset.data.accountCreated, false);
    assert.ok(/^\d{4}$/.test(generatedPinReset.data.generatedPin), 'Expected generated 4-digit PIN');

    await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { phone: '254700999999', pin: generatedPinReset.data.generatedPin },
      expectStatus: 200
    });

    const produceResponse = await request(baseUrl, '/api/produce', {
      method: 'POST',
      token: agentToken,
      body: {
        farmerId,
        kgs: 128.5,
        lotWeightKgs: 128.5,
        variety: 'Hass',
        sampleSize: 16,
        visualGrade: 'Pass',
        dryMatterPct: 23.8,
        firmnessValue: 75.2,
        firmnessUnit: 'N',
        avgFruitWeightG: 245.7,
        sizeCode: 'C20',
        qcDecision: 'Accept',
        inspector: 'Agent Jane',
        notes: 'Morning collection'
      },
      expectStatus: 201
    });
    assert.ok(produceResponse.data?.id, 'QC record ID missing');

    const purchaseResponse = await request(baseUrl, '/api/produce-purchases', {
      method: 'POST',
      token: agentToken,
      body: {
        farmerId,
        qcRecordId: '',
        variety: 'Hass',
        sizeCode: 'C20',
        purchasedKgs: 120.3,
        pricePerKgKes: 155.5,
        buyer: 'Agent Jane',
        notes: 'Accepted for pickup'
      },
      expectStatus: 201
    });
    assert.ok(purchaseResponse.data.id, 'Produce purchase ID missing');

    const owedBefore = await request(baseUrl, '/api/payments/owed?period=all', {
      token: adminToken,
      expectStatus: 200
    });
    const owedBeforeRow = (owedBefore.data || []).find((row) => row.farmerId === farmerId);
    assert.ok(owedBeforeRow, 'Expected farmer to appear in owed list');
    const owedBalanceBefore = Number(owedBeforeRow.balanceKes || 0);
    assert.ok(owedBalanceBefore > 1000, 'Expected positive owed balance before settlement');

    const pendingSettlement = await request(baseUrl, '/api/payments/settle', {
      method: 'POST',
      token: adminToken,
      body: {
        period: 'all',
        farmerIds: [farmerId],
        status: 'Pending'
      },
      expectStatus: 201
    });
    assert.strictEqual(pendingSettlement.data.createdCount, 1);
    assert.strictEqual(pendingSettlement.data.status, 'Pending');

    const owedAfterPending = await request(baseUrl, '/api/payments/owed?period=all', {
      token: adminToken,
      expectStatus: 200
    });
    const owedAfterPendingRow = (owedAfterPending.data || []).find((row) => row.farmerId === farmerId);
    assert.ok(owedAfterPendingRow, 'Farmer should remain owed after pending settlement');
    assert.ok(Number(owedAfterPendingRow.balanceKes || 0) >= owedBalanceBefore - 0.01);

    const receivedSettlement = await request(baseUrl, '/api/payments/settle', {
      method: 'POST',
      token: adminToken,
      body: {
        period: 'all',
        farmerIds: [farmerId],
        status: 'Received'
      },
      expectStatus: 201
    });
    assert.strictEqual(receivedSettlement.data.createdCount, 1);
    assert.strictEqual(receivedSettlement.data.status, 'Received');

    const owedAfterReceived = await request(baseUrl, '/api/payments/owed?period=all', {
      token: adminToken,
      expectStatus: 200
    });
    const owedAfterReceivedRow = (owedAfterReceived.data || []).find((row) => row.farmerId === farmerId);
    assert.ok(!owedAfterReceivedRow || Number(owedAfterReceivedRow.balanceKes || 0) < 0.01);

    await request(baseUrl, '/api/payments', {
      method: 'POST',
      token: agentToken,
      body: {
        farmerId,
        amount: 5400,
        ref: 'AGTFAIL001',
        status: 'Received'
      },
      expectStatus: 403
    });

    await request(baseUrl, '/api/payments', {
      method: 'POST',
      token: adminToken,
      body: {
        farmerId,
        amount: 5400,
        ref: 'ADMOK001',
        status: 'Received',
        notes: 'Approved payment'
      },
      expectStatus: 201
    });

    const qcIntelligence = await request(baseUrl, '/api/ai/qc-intelligence', {
      method: 'POST',
      token: agentToken,
      body: {
        qcRecordId: produceResponse.data.id
      },
      expectStatus: 200
    });
    assert.strictEqual(qcIntelligence.data.items.length, 1);
    assert.ok(
      ['Accept', 'Hold', 'Reject'].includes(qcIntelligence.data.items[0].recommendedDecision),
      'QC intelligence should return recommended decision'
    );
    assert.ok(qcIntelligence.data.items[0].summary, 'QC intelligence summary should be returned');

    await request(baseUrl, '/api/ai/payment-risk', {
      method: 'POST',
      token: agentToken,
      body: { period: 'week' },
      expectStatus: 403
    });

    const paymentRisk = await request(baseUrl, '/api/ai/payment-risk', {
      method: 'POST',
      token: adminToken,
      body: { period: 'week' },
      expectStatus: 200
    });
    assert.ok(paymentRisk.data.summary, 'Payment risk summary should be returned');
    assert.ok(Array.isArray(paymentRisk.data.flags), 'Payment risk flags should be an array');

    // Phase 2B (exact order): Executive Briefs -> Proposal Queue -> Ops Tasks -> Knowledge Base -> Feedback/Evals
    await request(baseUrl, '/api/ai/briefs/run-now', {
      method: 'POST',
      token: agentToken,
      body: {},
      expectStatus: 403
    });

    const runBrief = await request(baseUrl, '/api/ai/briefs/run-now', {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 201
    });
    assert.ok(runBrief.data?.id, 'Executive brief run-now should return brief id');
    assert.ok(runBrief.data?.overallRisk, 'Executive brief should include overall risk');

    const listBriefsAgent = await request(baseUrl, '/api/ai/briefs?limit=5&includeAlerts=true', {
      token: agentToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(listBriefsAgent.data?.briefs), 'Briefs endpoint should return briefs array');
    assert.ok(listBriefsAgent.data.briefs.length >= 1, 'Expected at least one executive brief');

    await request(baseUrl, '/api/ai/proposals', {
      token: agentToken,
      expectStatus: 403
    });

    const proposalCreate = await request(baseUrl, '/api/ai/proposals', {
      method: 'POST',
      token: adminToken,
      body: {
        actionType: 'create_ops_tasks',
        confidence: 0.91,
        title: 'Follow up pending collections',
        description: 'Create follow-up tasks for pending collections and payment verification.',
        payload: {
          tasks: [
            {
              title: 'Call Kiambu farmer cluster',
              description: 'Verify pending produce handover and confirm next pickup window.',
              severity: 'high',
              sourceType: 'manual-ai',
              sourceRef: 'cluster-kiambu'
            },
            {
              title: 'Reconcile payment evidence',
              description: 'Match payment references against settlement sheet before disbursement.',
              severity: 'medium',
              sourceType: 'manual-ai',
              sourceRef: 'settlement-batch-1'
            }
          ]
        }
      },
      expectStatus: 201
    });
    const proposalId = proposalCreate.data?.id;
    assert.ok(proposalId, 'Proposal id should be returned');
    assert.strictEqual(proposalCreate.data?.status, 'pending');

    const listProposals = await request(baseUrl, '/api/ai/proposals?status=pending&limit=20', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(listProposals.data), 'Proposals endpoint should return array');
    assert.ok((listProposals.data || []).some((row) => row.id === proposalId), 'Created proposal should appear in queue');

    const approvedProposal = await request(baseUrl, `/api/ai/proposals/${encodeURIComponent(proposalId)}/approve`, {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 200
    });
    assert.strictEqual(approvedProposal.data?.status, 'approved', 'Proposal should move to approved state');

    const executedProposal = await request(baseUrl, `/api/ai/proposals/${encodeURIComponent(proposalId)}/execute`, {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 200
    });
    assert.strictEqual(executedProposal.data?.status, 'executed', 'Proposal should move to executed state');
    assert.ok(
      Number(executedProposal.result?.tasksCreated || 0) >= 2,
      'Executed custom proposal should create ops tasks'
    );

    const proposalRejectCreate = await request(baseUrl, '/api/ai/proposals', {
      method: 'POST',
      token: adminToken,
      body: {
        actionType: 'generate_executive_brief',
        title: 'Reject flow coverage',
        description: 'Test reject lifecycle',
        payload: {}
      },
      expectStatus: 201
    });
    const rejectProposalId = proposalRejectCreate.data?.id;
    assert.ok(rejectProposalId, 'Reject test proposal should be created');

    const rejectedProposal = await request(baseUrl, `/api/ai/proposals/${encodeURIComponent(rejectProposalId)}/reject`, {
      method: 'POST',
      token: adminToken,
      body: { reason: 'Duplicate request' },
      expectStatus: 200
    });
    assert.strictEqual(rejectedProposal.data?.status, 'rejected', 'Proposal should move to rejected state');
    assert.strictEqual(rejectedProposal.data?.rejectionReason, 'Duplicate request');

    await request(baseUrl, `/api/ai/proposals/${encodeURIComponent(rejectProposalId)}/execute`, {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 409
    });

    await request(baseUrl, '/api/ops/tasks/from-ai', {
      method: 'POST',
      token: agentToken,
      body: { source: 'custom', tasks: [{ title: 'agent forbidden task' }] },
      expectStatus: 403
    });

    const opsCreateCustom = await request(baseUrl, '/api/ops/tasks/from-ai', {
      method: 'POST',
      token: adminToken,
      body: {
        source: 'custom',
        tasks: [
          {
            title: 'Verify packhouse readiness',
            description: 'Confirm grading line staffing for tomorrow.',
            severity: 'low',
            sourceType: 'ops-manual',
            sourceRef: 'packhouse'
          }
        ]
      },
      expectStatus: 201
    });
    assert.ok(Number(opsCreateCustom.meta?.createdCount || 0) >= 1, 'Expected at least one custom ops task');

    const opsListAgent = await request(baseUrl, '/api/ops/tasks?limit=100', {
      token: agentToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(opsListAgent.data), 'Ops tasks endpoint should return array');
    assert.ok((opsListAgent.data || []).length >= 3, 'Expected ops tasks from executed proposal and custom source');

    const firstOpenTask = (opsListAgent.data || []).find((row) => row.status === 'open') || opsListAgent.data[0];
    assert.ok(firstOpenTask?.id, 'Expected at least one ops task id');

    const taskPatchedByAgent = await request(baseUrl, `/api/ops/tasks/${encodeURIComponent(firstOpenTask.id)}`, {
      method: 'PATCH',
      token: agentToken,
      body: {
        status: 'in_progress',
        assignedTo: 'Field Agent East',
        notes: 'Started follow-up.'
      },
      expectStatus: 200
    });
    assert.strictEqual(taskPatchedByAgent.data?.status, 'in_progress', 'Agent should be able to update task status');

    const knowledgeCreated = await request(baseUrl, '/api/ai/knowledge', {
      method: 'POST',
      token: adminToken,
      body: {
        title: 'Payment Reconciliation SOP',
        source: 'Internal SOP v1',
        tags: 'payments,reconciliation,weekly',
        content: 'Always match M-PESA reference, farmer national ID, and owed balance before final disbursement.'
      },
      expectStatus: 201
    });
    const knowledgeDocId = knowledgeCreated.data?.id;
    assert.ok(knowledgeDocId, 'Knowledge document id should be returned');

    await request(baseUrl, '/api/ai/knowledge', {
      method: 'POST',
      token: agentToken,
      body: {
        title: 'Agent cannot write',
        content: 'Should fail'
      },
      expectStatus: 403
    });

    const knowledgeSearchAgent = await request(baseUrl, '/api/ai/knowledge?q=reconciliation&limit=10', {
      token: agentToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(knowledgeSearchAgent.data), 'Knowledge list should return array');
    assert.ok(
      (knowledgeSearchAgent.data || []).some((row) => row.id === knowledgeDocId),
      'Agent should see created knowledge document'
    );

    await request(baseUrl, `/api/ai/knowledge/${encodeURIComponent(knowledgeDocId)}`, {
      method: 'DELETE',
      token: agentToken,
      body: {},
      expectStatus: 403
    });

    await request(baseUrl, `/api/ai/knowledge/${encodeURIComponent(knowledgeDocId)}`, {
      method: 'DELETE',
      token: adminToken,
      body: {},
      expectStatus: 200
    });

    const feedbackAgent = await request(baseUrl, '/api/ai/feedback', {
      method: 'POST',
      token: agentToken,
      body: {
        tool: 'copilot',
        rating: 'up',
        responseId: runBrief.data.id,
        note: 'Helpful summary for field follow-up.'
      },
      expectStatus: 201
    });
    assert.ok(feedbackAgent.data?.id, 'Feedback entry id should be returned');

    const feedbackAdmin = await request(baseUrl, '/api/ai/feedback', {
      method: 'POST',
      token: adminToken,
      body: {
        tool: 'payment-risk',
        rating: 'neutral',
        responseId: paymentRisk.data?.responseId || '',
        note: 'Useful but needs more context on duplicates.'
      },
      expectStatus: 201
    });
    assert.ok(feedbackAdmin.data?.id, 'Admin feedback entry id should be returned');

    await request(baseUrl, '/api/ai/feedback/summary?limit=10', {
      token: agentToken,
      expectStatus: 403
    });

    const feedbackSummary = await request(baseUrl, '/api/ai/feedback/summary?limit=20', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(feedbackSummary.data?.totals?.total >= 2, 'Feedback summary should aggregate submitted entries');
    assert.ok(Array.isArray(feedbackSummary.data?.byTool), 'Feedback summary should include byTool array');

    await request(baseUrl, '/api/ai/evals/run', {
      method: 'POST',
      token: agentToken,
      body: {},
      expectStatus: 403
    });

    const evalRun = await request(baseUrl, '/api/ai/evals/run', {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 201
    });
    assert.ok(evalRun.data?.id, 'Eval run id should be returned');
    assert.ok(Number(evalRun.data?.totalChecks || 0) >= 1, 'Eval run should include checks');

    const evalRuns = await request(baseUrl, '/api/ai/evals?limit=10', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(evalRuns.data), 'Eval runs endpoint should return array');
    assert.ok((evalRuns.data || []).some((row) => row.id === evalRun.data.id), 'Latest eval run should be listed');

    await request(baseUrl, '/api/sms/send', {
      method: 'POST',
      token: agentToken,
      body: {
        farmerId,
        message: 'Your avocado payment has been processed.'
      },
      expectStatus: 201
    });

    const recipientSearch = await request(baseUrl, '/api/sms/recipients?q=grace&limit=10&offset=0', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(recipientSearch.data));
    assert.ok((recipientSearch.meta?.total || 0) >= 1);

    await request(baseUrl, '/api/sms/send-bulk', {
      method: 'POST',
      token: agentToken,
      body: {
        mode: 'all',
        message: 'Agent should not be allowed for bulk SMS.'
      },
      expectStatus: 403
    });

    const bulkSelected = await request(baseUrl, '/api/sms/send-bulk', {
      method: 'POST',
      token: adminToken,
      body: {
        mode: 'selected',
        farmerIds: [farmerId],
        message: 'Selected recipient broadcast.'
      },
      expectStatus: 201
    });
    assert.strictEqual(bulkSelected.data.sentCount, 1);

    const bulkAll = await request(baseUrl, '/api/sms/send-bulk', {
      method: 'POST',
      token: adminToken,
      body: {
        mode: 'all',
        message: 'Platform-wide broadcast.'
      },
      expectStatus: 201
    });
    assert.strictEqual(bulkAll.data.sentCount, 4);

    const summary = await request(baseUrl, '/api/reports/summary', {
      token: adminToken,
      expectStatus: 200
    });
    assert.strictEqual(summary.data.farmers, 4);
    assert.strictEqual(summary.data.produceRecords, 1);
    assert.strictEqual(summary.data.purchasedRecords, 1);
    assert.strictEqual(summary.data.paymentRecords, 3);
    assert.strictEqual(summary.data.smsSent, 8);
    assert.strictEqual(Number(summary.data.smsOwnerCostPerMessageKes), 0.25);
    assert.strictEqual(summary.data.smsSentLast24h, 8);
    assert.ok(Math.abs(Number(summary.data.smsSpentKes) - 2.0) < 0.0001, 'Expected total SMS spend to be KES 2.00');
    assert.ok(
      Math.abs(Number(summary.data.smsSpentLast24hKes) - 2.0) < 0.0001,
      'Expected 24h SMS spend to be KES 2.00'
    );
    assert.ok(Number(summary.data.totalPurchasedKg) > 120, 'Expected purchased produce kg in summary');
    assert.ok(Number(summary.data.totalOwedKes || 0) < 0.01, 'Expected owed balance to be settled');

    const ussdUnauthorized = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unauth',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        text: ''
      }).toString(),
      responseType: 'text',
      expectStatus: 403
    });
    assert.ok(
      ussdUnauthorized.startsWith('END Unauthorized'),
      'USSD callback should reject requests without shared secret header'
    );

    const ussdMenu = await request(baseUrl, `/api/ussd/callback?secret=${encodeURIComponent(ussdSecret)}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-menu',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        text: ''
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdMenu.startsWith('CON '), 'USSD menu should keep session open');
    assert.ok(ussdMenu.includes('Choose language'), 'USSD language menu should be shown first');

    const ussdEnglishMenu = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-menu-en',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        text: '1'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdEnglishMenu.startsWith('CON '), 'USSD English menu should keep session open');
    assert.ok(ussdEnglishMenu.includes('1. Payment Status'), 'USSD English menu missing payment option');

    await request(baseUrl, `/api/ussd/events?secret=${encodeURIComponent(ussdSecret)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      bodyRaw: new URLSearchParams({
        sessionId: 'evt-1',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        status: 'Completed'
      }).toString(),
      expectStatus: 200
    });

    const ussdPaymentStatus = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-pay',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        text: '1*1'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdPaymentStatus.startsWith('END '), 'USSD payment lookup should end session');
    assert.ok(ussdPaymentStatus.includes('Balance: KES'), 'USSD payment response missing balance');

    const ussdQcStatus = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-qc',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        text: '1*2'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdQcStatus.startsWith('END '), 'USSD QC lookup should end session');
    assert.ok(ussdQcStatus.includes('Decision:'), 'USSD QC response missing decision');

    const ussdProfile = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-profile',
        serviceCode: ussdCode,
        phoneNumber: '254711223344',
        text: '1*3'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdProfile.startsWith('END '), 'USSD profile lookup should end session');
    assert.ok(ussdProfile.includes('Name: Grace Njoki'), 'USSD profile response missing farmer name');

    const ussdUnknownLanguageSelected = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-en',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownLanguageSelected.startsWith('CON '), 'Unregistered USSD flow should continue to registration menu');
    assert.ok(
      ussdUnknownLanguageSelected.includes('Register Now'),
      'USSD unknown number should receive registration option'
    );

    const ussdUnknownStep1 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step1',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep1.includes('Step 1/7'), 'USSD registration should prompt for name');

    const ussdUnknownStep2 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step2',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep2.includes('Step 2/7'), 'USSD registration should prompt for National ID');

    const ussdUnknownStep3 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step3',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno*33445566'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep3.includes('Step 3/7'), 'USSD registration should prompt for location');

    const ussdUnknownStep4 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step4',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno*33445566*Kisumu'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep4.includes('Step 4/7'), 'USSD registration should prompt for total acres');

    const ussdUnknownStep5 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step5',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno*33445566*Kisumu*2.4'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep5.includes('Step 5/7'), 'USSD registration should prompt for avocado acres');

    const ussdUnknownStep6 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step6',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno*33445566*Kisumu*2.4*1.6'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep6.includes('Step 6/7'), 'USSD registration should prompt for PIN');

    const ussdUnknownStep7 = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-step7',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno*33445566*Kisumu*2.4*1.6*2468'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownStep7.includes('Step 7/7'), 'USSD registration should confirm PIN');

    const ussdUnknownComplete = await request(baseUrl, '/api/ussd/callback', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'x-ussd-secret': ussdSecret
      },
      bodyRaw: new URLSearchParams({
        sessionId: 'sess-unknown-complete',
        serviceCode: ussdCode,
        phoneNumber: '254799000000',
        text: '1*1*Jane Atieno*33445566*Kisumu*2.4*1.6*2468*2468'
      }).toString(),
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(ussdUnknownComplete.startsWith('END '), 'USSD registration should complete and end session');
    assert.ok(ussdUnknownComplete.includes('Registration complete'), 'USSD registration completion message missing');

    const farmersAfterUssdRegistration = await request(baseUrl, '/api/farmers?limit=50', {
      token: adminToken,
      expectStatus: 200
    });
    const jane = (farmersAfterUssdRegistration.data || []).find((row) => row.phone === '254799000000');
    assert.ok(jane, 'USSD registration should create a farmer profile');
    assert.strictEqual(jane.nationalId, '33445566', 'USSD registration should capture National ID');
    assert.strictEqual(jane.preferredLanguage, 'en', 'USSD registration should capture selected language');

    const smsAfterUssdRegistration = await request(baseUrl, '/api/sms?limit=50', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(
      (smsAfterUssdRegistration.data || []).some(
        (row) => row.phone === '254799000000' && String(row.message || '').includes('Registration complete')
      ),
      'USSD registration should log confirmation SMS'
    );

    const farmersCsv = await request(baseUrl, '/api/exports/farmers.csv', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(
      farmersCsv.includes('name,phone,nationalId,location,preferredLanguage,hectares,acres,squareFeet,avocadoHectares,avocadoAcres,avocadoSquareFeet'),
      'Farmers CSV header missing'
    );
    assert.ok(farmersCsv.includes('Grace Njoki'), 'Farmers CSV content missing');

    const produceCsv = await request(baseUrl, '/api/exports/produce.csv', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(
      produceCsv.includes('lotWeightKgs,variety,sampleSize,visualGrade,dryMatterPct,firmnessValue,firmnessUnit,avgFruitWeightG,sizeCode,qcDecision,inspector'),
      'Produce CSV header missing farm-gate QC columns'
    );

    const producePurchasesCsv = await request(baseUrl, '/api/exports/produce-purchases.csv', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(
      producePurchasesCsv.includes('qcRecordId,variety,sizeCode,purchasedKgs,pricePerKgKes,purchaseValueKes,paidAmountKes,balanceKes,settlementStatus,buyer'),
      'Produce purchases CSV header missing purchase columns'
    );

    const owedCsv = await request(baseUrl, '/api/exports/payments-owed.csv?period=all', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(
      owedCsv.includes('farmerId,farmerName,phone,nationalId,location,purchaseCount,purchasedKgs,totalValueKes,paidKes,balanceKes,lastPurchaseAt'),
      'Owed CSV header missing'
    );

    const smsCsv = await request(baseUrl, '/api/exports/sms.csv', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(
      smsCsv.includes('id,farmerId,farmerName,phone,message,provider,status,ownerCostKes,billable,createdBy,createdAt'),
      'SMS CSV header missing'
    );
    assert.ok(smsCsv.includes('0.25'), 'SMS CSV should include owner SMS cost values');

    const smsCsvToday = await request(baseUrl, '/api/exports/sms.csv?period=today', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(smsCsvToday.includes('ownerCostKes,billable'), 'SMS today export should include cost columns');

    const smsCsvFuture = await request(baseUrl, '/api/exports/sms.csv?period=custom&from=2099-01-01&to=2099-01-02', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    const smsCsvFutureLines = smsCsvFuture.trim().split('\n');
    assert.strictEqual(smsCsvFutureLines.length, 1, 'Future-range SMS export should return header only');

    const backupCreated = await request(baseUrl, '/api/admin/backup', {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 201
    });
    const backupFilename = backupCreated.data?.filename;
    assert.ok(backupFilename, 'Backup filename should be returned');

    const backups = await request(baseUrl, '/api/admin/backups', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(Array.isArray(backups.data));
    assert.ok(backups.data.length >= 1, 'Expected at least one backup');

    await request(baseUrl, '/api/admin/reset', {
      method: 'POST',
      token: adminToken,
      body: { confirm: true },
      expectStatus: 200
    });

    const summaryAfterReset = await request(baseUrl, '/api/reports/summary', {
      token: adminToken,
      expectStatus: 200
    });
    assert.strictEqual(summaryAfterReset.data.farmers, 0);
    assert.strictEqual(summaryAfterReset.data.purchasedRecords, 0);

    const restoreResponse = await request(baseUrl, '/api/admin/restore', {
      method: 'POST',
      token: adminToken,
      body: {
        filename: backupFilename,
        confirm: true
      },
      expectStatus: 200
    });
    assert.strictEqual(restoreResponse.data.restored, true);
    assert.strictEqual(restoreResponse.data.fromBackup, backupFilename);

    const summaryAfterRestore = await request(baseUrl, '/api/reports/summary', {
      token: adminToken,
      expectStatus: 200
    });
    assert.ok(summaryAfterRestore.data.farmers >= 4, 'Expected farmers after restore');
    assert.ok(summaryAfterRestore.data.purchasedRecords >= 1, 'Expected produce purchases after restore');

    await request(baseUrl, '/api/auth/logout', {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 200
    });

    await request(baseUrl, '/api/auth/me', {
      token: adminToken,
      expectStatus: 401
    });

    console.log('All integration checks passed.');
  } finally {
    server.kill('SIGTERM');
    await sleep(200);
    if (!server.killed) {
      server.kill('SIGKILL');
    }

    if (stderr.trim()) {
      // Keep stderr visible only when needed for debugging CI failures.
    }
  }
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
