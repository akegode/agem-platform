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
  const headers = {};
  if (options.token) headers.Authorization = `Bearer ${options.token}`;
  if (options.body !== undefined) headers['Content-Type'] = 'application/json';

  const response = await fetch(`${baseUrl}${route}`, {
    method: options.method || 'GET',
    headers,
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined
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
  const agentRecoveryCode = 'AgentRecovery#2026';
  const farmerSelfUsername = 'grower01';
  const farmerSelfPassword = 'GrowerPass#2026!';

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
      AGENT_RECOVERY_CODE: agentRecoveryCode
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

    const selfRegister = await request(baseUrl, '/api/auth/register-farmer', {
      method: 'POST',
      body: {
        name: 'Self Grower',
        phone: '254700111222',
        location: 'Muranga',
        hectares: 1.2,
        username: farmerSelfUsername,
        password: farmerSelfPassword,
        confirmPassword: farmerSelfPassword
      },
      expectStatus: 201
    });
    assert.strictEqual(selfRegister.data.user.role, 'farmer');
    assert.ok(selfRegister.data.farmerId, 'Farmer ID missing after self-registration');

    await request(baseUrl, '/api/auth/register-farmer', {
      method: 'POST',
      body: {
        name: 'Duplicate Grower',
        phone: '254700111222',
        location: 'Muranga',
        hectares: 1.1,
        username: 'grower02',
        password: farmerSelfPassword,
        confirmPassword: farmerSelfPassword
      },
      expectStatus: 409
    });

    const farmerLogin = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: farmerSelfUsername.toUpperCase(), password: farmerSelfPassword },
      expectStatus: 200
    });
    const farmerToken = farmerLogin.data.token;
    const farmerMe = await request(baseUrl, '/api/auth/me', {
      token: farmerToken,
      expectStatus: 200
    });
    assert.strictEqual(farmerMe.data.user.username, farmerSelfUsername);

    const farmer = await request(baseUrl, '/api/farmers', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Grace Njoki',
        phone: '254711223344',
        location: 'Kiambu',
        hectares: 2.35,
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
        phone: '254711223344',
        location: 'Kiambu',
        hectares: 1.2,
        trees: 10
      },
      expectStatus: 409
    });

    const farmerPatchedArea = await request(baseUrl, `/api/farmers/${encodeURIComponent(farmerId)}`, {
      method: 'PATCH',
      token: adminToken,
      body: {
        squareFeet: 107639.1
      },
      expectStatus: 200
    });
    assert.ok(Math.abs(Number(farmerPatchedArea.data.hectares) - 1.0) < 0.02, 'Square feet update should convert to hectares');

    await request(baseUrl, '/api/farmers/import', {
      method: 'POST',
      token: agentToken,
      body: {
        records: [
          {
            name: 'Agent Attempt',
            phone: '254700000001',
            location: 'Nakuru',
            hectares: 1.4
          }
        ]
      },
      expectStatus: 403
    });

    const importResponse = await request(baseUrl, '/api/farmers/import', {
      method: 'POST',
      token: adminToken,
      body: {
        records: [
          {
            'Farmer Name': 'Peter Mwangi',
            'Mobile Number': '254700123456',
            Ward: 'Kirinyaga',
            acreage: '4.5',
            'Tree Count': '30',
            'Random Extra Column': 'ignore this'
          },
          {
            name: 'Duplicate Grace',
            phone: '254711223344',
            location: 'Kiambu',
            hectares: 1.1,
            trees: 55
          },
          {
            Name: '',
            Phone: '254722000001',
            Location: 'Nyandarua'
          },
          {
            'Full Name': 'Alice Otieno',
            MSISDN: '254733445566',
            County: 'Nyeri',
            'Square Feet': '107639.1',
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
    assert.ok(Array.isArray(importResponse.data.errors), 'Expected row errors array');
    assert.ok(importResponse.data.errors.length >= 2, 'Expected duplicate and validation row errors');

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
    assert.ok(Math.abs(Number(peter.hectares) - 1.821) < 0.02, 'Acreage to hectares conversion should be applied');
    assert.ok(Math.abs(Number(alice.hectares) - 1.0) < 0.02, 'Square-feet to hectares conversion should be applied');

    const overwriteImport = await request(baseUrl, '/api/farmers/import', {
      method: 'POST',
      token: adminToken,
      body: {
        onDuplicate: 'overwrite',
        records: [
          {
            name: 'Peter Mwangi Updated',
            phone: '254700123456',
            location: 'Embu',
            hectares: 2.2,
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
    const updatedPeter = (farmersAfterOverwrite.data || []).find((row) => row.phone === '254700123456');
    assert.ok(updatedPeter, 'Expected updated farmer by phone');
    assert.strictEqual(updatedPeter.location, 'Embu');
    assert.ok(Math.abs(Number(updatedPeter.hectares) - 2.2) < 0.01);

    await request(baseUrl, '/api/produce', {
      method: 'POST',
      token: agentToken,
      body: {
        farmerId,
        kgs: 128.5,
        quality: 'A',
        agent: 'Agent Jane',
        notes: 'Morning collection'
      },
      expectStatus: 201
    });

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
    assert.strictEqual(summary.data.paymentRecords, 1);
    assert.strictEqual(summary.data.smsSent, 6);

    const farmersCsv = await request(baseUrl, '/api/exports/farmers.csv', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(farmersCsv.includes('name,phone,location,hectares,acres,squareFeet'), 'Farmers CSV header missing');
    assert.ok(farmersCsv.includes('Grace Njoki'), 'Farmers CSV content missing');

    await request(baseUrl, '/api/admin/backup', {
      method: 'POST',
      token: adminToken,
      body: {},
      expectStatus: 201
    });

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
