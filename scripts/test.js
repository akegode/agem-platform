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
  const agentUsername = 'fieldagent';
  const agentPassword = 'FieldAgent#2026!';

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
      AGENT_PASSWORD: agentPassword
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
    const adminToken = adminLogin.data.token;
    assert.ok(adminToken, 'Admin token missing');

    const agentLogin = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: agentUsername, password: agentPassword },
      expectStatus: 200
    });
    const agentToken = agentLogin.data.token;

    const farmer = await request(baseUrl, '/api/farmers', {
      method: 'POST',
      token: adminToken,
      body: {
        name: 'Grace Njoki',
        phone: '254711223344',
        location: 'Kiambu',
        trees: 44,
        notes: 'Integration test farmer'
      },
      expectStatus: 201
    });

    const farmerId = farmer.data.id;
    assert.ok(farmerId, 'Farmer ID missing after create');

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

    const summary = await request(baseUrl, '/api/reports/summary', {
      token: adminToken,
      expectStatus: 200
    });
    assert.strictEqual(summary.data.farmers, 1);
    assert.strictEqual(summary.data.produceRecords, 1);
    assert.strictEqual(summary.data.paymentRecords, 1);
    assert.strictEqual(summary.data.smsSent, 1);

    const farmersCsv = await request(baseUrl, '/api/exports/farmers.csv', {
      token: adminToken,
      responseType: 'text',
      expectStatus: 200
    });
    assert.ok(farmersCsv.includes('name,phone,location'), 'Farmers CSV header missing');
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
