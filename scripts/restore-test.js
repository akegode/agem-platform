const assert = require('assert');
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServer(baseUrl, timeoutMs = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(`${baseUrl}/api/health`);
      if (res.ok) {
        return;
      }
    } catch {
      // retry
    }
    await sleep(250);
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
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined
  });

  const data = await response.json().catch(() => ({}));
  if (options.expectStatus !== undefined) {
    assert.strictEqual(
      response.status,
      options.expectStatus,
      `Expected ${options.expectStatus} for ${route}, got ${response.status} (${JSON.stringify(data)})`
    );
  }
  return data;
}

async function run() {
  const root = path.resolve(__dirname, '..');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'agem-restore-test-'));
  const port = 19000 + Math.floor(Math.random() * 1000);
  const baseUrl = `http://127.0.0.1:${port}`;

  const adminUsername = 'restoreadmin';
  const adminPassword = 'RestoreAdmin#2026!';

  const storageBackend = process.env.STORAGE_BACKEND || (process.env.DATABASE_URL ? '' : 'json');
  const env = {
    ...process.env,
    HOST: '127.0.0.1',
    PORT: String(port),
    ALLOW_DEMO_USERS: 'false',
    ADMIN_USERNAME: adminUsername,
    ADMIN_PASSWORD: adminPassword,
    BACKUP_INTERVAL_HOURS: '24',
    BACKUP_RETENTION_DAYS: '30'
  };

  if (storageBackend) {
    env.STORAGE_BACKEND = storageBackend;
  }
  if (!process.env.DATABASE_URL && (!storageBackend || storageBackend === 'json')) {
    env.STORE_PATH = path.join(tmp, 'store.json');
  }

  const server = spawn('node', ['backend/server.js'], {
    cwd: root,
    env,
    stdio: ['ignore', 'pipe', 'pipe']
  });

  let stderr = '';
  server.stderr.on('data', (chunk) => {
    stderr += chunk.toString();
  });

  try {
    await waitForServer(baseUrl);

    const login = await request(baseUrl, '/api/auth/login', {
      method: 'POST',
      body: { username: adminUsername, password: adminPassword },
      expectStatus: 200
    });
    const token = login.data?.token;
    assert.ok(token, 'Admin token missing');

    const createdFarmer = await request(baseUrl, '/api/farmers', {
      method: 'POST',
      token,
      body: {
        name: 'Restore Test Farmer',
        phone: '254799888777',
        nationalId: '99887766',
        location: 'Nyeri',
        hectares: 1.75,
        avocadoHectares: 1.2,
        trees: 42
      },
      expectStatus: 201
    });
    const farmerId = createdFarmer.data?.id;
    assert.ok(farmerId, 'Farmer create failed before backup');

    const backup = await request(baseUrl, '/api/admin/backup', {
      method: 'POST',
      token,
      body: {},
      expectStatus: 201
    });
    const filename = backup.data?.filename;
    assert.ok(filename, 'Backup filename missing');

    await request(baseUrl, '/api/admin/reset', {
      method: 'POST',
      token,
      body: { confirm: true },
      expectStatus: 200
    });

    const afterReset = await request(baseUrl, '/api/farmers', {
      token,
      expectStatus: 200
    });
    assert.strictEqual(afterReset.data.length, 0, 'Expected no farmers after reset');

    await request(baseUrl, '/api/admin/restore', {
      method: 'POST',
      token,
      body: { filename, confirm: true },
      expectStatus: 200
    });

    const afterRestore = await request(baseUrl, '/api/farmers', {
      token,
      expectStatus: 200
    });
    const restored = (afterRestore.data || []).find((row) => row.id === farmerId);
    assert.ok(restored, 'Expected farmer record after restore');

    console.log('Restore test passed:', {
      storageBackend: env.STORAGE_BACKEND || 'auto',
      backup: filename,
      restoredFarmerId: farmerId
    });
  } finally {
    server.kill('SIGTERM');
    await sleep(400);
    if (stderr.trim()) {
      console.error(stderr);
    }
  }
}

run().catch((error) => {
  console.error('Restore test failed:', error.message);
  process.exit(1);
});
