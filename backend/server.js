const crypto = require('crypto');
const fs = require('fs');
const http = require('http');
const path = require('path');
const { URL } = require('url');

const PORT = process.env.PORT || 8080;
const HOST = process.env.HOST || '127.0.0.1';
const ROOT = path.resolve(__dirname, '..');
const FRONTEND_ROOT = path.join(ROOT, 'frontend');
const DATA_ROOT = path.join(__dirname, 'data');
const STORE_PATH = process.env.STORE_PATH || path.join(DATA_ROOT, 'store.json');
const BACKUP_DIR = path.join(DATA_ROOT, 'backups');
const SESSION_TTL_MS = 1000 * 60 * 60 * 12;
const ACTIVITY_CAP = 1000;

const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml'
};

function nowIso() {
  return new Date().toISOString();
}

function id(prefix) {
  const rand = crypto.randomBytes(3).toString('hex');
  return `${prefix}-${Date.now()}-${rand}`;
}

function randomToken() {
  return crypto.randomBytes(24).toString('hex');
}

function hashPassword(plain, saltHex) {
  const salt = saltHex || crypto.randomBytes(16).toString('hex');
  const iterations = 120000;
  const keylen = 64;
  const digest = 'sha512';
  const hash = crypto.pbkdf2Sync(plain, salt, iterations, keylen, digest).toString('hex');
  return { salt, iterations, keylen, digest, hash };
}

function verifyPassword(plain, saved) {
  if (!saved || !saved.salt || !saved.hash) return false;
  const derived = crypto
    .pbkdf2Sync(plain, saved.salt, saved.iterations, saved.keylen, saved.digest)
    .toString('hex');
  const left = Buffer.from(saved.hash, 'hex');
  const right = Buffer.from(derived, 'hex');
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function seededUsers() {
  const defaults = [
    { username: 'admin', password: 'admin123', role: 'admin', name: 'Platform Administrator' },
    { username: 'agent', password: 'agent123', role: 'agent', name: 'Field Agent' },
    { username: 'farmer', password: 'farmer123', role: 'farmer', name: 'Registered Farmer' }
  ];

  return defaults.map((user) => ({
    id: id('USR'),
    username: user.username,
    name: user.name,
    role: user.role,
    status: 'active',
    password: hashPassword(user.password),
    createdAt: nowIso()
  }));
}

function freshStore() {
  return {
    meta: {
      version: 2,
      createdAt: nowIso(),
      lastWriteAt: nowIso()
    },
    users: seededUsers(),
    farmers: [],
    produce: [],
    payments: [],
    smsLogs: [],
    activityLogs: [],
    sessions: {}
  };
}

function normalizeStore(raw) {
  const store = {
    meta: raw.meta && typeof raw.meta === 'object' ? raw.meta : {},
    users: Array.isArray(raw.users) ? raw.users : [],
    farmers: Array.isArray(raw.farmers) ? raw.farmers : [],
    produce: Array.isArray(raw.produce) ? raw.produce : [],
    payments: Array.isArray(raw.payments) ? raw.payments : [],
    smsLogs: Array.isArray(raw.smsLogs) ? raw.smsLogs : [],
    activityLogs: Array.isArray(raw.activityLogs) ? raw.activityLogs : [],
    sessions: raw.sessions && typeof raw.sessions === 'object' ? raw.sessions : {}
  };

  if (!store.meta.createdAt) store.meta.createdAt = nowIso();
  store.meta.version = 2;

  const usersValid = store.users.some((u) => u && u.username && u.password && u.password.hash);
  if (!usersValid) {
    store.users = seededUsers();
  }

  pruneExpiredSessions(store);
  return store;
}

function ensureDataPaths() {
  fs.mkdirSync(path.dirname(STORE_PATH), { recursive: true });
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
}

function ensureStore() {
  ensureDataPaths();
  if (!fs.existsSync(STORE_PATH)) {
    fs.writeFileSync(STORE_PATH, JSON.stringify(freshStore(), null, 2));
    return;
  }

  const store = readStore();
  writeStore(store);
}

function readStore() {
  ensureDataPaths();
  try {
    const raw = fs.readFileSync(STORE_PATH, 'utf-8');
    return normalizeStore(JSON.parse(raw));
  } catch {
    return freshStore();
  }
}

function writeStore(store) {
  store.meta.lastWriteAt = nowIso();
  fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2));
}

function backupFilename(label) {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  return `${stamp}-${label}.json`;
}

function createBackupSnapshot(store, label) {
  ensureDataPaths();
  const filename = backupFilename(label);
  const filepath = path.join(BACKUP_DIR, filename);
  fs.writeFileSync(filepath, JSON.stringify(store, null, 2));
  return { filename, filepath, createdAt: nowIso() };
}

function addActivity(store, entry) {
  store.activityLogs.unshift({
    id: id('ACT'),
    at: nowIso(),
    actor: entry.actor || 'system',
    role: entry.role || 'system',
    action: entry.action,
    entity: entry.entity || '',
    entityId: entry.entityId || '',
    details: entry.details || ''
  });

  if (store.activityLogs.length > ACTIVITY_CAP) {
    store.activityLogs.length = ACTIVITY_CAP;
  }
}

function json(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,PATCH,DELETE,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization,X-Role,X-Auth-Token'
  });
  res.end(JSON.stringify(payload));
}

function text(res, status, body, headers = {}) {
  res.writeHead(status, {
    'Content-Type': 'text/plain; charset=utf-8',
    ...headers
  });
  res.end(body);
}

function csv(res, filename, data) {
  res.writeHead(200, {
    'Content-Type': 'text/csv; charset=utf-8',
    'Content-Disposition': `attachment; filename="${filename}"`
  });
  res.end(data);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) reject(new Error('Request too large'));
    });
    req.on('end', () => {
      if (!body) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(body));
      } catch {
        reject(new Error('Invalid JSON body'));
      }
    });
    req.on('error', reject);
  });
}

function extractToken(req) {
  const authHeader = (req.headers.authorization || '').toString().trim();
  if (authHeader.toLowerCase().startsWith('bearer ')) {
    return authHeader.slice(7).trim();
  }
  return (req.headers['x-auth-token'] || '').toString().trim();
}

function pruneExpiredSessions(store) {
  const now = Date.now();
  for (const [token, session] of Object.entries(store.sessions)) {
    if (!session || !session.expiresAt || Date.parse(session.expiresAt) <= now) {
      delete store.sessions[token];
    }
  }
}

function currentSession(req, store) {
  pruneExpiredSessions(store);
  const token = extractToken(req);
  if (!token) return null;

  const session = store.sessions[token];
  if (!session) return null;

  return {
    token,
    userId: session.userId,
    username: session.username,
    name: session.name,
    role: session.role,
    createdAt: session.createdAt,
    expiresAt: session.expiresAt
  };
}

function requireSession(req, store, allowedRoles) {
  const session = currentSession(req, store);
  if (!session) {
    return { ok: false, status: 401, error: 'Authentication required. Please sign in.' };
  }

  if (allowedRoles && !allowedRoles.includes(session.role)) {
    return { ok: false, status: 403, error: 'You do not have permission for this action.' };
  }

  return { ok: true, session };
}

function clean(value) {
  return String(value || '').trim();
}

function parseNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : NaN;
}

function findById(list, entityId) {
  return list.find((row) => row.id === entityId);
}

function safeQueryInt(searchParams, key, fallback, max = 1000) {
  const raw = searchParams.get(key);
  if (!raw) return fallback;
  const num = Number(raw);
  if (!Number.isFinite(num) || num < 1) return fallback;
  return Math.min(Math.floor(num), max);
}

function filterByQuery(list, searchParams, fields) {
  const q = clean(searchParams.get('q')).toLowerCase();
  if (!q) return list;
  return list.filter((row) => fields.some((field) => clean(row[field]).toLowerCase().includes(q)));
}

function validateFarmer(payload) {
  if (!clean(payload.name)) return 'name is required';
  if (!clean(payload.phone)) return 'phone is required';
  if (!clean(payload.location)) return 'location is required';
  if (Number.isNaN(parseNumber(payload.trees || 0))) return 'trees must be a number';
  return '';
}

function validateProduce(payload) {
  if (!clean(payload.farmerId)) return 'farmerId is required';
  if (!(parseNumber(payload.kgs) > 0)) return 'kgs must be greater than 0';
  if (!clean(payload.quality)) return 'quality is required';
  if (!clean(payload.agent)) return 'agent is required';
  return '';
}

function validatePayment(payload) {
  if (!clean(payload.farmerId)) return 'farmerId is required';
  if (!(parseNumber(payload.amount) > 0)) return 'amount must be greater than 0';
  if (!clean(payload.ref)) return 'ref is required';
  if (!clean(payload.status)) return 'status is required';
  return '';
}

function normalizePaymentStatus(status) {
  const lower = clean(status).toLowerCase();
  if (lower === 'received') return 'Received';
  if (lower === 'pending') return 'Pending';
  return 'Failed';
}

function toCsv(rows, headers) {
  const head = headers.map((h) => h.label).join(',');
  const body = rows.map((row) => headers.map((h) => csvCell(row[h.key])).join(',')).join('\n');
  return `${head}\n${body}`;
}

function csvCell(value) {
  const raw = value == null ? '' : String(value);
  if (/[",\n]/.test(raw)) {
    return `"${raw.replace(/"/g, '""')}"`;
  }
  return raw;
}

function routeParam(pathname, prefix) {
  if (!pathname.startsWith(prefix)) return '';
  const remainder = pathname.slice(prefix.length);
  const parts = remainder.split('/').filter(Boolean);
  return parts[0] || '';
}

function resolveStaticTarget(pathname) {
  const resolved = pathname === '/' ? '/index.html' : pathname;
  const relativePath = resolved.replace(/^\/+/, '');

  const frontendTarget = path.join(FRONTEND_ROOT, relativePath);
  if (
    frontendTarget.startsWith(FRONTEND_ROOT) &&
    fs.existsSync(frontendTarget) &&
    fs.statSync(frontendTarget).isFile()
  ) {
    return frontendTarget;
  }

  const rootTarget = path.join(ROOT, relativePath);
  if (
    rootTarget.startsWith(ROOT) &&
    fs.existsSync(rootTarget) &&
    fs.statSync(rootTarget).isFile()
  ) {
    return rootTarget;
  }

  return null;
}

function sendFile(res, filePath) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      text(res, 404, 'Not found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      'Content-Type': CONTENT_TYPES[ext] || 'application/octet-stream'
    });
    res.end(data);
  });
}

function makeSummary(store) {
  const totalProduceKg = store.produce.reduce((sum, row) => sum + Number(row.kgs || 0), 0);
  const paymentsReceived = store.payments
    .filter((row) => row.status === 'Received')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);

  const successfulTx = store.payments.filter((row) => row.status === 'Received').length;
  const paymentSuccessRate = store.payments.length
    ? Math.round((successfulTx / store.payments.length) * 100)
    : 0;

  return {
    farmers: store.farmers.length,
    produceRecords: store.produce.length,
    paymentRecords: store.payments.length,
    smsSent: store.smsLogs.length,
    totalProduceKg,
    paymentsReceived,
    paymentSuccessRate,
    launchReady:
      store.farmers.length > 0 &&
      store.produce.length > 0 &&
      store.payments.some((row) => row.status === 'Received')
  };
}

function makeAgentStats(store) {
  const byActor = {};

  for (const farmer of store.farmers) {
    if (!farmer.createdBy) continue;
    if (!byActor[farmer.createdBy]) byActor[farmer.createdBy] = { actor: farmer.createdBy, farmers: 0, produceKg: 0, sms: 0 };
    byActor[farmer.createdBy].farmers += 1;
  }

  for (const produce of store.produce) {
    const actor = produce.createdBy || produce.agent || 'unknown';
    if (!byActor[actor]) byActor[actor] = { actor, farmers: 0, produceKg: 0, sms: 0 };
    byActor[actor].produceKg += Number(produce.kgs || 0);
  }

  for (const sms of store.smsLogs) {
    const actor = sms.createdBy || 'unknown';
    if (!byActor[actor]) byActor[actor] = { actor, farmers: 0, produceKg: 0, sms: 0 };
    byActor[actor].sms += 1;
  }

  return Object.values(byActor)
    .sort((a, b) => b.produceKg - a.produceKg)
    .map((row) => ({ ...row, produceKg: Number(row.produceKg.toFixed(2)) }));
}

const server = http.createServer(async (req, res) => {
  const reqUrl = new URL(req.url, `http://${req.headers.host}`);
  const pathname = reqUrl.pathname;

  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET,POST,PATCH,DELETE,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type,Authorization,X-Role,X-Auth-Token'
    });
    res.end();
    return;
  }

  if (pathname === '/api/health' && req.method === 'GET') {
    json(res, 200, {
      service: 'agem-platform-api',
      status: 'ok',
      now: nowIso()
    });
    return;
  }

  if (pathname === '/api/auth/login' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      const username = clean(payload.username);
      const password = String(payload.password || '');
      const store = readStore();

      const user = store.users.find((row) => row.username === username && row.status === 'active');
      if (!user || !verifyPassword(password, user.password)) {
        json(res, 401, { error: 'Invalid username or password' });
        return;
      }

      const token = randomToken();
      const createdAt = nowIso();
      const expiresAt = new Date(Date.now() + SESSION_TTL_MS).toISOString();

      store.sessions[token] = {
        userId: user.id,
        username: user.username,
        role: user.role,
        name: user.name,
        createdAt,
        expiresAt
      };

      addActivity(store, {
        actor: user.username,
        role: user.role,
        action: 'auth.login',
        entity: 'session',
        details: `User ${user.username} signed in`
      });

      writeStore(store);
      json(res, 200, {
        data: {
          token,
          expiresAt,
          user: {
            id: user.id,
            username: user.username,
            role: user.role,
            name: user.name
          }
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/auth/me' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    json(res, 200, {
      data: {
        user: {
          id: auth.session.userId,
          username: auth.session.username,
          role: auth.session.role,
          name: auth.session.name
        },
        expiresAt: auth.session.expiresAt
      }
    });
    return;
  }

  if (pathname === '/api/auth/logout' && req.method === 'POST') {
    const store = readStore();
    const auth = requireSession(req, store);
    const token = extractToken(req);

    if (auth.ok && token && store.sessions[token]) {
      delete store.sessions[token];
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'auth.logout',
        entity: 'session',
        details: `User ${auth.session.username} signed out`
      });
      writeStore(store);
    }

    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/farmers' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.farmers];
    rows = filterByQuery(rows, reqUrl.searchParams, ['name', 'phone', 'location', 'notes']);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 200, 2000);
    rows = rows.slice(0, limit);
    json(res, 200, { data: rows });
    return;
  }

  if (pathname === '/api/farmers' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const invalid = validateFarmer(payload);
      if (invalid) {
        json(res, 422, { error: invalid });
        return;
      }

      const record = {
        id: id('F'),
        name: clean(payload.name),
        phone: clean(payload.phone),
        location: clean(payload.location),
        trees: Number(parseNumber(payload.trees || 0).toFixed(2)),
        notes: clean(payload.notes),
        createdBy: auth.session.username,
        createdAt: nowIso(),
        updatedAt: nowIso()
      };

      store.farmers.unshift(record);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'farmer.create',
        entity: 'farmer',
        entityId: record.id,
        details: `Created farmer ${record.name}`
      });
      writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/farmers/') && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const farmerId = routeParam(pathname, '/api/farmers/');
    const farmer = findById(store.farmers, farmerId);
    if (!farmer) {
      json(res, 404, { error: 'Farmer not found' });
      return;
    }

    const produce = store.produce.filter((row) => row.farmerId === farmer.id);
    const payments = store.payments.filter((row) => row.farmerId === farmer.id);

    json(res, 200, {
      data: {
        farmer,
        produce,
        payments
      }
    });
    return;
  }

  if (pathname.startsWith('/api/farmers/') && req.method === 'PATCH') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const farmerId = routeParam(pathname, '/api/farmers/');
      const farmer = findById(store.farmers, farmerId);
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found' });
        return;
      }

      const payload = await readBody(req);
      if (payload.name !== undefined) farmer.name = clean(payload.name);
      if (payload.phone !== undefined) farmer.phone = clean(payload.phone);
      if (payload.location !== undefined) farmer.location = clean(payload.location);
      if (payload.notes !== undefined) farmer.notes = clean(payload.notes);
      if (payload.trees !== undefined) {
        const trees = parseNumber(payload.trees);
        if (!Number.isFinite(trees)) {
          json(res, 422, { error: 'trees must be a number' });
          return;
        }
        farmer.trees = Number(trees.toFixed(2));
      }
      farmer.updatedAt = nowIso();

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'farmer.update',
        entity: 'farmer',
        entityId: farmer.id,
        details: `Updated farmer ${farmer.name}`
      });
      writeStore(store);
      json(res, 200, { data: farmer });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/farmers/') && req.method === 'DELETE') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const farmerId = routeParam(pathname, '/api/farmers/');
    const farmerIndex = store.farmers.findIndex((row) => row.id === farmerId);
    if (farmerIndex === -1) {
      json(res, 404, { error: 'Farmer not found' });
      return;
    }

    const linkedProduce = store.produce.some((row) => row.farmerId === farmerId);
    const linkedPayments = store.payments.some((row) => row.farmerId === farmerId);
    if (linkedProduce || linkedPayments) {
      json(res, 409, { error: 'Cannot delete farmer with produce/payment history' });
      return;
    }

    const [removed] = store.farmers.splice(farmerIndex, 1);
    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'farmer.delete',
      entity: 'farmer',
      entityId: removed.id,
      details: `Deleted farmer ${removed.name}`
    });
    writeStore(store);
    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/produce' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.produce];
    const farmerId = clean(reqUrl.searchParams.get('farmerId'));
    if (farmerId) rows = rows.filter((row) => row.farmerId === farmerId);
    rows = filterByQuery(rows, reqUrl.searchParams, ['farmerName', 'quality', 'agent', 'createdBy']);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 300, 2000);
    json(res, 200, { data: rows.slice(0, limit) });
    return;
  }

  if (pathname === '/api/produce' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const invalid = validateProduce(payload);
      if (invalid) {
        json(res, 422, { error: invalid });
        return;
      }

      const farmer = findById(store.farmers, clean(payload.farmerId));
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found' });
        return;
      }

      const record = {
        id: id('P'),
        farmerId: farmer.id,
        farmerName: farmer.name,
        kgs: Number(parseNumber(payload.kgs).toFixed(2)),
        quality: clean(payload.quality),
        agent: clean(payload.agent),
        notes: clean(payload.notes),
        createdBy: auth.session.username,
        createdAt: nowIso()
      };

      store.produce.unshift(record);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'produce.create',
        entity: 'produce',
        entityId: record.id,
        details: `Logged ${record.kgs}kg for ${record.farmerName}`
      });
      writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/produce/') && req.method === 'DELETE') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const produceId = routeParam(pathname, '/api/produce/');
    const index = store.produce.findIndex((row) => row.id === produceId);
    if (index === -1) {
      json(res, 404, { error: 'Produce record not found' });
      return;
    }

    const [removed] = store.produce.splice(index, 1);
    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'produce.delete',
      entity: 'produce',
      entityId: removed.id,
      details: `Deleted produce record ${removed.id}`
    });
    writeStore(store);
    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/payments' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.payments];
    const farmerId = clean(reqUrl.searchParams.get('farmerId'));
    const status = clean(reqUrl.searchParams.get('status'));

    if (farmerId) rows = rows.filter((row) => row.farmerId === farmerId);
    if (status) rows = rows.filter((row) => row.status.toLowerCase() === status.toLowerCase());

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 300, 2000);
    json(res, 200, { data: rows.slice(0, limit) });
    return;
  }

  if (pathname === '/api/payments' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const invalid = validatePayment(payload);
      if (invalid) {
        json(res, 422, { error: invalid });
        return;
      }

      const farmer = findById(store.farmers, clean(payload.farmerId));
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found' });
        return;
      }

      const record = {
        id: id('TX'),
        farmerId: farmer.id,
        farmerName: farmer.name,
        amount: Number(parseNumber(payload.amount).toFixed(2)),
        ref: clean(payload.ref),
        status: normalizePaymentStatus(payload.status),
        method: clean(payload.method) || 'M-PESA',
        notes: clean(payload.notes),
        createdBy: auth.session.username,
        createdAt: nowIso(),
        updatedAt: nowIso()
      };

      store.payments.unshift(record);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'payment.create',
        entity: 'payment',
        entityId: record.id,
        details: `Logged payment ${record.ref} for ${record.farmerName}`
      });
      writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/payments/') && pathname.endsWith('/status') && req.method === 'PATCH') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const paymentId = pathname.split('/')[3] || '';
      const payment = findById(store.payments, paymentId);
      if (!payment) {
        json(res, 404, { error: 'Payment not found' });
        return;
      }

      const payload = await readBody(req);
      payment.status = normalizePaymentStatus(payload.status);
      payment.updatedAt = nowIso();

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'payment.status',
        entity: 'payment',
        entityId: payment.id,
        details: `Changed payment ${payment.ref} to ${payment.status}`
      });
      writeStore(store);
      json(res, 200, { data: payment });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/sms' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.smsLogs];
    rows = filterByQuery(rows, reqUrl.searchParams, ['phone', 'message', 'createdBy']);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 300, 2000);
    json(res, 200, { data: rows.slice(0, limit) });
    return;
  }

  if (pathname === '/api/sms/send' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const farmerId = clean(payload.farmerId);
      let phone = clean(payload.phone);
      const message = clean(payload.message);
      if (!message) {
        json(res, 422, { error: 'message is required' });
        return;
      }

      let farmerName = '';
      if (farmerId) {
        const farmer = findById(store.farmers, farmerId);
        if (!farmer) {
          json(res, 404, { error: 'Farmer not found' });
          return;
        }
        phone = farmer.phone;
        farmerName = farmer.name;
      }

      if (!phone) {
        json(res, 422, { error: 'phone is required when farmerId is not provided' });
        return;
      }

      const log = {
        id: id('SMS'),
        farmerId: farmerId || '',
        farmerName,
        phone,
        message,
        provider: 'AfricaTalking(Mock)',
        status: 'Sent',
        createdBy: auth.session.username,
        createdAt: nowIso()
      };

      store.smsLogs.unshift(log);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'sms.send',
        entity: 'sms',
        entityId: log.id,
        details: `Sent SMS to ${phone}`
      });
      writeStore(store);
      json(res, 201, { data: log });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/reports/summary' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    json(res, 200, { data: makeSummary(store) });
    return;
  }

  if (pathname === '/api/reports/agents' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    json(res, 200, { data: makeAgentStats(store) });
    return;
  }

  if (pathname === '/api/integrations/mpesa/disburse' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      if (!clean(payload.farmerId)) {
        json(res, 422, { error: 'farmerId is required' });
        return;
      }
      if (!(parseNumber(payload.amount) > 0)) {
        json(res, 422, { error: 'amount must be greater than 0' });
        return;
      }

      const farmer = findById(store.farmers, clean(payload.farmerId));
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found' });
        return;
      }

      const ref = `MP${Date.now().toString().slice(-8)}${Math.floor(Math.random() * 90 + 10)}`;
      const payment = {
        id: id('TX'),
        farmerId: farmer.id,
        farmerName: farmer.name,
        amount: Number(parseNumber(payload.amount).toFixed(2)),
        ref,
        status: 'Received',
        method: 'M-PESA(Mock)',
        notes: clean(payload.narration),
        createdBy: auth.session.username,
        createdAt: nowIso(),
        updatedAt: nowIso()
      };

      store.payments.unshift(payment);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'mpesa.disburse',
        entity: 'payment',
        entityId: payment.id,
        details: `Disbursed KES ${payment.amount} to ${farmer.name}`
      });
      writeStore(store);

      json(res, 201, {
        data: payment,
        meta: {
          provider: 'M-PESA(Mock)',
          resultCode: '0',
          resultDesc: 'Accepted'
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/exports/farmers.csv' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const csvData = toCsv(store.farmers, [
      { key: 'id', label: 'id' },
      { key: 'name', label: 'name' },
      { key: 'phone', label: 'phone' },
      { key: 'location', label: 'location' },
      { key: 'trees', label: 'trees' },
      { key: 'notes', label: 'notes' },
      { key: 'createdBy', label: 'createdBy' },
      { key: 'createdAt', label: 'createdAt' }
    ]);
    csv(res, 'farmers.csv', csvData);
    return;
  }

  if (pathname === '/api/exports/produce.csv' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const csvData = toCsv(store.produce, [
      { key: 'id', label: 'id' },
      { key: 'farmerId', label: 'farmerId' },
      { key: 'farmerName', label: 'farmerName' },
      { key: 'kgs', label: 'kgs' },
      { key: 'quality', label: 'quality' },
      { key: 'agent', label: 'agent' },
      { key: 'createdBy', label: 'createdBy' },
      { key: 'createdAt', label: 'createdAt' }
    ]);
    csv(res, 'produce.csv', csvData);
    return;
  }

  if (pathname === '/api/exports/payments.csv' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const csvData = toCsv(store.payments, [
      { key: 'id', label: 'id' },
      { key: 'farmerId', label: 'farmerId' },
      { key: 'farmerName', label: 'farmerName' },
      { key: 'amount', label: 'amount' },
      { key: 'ref', label: 'ref' },
      { key: 'status', label: 'status' },
      { key: 'method', label: 'method' },
      { key: 'createdBy', label: 'createdBy' },
      { key: 'createdAt', label: 'createdAt' }
    ]);
    csv(res, 'payments.csv', csvData);
    return;
  }

  if (pathname === '/api/exports/activity.csv' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const csvData = toCsv(store.activityLogs, [
      { key: 'id', label: 'id' },
      { key: 'at', label: 'at' },
      { key: 'actor', label: 'actor' },
      { key: 'role', label: 'role' },
      { key: 'action', label: 'action' },
      { key: 'entity', label: 'entity' },
      { key: 'entityId', label: 'entityId' },
      { key: 'details', label: 'details' }
    ]);
    csv(res, 'activity.csv', csvData);
    return;
  }

  if (pathname === '/api/admin/backup' && req.method === 'POST') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const snapshot = createBackupSnapshot(store, 'manual');
    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'admin.backup',
      entity: 'backup',
      entityId: snapshot.filename,
      details: `Created backup ${snapshot.filename}`
    });
    writeStore(store);

    json(res, 201, {
      data: {
        filename: snapshot.filename,
        createdAt: snapshot.createdAt
      }
    });
    return;
  }

  if (pathname === '/api/admin/backups' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    ensureDataPaths();
    const files = fs
      .readdirSync(BACKUP_DIR)
      .filter((name) => name.endsWith('.json'))
      .map((name) => {
        const stat = fs.statSync(path.join(BACKUP_DIR, name));
        return {
          filename: name,
          sizeBytes: stat.size,
          createdAt: stat.mtime.toISOString()
        };
      })
      .sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));

    json(res, 200, { data: files });
    return;
  }

  if (pathname === '/api/admin/reset' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      if (payload.confirm !== true) {
        json(res, 422, { error: 'confirm must be true' });
        return;
      }

      const snapshot = createBackupSnapshot(store, 'pre-reset');

      store.farmers = [];
      store.produce = [];
      store.payments = [];
      store.smsLogs = [];
      store.activityLogs = [];
      const currentToken = auth.session.token;
      store.sessions = currentToken && store.sessions[currentToken]
        ? { [currentToken]: store.sessions[currentToken] }
        : {};

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'admin.reset',
        entity: 'system',
        details: `Reset dataset. Backup: ${snapshot.filename}`
      });
      writeStore(store);

      json(res, 200, {
        data: {
          success: true,
          backup: snapshot.filename
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/seed' && req.method === 'POST') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    if (store.farmers.length || store.produce.length || store.payments.length) {
      json(res, 409, { error: 'Store already has data' });
      return;
    }

    const farmerA = {
      id: id('F'),
      name: 'Mercy Achieng',
      phone: '254712330001',
      location: 'Muranga',
      trees: 48,
      notes: 'Group A',
      createdBy: auth.session.username,
      createdAt: nowIso(),
      updatedAt: nowIso()
    };
    const farmerB = {
      id: id('F'),
      name: 'David Mwangi',
      phone: '254712330002',
      location: 'Nyeri',
      trees: 62,
      notes: 'Organic',
      createdBy: auth.session.username,
      createdAt: nowIso(),
      updatedAt: nowIso()
    };

    store.farmers = [farmerA, farmerB];
    store.produce = [
      {
        id: id('P'),
        farmerId: farmerA.id,
        farmerName: farmerA.name,
        kgs: 320.4,
        quality: 'A',
        agent: 'Agent Njoroge',
        notes: 'First harvest batch',
        createdBy: auth.session.username,
        createdAt: nowIso()
      }
    ];
    store.payments = [
      {
        id: id('TX'),
        farmerId: farmerA.id,
        farmerName: farmerA.name,
        amount: 45800,
        ref: `MP${Date.now().toString().slice(-8)}11`,
        status: 'Received',
        method: 'M-PESA(Mock)',
        notes: 'Seed sample payment',
        createdBy: auth.session.username,
        createdAt: nowIso(),
        updatedAt: nowIso()
      }
    ];

    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'seed.load',
      entity: 'system',
      details: 'Loaded starter data'
    });
    writeStore(store);

    json(res, 201, {
      data: {
        farmers: store.farmers,
        produce: store.produce,
        payments: store.payments,
        summary: makeSummary(store)
      }
    });
    return;
  }

  const staticTarget = resolveStaticTarget(pathname);
  if (!staticTarget) {
    text(res, 404, 'Not found');
    return;
  }
  sendFile(res, staticTarget);
});

server.listen(PORT, HOST, () => {
  ensureStore();
  console.log(`Agem MVP server running on http://${HOST}:${PORT}`);
});
