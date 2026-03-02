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
const HECTARES_PER_ACRE = 0.40468564224;
const SQFT_PER_HECTARE = 107639.1041671;
const MIN_PASSWORD_LENGTH = 10;

function envValue(key) {
  return String(process.env[key] || '').trim();
}

const RUNNING_ON_RENDER = Boolean(envValue('RENDER'));
const ALLOW_DEMO_USERS =
  envValue('ALLOW_DEMO_USERS') === 'true' ||
  (!RUNNING_ON_RENDER && envValue('ALLOW_DEMO_USERS') !== 'false');
const ALLOW_FARMER_REGISTRATION = envValue('ALLOW_FARMER_REGISTRATION') !== 'false';
const SYNC_PASSWORDS_FROM_ENV = envValue('SYNC_PASSWORDS_FROM_ENV') === 'true';

const INSECURE_DEMO_ACCOUNTS = [
  { username: 'admin', password: 'admin123', role: 'admin', name: 'Platform Administrator' },
  { username: 'agent', password: 'agent123', role: 'agent', name: 'Field Agent' },
  { username: 'farmer', password: 'farmer123', role: 'farmer', name: 'Registered Farmer' }
];
const RECOVERY_CODE_BY_ROLE = {
  admin: 'ADMIN_RECOVERY_CODE',
  agent: 'AGENT_RECOVERY_CODE'
};

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

function secureEquals(left, right) {
  const leftBuf = Buffer.from(String(left || ''), 'utf-8');
  const rightBuf = Buffer.from(String(right || ''), 'utf-8');
  if (leftBuf.length !== rightBuf.length) return false;
  return crypto.timingSafeEqual(leftBuf, rightBuf);
}

function validateNewPassword(password) {
  const value = String(password || '');
  if (value.length < MIN_PASSWORD_LENGTH) {
    return `Password must be at least ${MIN_PASSWORD_LENGTH} characters long.`;
  }
  if (!/[A-Za-z]/.test(value) || !/\d/.test(value)) {
    return 'Password must include at least one letter and one number.';
  }
  return '';
}

function seededDemoUsers() {
  return INSECURE_DEMO_ACCOUNTS.map((user) => ({
    id: id('USR'),
    username: user.username,
    name: user.name,
    role: user.role,
    status: 'active',
    password: hashPassword(user.password),
    createdAt: nowIso()
  }));
}

function configuredStaffAccounts() {
  const accounts = [];
  const adminUsername = envValue('ADMIN_USERNAME');
  const adminPassword = process.env.ADMIN_PASSWORD || '';
  const adminName = envValue('ADMIN_NAME') || 'Platform Administrator';

  if (adminUsername && adminPassword) {
    accounts.push({
      username: adminUsername,
      password: adminPassword,
      role: 'admin',
      name: adminName
    });
  }

  const agentUsername = envValue('AGENT_USERNAME');
  const agentPassword = process.env.AGENT_PASSWORD || '';
  const agentName = envValue('AGENT_NAME') || 'Field Agent';

  if (agentUsername && agentPassword) {
    accounts.push({
      username: agentUsername,
      password: agentPassword,
      role: 'agent',
      name: agentName
    });
  }

  return accounts;
}

function isInsecureDemoUser(user) {
  const matchedDemo = INSECURE_DEMO_ACCOUNTS.find(
    (demo) => demo.username === user.username && demo.role === user.role
  );
  if (!matchedDemo) return false;
  return verifyPassword(matchedDemo.password, user.password);
}

function disableInsecureDemoUsers(store) {
  let changed = false;

  for (const user of store.users) {
    if (!user || !user.password || !user.password.hash) continue;
    if (!isInsecureDemoUser(user)) continue;
    if (user.status === 'disabled') continue;

    user.status = 'disabled';
    user.updatedAt = nowIso();
    changed = true;
  }

  return changed;
}

function upsertConfiguredStaffUsers(store) {
  const accounts = configuredStaffAccounts();
  let changed = false;

  for (const account of accounts) {
    const accountKey = normalizedUsername(account.username);
    const existing = store.users.find((row) => normalizedUsername(row.username) === accountKey);
    if (!existing) {
      store.users.push({
        id: id('USR'),
        username: account.username,
        name: account.name,
        role: account.role,
        status: 'active',
        password: hashPassword(account.password),
        createdAt: nowIso(),
        updatedAt: nowIso()
      });
      changed = true;
      continue;
    }

    const roleChanged = existing.role !== account.role;
    const nameChanged = existing.name !== account.name;
    const statusChanged = existing.status !== 'active';
    const hasPasswordHash = Boolean(existing.password && existing.password.hash);
    const passwordChanged = SYNC_PASSWORDS_FROM_ENV && hasPasswordHash
      ? !verifyPassword(account.password, existing.password)
      : false;

    if (roleChanged || nameChanged || statusChanged || passwordChanged || !hasPasswordHash) {
      existing.role = account.role;
      existing.name = account.name;
      existing.status = 'active';
      if (!hasPasswordHash || passwordChanged) {
        existing.password = hashPassword(account.password);
      }
      existing.updatedAt = nowIso();
      changed = true;
    }

    for (const user of store.users) {
      if (!user || user.id === existing.id) continue;
      if (user.role !== account.role) continue;
      if (user.status !== 'active') continue;
      user.status = 'disabled';
      user.updatedAt = nowIso();
      changed = true;
    }
  }

  return changed;
}

function activeUserCount(store) {
  return store.users.filter((u) => u && u.status === 'active' && u.password && u.password.hash).length;
}

function applyUserPolicy(store) {
  let changed = false;

  if (!ALLOW_DEMO_USERS) {
    changed = disableInsecureDemoUsers(store) || changed;
  }

  changed = upsertConfiguredStaffUsers(store) || changed;
  return changed;
}

function accountByRole(store, role) {
  return store.users.find((user) => user && user.status === 'active' && user.role === role);
}

function recoveryCodeForRole(role) {
  const envKey = RECOVERY_CODE_BY_ROLE[role];
  if (!envKey) return '';
  return process.env[envKey] || '';
}

function invalidateUserSessions(store, userId, exceptToken = '') {
  for (const [token, session] of Object.entries(store.sessions)) {
    if (!session || session.userId !== userId) continue;
    if (token === exceptToken) continue;
    delete store.sessions[token];
  }
}

function freshStore() {
  return {
    meta: {
      version: 3,
      createdAt: nowIso(),
      lastWriteAt: nowIso(),
      authMode: ALLOW_DEMO_USERS ? 'demo' : 'private'
    },
    users: ALLOW_DEMO_USERS ? seededDemoUsers() : [],
    farmers: [],
    produce: [],
    producePurchases: [],
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
    producePurchases: Array.isArray(raw.producePurchases) ? raw.producePurchases : [],
    payments: Array.isArray(raw.payments) ? raw.payments : [],
    smsLogs: Array.isArray(raw.smsLogs) ? raw.smsLogs : [],
    activityLogs: Array.isArray(raw.activityLogs) ? raw.activityLogs : [],
    sessions: raw.sessions && typeof raw.sessions === 'object' ? raw.sessions : {}
  };

  if (!store.meta.createdAt) store.meta.createdAt = nowIso();
  store.meta.version = 3;
  store.meta.authMode = ALLOW_DEMO_USERS ? 'demo' : 'private';

  const usersValid = store.users.some((u) => u && u.username && u.password && u.password.hash);
  if (!usersValid) {
    store.users = ALLOW_DEMO_USERS ? seededDemoUsers() : [];
  }

  for (const farmer of store.farmers) {
    if (!farmer || typeof farmer !== 'object') continue;
    const totalHectares = Number(farmer.hectares);
    if (!Number.isFinite(totalHectares) || totalHectares <= 0) continue;

    const avocadoHectares = Number(farmer.avocadoHectares);
    if (!Number.isFinite(avocadoHectares) || avocadoHectares <= 0 || avocadoHectares > totalHectares) {
      farmer.avocadoHectares = Number(totalHectares.toFixed(3));
    } else {
      farmer.avocadoHectares = Number(avocadoHectares.toFixed(3));
    }
  }

  reconcileProducePurchases(store);
  applyUserPolicy(store);
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

function readBody(req, maxBytes = 1_000_000) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > maxBytes) reject(new Error('Request too large'));
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

  const currentUser = store.users.find(
    (user) => user && user.status === 'active' && (user.id === session.userId || user.username === session.username)
  );
  if (!currentUser) {
    delete store.sessions[token];
    return null;
  }

  return {
    token,
    userId: currentUser.id,
    username: currentUser.username,
    name: currentUser.name,
    role: currentUser.role,
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

function normalizedUsername(value) {
  return clean(value).toLowerCase();
}

function findUserByUsername(store, username) {
  const target = normalizedUsername(username);
  return store.users.find((row) => normalizedUsername(row.username) === target);
}

function parseNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : NaN;
}

function parseTrees(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return 0;
  const num = Number(raw.replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function parseArea(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return NaN;
  const num = Number(raw.replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function resolveHectaresFromValues(hectaresRaw, acresRaw, squareFeetRaw) {
  const hasHectares = clean(hectaresRaw) !== '';
  const hasAcres = clean(acresRaw) !== '';
  const hasSquareFeet = clean(squareFeetRaw) !== '';

  if (hasHectares) return parseArea(hectaresRaw);
  if (hasAcres) {
    const acres = parseArea(acresRaw);
    if (!Number.isFinite(acres)) return NaN;
    return acres * HECTARES_PER_ACRE;
  }
  if (hasSquareFeet) {
    const squareFeet = parseArea(squareFeetRaw);
    if (!Number.isFinite(squareFeet)) return NaN;
    return squareFeet / SQFT_PER_HECTARE;
  }
  return NaN;
}

function resolveHectares(payload) {
  const hectaresRaw = payload?.hectares;
  const acresRaw = payload?.acres;
  const squareFeetRaw =
    payload?.squareFeet ?? payload?.squarefeet ?? payload?.square_feet ?? payload?.sqft ?? payload?.squareFt;
  return resolveHectaresFromValues(hectaresRaw, acresRaw, squareFeetRaw);
}

function resolveAvocadoHectares(payload) {
  const hectaresRaw =
    payload?.avocadoHectares ??
    payload?.avocadohectares ??
    payload?.areaUnderAvocadoHectares ??
    payload?.areaunderavocadohectares;
  const acresRaw =
    payload?.avocadoAcres ??
    payload?.avocadoacres ??
    payload?.areaUnderAvocadoAcres ??
    payload?.areaunderavocadoacres ??
    payload?.avocadoAcreage ??
    payload?.avocadoacreage;
  const squareFeetRaw =
    payload?.avocadoSquareFeet ??
    payload?.avocadosquarefeet ??
    payload?.avocado_square_feet ??
    payload?.avocadoSqft ??
    payload?.avocadosqft ??
    payload?.areaUnderAvocadoSquareFeet ??
    payload?.areaunderavocadosquarefeet ??
    payload?.areaUnderAvocadoSqft ??
    payload?.areaunderavocadosqft;
  return resolveHectaresFromValues(hectaresRaw, acresRaw, squareFeetRaw);
}

function findById(list, entityId) {
  return list.find((row) => row.id === entityId);
}

function findFarmerByPhone(list, phone, excludeId = '') {
  const target = clean(phone);
  if (!target) return null;
  return list.find((row) => clean(row.phone) === target && row.id !== excludeId) || null;
}

function normalizedNationalId(value) {
  return clean(value)
    .toUpperCase()
    .replace(/[\s-]/g, '');
}

function cleanNationalId(value) {
  return clean(value).toUpperCase();
}

function findFarmerByNationalId(list, nationalId, excludeId = '') {
  const target = normalizedNationalId(nationalId);
  if (!target) return null;
  return list.find((row) => normalizedNationalId(row.nationalId) === target && row.id !== excludeId) || null;
}

function safeQueryInt(searchParams, key, fallback, max = 1000) {
  const raw = searchParams.get(key);
  if (!raw) return fallback;
  const num = Number(raw);
  if (!Number.isFinite(num) || num < 1) return fallback;
  return Math.min(Math.floor(num), max);
}

function safeQueryOffset(searchParams, key, fallback = 0, max = 2_000_000) {
  const raw = searchParams.get(key);
  if (raw === null || raw === undefined || raw === '') return fallback;
  const num = Number(raw);
  if (!Number.isFinite(num) || num < 0) return fallback;
  return Math.min(Math.floor(num), max);
}

function filterByTextQuery(list, query, fields) {
  const q = clean(query).toLowerCase();
  if (!q) return list;
  return list.filter((row) => fields.some((field) => clean(row[field]).toLowerCase().includes(q)));
}

function filterByQuery(list, searchParams, fields) {
  return filterByTextQuery(list, searchParams.get('q'), fields);
}

function validateFarmer(payload) {
  if (!clean(payload.name)) return 'name is required';
  if (!clean(payload.phone)) return 'phone is required';
  if (!clean(payload.nationalId)) return 'nationalId is required';
  if (!clean(payload.location)) return 'location is required';
  if (Number.isNaN(parseTrees(payload.trees))) return 'trees must be a number';
  const hectares = resolveHectares(payload);
  const avocadoHectares = resolveAvocadoHectares(payload);
  if (!Number.isFinite(hectares)) return 'hectares/acres/square feet is required';
  if (!Number.isFinite(avocadoHectares)) return 'avocadoHectares/avocadoAcres/avocadoSquareFeet is required';
  if (hectares <= 0) return 'hectares must be greater than 0';
  if (avocadoHectares <= 0) return 'area under avocado must be greater than 0';
  if (avocadoHectares > hectares) return 'area under avocado cannot be greater than total farm size';
  return '';
}

function normalizeAvocadoVariety(value) {
  const upper = clean(value).toUpperCase();
  if (upper === 'HASS') return 'Hass';
  if (upper === 'FUERTE') return 'Fuerte';
  return '';
}

function normalizeVisualGrade(value) {
  const upper = clean(value).toUpperCase();
  if (upper === 'PASS') return 'Pass';
  if (upper === 'BORDERLINE') return 'Borderline';
  if (upper === 'REJECT') return 'Reject';
  return '';
}

function normalizeQcDecision(value) {
  const upper = clean(value).toUpperCase();
  if (upper === 'ACCEPT') return 'Accept';
  if (upper === 'HOLD') return 'Hold';
  if (upper === 'REJECT') return 'Reject';
  return '';
}

function normalizeFirmnessUnit(value) {
  const upper = clean(value).toUpperCase();
  if (upper === 'N') return 'N';
  if (upper === 'KGF') return 'kgf';
  return '';
}

function normalizeSizeCode(value) {
  const upper = clean(value).toUpperCase();
  if (!upper) return '';
  const numeric = upper.startsWith('C') ? upper.slice(1) : upper;
  const allowed = new Set(['12', '14', '16', '18', '20', '22', '24', '26', '28']);
  if (!allowed.has(numeric)) return '';
  return `C${numeric}`;
}

function validateProduce(payload) {
  if (!clean(payload.farmerId)) return 'farmerId is required';
  if (!(parseNumber(payload.kgs ?? payload.lotWeightKgs) > 0)) return 'lotWeightKgs must be greater than 0';
  if (!normalizeAvocadoVariety(payload.variety)) return 'variety must be Hass or Fuerte';

  const sampleSize = parseNumber(payload.sampleSize);
  if (!Number.isFinite(sampleSize) || sampleSize < 1 || !Number.isInteger(sampleSize)) {
    return 'sampleSize must be a whole number greater than 0';
  }

  if (!normalizeVisualGrade(payload.visualGrade)) return 'visualGrade must be Pass, Borderline, or Reject';

  const dryMatterRaw = clean(payload.dryMatterPct);
  if (dryMatterRaw) {
    const dryMatterPct = parseNumber(dryMatterRaw);
    if (!Number.isFinite(dryMatterPct) || dryMatterPct <= 0 || dryMatterPct > 100) {
      return 'dryMatterPct must be between 0 and 100 when provided';
    }
  }

  if (!(parseNumber(payload.firmnessValue) > 0)) return 'firmnessValue must be greater than 0';
  if (!normalizeFirmnessUnit(payload.firmnessUnit)) return 'firmnessUnit must be N or kgf';
  if (!(parseNumber(payload.avgFruitWeightG) > 0)) return 'avgFruitWeightG must be greater than 0';
  if (!normalizeSizeCode(payload.sizeCode)) return 'sizeCode must be one of C12, C14, C16, C18, C20, C22, C24, C26, C28';
  if (!normalizeQcDecision(payload.qcDecision)) return 'qcDecision must be Accept, Hold, or Reject';
  if (!clean(payload.inspector ?? payload.agent)) return 'inspector is required';
  return '';
}

function validateProducePurchase(payload) {
  if (!clean(payload.farmerId)) return 'farmerId is required';
  if (!(parseNumber(payload.purchasedKgs) > 0)) return 'purchasedKgs must be greater than 0';

  const varietyRaw = clean(payload.variety);
  if (varietyRaw && !normalizeAvocadoVariety(varietyRaw)) return 'variety must be Hass or Fuerte when provided';

  const unitPriceRaw = clean(payload.pricePerKgKes);
  if (unitPriceRaw && !(parseNumber(unitPriceRaw) > 0)) return 'pricePerKgKes must be greater than 0 when provided';

  const valueRaw = clean(payload.purchaseValueKes);
  if (valueRaw && !(parseNumber(valueRaw) > 0)) return 'purchaseValueKes must be greater than 0 when provided';

  if (!(parseNumber(payload.pricePerKgKes) > 0) && !(parseNumber(payload.purchaseValueKes) > 0)) {
    return 'Provide either pricePerKgKes or purchaseValueKes';
  }

  if (!clean(payload.buyer)) return 'buyer is required';
  return '';
}

function validatePayment(payload) {
  if (!clean(payload.farmerId)) return 'farmerId is required';
  if (!(parseNumber(payload.amount) > 0)) return 'amount must be greater than 0';
  if (!clean(payload.ref)) return 'ref is required';
  if (!clean(payload.status)) return 'status is required';
  return '';
}

function normalizeImportHeader(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function firstNonEmpty(values) {
  for (const value of values) {
    const cleaned = clean(value);
    if (cleaned) return cleaned;
  }
  return '';
}

function mapFarmerImportRecord(raw) {
  const flat = {};
  for (const [key, value] of Object.entries(raw || {})) {
    flat[normalizeImportHeader(key)] = value;
  }

  const name = firstNonEmpty([
    flat.name,
    flat.farmername,
    flat.fullname,
    flat.farmerfullname,
    flat.growername
  ]);
  const phone = firstNonEmpty([
    flat.phone,
    flat.phonenumber,
    flat.mobile,
    flat.mobilenumber,
    flat.msisdn,
    flat.contactnumber
  ]);
  const nationalId = firstNonEmpty([
    flat.nationalid,
    flat.nationalidnumber,
    flat.idnumber,
    flat.idno,
    flat.nationalidno,
    flat.governmentid,
    flat.governmentidnumber
  ]);
  const location = firstNonEmpty([
    flat.location,
    flat.area,
    flat.ward,
    flat.county,
    flat.village,
    flat.region
  ]);
  const notes = firstNonEmpty([flat.notes, flat.note, flat.comments, flat.remarks, flat.description]);
  const treesRaw = firstNonEmpty([flat.trees, flat.treecount, flat.numberoftrees, flat.treequantity, flat.treenumber]);
  const trees = parseTrees(treesRaw);
  const hectaresRaw = firstNonEmpty([flat.hectares, flat.hectare, flat.farmsizeha, flat.farmsizehectares, flat.landsizehectares]);
  const acresRaw = firstNonEmpty([flat.acres, flat.acre, flat.acreage, flat.farmsizeacres, flat.landsizeacres]);
  const squareFeetRaw = firstNonEmpty([
    flat.squarefeet,
    flat.squarefoot,
    flat.squareft,
    flat.sqft,
    flat.ft2,
    flat.squarefeetarea,
    flat.farmsizesquarefeet,
    flat.landsizesquarefeet
  ]);
  const avocadoHectaresRaw = firstNonEmpty([
    flat.avocadohectares,
    flat.areaunderavocadohectares,
    flat.avocadoareahectares,
    flat.avocadoplothectares
  ]);
  const avocadoAcresRaw = firstNonEmpty([
    flat.avocadoacres,
    flat.avocadoacreage,
    flat.areaunderavocadoacres,
    flat.avocadoareaacres,
    flat.avocadoplotacres
  ]);
  const avocadoSquareFeetRaw = firstNonEmpty([
    flat.avocadosquarefeet,
    flat.avocadosquarefoot,
    flat.avocadosqft,
    flat.areaunderavocadosquarefeet,
    flat.areaunderavocadosqft
  ]);
  const hectares = resolveHectares({ hectares: hectaresRaw, acres: acresRaw, squareFeet: squareFeetRaw });
  const avocadoHectares = resolveAvocadoHectares({
    avocadoHectares: avocadoHectaresRaw,
    avocadoAcres: avocadoAcresRaw,
    avocadoSquareFeet: avocadoSquareFeetRaw
  });

  return {
    name,
    phone,
    nationalId,
    location,
    trees,
    hectares,
    avocadoHectares,
    notes
  };
}

function normalizePaymentStatus(status) {
  const lower = clean(status).toLowerCase();
  if (lower === 'received') return 'Received';
  if (lower === 'pending') return 'Pending';
  return 'Failed';
}

function money(value) {
  return Number(parseNumber(value).toFixed(2));
}

function valueFromPurchase(purchase) {
  const explicitValue = parseNumber(purchase.purchaseValueKes);
  if (explicitValue > 0) return money(explicitValue);
  const pricePerKg = parseNumber(purchase.pricePerKgKes);
  const purchasedKgs = parseNumber(purchase.purchasedKgs);
  if (pricePerKg > 0 && purchasedKgs > 0) return money(pricePerKg * purchasedKgs);
  return 0;
}

function dateMs(value) {
  const ms = new Date(value).getTime();
  return Number.isFinite(ms) ? ms : 0;
}

function parseDateBound(raw, isEnd) {
  const value = clean(raw);
  if (!value) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const stamp = isEnd ? `${value}T23:59:59.999Z` : `${value}T00:00:00.000Z`;
    const ms = new Date(stamp).getTime();
    return Number.isFinite(ms) ? ms : null;
  }

  const ms = new Date(value).getTime();
  return Number.isFinite(ms) ? ms : null;
}

function endOfDayUtc(now) {
  const date = new Date(now);
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), 23, 59, 59, 999);
}

function startOfDayUtc(now) {
  const date = new Date(now);
  return Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), 0, 0, 0, 0);
}

function resolveRange(period, fromRaw, toRaw) {
  const normalizedPeriod = clean(period).toLowerCase();
  const now = Date.now();

  let from = parseDateBound(fromRaw, false);
  let to = parseDateBound(toRaw, true);

  if (from !== null && to !== null && from > to) {
    const swap = from;
    from = to;
    to = swap;
  }

  if (from === null && to === null) {
    if (normalizedPeriod === 'today' || normalizedPeriod === 'day') {
      from = startOfDayUtc(now);
      to = endOfDayUtc(now);
    } else if (normalizedPeriod === 'week') {
      to = endOfDayUtc(now);
      from = startOfDayUtc(now - 6 * 24 * 60 * 60 * 1000);
    } else if (normalizedPeriod === 'month') {
      const nowDate = new Date(now);
      from = Date.UTC(nowDate.getUTCFullYear(), nowDate.getUTCMonth(), 1, 0, 0, 0, 0);
      to = endOfDayUtc(now);
    }
  }

  return {
    period: normalizedPeriod || 'all',
    from,
    to
  };
}

function isInRange(createdAt, range) {
  const ms = dateMs(createdAt);
  if (range.from !== null && ms < range.from) return false;
  if (range.to !== null && ms > range.to) return false;
  return true;
}

function reconcileProducePurchases(store) {
  if (!Array.isArray(store.producePurchases)) {
    store.producePurchases = [];
  }

  const receivedByFarmer = {};
  for (const payment of store.payments || []) {
    if (normalizePaymentStatus(payment.status) !== 'Received') continue;
    const farmerId = clean(payment.farmerId);
    if (!farmerId) continue;
    receivedByFarmer[farmerId] = money((receivedByFarmer[farmerId] || 0) + money(payment.amount));
  }

  const purchasesChronological = [...store.producePurchases].sort((a, b) => dateMs(a.createdAt) - dateMs(b.createdAt));
  for (const purchase of purchasesChronological) {
    if (!purchase || typeof purchase !== 'object') continue;
    const farmerId = clean(purchase.farmerId);
    const purchasedKgs = money(purchase.purchasedKgs);
    purchase.purchasedKgs = purchasedKgs;

    const pricePerKg = parseNumber(purchase.pricePerKgKes);
    purchase.pricePerKgKes = pricePerKg > 0 ? money(pricePerKg) : null;

    const totalValue = valueFromPurchase(purchase);
    purchase.purchaseValueKes = totalValue > 0 ? totalValue : null;

    const available = money(receivedByFarmer[farmerId] || 0);
    const paidAmountKes = totalValue > 0 ? money(Math.min(totalValue, available)) : 0;
    const balanceKes = totalValue > 0 ? money(totalValue - paidAmountKes) : 0;

    purchase.paidAmountKes = paidAmountKes;
    purchase.balanceKes = balanceKes;

    if (totalValue <= 0) {
      purchase.settlementStatus = 'Unpriced';
    } else if (paidAmountKes <= 0) {
      purchase.settlementStatus = 'Unpaid';
    } else if (balanceKes > 0) {
      purchase.settlementStatus = 'Partially Paid';
    } else {
      purchase.settlementStatus = 'Paid';
    }

    receivedByFarmer[farmerId] = money(Math.max(0, available - paidAmountKes));
  }
}

function buildOwedRows(store, options = {}) {
  reconcileProducePurchases(store);

  const range = resolveRange(options.period, options.from, options.to);
  const farmersById = new Map((store.farmers || []).map((row) => [row.id, row]));
  const byFarmer = new Map();

  for (const purchase of store.producePurchases || []) {
    if (!isInRange(purchase.createdAt, range)) continue;

    const farmerId = clean(purchase.farmerId);
    if (!farmerId) continue;

    const balanceKes = money(purchase.balanceKes);
    if (!(balanceKes > 0)) continue;

    const farmer = farmersById.get(farmerId) || {};
    if (!byFarmer.has(farmerId)) {
      byFarmer.set(farmerId, {
        farmerId,
        farmerName: clean(purchase.farmerName) || clean(farmer.name),
        phone: clean(farmer.phone),
        nationalId: clean(farmer.nationalId),
        location: clean(farmer.location),
        purchaseCount: 0,
        purchasedKgs: 0,
        totalValueKes: 0,
        paidKes: 0,
        balanceKes: 0,
        lastPurchaseAt: ''
      });
    }

    const row = byFarmer.get(farmerId);
    row.purchaseCount += 1;
    row.purchasedKgs = money(row.purchasedKgs + money(purchase.purchasedKgs));
    row.totalValueKes = money(row.totalValueKes + money(purchase.purchaseValueKes));
    row.paidKes = money(row.paidKes + money(purchase.paidAmountKes));
    row.balanceKes = money(row.balanceKes + balanceKes);

    if (!row.lastPurchaseAt || dateMs(purchase.createdAt) > dateMs(row.lastPurchaseAt)) {
      row.lastPurchaseAt = purchase.createdAt;
    }
  }

  let rows = Array.from(byFarmer.values())
    .filter((row) => row.balanceKes > 0)
    .sort((a, b) => {
      if (b.balanceKes !== a.balanceKes) return b.balanceKes - a.balanceKes;
      return dateMs(b.lastPurchaseAt) - dateMs(a.lastPurchaseAt);
    });

  const q = clean(options.q).toLowerCase();
  if (q) {
    rows = rows.filter((row) =>
      [row.farmerName, row.phone, row.nationalId, row.location].some((value) => clean(value).toLowerCase().includes(q))
    );
  }

  const totalBalanceKes = money(rows.reduce((sum, row) => sum + money(row.balanceKes), 0));
  return {
    rows,
    meta: {
      period: range.period,
      from: range.from ? new Date(range.from).toISOString() : '',
      to: range.to ? new Date(range.to).toISOString() : '',
      count: rows.length,
      totalBalanceKes
    }
  };
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
  reconcileProducePurchases(store);
  const totalProduceKg = store.produce.reduce((sum, row) => sum + Number(row.kgs || 0), 0);
  const totalPurchasedKg = store.producePurchases.reduce((sum, row) => sum + Number(row.purchasedKgs || 0), 0);
  const purchasedValueKes = store.producePurchases.reduce((sum, row) => sum + Number(row.purchaseValueKes || 0), 0);
  const owed = buildOwedRows(store, { period: 'all' });
  const paymentsReceived = store.payments
    .filter((row) => row.status === 'Received')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);

  const successfulTx = store.payments.filter((row) => row.status === 'Received').length;
  const paymentSuccessRate = store.payments.length
    ? Math.round((successfulTx / store.payments.length) * 100)
    : 0;

  return {
    farmers: store.farmers.length,
    qcRecords: store.produce.length,
    produceRecords: store.produce.length,
    purchasedRecords: store.producePurchases.length,
    paymentRecords: store.payments.length,
    smsSent: store.smsLogs.length,
    totalProduceKg,
    totalPurchasedKg,
    purchasedValueKes,
    owedFarmers: owed.meta.count || 0,
    totalOwedKes: owed.meta.totalBalanceKes || 0,
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
    if (!byActor[farmer.createdBy]) {
      byActor[farmer.createdBy] = { actor: farmer.createdBy, farmers: 0, produceKg: 0, purchasedKg: 0, sms: 0 };
    }
    byActor[farmer.createdBy].farmers += 1;
  }

  for (const produce of store.produce) {
    const actor = produce.createdBy || produce.agent || 'unknown';
    if (!byActor[actor]) byActor[actor] = { actor, farmers: 0, produceKg: 0, purchasedKg: 0, sms: 0 };
    byActor[actor].produceKg += Number(produce.kgs || 0);
  }

  for (const purchase of store.producePurchases) {
    const actor = purchase.createdBy || purchase.buyer || 'unknown';
    if (!byActor[actor]) byActor[actor] = { actor, farmers: 0, produceKg: 0, purchasedKg: 0, sms: 0 };
    byActor[actor].purchasedKg += Number(purchase.purchasedKgs || 0);
  }

  for (const sms of store.smsLogs) {
    const actor = sms.createdBy || 'unknown';
    if (!byActor[actor]) byActor[actor] = { actor, farmers: 0, produceKg: 0, purchasedKg: 0, sms: 0 };
    byActor[actor].sms += 1;
  }

  return Object.values(byActor)
    .sort((a, b) => b.produceKg - a.produceKg)
    .map((row) => ({
      ...row,
      produceKg: Number(row.produceKg.toFixed(2)),
      purchasedKg: Number((row.purchasedKg || 0).toFixed(2))
    }));
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
      const username = normalizedUsername(payload.username);
      const password = String(payload.password || '');
      const store = readStore();

      if (activeUserCount(store) === 0) {
        json(res, 503, {
          error:
            'No active staff accounts configured. Set ADMIN_USERNAME and ADMIN_PASSWORD in environment variables, then redeploy.'
        });
        return;
      }

      const user = findUserByUsername(store, username);
      if (!user || user.status !== 'active' || !verifyPassword(password, user.password)) {
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

  if (pathname === '/api/auth/register-farmer' && req.method === 'POST') {
    try {
      if (!ALLOW_FARMER_REGISTRATION) {
        json(res, 403, { error: 'Farmer self-registration is disabled.' });
        return;
      }

      const payload = await readBody(req);
      const name = clean(payload.name);
      const phone = clean(payload.phone);
      const nationalId = cleanNationalId(payload.nationalId);
      const location = clean(payload.location);
      const username = normalizedUsername(payload.username);
      const password = String(payload.password || '');
      const confirmPassword = String(payload.confirmPassword || '');
      const notes = clean(payload.notes);
      const treesValue = payload.trees === undefined ? 0 : parseTrees(payload.trees);
      const trees = Number.isFinite(treesValue) ? Number(treesValue.toFixed(2)) : NaN;
      const hectares = resolveHectares(payload);
      const avocadoHectares = resolveAvocadoHectares(payload);

      if (!name || !phone || !nationalId || !location || !username || !password || !confirmPassword) {
        json(res, 422, { error: 'name, phone, nationalId, location, total area, area under avocado, username, password, and confirmPassword are required.' });
        return;
      }
      if (!/^[a-z0-9._-]{3,32}$/.test(username)) {
        json(res, 422, { error: 'Username must be 3-32 chars and contain only letters, numbers, dot, underscore, or dash.' });
        return;
      }
      if (!secureEquals(password, confirmPassword)) {
        json(res, 422, { error: 'Password confirmation does not match.' });
        return;
      }
      if (!Number.isFinite(trees) || trees < 0) {
        json(res, 422, { error: 'trees must be a number greater than or equal to 0.' });
        return;
      }
      if (!Number.isFinite(hectares) || hectares <= 0) {
        json(res, 422, { error: 'hectares/acres/square feet must be a number greater than 0.' });
        return;
      }
      if (!Number.isFinite(avocadoHectares) || avocadoHectares <= 0) {
        json(res, 422, { error: 'avocadoHectares/avocadoAcres/avocadoSquareFeet must be a number greater than 0.' });
        return;
      }
      if (avocadoHectares > hectares) {
        json(res, 422, { error: 'area under avocado cannot be greater than total farm size.' });
        return;
      }

      const invalidPassword = validateNewPassword(password);
      if (invalidPassword) {
        json(res, 422, { error: invalidPassword });
        return;
      }

      const store = readStore();
      if (findUserByUsername(store, username)) {
        json(res, 409, { error: 'Username is already taken.' });
        return;
      }
      if (findFarmerByPhone(store.farmers, phone)) {
        json(res, 409, { error: 'A farmer with this phone number already exists.' });
        return;
      }
      if (findFarmerByNationalId(store.farmers, nationalId)) {
        json(res, 409, { error: 'A farmer with this National ID already exists.' });
        return;
      }

      const user = {
        id: id('USR'),
        username,
        name,
        role: 'farmer',
        status: 'active',
        password: hashPassword(password),
        createdAt: nowIso(),
        updatedAt: nowIso()
      };
      store.users.push(user);

      const farmer = {
        id: id('F'),
        name,
        phone,
        nationalId,
        location,
        trees,
        hectares: Number(hectares.toFixed(3)),
        avocadoHectares: Number(avocadoHectares.toFixed(3)),
        notes,
        createdBy: user.username,
        createdAt: nowIso(),
        updatedAt: nowIso()
      };
      store.farmers.unshift(farmer);

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
        action: 'auth.register',
        entity: 'user',
        entityId: user.id,
        details: `Self-registered farmer account ${user.username}`
      });
      addActivity(store, {
        actor: user.username,
        role: user.role,
        action: 'farmer.create',
        entity: 'farmer',
        entityId: farmer.id,
        details: `Self-registered farmer profile ${farmer.name}`
      });

      writeStore(store);
      json(res, 201, {
        data: {
          token,
          expiresAt,
          user: {
            id: user.id,
            username: user.username,
            role: user.role,
            name: user.name
          },
          farmerId: farmer.id
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/auth/change-password' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const currentPassword = String(payload.currentPassword || '');
      const newPassword = String(payload.newPassword || '');
      const confirmPassword = payload.confirmPassword === undefined
        ? newPassword
        : String(payload.confirmPassword || '');

      if (!currentPassword || !newPassword) {
        json(res, 422, { error: 'Current and new password are required.' });
        return;
      }
      if (!secureEquals(newPassword, confirmPassword)) {
        json(res, 422, { error: 'New password and confirmation do not match.' });
        return;
      }
      if (secureEquals(currentPassword, newPassword)) {
        json(res, 422, { error: 'New password must be different from current password.' });
        return;
      }

      const invalidPassword = validateNewPassword(newPassword);
      if (invalidPassword) {
        json(res, 422, { error: invalidPassword });
        return;
      }

      const user = store.users.find((row) => row.id === auth.session.userId && row.status === 'active');
      if (!user) {
        json(res, 401, { error: 'Session is no longer valid. Please sign in again.' });
        return;
      }
      if (!verifyPassword(currentPassword, user.password)) {
        json(res, 401, { error: 'Current password is incorrect.' });
        return;
      }

      user.password = hashPassword(newPassword);
      user.updatedAt = nowIso();

      const currentToken = extractToken(req);
      invalidateUserSessions(store, user.id, currentToken);

      addActivity(store, {
        actor: user.username,
        role: user.role,
        action: 'auth.password_change',
        entity: 'user',
        entityId: user.id,
        details: 'Password changed by signed-in user.'
      });

      writeStore(store);
      json(res, 200, { data: { success: true } });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/auth/recover-username' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      const role = clean(payload.role).toLowerCase();
      const recoveryCode = String(payload.recoveryCode || '');

      if (!role || !RECOVERY_CODE_BY_ROLE[role]) {
        json(res, 422, { error: 'Role must be admin or agent.' });
        return;
      }
      if (!recoveryCode) {
        json(res, 422, { error: 'Recovery code is required.' });
        return;
      }

      const expectedCode = recoveryCodeForRole(role);
      if (!expectedCode) {
        json(res, 503, { error: `Recovery is not configured for role "${role}".` });
        return;
      }
      if (!secureEquals(recoveryCode, expectedCode)) {
        json(res, 401, { error: 'Recovery code is invalid.' });
        return;
      }

      const store = readStore();
      const user = accountByRole(store, role);
      if (!user) {
        json(res, 404, { error: `No active ${role} account found.` });
        return;
      }

      addActivity(store, {
        actor: `${role}:recovery`,
        role,
        action: 'auth.recover_username',
        entity: 'user',
        entityId: user.id,
        details: `Username recovery completed for ${role} account.`
      });
      writeStore(store);

      json(res, 200, {
        data: {
          username: user.username,
          role: user.role
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/auth/recover-password' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      const role = clean(payload.role).toLowerCase();
      const recoveryCode = String(payload.recoveryCode || '');
      const newPassword = String(payload.newPassword || '');
      const confirmPassword = payload.confirmPassword === undefined
        ? newPassword
        : String(payload.confirmPassword || '');

      if (!role || !RECOVERY_CODE_BY_ROLE[role]) {
        json(res, 422, { error: 'Role must be admin or agent.' });
        return;
      }
      if (!recoveryCode) {
        json(res, 422, { error: 'Recovery code is required.' });
        return;
      }
      if (!newPassword) {
        json(res, 422, { error: 'New password is required.' });
        return;
      }
      if (!secureEquals(newPassword, confirmPassword)) {
        json(res, 422, { error: 'New password and confirmation do not match.' });
        return;
      }

      const invalidPassword = validateNewPassword(newPassword);
      if (invalidPassword) {
        json(res, 422, { error: invalidPassword });
        return;
      }

      const expectedCode = recoveryCodeForRole(role);
      if (!expectedCode) {
        json(res, 503, { error: `Recovery is not configured for role "${role}".` });
        return;
      }
      if (!secureEquals(recoveryCode, expectedCode)) {
        json(res, 401, { error: 'Recovery code is invalid.' });
        return;
      }

      const store = readStore();
      const user = accountByRole(store, role);
      if (!user) {
        json(res, 404, { error: `No active ${role} account found.` });
        return;
      }

      user.password = hashPassword(newPassword);
      user.updatedAt = nowIso();
      invalidateUserSessions(store, user.id);

      addActivity(store, {
        actor: `${role}:recovery`,
        role,
        action: 'auth.recover_password',
        entity: 'user',
        entityId: user.id,
        details: `Password reset completed for ${role} account.`
      });
      writeStore(store);

      json(res, 200, {
        data: {
          success: true,
          username: user.username,
          role: user.role
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
    rows = filterByQuery(rows, reqUrl.searchParams, ['name', 'phone', 'nationalId', 'location', 'notes']);

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
      if (findFarmerByPhone(store.farmers, payload.phone)) {
        json(res, 409, { error: 'A farmer with this phone number already exists.' });
        return;
      }
      if (findFarmerByNationalId(store.farmers, payload.nationalId)) {
        json(res, 409, { error: 'A farmer with this National ID already exists.' });
        return;
      }

      const hectares = resolveHectares(payload);
      const avocadoHectares = resolveAvocadoHectares(payload);
      const record = {
        id: id('F'),
        name: clean(payload.name),
        phone: clean(payload.phone),
        nationalId: cleanNationalId(payload.nationalId),
        location: clean(payload.location),
        trees: Number(parseTrees(payload.trees).toFixed(2)),
        hectares: Number(hectares.toFixed(3)),
        avocadoHectares: Number(avocadoHectares.toFixed(3)),
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

  if (pathname === '/api/farmers/import' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 10_000_000);
      const records = Array.isArray(payload.records) ? payload.records : [];
      if (!records.length) {
        json(res, 422, { error: 'records array is required' });
        return;
      }

      const onDuplicate = clean(payload.onDuplicate).toLowerCase() === 'overwrite' ? 'overwrite' : 'skip';
      const farmersByPhone = new Map(
        store.farmers
          .map((row) => [clean(row.phone), row])
          .filter(([phone]) => Boolean(phone))
      );
      const farmersByNationalId = new Map(
        store.farmers
          .map((row) => [normalizedNationalId(row.nationalId), row])
          .filter(([nationalId]) => Boolean(nationalId))
      );
      const created = [];
      const updated = [];
      const errors = [];

      records.forEach((raw, idx) => {
        if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
          errors.push({ row: idx + 1, error: 'Row is not a valid object' });
          return;
        }

        const mapped = mapFarmerImportRecord(raw);
        const invalid = validateFarmer(mapped);
        if (invalid) {
          errors.push({ row: idx + 1, error: invalid });
          return;
        }

        const phone = clean(mapped.phone);
        const nationalId = cleanNationalId(mapped.nationalId);
        const nationalIdKey = normalizedNationalId(nationalId);
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

            const oldPhoneKey = clean(existing.phone);
            const oldNationalIdKey = normalizedNationalId(existing.nationalId);
            if (oldPhoneKey && farmersByPhone.get(oldPhoneKey)?.id === existing.id) {
              farmersByPhone.delete(oldPhoneKey);
            }
            if (oldNationalIdKey && farmersByNationalId.get(oldNationalIdKey)?.id === existing.id) {
              farmersByNationalId.delete(oldNationalIdKey);
            }

            existing.name = clean(mapped.name);
            existing.phone = phone;
            existing.nationalId = nationalId;
            existing.location = clean(mapped.location);
            existing.trees = Number(parseTrees(mapped.trees).toFixed(2));
            existing.hectares = Number(resolveHectares(mapped).toFixed(3));
            existing.avocadoHectares = Number(resolveAvocadoHectares(mapped).toFixed(3));
            existing.notes = clean(mapped.notes);
            existing.updatedAt = nowIso();

            farmersByPhone.set(phone, existing);
            farmersByNationalId.set(nationalIdKey, existing);
            updated.push(existing.id);
          } else {
            errors.push({ row: idx + 1, error: 'Duplicate farmer (matching phone or National ID)' });
          }
          return;
        }

        const record = {
          id: id('F'),
          name: clean(mapped.name),
          phone,
          nationalId,
          location: clean(mapped.location),
          trees: Number(parseTrees(mapped.trees).toFixed(2)),
          hectares: Number(resolveHectares(mapped).toFixed(3)),
          avocadoHectares: Number(resolveAvocadoHectares(mapped).toFixed(3)),
          notes: clean(mapped.notes),
          createdBy: auth.session.username,
          createdAt: nowIso(),
          updatedAt: nowIso()
        };

        farmersByPhone.set(phone, record);
        farmersByNationalId.set(nationalIdKey, record);
        created.push(record);
      });

      if (created.length || updated.length) {
        store.farmers = created.reverse().concat(store.farmers);
        addActivity(store, {
          actor: auth.session.username,
          role: auth.session.role,
          action: 'farmer.import',
          entity: 'farmer',
          details:
            `Imported ${created.length}, updated ${updated.length}, skipped ${records.length - created.length - updated.length}` +
            ` (duplicate mode: ${onDuplicate})`
        });
        writeStore(store);
      }

      json(res, 200, {
        data: {
          totalRows: records.length,
          imported: created.length,
          updated: updated.length,
          skipped: records.length - created.length - updated.length,
          duplicateMode: onDuplicate,
          errors: errors.slice(0, 250)
        }
      });
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
    const producePurchases = store.producePurchases.filter((row) => row.farmerId === farmer.id);
    const payments = store.payments.filter((row) => row.farmerId === farmer.id);

    json(res, 200, {
      data: {
        farmer,
        produce,
        producePurchases,
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
      const nextName = payload.name !== undefined ? clean(payload.name) : clean(farmer.name);
      const nextPhone = payload.phone !== undefined ? clean(payload.phone) : clean(farmer.phone);
      const nextNationalId =
        payload.nationalId !== undefined ? cleanNationalId(payload.nationalId) : cleanNationalId(farmer.nationalId);
      const nextLocation = payload.location !== undefined ? clean(payload.location) : clean(farmer.location);
      const nextNotes = payload.notes !== undefined ? clean(payload.notes) : clean(farmer.notes);

      if (!nextName) {
        json(res, 422, { error: 'name is required' });
        return;
      }
      if (!nextPhone) {
        json(res, 422, { error: 'phone is required' });
        return;
      }
      if (!nextNationalId) {
        json(res, 422, { error: 'nationalId is required' });
        return;
      }
      if (!nextLocation) {
        json(res, 422, { error: 'location is required' });
        return;
      }

      if (payload.trees !== undefined) {
        const trees = parseTrees(payload.trees);
        if (!Number.isFinite(trees)) {
          json(res, 422, { error: 'trees must be a number' });
          return;
        }
        farmer.trees = Number(trees.toFixed(2));
      }
      const duplicatePhone = findFarmerByPhone(store.farmers, nextPhone, farmer.id);
      if (duplicatePhone) {
        json(res, 409, { error: 'A farmer with this phone number already exists.' });
        return;
      }
      const duplicateNationalId = findFarmerByNationalId(store.farmers, nextNationalId, farmer.id);
      if (duplicateNationalId) {
        json(res, 409, { error: 'A farmer with this National ID already exists.' });
        return;
      }
      const hasFarmAreaUpdate =
        payload.hectares !== undefined ||
        payload.acres !== undefined ||
        payload.squareFeet !== undefined ||
        payload.squarefeet !== undefined ||
        payload.square_feet !== undefined ||
        payload.sqft !== undefined ||
        payload.squareFt !== undefined;
      const hasAvocadoAreaUpdate =
        payload.avocadoHectares !== undefined ||
        payload.avocadohectares !== undefined ||
        payload.areaUnderAvocadoHectares !== undefined ||
        payload.areaunderavocadohectares !== undefined ||
        payload.avocadoAcres !== undefined ||
        payload.avocadoacres !== undefined ||
        payload.areaUnderAvocadoAcres !== undefined ||
        payload.areaunderavocadoacres !== undefined ||
        payload.avocadoAcreage !== undefined ||
        payload.avocadoacreage !== undefined ||
        payload.avocadoSquareFeet !== undefined ||
        payload.avocadosquarefeet !== undefined ||
        payload.avocado_square_feet !== undefined ||
        payload.avocadoSqft !== undefined ||
        payload.avocadosqft !== undefined ||
        payload.areaUnderAvocadoSquareFeet !== undefined ||
        payload.areaunderavocadosquarefeet !== undefined ||
        payload.areaUnderAvocadoSqft !== undefined ||
        payload.areaunderavocadosqft !== undefined;

      const currentHectares = Number(farmer.hectares);
      const currentAvocadoHectares = Number(farmer.avocadoHectares);
      const nextHectares = hasFarmAreaUpdate ? resolveHectares(payload) : currentHectares;
      const nextAvocadoHectares = hasAvocadoAreaUpdate
        ? resolveAvocadoHectares(payload)
        : Number.isFinite(currentAvocadoHectares) && currentAvocadoHectares > 0
          ? currentAvocadoHectares
          : currentHectares;

      if (!Number.isFinite(nextHectares) || nextHectares <= 0) {
        json(res, 422, { error: 'hectares/acres/square feet must be greater than 0' });
        return;
      }
      if (!Number.isFinite(nextAvocadoHectares) || nextAvocadoHectares <= 0) {
        json(res, 422, { error: 'avocadoHectares/avocadoAcres/avocadoSquareFeet must be greater than 0' });
        return;
      }
      if (nextAvocadoHectares > nextHectares) {
        json(res, 422, { error: 'area under avocado cannot be greater than total farm size' });
        return;
      }

      farmer.hectares = Number(nextHectares.toFixed(3));
      farmer.avocadoHectares = Number(nextAvocadoHectares.toFixed(3));
      farmer.name = nextName;
      farmer.phone = nextPhone;
      farmer.nationalId = nextNationalId;
      farmer.location = nextLocation;
      farmer.notes = nextNotes;
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
    const linkedProducePurchases = store.producePurchases.some((row) => row.farmerId === farmerId);
    const linkedPayments = store.payments.some((row) => row.farmerId === farmerId);
    if (linkedProduce || linkedProducePurchases || linkedPayments) {
      json(res, 409, { error: 'Cannot delete farmer with produce/purchase/payment history' });
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
    rows = filterByQuery(rows, reqUrl.searchParams, [
      'farmerName',
      'variety',
      'visualGrade',
      'sizeCode',
      'qcDecision',
      'inspector',
      'quality',
      'agent',
      'notes',
      'createdBy'
    ]);

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

      const lotWeightKgs = Number(parseNumber(payload.kgs ?? payload.lotWeightKgs).toFixed(2));
      const variety = normalizeAvocadoVariety(payload.variety);
      const sampleSize = Math.floor(parseNumber(payload.sampleSize));
      const visualGrade = normalizeVisualGrade(payload.visualGrade);
      const dryMatterRaw = clean(payload.dryMatterPct);
      const dryMatterPct = dryMatterRaw ? Number(parseNumber(dryMatterRaw).toFixed(2)) : null;
      const firmnessValue = Number(parseNumber(payload.firmnessValue).toFixed(2));
      const firmnessUnit = normalizeFirmnessUnit(payload.firmnessUnit);
      const avgFruitWeightG = Number(parseNumber(payload.avgFruitWeightG).toFixed(1));
      const sizeCode = normalizeSizeCode(payload.sizeCode);
      const qcDecision = normalizeQcDecision(payload.qcDecision);
      const inspector = clean(payload.inspector ?? payload.agent);

      const record = {
        id: id('P'),
        farmerId: farmer.id,
        farmerName: farmer.name,
        kgs: lotWeightKgs,
        lotWeightKgs,
        variety,
        sampleSize,
        visualGrade,
        dryMatterPct,
        firmnessValue,
        firmnessUnit,
        avgFruitWeightG,
        sizeCode,
        qcDecision,
        inspector,
        quality: visualGrade,
        agent: inspector,
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
        details: `Logged farm-gate QC ${record.variety} lot (${record.kgs}kg) for ${record.farmerName}`
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

  if (pathname === '/api/produce-purchases' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.producePurchases];
    const farmerId = clean(reqUrl.searchParams.get('farmerId'));
    if (farmerId) rows = rows.filter((row) => row.farmerId === farmerId);
    rows = filterByQuery(rows, reqUrl.searchParams, [
      'farmerName',
      'variety',
      'sizeCode',
      'buyer',
      'qcRecordId',
      'notes',
      'createdBy'
    ]);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 300, 2000);
    json(res, 200, { data: rows.slice(0, limit) });
    return;
  }

  if (pathname === '/api/produce-purchases' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const invalid = validateProducePurchase(payload);
      if (invalid) {
        json(res, 422, { error: invalid });
        return;
      }

      const farmer = findById(store.farmers, clean(payload.farmerId));
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found' });
        return;
      }

      const qcRecordId = clean(payload.qcRecordId);
      let qcRecord = null;
      if (qcRecordId) {
        qcRecord = findById(store.produce, qcRecordId);
        if (!qcRecord) {
          json(res, 404, { error: 'QC record not found' });
          return;
        }
        if (qcRecord.farmerId !== farmer.id) {
          json(res, 422, { error: 'qcRecordId must belong to the selected farmer' });
          return;
        }
      }

      const purchasedKgs = Number(parseNumber(payload.purchasedKgs).toFixed(2));
      const normalizedVariety = normalizeAvocadoVariety(payload.variety);
      const variety = normalizedVariety || (qcRecord ? clean(qcRecord.variety) : '');
      if (!variety) {
        json(res, 422, { error: 'variety is required when no QC record is linked' });
        return;
      }

      const rawSizeCode = clean(payload.sizeCode);
      const sizeCode = rawSizeCode ? normalizeSizeCode(rawSizeCode) : normalizeSizeCode(qcRecord?.sizeCode);
      if (rawSizeCode && !sizeCode) {
        json(res, 422, { error: 'sizeCode must be one of C12, C14, C16, C18, C20, C22, C24, C26, C28' });
        return;
      }

      const pricePerKgRaw = clean(payload.pricePerKgKes);
      const pricePerKgKes = pricePerKgRaw ? Number(parseNumber(pricePerKgRaw).toFixed(2)) : null;
      const purchaseValueRaw = clean(payload.purchaseValueKes);
      const computedValue = pricePerKgKes !== null ? purchasedKgs * pricePerKgKes : null;
      const purchaseValueKes = purchaseValueRaw
        ? Number(parseNumber(purchaseValueRaw).toFixed(2))
        : computedValue !== null
          ? Number(computedValue.toFixed(2))
          : null;

      const record = {
        id: id('PR'),
        farmerId: farmer.id,
        farmerName: farmer.name,
        qcRecordId: qcRecord ? qcRecord.id : '',
        variety,
        sizeCode: sizeCode || '',
        purchasedKgs,
        pricePerKgKes,
        purchaseValueKes,
        paidAmountKes: 0,
        balanceKes: purchaseValueKes || 0,
        settlementStatus: purchaseValueKes ? 'Unpaid' : 'Unpriced',
        buyer: clean(payload.buyer),
        notes: clean(payload.notes),
        createdBy: auth.session.username,
        createdAt: nowIso()
      };

      store.producePurchases.unshift(record);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'produce.purchase.create',
        entity: 'producePurchase',
        entityId: record.id,
        details: `Logged purchase ${record.purchasedKgs}kg (${record.variety}) for ${record.farmerName}`
      });
      writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/produce-purchases/') && req.method === 'DELETE') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const purchaseId = routeParam(pathname, '/api/produce-purchases/');
    const index = store.producePurchases.findIndex((row) => row.id === purchaseId);
    if (index === -1) {
      json(res, 404, { error: 'Produce purchase record not found' });
      return;
    }

    const [removed] = store.producePurchases.splice(index, 1);
    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'produce.purchase.delete',
      entity: 'producePurchase',
      entityId: removed.id,
      details: `Deleted produce purchase record ${removed.id}`
    });
    writeStore(store);
    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/payments/owed' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const owed = buildOwedRows(store, {
      period: reqUrl.searchParams.get('period'),
      from: reqUrl.searchParams.get('from'),
      to: reqUrl.searchParams.get('to'),
      q: reqUrl.searchParams.get('q')
    });
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 500, 5000);

    json(res, 200, {
      data: owed.rows.slice(0, limit),
      meta: {
        ...owed.meta,
        limit,
        returned: Math.min(limit, owed.rows.length)
      }
    });
    return;
  }

  if (pathname === '/api/payments/settle' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const farmerIds = Array.isArray(payload.farmerIds) ? payload.farmerIds.map(clean).filter(Boolean) : [];
      const payAll = payload.payAll === true;
      const owed = buildOwedRows(store, {
        period: payload.period,
        from: payload.from,
        to: payload.to,
        q: payload.q
      });

      let targets = [];
      if (payAll) {
        targets = owed.rows;
      } else {
        const selected = new Set(farmerIds);
        targets = owed.rows.filter((row) => selected.has(row.farmerId));
      }

      if (!targets.length) {
        json(res, 422, { error: 'No farmers with outstanding balances selected in this date range.' });
        return;
      }

      const status = normalizePaymentStatus(payload.status || 'Pending');
      const method = clean(payload.method) || (status === 'Received' ? 'M-PESA(Mock)' : 'M-PESA');
      const refPrefix = clean(payload.refPrefix) || (status === 'Received' ? 'MPB' : 'SET');
      const noteBase = clean(payload.notes) || `Settlement from owed list (${owed.meta.period || 'all'})`;

      const created = [];
      let totalAmount = 0;
      const stamp = Date.now().toString().slice(-8);
      for (let index = 0; index < targets.length; index += 1) {
        const row = targets[index];
        const amount = money(row.balanceKes);
        if (!(amount > 0)) continue;

        totalAmount = money(totalAmount + amount);
        const ref = `${refPrefix}${stamp}${String(index + 1).padStart(2, '0')}`;
        const record = {
          id: id('TX'),
          farmerId: row.farmerId,
          farmerName: row.farmerName,
          amount,
          ref,
          status,
          method,
          notes: noteBase,
          source: 'purchase-owed',
          settlementPeriod: owed.meta.period || 'all',
          settlementFrom: owed.meta.from || '',
          settlementTo: owed.meta.to || '',
          createdBy: auth.session.username,
          createdAt: nowIso(),
          updatedAt: nowIso()
        };

        store.payments.unshift(record);
        created.push(record);
      }

      if (!created.length) {
        json(res, 422, { error: 'No payable balances found for the selected farmers.' });
        return;
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'payment.bulk.settle',
        entity: 'payment',
        details: `Created ${created.length} settlement payments (KES ${totalAmount}) with status ${status}`
      });
      writeStore(store);

      const refreshed = buildOwedRows(store, {
        period: payload.period,
        from: payload.from,
        to: payload.to,
        q: payload.q
      });

      json(res, 201, {
        data: {
          createdCount: created.length,
          totalAmount,
          status,
          remainingOwedKes: refreshed.meta.totalBalanceKes,
          payments: created
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
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
    rows = filterByQuery(rows, reqUrl.searchParams, ['phone', 'message', 'createdBy', 'farmerName']);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 300, 2000);
    json(res, 200, { data: rows.slice(0, limit) });
    return;
  }

  if (pathname === '/api/sms/recipients' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const q = clean(reqUrl.searchParams.get('q'));
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 50, 200);
    const offset = safeQueryOffset(reqUrl.searchParams, 'offset', 0, 2_000_000);

    let recipients = store.farmers.filter((row) => clean(row.phone));
    recipients = filterByTextQuery(recipients, q, ['id', 'name', 'phone', 'nationalId', 'location', 'notes']);

    const total = recipients.length;
    const rows = recipients.slice(offset, offset + limit).map((row) => ({
      id: row.id,
      name: row.name,
      phone: row.phone,
      nationalId: row.nationalId,
      location: row.location
    }));

    json(res, 200, {
      data: rows,
      meta: {
        q,
        total,
        limit,
        offset,
        hasMore: offset + rows.length < total
      }
    });
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

  if (pathname === '/api/sms/send-bulk' && req.method === 'POST') {
    try {
      const store = readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const mode = clean(payload.mode).toLowerCase();
      const message = clean(payload.message);
      const selectedIds = Array.isArray(payload.farmerIds) ? payload.farmerIds.map((row) => clean(row)).filter(Boolean) : [];

      if (!message) {
        json(res, 422, { error: 'message is required' });
        return;
      }
      if (!['all', 'selected'].includes(mode)) {
        json(res, 422, { error: 'mode must be all or selected' });
        return;
      }
      if (mode === 'selected' && selectedIds.length === 0) {
        json(res, 422, { error: 'farmerIds is required when mode is selected' });
        return;
      }

      let targetFarmers = [];
      if (mode === 'all') {
        targetFarmers = store.farmers.filter((row) => clean(row.phone));
      } else {
        const selectedSet = new Set(selectedIds);
        targetFarmers = store.farmers.filter((row) => selectedSet.has(row.id) && clean(row.phone));
      }

      if (!targetFarmers.length) {
        json(res, 422, { error: 'No valid farmer recipients with mobile numbers were found.' });
        return;
      }

      const byPhone = new Map();
      for (const farmer of targetFarmers) {
        const phone = clean(farmer.phone);
        if (!phone || byPhone.has(phone)) continue;
        byPhone.set(phone, farmer);
      }

      const createdAt = nowIso();
      const logs = [];
      for (const [phone, farmer] of byPhone.entries()) {
        logs.push({
          id: id('SMS'),
          farmerId: farmer.id || '',
          farmerName: farmer.name || '',
          phone,
          message,
          provider: 'AfricaTalking(Mock)',
          status: 'Sent',
          createdBy: auth.session.username,
          createdAt
        });
      }

      if (!logs.length) {
        json(res, 422, { error: 'No valid mobile numbers found after deduplication.' });
        return;
      }

      store.smsLogs = logs.reverse().concat(store.smsLogs);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'sms.send_bulk',
        entity: 'sms',
        details: `Sent bulk SMS to ${logs.length} unique mobile numbers (${mode} mode)`
      });

      writeStore(store);
      json(res, 201, {
        data: {
          mode,
          sentCount: logs.length,
          requestedFarmerCount: targetFarmers.length,
          uniquePhoneCount: logs.length
        }
      });
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

    const farmerRows = store.farmers.map((row) => {
      const hectares = Number(row.hectares || 0);
      const avocadoHectares = Number(row.avocadoHectares || 0);
      const acres = Number.isFinite(hectares) ? hectares / HECTARES_PER_ACRE : 0;
      const squareFeet = Number.isFinite(hectares) ? hectares * SQFT_PER_HECTARE : 0;
      const avocadoAcres = Number.isFinite(avocadoHectares) ? avocadoHectares / HECTARES_PER_ACRE : 0;
      const avocadoSquareFeet = Number.isFinite(avocadoHectares) ? avocadoHectares * SQFT_PER_HECTARE : 0;

      return {
        ...row,
        acres: Number(acres.toFixed(3)),
        squareFeet: Number(squareFeet.toFixed(2)),
        avocadoAcres: Number(avocadoAcres.toFixed(3)),
        avocadoSquareFeet: Number(avocadoSquareFeet.toFixed(2))
      };
    });

    const csvData = toCsv(farmerRows, [
      { key: 'id', label: 'id' },
      { key: 'name', label: 'name' },
      { key: 'phone', label: 'phone' },
      { key: 'nationalId', label: 'nationalId' },
      { key: 'location', label: 'location' },
      { key: 'hectares', label: 'hectares' },
      { key: 'acres', label: 'acres' },
      { key: 'squareFeet', label: 'squareFeet' },
      { key: 'avocadoHectares', label: 'avocadoHectares' },
      { key: 'avocadoAcres', label: 'avocadoAcres' },
      { key: 'avocadoSquareFeet', label: 'avocadoSquareFeet' },
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

    const produceRows = store.produce.map((row) => ({
      ...row,
      lotWeightKgs: Number.isFinite(Number(row.lotWeightKgs)) ? Number(row.lotWeightKgs) : Number(row.kgs || 0),
      visualGrade: clean(row.visualGrade) || clean(row.quality),
      inspector: clean(row.inspector) || clean(row.agent)
    }));

    const csvData = toCsv(produceRows, [
      { key: 'id', label: 'id' },
      { key: 'farmerId', label: 'farmerId' },
      { key: 'farmerName', label: 'farmerName' },
      { key: 'lotWeightKgs', label: 'lotWeightKgs' },
      { key: 'variety', label: 'variety' },
      { key: 'sampleSize', label: 'sampleSize' },
      { key: 'visualGrade', label: 'visualGrade' },
      { key: 'dryMatterPct', label: 'dryMatterPct' },
      { key: 'firmnessValue', label: 'firmnessValue' },
      { key: 'firmnessUnit', label: 'firmnessUnit' },
      { key: 'avgFruitWeightG', label: 'avgFruitWeightG' },
      { key: 'sizeCode', label: 'sizeCode' },
      { key: 'qcDecision', label: 'qcDecision' },
      { key: 'inspector', label: 'inspector' },
      { key: 'notes', label: 'notes' },
      { key: 'createdBy', label: 'createdBy' },
      { key: 'createdAt', label: 'createdAt' }
    ]);
    csv(res, 'produce.csv', csvData);
    return;
  }

  if (pathname === '/api/exports/produce-purchases.csv' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const csvData = toCsv(store.producePurchases, [
      { key: 'id', label: 'id' },
      { key: 'farmerId', label: 'farmerId' },
      { key: 'farmerName', label: 'farmerName' },
      { key: 'qcRecordId', label: 'qcRecordId' },
      { key: 'variety', label: 'variety' },
      { key: 'sizeCode', label: 'sizeCode' },
      { key: 'purchasedKgs', label: 'purchasedKgs' },
      { key: 'pricePerKgKes', label: 'pricePerKgKes' },
      { key: 'purchaseValueKes', label: 'purchaseValueKes' },
      { key: 'paidAmountKes', label: 'paidAmountKes' },
      { key: 'balanceKes', label: 'balanceKes' },
      { key: 'settlementStatus', label: 'settlementStatus' },
      { key: 'buyer', label: 'buyer' },
      { key: 'notes', label: 'notes' },
      { key: 'createdBy', label: 'createdBy' },
      { key: 'createdAt', label: 'createdAt' }
    ]);
    csv(res, 'produce-purchases.csv', csvData);
    return;
  }

  if (pathname === '/api/exports/payments-owed.csv' && req.method === 'GET') {
    const store = readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const owed = buildOwedRows(store, {
      period: reqUrl.searchParams.get('period'),
      from: reqUrl.searchParams.get('from'),
      to: reqUrl.searchParams.get('to'),
      q: reqUrl.searchParams.get('q')
    });

    const csvData = toCsv(owed.rows, [
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
    csv(res, 'payments-owed.csv', csvData);
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
      { key: 'source', label: 'source' },
      { key: 'settlementPeriod', label: 'settlementPeriod' },
      { key: 'settlementFrom', label: 'settlementFrom' },
      { key: 'settlementTo', label: 'settlementTo' },
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
      store.producePurchases = [];
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

    if (store.farmers.length || store.produce.length || store.producePurchases.length || store.payments.length) {
      json(res, 409, { error: 'Store already has data' });
      return;
    }

    const farmerA = {
      id: id('F'),
      name: 'Mercy Achieng',
      phone: '254712330001',
      nationalId: '28643197',
      location: 'Muranga',
      hectares: 1.9,
      avocadoHectares: 1.2,
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
      nationalId: '30984522',
      location: 'Nyeri',
      hectares: 2.4,
      avocadoHectares: 1.6,
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
        notes: 'Farm-gate QC complete',
        createdBy: auth.session.username,
        createdAt: nowIso()
      }
    ];
    store.producePurchases = [
      {
        id: id('PR'),
        farmerId: farmerA.id,
        farmerName: farmerA.name,
        qcRecordId: store.produce[0].id,
        variety: 'Hass',
        sizeCode: 'C20',
        purchasedKgs: 302.5,
        pricePerKgKes: 152.5,
        purchaseValueKes: Number((302.5 * 152.5).toFixed(2)),
        buyer: 'Agent Njoroge',
        notes: 'Accepted lot from farm-gate QC',
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
        producePurchases: store.producePurchases,
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
