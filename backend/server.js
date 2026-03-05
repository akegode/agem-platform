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
const ACRES_PER_HECTARE = 1 / HECTARES_PER_ACRE;
const SQFT_PER_HECTARE = 107639.1041671;
const MIN_PASSWORD_LENGTH = 10;
const FARMER_PIN_PATTERN = /^\d{4}$/;
const DEFAULT_LANGUAGE = 'en';
const IMPORT_ONBOARDING_SMS_DEFAULT_EN =
  'Agem Portal: Hello {{name}}. You are now registered in the AGEM farmer system. Use USSD {{ussd}} (once active) for farmer services.';
const IMPORT_ONBOARDING_SMS_DEFAULT_SW =
  'Agem Portal: Habari {{name}}. Umesajiliwa kwenye mfumo wa wakulima wa AGEM. Tumia USSD {{ussd}} (ukishawashwa) kupata huduma.';
const IMPORT_ONBOARDING_SMS_MAX_LENGTH = 500;

function envValue(key) {
  return String(process.env[key] || '').trim();
}

function parseEnvInt(key, fallback, min = 1, max = Number.MAX_SAFE_INTEGER) {
  const raw = envValue(key);
  if (!raw) return fallback;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return fallback;
  const normalized = Math.floor(parsed);
  if (normalized < min || normalized > max) return fallback;
  return normalized;
}

function sanitizeSqlIdentifier(value, fallback = 'agem_store') {
  const candidate = String(value || '').trim();
  if (!candidate) return fallback;
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(candidate)) return fallback;
  return candidate;
}

function detectStorageBackend() {
  const explicit = envValue('STORAGE_BACKEND').toLowerCase();
  if (explicit === 'json' || explicit === 'postgres' || explicit === 'mysql') {
    return explicit;
  }

  const databaseUrl = envValue('DATABASE_URL');
  if (databaseUrl.startsWith('postgres://') || databaseUrl.startsWith('postgresql://')) {
    return 'postgres';
  }
  if (databaseUrl.startsWith('mysql://') || databaseUrl.startsWith('mysql2://')) {
    return 'mysql';
  }
  return 'json';
}

const RUNNING_ON_RENDER = Boolean(envValue('RENDER'));
const STORAGE_BACKEND = detectStorageBackend();
const DATABASE_URL = envValue('DATABASE_URL');
const SQL_STORE_TABLE = sanitizeSqlIdentifier(envValue('SQL_STORE_TABLE') || 'agem_store');
const BACKUP_INTERVAL_HOURS = parseEnvInt('BACKUP_INTERVAL_HOURS', 24, 1, 24 * 365);
const BACKUP_RETENTION_DAYS = parseEnvInt('BACKUP_RETENTION_DAYS', 30, 1, 3650);
const ALLOW_DEMO_USERS =
  envValue('ALLOW_DEMO_USERS') === 'true' ||
  (!RUNNING_ON_RENDER && envValue('ALLOW_DEMO_USERS') !== 'false');
const ALLOW_FARMER_REGISTRATION = envValue('ALLOW_FARMER_REGISTRATION') !== 'false';
const SYNC_PASSWORDS_FROM_ENV = envValue('SYNC_PASSWORDS_FROM_ENV') === 'true';
const USSD_ENABLED = envValue('USSD_ENABLED') !== 'false';
const USSD_CODE = envValue('USSD_CODE') || '*483#';
const USSD_HELP_PHONE = envValue('USSD_HELP_PHONE') || '+254700000000';
const USSD_SHARED_SECRET = process.env.USSD_SHARED_SECRET || '';
const SMS_OWNER_COST_PER_MESSAGE_KES = (() => {
  const raw = envValue('SMS_OWNER_COST_PER_MESSAGE_KES');
  if (!raw) return 0.25;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed < 0) return 0.25;
  return Number(parsed.toFixed(4));
})();
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const OPENAI_MODEL = envValue('OPENAI_MODEL') || 'gpt-4.1-mini';
const OPENAI_TIMEOUT_MS = parseEnvInt('OPENAI_TIMEOUT_MS', 12000, 1000, 120000);
const AI_EXEC_BRIEF_ENABLED = envValue('AI_EXEC_BRIEF_ENABLED') !== 'false';
const AI_EXEC_BRIEF_INTERVAL_HOURS = parseEnvInt('AI_EXEC_BRIEF_INTERVAL_HOURS', 24, 1, 24 * 30);
const AI_EXEC_BRIEF_MAX_HISTORY = parseEnvInt('AI_EXEC_BRIEF_MAX_HISTORY', 120, 10, 2000);
const AI_PROPOSAL_CAP = parseEnvInt('AI_PROPOSAL_CAP', 500, 50, 10000);
const AI_FEEDBACK_CAP = parseEnvInt('AI_FEEDBACK_CAP', 2000, 100, 20000);
const AI_EVAL_CAP = parseEnvInt('AI_EVAL_CAP', 500, 20, 5000);
const AI_KNOWLEDGE_CAP = parseEnvInt('AI_KNOWLEDGE_CAP', 500, 10, 5000);
const OPS_TASK_CAP = parseEnvInt('OPS_TASK_CAP', 5000, 100, 50000);
const PAYMENT_RECOMMENDATION_CAP = parseEnvInt('PAYMENT_RECOMMENDATION_CAP', 5000, 100, 50000);
const AI_MSAIDIZI_SYNC_ENABLED = envValue('AI_MSAIDIZI_SYNC_ENABLED') !== 'false';
const AI_MSAIDIZI_SYNC_INTERVAL_HOURS = parseEnvInt('AI_MSAIDIZI_SYNC_INTERVAL_HOURS', 6, 1, 24 * 30);
const AI_MSAIDIZI_MODULE_CAP = parseEnvInt('AI_MSAIDIZI_MODULE_CAP', 80, 10, 500);

let sqlStoreAdapter = null;
let backupScheduleHandle = null;
let executiveBriefScheduleHandle = null;
let msaidiziScheduleHandle = null;

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

function validateFarmerPin(pin) {
  const value = String(pin || '').trim();
  if (!FARMER_PIN_PATTERN.test(value)) {
    return 'PIN must be exactly 4 digits.';
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
        provisioning: 'environment',
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
      existing.provisioning = 'environment';
      if (!hasPasswordHash || passwordChanged) {
        existing.password = hashPassword(account.password);
      }
      existing.updatedAt = nowIso();
      changed = true;
    }

    if (account.role === 'admin') {
      for (const user of store.users) {
        if (!user || user.id === existing.id) continue;
        if (user.role !== account.role) continue;
        if (user.status !== 'active') continue;
        user.status = 'disabled';
        user.updatedAt = nowIso();
        changed = true;
      }
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
      version: 4,
      createdAt: nowIso(),
      lastWriteAt: nowIso(),
      authMode: ALLOW_DEMO_USERS ? 'demo' : 'private'
    },
    users: ALLOW_DEMO_USERS ? seededDemoUsers() : [],
    farmers: [],
    produce: [],
    producePurchases: [],
    payments: [],
    paymentRecommendations: [],
    smsLogs: [],
    activityLogs: [],
    sessions: {},
    aiProposals: [],
    aiBriefs: [],
    aiAlerts: [],
    aiFeedback: [],
    aiEvalRuns: [],
    aiKnowledgeDocs: [],
    aiMsaidizi: {
      lastSyncAt: '',
      sourceSignature: '',
      status: 'never',
      modules: [],
      syncHistory: []
    },
    aiPromptConfig: {},
    opsTasks: []
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
    paymentRecommendations: Array.isArray(raw.paymentRecommendations) ? raw.paymentRecommendations : [],
    smsLogs: Array.isArray(raw.smsLogs) ? raw.smsLogs : [],
    activityLogs: Array.isArray(raw.activityLogs) ? raw.activityLogs : [],
    sessions: raw.sessions && typeof raw.sessions === 'object' ? raw.sessions : {},
    aiProposals: Array.isArray(raw.aiProposals) ? raw.aiProposals : [],
    aiBriefs: Array.isArray(raw.aiBriefs) ? raw.aiBriefs : [],
    aiAlerts: Array.isArray(raw.aiAlerts) ? raw.aiAlerts : [],
    aiFeedback: Array.isArray(raw.aiFeedback) ? raw.aiFeedback : [],
    aiEvalRuns: Array.isArray(raw.aiEvalRuns) ? raw.aiEvalRuns : [],
    aiKnowledgeDocs: Array.isArray(raw.aiKnowledgeDocs) ? raw.aiKnowledgeDocs : [],
    aiMsaidizi: raw.aiMsaidizi && typeof raw.aiMsaidizi === 'object' ? raw.aiMsaidizi : {},
    aiPromptConfig: raw.aiPromptConfig && typeof raw.aiPromptConfig === 'object' ? raw.aiPromptConfig : {},
    opsTasks: Array.isArray(raw.opsTasks) ? raw.opsTasks : []
  };

  if (!store.meta.createdAt) store.meta.createdAt = nowIso();
  store.meta.version = 4;
  store.meta.authMode = ALLOW_DEMO_USERS ? 'demo' : 'private';

  const usersValid = store.users.some((u) => u && u.username && u.password && u.password.hash);
  if (!usersValid) {
    store.users = ALLOW_DEMO_USERS ? seededDemoUsers() : [];
  }

  for (const farmer of store.farmers) {
    if (!farmer || typeof farmer !== 'object') continue;
    farmer.phone = clean(farmer.phone);
    farmer.nationalId = cleanNationalId(farmer.nationalId);
    farmer.preferredLanguage = languageOrDefault(farmer.preferredLanguage);
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

  store.aiProposals = store.aiProposals.slice(0, AI_PROPOSAL_CAP).map((row) => ({
    id: clean(row?.id) || id('AIP'),
    status: normalizeAiProposalStatus(row?.status),
    actionType: clean(row?.actionType),
    title: clean(row?.title),
    description: clean(row?.description),
    confidence: Number.isFinite(Number(row?.confidence)) ? Number(row.confidence) : '',
    source: clean(row?.source) || 'local-rules',
    payload: row?.payload && typeof row.payload === 'object' ? row.payload : {},
    createdBy: clean(row?.createdBy) || 'system',
    createdAt: clean(row?.createdAt) || nowIso(),
    updatedAt: clean(row?.updatedAt) || clean(row?.createdAt) || nowIso(),
    approvedBy: clean(row?.approvedBy),
    approvedAt: clean(row?.approvedAt),
    rejectedBy: clean(row?.rejectedBy),
    rejectedAt: clean(row?.rejectedAt),
    rejectionReason: clean(row?.rejectionReason),
    executedBy: clean(row?.executedBy),
    executedAt: clean(row?.executedAt),
    executionSummary: clean(row?.executionSummary)
  }));
  store.aiBriefs = store.aiBriefs.slice(0, AI_EXEC_BRIEF_MAX_HISTORY);
  store.aiAlerts = store.aiAlerts.slice(0, AI_EXEC_BRIEF_MAX_HISTORY);
  store.aiFeedback = store.aiFeedback.slice(0, AI_FEEDBACK_CAP);
  store.aiEvalRuns = store.aiEvalRuns.slice(0, AI_EVAL_CAP);
  store.aiKnowledgeDocs = store.aiKnowledgeDocs.slice(0, AI_KNOWLEDGE_CAP);
  const rawMsaidizi = store.aiMsaidizi && typeof store.aiMsaidizi === 'object' ? store.aiMsaidizi : {};
  const rawModules = Array.isArray(rawMsaidizi.modules) ? rawMsaidizi.modules : [];
  const rawHistory = Array.isArray(rawMsaidizi.syncHistory) ? rawMsaidizi.syncHistory : [];
  store.aiMsaidizi = {
    lastSyncAt: clean(rawMsaidizi.lastSyncAt),
    sourceSignature: clean(rawMsaidizi.sourceSignature),
    status: clean(rawMsaidizi.status) || 'never',
    modules: rawModules
      .slice(0, AI_MSAIDIZI_MODULE_CAP)
      .map((module) => ({
        id: clean(module?.id) || id('MSMOD'),
        title: clean(module?.title) || 'Untitled module',
        short: clean(module?.short),
        panes: Array.isArray(module?.panes) ? module.panes.map((row) => clean(row)).filter(Boolean).slice(0, 12) : [],
        roles: Array.isArray(module?.roles) ? module.roles.map((row) => clean(row)).filter(Boolean).slice(0, 8) : [],
        keywords: Array.isArray(module?.keywords)
          ? module.keywords.map((row) => clean(row).toLowerCase()).filter(Boolean).slice(0, 24)
          : [],
        steps: Array.isArray(module?.steps) ? module.steps.map((row) => clean(row)).filter(Boolean).slice(0, 20) : [],
        troubleshooting: Array.isArray(module?.troubleshooting)
          ? module.troubleshooting.map((row) => clean(row)).filter(Boolean).slice(0, 20)
          : [],
        doNow: Array.isArray(module?.doNow) ? module.doNow.map((row) => clean(row)).filter(Boolean).slice(0, 12) : [],
        avoid: Array.isArray(module?.avoid) ? module.avoid.map((row) => clean(row)).filter(Boolean).slice(0, 12) : [],
        waitFor: Array.isArray(module?.waitFor) ? module.waitFor.map((row) => clean(row)).filter(Boolean).slice(0, 12) : [],
        updatedAt: clean(module?.updatedAt) || nowIso(),
        sourceSignature: clean(module?.sourceSignature)
      })),
    syncHistory: rawHistory.slice(0, 120).map((row) => ({
      at: clean(row?.at) || nowIso(),
      reason: clean(row?.reason) || 'unknown',
      actor: clean(row?.actor) || 'system',
      updated: Boolean(row?.updated),
      sourceSignature: clean(row?.sourceSignature)
    }))
  };
  store.paymentRecommendations = store.paymentRecommendations
    .slice(0, PAYMENT_RECOMMENDATION_CAP)
    .map((row) => ({
      id: clean(row?.id) || id('PREQ'),
      farmerId: clean(row?.farmerId),
      farmerName: clean(row?.farmerName),
      farmerPhone: clean(row?.farmerPhone),
      amount: money(row?.amount),
      requestedOwedKes: money(row?.requestedOwedKes),
      status: normalizePaymentRecommendationStatus(row?.status),
      reason: clean(row?.reason),
      decisionNote: clean(row?.decisionNote),
      createdBy: clean(row?.createdBy) || 'unknown-agent',
      createdAt: clean(row?.createdAt) || nowIso(),
      updatedAt: clean(row?.updatedAt) || clean(row?.createdAt) || nowIso(),
      approvedBy: clean(row?.approvedBy),
      approvedAt: clean(row?.approvedAt),
      rejectedBy: clean(row?.rejectedBy),
      rejectedAt: clean(row?.rejectedAt),
      rejectionReason: clean(row?.rejectionReason),
      paymentId: clean(row?.paymentId)
    }));
  store.opsTasks = store.opsTasks.slice(0, OPS_TASK_CAP).map((row) => ({
    id: clean(row?.id) || id('TASK'),
    title: clean(row?.title) || 'Untitled task',
    description: clean(row?.description),
    severity: normalizeRiskSeverity(row?.severity, 'medium'),
    status: normalizeOpsTaskStatus(row?.status),
    sourceType: clean(row?.sourceType) || 'ai',
    sourceRef: clean(row?.sourceRef),
    assignedTo: clean(row?.assignedTo),
    notes: clean(row?.notes),
    createdBy: clean(row?.createdBy) || 'system',
    createdAt: clean(row?.createdAt) || nowIso(),
    updatedAt: clean(row?.updatedAt) || clean(row?.createdAt) || nowIso(),
    resolvedAt: clean(row?.resolvedAt)
  }));

  applyUserPolicy(store);
  pruneExpiredSessions(store);
  return store;
}

function ensureDataPaths() {
  fs.mkdirSync(DATA_ROOT, { recursive: true });
  fs.mkdirSync(path.dirname(STORE_PATH), { recursive: true });
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
}

function readStoreFromDisk() {
  try {
    if (!fs.existsSync(STORE_PATH)) return null;
    const raw = fs.readFileSync(STORE_PATH, 'utf-8');
    return normalizeStore(JSON.parse(raw));
  } catch {
    return null;
  }
}

function writeStoreToDisk(store) {
  fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2));
}

function requireStorageDriver(pkgName, backendLabel) {
  try {
    return require(pkgName);
  } catch (error) {
    throw new Error(
      `${backendLabel} storage backend requires "${pkgName}". Run "npm install" to install dependencies. (${error.message})`
    );
  }
}

async function buildPostgresStoreAdapter() {
  const { Pool } = requireStorageDriver('pg', 'PostgreSQL');
  if (!DATABASE_URL) {
    throw new Error('DATABASE_URL is required for PostgreSQL storage backend.');
  }

  const sslMode = envValue('PGSSLMODE').toLowerCase();
  const ssl =
    sslMode === 'disable'
      ? false
      : sslMode === 'require' || RUNNING_ON_RENDER
        ? { rejectUnauthorized: false }
        : undefined;

  const pool = new Pool({
    connectionString: DATABASE_URL,
    ssl
  });

  await pool.query(
    `CREATE TABLE IF NOT EXISTS ${SQL_STORE_TABLE} (
      id SMALLINT PRIMARY KEY,
      payload JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )`
  );

  return {
    backend: 'postgres',
    async read() {
      const result = await pool.query(
        `SELECT payload FROM ${SQL_STORE_TABLE} WHERE id = $1 LIMIT 1`,
        [1]
      );
      if (!result.rows.length) return null;
      const value = result.rows[0].payload;
      return typeof value === 'string' ? JSON.parse(value) : value;
    },
    async write(payload) {
      await pool.query(
        `INSERT INTO ${SQL_STORE_TABLE} (id, payload, updated_at)
         VALUES ($1, $2::jsonb, NOW())
         ON CONFLICT (id)
         DO UPDATE SET payload = EXCLUDED.payload, updated_at = NOW()`,
        [1, JSON.stringify(payload)]
      );
    },
    async close() {
      await pool.end();
    }
  };
}

async function buildMysqlStoreAdapter() {
  const mysql = requireStorageDriver('mysql2/promise', 'MySQL');
  const connectionConfig = DATABASE_URL
    ? DATABASE_URL
    : {
        host: envValue('MYSQL_HOST') || '127.0.0.1',
        port: parseEnvInt('MYSQL_PORT', 3306, 1, 65535),
        user: envValue('MYSQL_USER'),
        password: process.env.MYSQL_PASSWORD || '',
        database: envValue('MYSQL_DATABASE'),
        waitForConnections: true,
        connectionLimit: parseEnvInt('MYSQL_POOL_SIZE', 10, 1, 100),
        queueLimit: 0
      };

  if (!DATABASE_URL && (!connectionConfig.user || !connectionConfig.database)) {
    throw new Error('MYSQL_USER and MYSQL_DATABASE are required for MySQL storage backend.');
  }

  const pool = mysql.createPool(connectionConfig);
  await pool.query(
    `CREATE TABLE IF NOT EXISTS ${SQL_STORE_TABLE} (
      id TINYINT PRIMARY KEY,
      payload JSON NOT NULL,
      updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
    )`
  );

  return {
    backend: 'mysql',
    async read() {
      const [rows] = await pool.query(
        `SELECT payload FROM ${SQL_STORE_TABLE} WHERE id = ? LIMIT 1`,
        [1]
      );
      if (!rows.length) return null;
      const value = rows[0].payload;
      return typeof value === 'string' ? JSON.parse(value) : value;
    },
    async write(payload) {
      await pool.query(
        `INSERT INTO ${SQL_STORE_TABLE} (id, payload, updated_at)
         VALUES (?, ?, CURRENT_TIMESTAMP)
         ON DUPLICATE KEY UPDATE payload = VALUES(payload), updated_at = CURRENT_TIMESTAMP`,
        [1, JSON.stringify(payload)]
      );
    },
    async close() {
      await pool.end();
    }
  };
}

async function ensureSqlStoreAdapter() {
  if (sqlStoreAdapter) return sqlStoreAdapter;
  if (STORAGE_BACKEND === 'postgres') {
    sqlStoreAdapter = await buildPostgresStoreAdapter();
    return sqlStoreAdapter;
  }
  if (STORAGE_BACKEND === 'mysql') {
    sqlStoreAdapter = await buildMysqlStoreAdapter();
    return sqlStoreAdapter;
  }
  return null;
}

async function ensureStore() {
  ensureDataPaths();

  if (STORAGE_BACKEND === 'json') {
    if (!fs.existsSync(STORE_PATH)) {
      writeStoreToDisk(freshStore());
      return;
    }
    const store = readStoreFromDisk() || freshStore();
    await writeStore(store);
    return;
  }

  const adapter = await ensureSqlStoreAdapter();
  let existing = await adapter.read();
  if (!existing) {
    existing = readStoreFromDisk() || freshStore();
    const snapshot = createBackupSnapshot(existing, 'migration-seed');
    console.log(`[storage] Seeded ${STORAGE_BACKEND} store from snapshot ${snapshot.filename}`);
  }
  await writeStore(existing);
}

async function readStore() {
  ensureDataPaths();

  if (STORAGE_BACKEND === 'json') {
    return readStoreFromDisk() || freshStore();
  }

  const adapter = await ensureSqlStoreAdapter();
  const persisted = await adapter.read();
  if (persisted) return normalizeStore(persisted);

  const fallback = readStoreFromDisk() || freshStore();
  await adapter.write(fallback);
  return normalizeStore(fallback);
}

async function writeStore(store) {
  const nextStore = normalizeStore(store || freshStore());
  nextStore.meta.lastWriteAt = nowIso();

  if (STORAGE_BACKEND === 'json') {
    writeStoreToDisk(nextStore);
    return;
  }

  const adapter = await ensureSqlStoreAdapter();
  await adapter.write(nextStore);
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

function pruneOldBackups(retentionDays = BACKUP_RETENTION_DAYS) {
  ensureDataPaths();
  const cutoffMs = Date.now() - retentionDays * 24 * 60 * 60 * 1000;
  const removed = [];

  for (const name of fs.readdirSync(BACKUP_DIR)) {
    if (!name.endsWith('.json')) continue;
    const filepath = path.join(BACKUP_DIR, name);
    const stat = fs.statSync(filepath);
    if (stat.mtimeMs >= cutoffMs) continue;
    fs.unlinkSync(filepath);
    removed.push(name);
  }

  return removed;
}

async function runScheduledBackup() {
  const store = await readStore();
  const snapshot = createBackupSnapshot(store, 'scheduled');
  const removed = pruneOldBackups();
  console.log(
    `[backup] Scheduled snapshot ${snapshot.filename} created (retention ${BACKUP_RETENTION_DAYS}d, pruned ${removed.length}).`
  );
  return { snapshot, removed };
}

function scheduleAutomaticBackups() {
  if (backupScheduleHandle) {
    clearInterval(backupScheduleHandle);
    backupScheduleHandle = null;
  }

  const intervalMs = BACKUP_INTERVAL_HOURS * 60 * 60 * 1000;
  backupScheduleHandle = setInterval(() => {
    runScheduledBackup().catch((error) => {
      console.error(`[backup] Scheduled backup failed: ${error.message}`);
    });
  }, intervalMs);
  if (typeof backupScheduleHandle.unref === 'function') {
    backupScheduleHandle.unref();
  }
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

function pushWithCap(list, item, cap) {
  if (!Array.isArray(list)) return;
  list.unshift(item);
  if (list.length > cap) list.length = cap;
}

function normalizeRiskSeverity(value, fallback = 'medium') {
  const normalized = clean(value).toLowerCase();
  if (['critical', 'high', 'medium', 'low'].includes(normalized)) return normalized;
  return fallback;
}

function normalizeAiProposalStatus(value) {
  const normalized = clean(value).toLowerCase();
  if (['pending', 'approved', 'rejected', 'executed'].includes(normalized)) return normalized;
  return 'pending';
}

function normalizeOpsTaskStatus(value) {
  const normalized = clean(value).toLowerCase();
  if (['open', 'in_progress', 'resolved', 'closed'].includes(normalized)) return normalized;
  return 'open';
}

function normalizePaymentRecommendationStatus(value) {
  const normalized = clean(value).toLowerCase();
  if (['pending', 'approved', 'rejected'].includes(normalized)) return normalized;
  return 'pending';
}

function aiPromptPolicy(store, key, fallback) {
  if (!store || !store.aiPromptConfig || typeof store.aiPromptConfig !== 'object') return fallback;
  const configured = clean(store.aiPromptConfig[key]);
  return configured || fallback;
}

function routeParamWithSuffix(pathname, prefix, suffix) {
  if (!pathname.startsWith(prefix) || !pathname.endsWith(suffix)) return '';
  const middle = pathname.slice(prefix.length, pathname.length - suffix.length);
  const parts = middle.split('/').filter(Boolean);
  return parts.length === 1 ? decodeURIComponent(parts[0]) : '';
}

function tokenizeSearchText(value) {
  return clean(value)
    .toLowerCase()
    .split(/[^a-z0-9]+/g)
    .map((row) => row.trim())
    .filter((row) => row.length >= 2)
    .slice(0, 40);
}

function compactWhitespace(value) {
  return clean(value).replace(/\s+/g, ' ').trim();
}

function buildKnowledgeSnippet(content, tokens) {
  const raw = compactWhitespace(content);
  if (!raw) return '';
  if (!tokens.length) return raw.slice(0, 220);
  const lower = raw.toLowerCase();
  for (const token of tokens) {
    const idx = lower.indexOf(token);
    if (idx >= 0) {
      const start = Math.max(0, idx - 70);
      const end = Math.min(raw.length, idx + token.length + 150);
      return (start > 0 ? '…' : '') + raw.slice(start, end) + (end < raw.length ? '…' : '');
    }
  }
  return raw.slice(0, 220);
}

function findKnowledgeMatches(store, query, limit = 4) {
  const docs = Array.isArray(store?.aiKnowledgeDocs) ? store.aiKnowledgeDocs : [];
  if (!docs.length) return [];
  const tokens = tokenizeSearchText(query);
  if (!tokens.length) {
    return docs
      .slice(0, Math.max(1, limit))
      .map((doc) => ({
        id: doc.id,
        title: doc.title || 'Untitled',
        source: doc.source || '',
        tags: Array.isArray(doc.tags) ? doc.tags : [],
        snippet: buildKnowledgeSnippet(doc.content, []),
        score: 0
      }));
  }

  const scored = [];
  for (const doc of docs) {
    const haystack = [
      clean(doc.title).toLowerCase(),
      clean(doc.source).toLowerCase(),
      Array.isArray(doc.tags) ? doc.tags.map((row) => clean(row).toLowerCase()).join(' ') : '',
      clean(doc.content).toLowerCase()
    ].join(' ');

    let score = 0;
    for (const token of tokens) {
      if (!token) continue;
      if (haystack.includes(token)) score += 1;
      if (clean(doc.title).toLowerCase().includes(token)) score += 2;
      if (Array.isArray(doc.tags) && doc.tags.some((tag) => clean(tag).toLowerCase().includes(token))) {
        score += 2;
      }
    }

    if (score <= 0) continue;
    scored.push({
      id: doc.id,
      title: doc.title || 'Untitled',
      source: doc.source || '',
      tags: Array.isArray(doc.tags) ? doc.tags : [],
      snippet: buildKnowledgeSnippet(doc.content, tokens),
      score
    });
  }

  return scored.sort((a, b) => b.score - a.score).slice(0, Math.max(1, Math.min(limit, 8)));
}

function normalizeMsaidiziMode(value) {
  const normalized = clean(value).toLowerCase();
  if (['simple', 'dumb', 'dumb-down'].includes(normalized)) return 'simple';
  if (['steps', 'step', 'step_by_step', 'step-by-step'].includes(normalized)) return 'step_by_step';
  return 'normal';
}

function msaidiziSourceFiles() {
  return [
    path.join(FRONTEND_ROOT, 'index.html'),
    path.join(FRONTEND_ROOT, 'app.js'),
    path.join(FRONTEND_ROOT, 'styles.css'),
    path.join(__dirname, 'server.js')
  ];
}

function buildMsaidiziSourceSignature() {
  const hash = crypto.createHash('sha256');
  for (const filePath of msaidiziSourceFiles()) {
    try {
      const content = fs.readFileSync(filePath);
      hash.update(content);
    } catch {
      hash.update(`missing:${filePath}`);
    }
  }
  return hash.digest('hex').slice(0, 24);
}

function msaidiziModuleDefinitions(store) {
  const summary = makeSummary(store);
  const farmersCount = Number(summary?.farmers) || 0;
  const purchasedCount = Number(summary?.purchasedRecords) || 0;
  const owedCount = Number(summary?.owedFarmers) || 0;
  const sms24h = Number(summary?.smsSentLast24h) || 0;

  return [
    {
      id: 'account-setup',
      title: 'Account Setup & Sign In',
      panes: ['auth', 'overview'],
      roles: ['guest', 'farmer', 'agent', 'admin'],
      keywords: ['sign in', 'login', 'username', 'pin', 'register', 'password', 'account settings'],
      short:
        'Use username/password for staff accounts and phone + 4-digit PIN for farmers. Account Settings lets users update photo and password.',
      steps: [
        'Open the sign-in page and enter username or phone number.',
        'Enter password (staff) or 4-digit PIN (farmer) then click Sign In.',
        'For new farmers, open "New farmer? Register here" and complete all required fields.',
        'After sign-in, open Account Settings to change password and upload/remove profile photo.'
      ],
      troubleshooting: [
        'If sign-in fails, verify phone format is 2547xxxxxxxx.',
        'If password reset is needed, use recovery only for Administrator and Agent accounts.',
        'If farmer cannot sign in, admin can set/reset farmer PIN in Farmers > Portal Access.'
      ],
      doNow: ['Keep phone numbers and National ID accurate before first login.'],
      avoid: ['Do not share passwords, PINs, or recovery codes over unsecured channels.'],
      waitFor: ['If backend is offline, sign-in and password updates will not complete until API reconnects.']
    },
    {
      id: 'farmer-management',
      title: 'Farmer Registration & Data Management',
      panes: ['farmers'],
      roles: ['agent', 'admin'],
      keywords: ['farmer', 'register farmer', 'edit farmer', 'bulk import', 'csv', 'excel', 'pin', 'national id', 'area'],
      short:
        `Manage farmer records with mandatory identity and area fields. Current live count: ${farmersCount} farmer(s). Admin handles bulk imports and PIN reset tools.`,
      steps: [
        'Open Farmers > Register / Edit Farmer.',
        'Enter full name, phone, location, National ID, language, tree count, total farm area, and avocado area.',
        'Save to create or update a farmer profile.',
        'Admin only: use Bulk Import for CSV/Excel and resolve anomalies before final import.',
        'Admin only: use Portal Access (Set/Reset Farmer PIN) for imported farmers.'
      ],
      troubleshooting: [
        'If avocado area is greater than total farm area, correct area fields before saving.',
        'If duplicate phone or National ID is detected, decide whether to keep existing data or overwrite.',
        'If import QA flags errors, fix blocked rows first and re-run Smart QA.'
      ],
      doNow: ['Capture National ID and phone consistently because payments and notifications depend on them.'],
      avoid: ['Do not leave mandatory identity fields blank.'],
      waitFor: ['SMS onboarding notices send only when SMS service is available and notification option is enabled.']
    },
    {
      id: 'agent-management',
      title: 'Agent Management (Admin Only)',
      panes: ['agents'],
      roles: ['admin'],
      keywords: ['agent', 'create agent', 'agent phone', 'agent email', 'temporary credentials'],
      short:
        'Only administrators can create and manage agent accounts. The system generates temporary username/password for each new agent.',
      steps: [
        'Open Agents > Create Agent Account.',
        'Enter agent full name, email, and phone number.',
        'Click Create Agent and record the generated temporary credentials securely.',
        'Ask the new agent to sign in and change password immediately.'
      ],
      troubleshooting: [
        'If account creation fails, verify email and phone formats.',
        'If credentials are lost, admin should reset the agent password and re-share securely.',
        'If non-admin role cannot view Agents pane, this is expected role restriction.'
      ],
      doNow: ['Collect valid agent phone numbers during registration for communication and approvals.'],
      avoid: ['Do not create duplicate agent accounts for the same person.'],
      waitFor: ['Wait for admin approval before giving production-sensitive permissions to new agents.']
    },
    {
      id: 'produce-qc',
      title: 'Produce QC & Purchase Workflow',
      panes: ['produce'],
      roles: ['agent', 'admin'],
      keywords: ['produce', 'qc', 'dry matter', 'firmness', 'size', 'purchase', 'hass', 'fuerte'],
      short:
        `Capture farm-gate avocado QC first, then record purchased produce. Purchased records currently: ${purchasedCount}.`,
      steps: [
        'Open Produce > Record Farm-Gate QC and capture variety, weight, sample size, visual grade, dry matter, firmness, average fruit weight, and size code.',
        'Save QC entry and confirm decision (Accept / Hold / Reject).',
        'Open Record Purchased Produce for quantities actually bought.',
        'Link purchase to QC lot when available and enter agreed/contract price per kg.',
        'Review Purchased Produce Log for audit history.'
      ],
      troubleshooting: [
        'If quality metrics look inconsistent, run AI QC Intelligence to review risk signals.',
        'If purchase totals do not match expected amount, verify kg and price-per-kg fields.',
        'If farmer is missing in dropdown, confirm the farmer was registered first.'
      ],
      doNow: ['Collect QC values at farm gate before creating payout obligations.'],
      avoid: ['Do not pay based on unverified or rejected QC lots.'],
      waitFor: ['Hold payment recommendation until QC decision and purchase records are complete.']
    },
    {
      id: 'payments-settlement',
      title: 'Payments, Recommendations, and Settlement',
      panes: ['payments'],
      roles: ['agent', 'admin'],
      keywords: ['payment', 'owed', 'mpesa', 'recommend', 'approve', 'reject', 'disbursement'],
      short:
        `Agents recommend payments; admins approve/reject and execute payment. Current farmers owed: ${owedCount}.`,
      steps: [
        'Agent: open Payments and submit Recommend For Payment with amount and reason.',
        'Admin: open Agent Payment Recommendations and review pending requests.',
        'Admin: Approve to proceed with payment logging/disbursement, or Reject with reason.',
        'Admin: use Farmers Owed filters, select beneficiaries, then process payment batch.',
        'Monitor Payment Transactions for status, references, and reconciliation.'
      ],
      troubleshooting: [
        'If owed list is blank for agent, this is expected: owed balances are admin-visible only.',
        'If recommendation cannot be approved, verify farmer exists and amount is valid.',
        'If payment status is pending/failed, reconcile M-PESA reference and retry policy.'
      ],
      doNow: ['Require clear recommendation reasons before admin approval.'],
      avoid: ['Agents must not execute direct payments.'],
      waitFor: ['Payment completion depends on successful disbursement and transaction logging.']
    },
    {
      id: 'sms-comms',
      title: 'SMS Communication & Engagement',
      panes: ['sms'],
      roles: ['agent', 'admin'],
      keywords: ['sms', 'bulk sms', 'select all', 'recipient search', 'kiswahili', 'english', 'sms spend'],
      short:
        `Send individual or bulk farmer SMS with recipient search and language support. Last 24h SMS count: ${sms24h}.`,
      steps: [
        'Open SMS > Send SMS.',
        'Choose target mode: Selected Farmers, All Farmers, or Single Number.',
        'Use recipient search for large lists and select visible rows efficiently.',
        'Enter message (or generate draft with AI) and send.',
        'Review SMS log and spend metrics in Reports/Overview.'
      ],
      troubleshooting: [
        'If no recipients appear, refresh farmers data and confirm phone numbers are valid.',
        'If message is too long, shorten text or generate concise AI draft.',
        'If SMS delivery log is empty, confirm SMS gateway credentials and service availability.'
      ],
      doNow: ['Prefer farmer preferred-language messaging (English/Kiswahili).'],
      avoid: ['Do not send bulk messages without reviewing recipient filters.'],
      waitFor: ['Production SMS requires active provider configuration and sufficient balance.']
    },
    {
      id: 'reports-analytics',
      title: 'Reports, AI Insights, and Monitoring',
      panes: ['reports'],
      roles: ['agent', 'admin'],
      keywords: ['reports', 'kpi', 'copilot', 'risk', 'brief', 'alerts', 'ops tasks', 'knowledge'],
      short:
        'Use operational KPIs plus AI cards (Copilot, QC Intelligence, Payment Risk, Briefs) for decision support with approval controls.',
      steps: [
        'Open Reports to review KPI cards, readiness narrative, and agent performance.',
        'Admin: ask Admin Copilot natural language questions.',
        'Run QC Intelligence and Payment Risk checks for actionable flags.',
        'Convert high-risk findings into Ops Tasks and track status.',
        'Maintain AI knowledge docs and feedback to improve response quality.'
      ],
      troubleshooting: [
        'If AI responses fall back to local rules, verify OPENAI_API_KEY and connectivity.',
        'If insights look outdated, refresh data and rerun report analyses.',
        'If proposals/tasks are missing, verify role permissions and filters.'
      ],
      doNow: ['Use AI outputs as recommendations, then require human approval for sensitive actions.'],
      avoid: ['Do not treat AI suggestions as auto-executed transactions.'],
      waitFor: ['Scheduled executive briefs run on configured interval when AI scheduling is enabled.']
    },
    {
      id: 'exports-admin',
      title: 'Exports, Date Ranges, and Admin Tools',
      panes: ['exports'],
      roles: ['admin', 'agent'],
      keywords: ['export', 'csv', 'excel', 'pdf', 'date range', 'backup', 'reset'],
      short:
        'Export operational data by range and format (CSV/Excel/PDF). Admin tools include backups and controlled reset operations.',
      steps: [
        'Open Exports and select date range (today, 7 days, month, all, or custom).',
        'Choose format: CSV, Excel, or PDF.',
        'Click the dataset button to download farmers, QC, purchases, payments, SMS, or activity logs.',
        'Admin: use backup snapshot and backup list before any major changes.',
        'Admin: reset platform data only with confirmed backup and authorization.'
      ],
      troubleshooting: [
        'If export is empty, verify selected range contains records.',
        'If PDF download is blocked, allow browser downloads/popups for this site.',
        'If backup list fails, verify backend storage health.'
      ],
      doNow: ['Use custom date range for weekly/monthly reconciliation packs.'],
      avoid: ['Do not run reset without verified backup recovery path.'],
      waitFor: ['Large exports may take longer on slower connections.']
    },
    {
      id: 'ussd-integration',
      title: 'USSD & Farmer Phone Access',
      panes: ['overview', 'sms', 'reports'],
      roles: ['admin', 'agent', 'farmer'],
      keywords: ['ussd', '384', '483', 'registration via phone', 'kiswahili', 'english'],
      short:
        'USSD supports low-data farmer access via phone menu, with bilingual English/Kiswahili pathways aligned to portal records.',
      steps: [
        'Configure provider callback URL and shared code/channel in the SMS/USSD gateway.',
        'Keep farmer phone and PIN mapping synchronized with portal records.',
        'Use onboarding SMS to notify farmers how to access USSD.',
        'Track usage and troubleshoot via provider portal plus AGEM logs.'
      ],
      troubleshooting: [
        'If USSD sessions fail, validate callback URL reachability and shared secret.',
        'If farmer records do not appear via USSD, verify phone normalization and PIN setup.',
        'If bilingual prompts mismatch, review preferred-language mappings.'
      ],
      doNow: ['Test full USSD registration and payment-status journey in sandbox before production launch.'],
      avoid: ['Do not expose admin-only actions through USSD menus.'],
      waitFor: ['Production USSD activation depends on provider approval and billing setup.']
    },
    {
      id: 'security-governance',
      title: 'Security, Compliance, and Access Control',
      panes: ['overview', 'reports', 'exports'],
      roles: ['admin', 'agent'],
      keywords: ['security', 'audit', 'roles', 'compliance', 'backup', 'approval'],
      short:
        'Platform uses role-based permissions, audit logs, approvals, and backups to protect operations and payment integrity.',
      steps: [
        'Assign roles by responsibility: farmer, agent, admin.',
        'Review activity logs and recommendation decisions regularly.',
        'Rotate credentials and enforce strong passwords for staff.',
        'Keep backups current and run restore tests on schedule.'
      ],
      troubleshooting: [
        'If unauthorized actions appear, check role assignments and session validity.',
        'If account recovery is required, use recovery flow only for admin/agent accounts.',
        'If restore fails, validate backup file integrity and storage backend settings.'
      ],
      doNow: ['Keep approval trail complete for payment-related decisions.'],
      avoid: ['Do not grant admin rights for convenience.'],
      waitFor: ['Production compliance controls should be verified before hard launch.']
    }
  ];
}

function buildMsaidiziModules(store, sourceSignature) {
  const now = nowIso();
  return msaidiziModuleDefinitions(store).map((module) => ({
    ...module,
    updatedAt: now,
    sourceSignature
  }));
}

function syncMsaidiziInStore(store, options = {}) {
  if (!store.aiMsaidizi || typeof store.aiMsaidizi !== 'object') {
    store.aiMsaidizi = {
      lastSyncAt: '',
      sourceSignature: '',
      status: 'never',
      modules: [],
      syncHistory: []
    };
  }

  const reason = clean(options.reason) || 'manual';
  const actor = clean(options.actor) || 'system';
  const force = Boolean(options.force);
  const sourceSignature = buildMsaidiziSourceSignature();
  const existingSignature = clean(store.aiMsaidizi.sourceSignature);
  const hasModules = Array.isArray(store.aiMsaidizi.modules) && store.aiMsaidizi.modules.length > 0;

  let updated = false;
  if (force || !hasModules || sourceSignature !== existingSignature) {
    store.aiMsaidizi.modules = buildMsaidiziModules(store, sourceSignature).slice(0, AI_MSAIDIZI_MODULE_CAP);
    store.aiMsaidizi.sourceSignature = sourceSignature;
    store.aiMsaidizi.lastSyncAt = nowIso();
    store.aiMsaidizi.status = 'ready';
    updated = true;
  } else if (!clean(store.aiMsaidizi.lastSyncAt)) {
    store.aiMsaidizi.lastSyncAt = nowIso();
  }

  const historyEntry = {
    at: nowIso(),
    reason,
    actor,
    updated,
    sourceSignature
  };
  store.aiMsaidizi.syncHistory = [historyEntry, ...(store.aiMsaidizi.syncHistory || [])].slice(0, 120);

  return {
    updated,
    sourceSignature,
    modules: Array.isArray(store.aiMsaidizi.modules) ? store.aiMsaidizi.modules : [],
    status: clean(store.aiMsaidizi.status) || 'ready',
    lastSyncAt: clean(store.aiMsaidizi.lastSyncAt)
  };
}

function msaidiziAccessibleModules(store, role = 'guest') {
  const modules = Array.isArray(store?.aiMsaidizi?.modules) ? store.aiMsaidizi.modules : [];
  const normalizedRole = clean(role).toLowerCase() || 'guest';
  return modules.filter((module) => {
    const allowedRoles = Array.isArray(module?.roles) ? module.roles.map((row) => clean(row).toLowerCase()) : [];
    if (!allowedRoles.length) return true;
    return allowedRoles.includes(normalizedRole) || allowedRoles.includes('all');
  });
}

function msaidiziScoreModule(module, queryTokens, pane = '') {
  const paneKey = clean(pane).toLowerCase();
  const title = clean(module?.title).toLowerCase();
  const short = clean(module?.short).toLowerCase();
  const keywords = Array.isArray(module?.keywords) ? module.keywords.map((row) => clean(row).toLowerCase()) : [];
  const panes = Array.isArray(module?.panes) ? module.panes.map((row) => clean(row).toLowerCase()) : [];

  let score = 0;
  if (paneKey && panes.includes(paneKey)) score += 3;
  if (paneKey && panes.includes('all')) score += 1;
  for (const token of queryTokens) {
    if (!token) continue;
    if (title.includes(token)) score += 4;
    if (short.includes(token)) score += 2;
    if (keywords.some((keyword) => keyword.includes(token))) score += 3;
  }
  return score;
}

function findMsaidiziModules(store, query, pane, role, limit = 4) {
  const modules = msaidiziAccessibleModules(store, role);
  const queryTokens = tokenizeSearchText(query);
  const paneKey = clean(pane).toLowerCase();

  const scored = modules
    .map((module) => ({
      module,
      score: msaidiziScoreModule(module, queryTokens, paneKey)
    }))
    .filter((row) => row.score > 0 || (!queryTokens.length && paneKey));

  if (scored.length) {
    return scored
      .sort((a, b) => b.score - a.score)
      .slice(0, Math.max(1, Math.min(limit, 8)))
      .map((row) => row.module);
  }

  if (paneKey) {
    const paneMatches = modules.filter((module) =>
      Array.isArray(module?.panes)
        ? module.panes.map((row) => clean(row).toLowerCase()).includes(paneKey)
        : false
    );
    if (paneMatches.length) return paneMatches.slice(0, Math.max(1, Math.min(limit, 8)));
  }

  return modules.slice(0, Math.max(1, Math.min(limit, 8)));
}

function buildMsaidiziLocalAnswer(modules, mode = 'normal') {
  const normalizedMode = normalizeMsaidiziMode(mode);
  const primary = modules[0];
  if (!primary) {
    return {
      answerShort: 'No instruction module matched this request yet. Try asking with page name or action.',
      steps: [],
      troubleshoot: ['Open the relevant section first, then ask again with exact action.'],
      doNow: ['Use plain language, for example: "How do I register a farmer?"'],
      avoid: [],
      waitFor: [],
      relatedModules: []
    };
  }

  const answerShort =
    normalizedMode === 'simple'
      ? `Simple guide: ${primary.short}`
      : normalizedMode === 'step_by_step'
        ? `Step-by-step guide for ${primary.title}.`
        : primary.short;

  const steps =
    normalizedMode === 'normal'
      ? (primary.steps || []).slice(0, 5)
      : normalizedMode === 'simple'
        ? (primary.steps || []).slice(0, 4).map((step) => step.replace(/\b(verify|synchronize|reconcile)\b/gi, 'check'))
        : primary.steps || [];

  const relatedModules = modules.slice(0, 4).map((module) => module.title);
  return {
    answerShort,
    steps,
    troubleshoot: (primary.troubleshooting || []).slice(0, 6),
    doNow: (primary.doNow || []).slice(0, 6),
    avoid: (primary.avoid || []).slice(0, 6),
    waitFor: (primary.waitFor || []).slice(0, 6),
    relatedModules
  };
}

async function runScheduledMsaidiziSync() {
  const store = await readStore();
  const result = syncMsaidiziInStore(store, {
    reason: 'scheduled',
    actor: 'system',
    force: false
  });
  if (!result.updated) return result;

  addActivity(store, {
    actor: 'system',
    role: 'system',
    action: 'ai.msaidizi.sync',
    entity: 'ai',
    details: `Msaidizi docs auto-synced (signature ${result.sourceSignature}).`
  });
  await writeStore(store);
  return result;
}

function scheduleMsaidiziSync() {
  if (msaidiziScheduleHandle) {
    clearInterval(msaidiziScheduleHandle);
    msaidiziScheduleHandle = null;
  }
  if (!AI_MSAIDIZI_SYNC_ENABLED) return;

  const intervalMs = AI_MSAIDIZI_SYNC_INTERVAL_HOURS * 60 * 60 * 1000;
  msaidiziScheduleHandle = setInterval(() => {
    runScheduledMsaidiziSync().catch((error) => {
      console.error(`[ai] Scheduled Msaidizi sync failed: ${error.message}`);
    });
  }, intervalMs);
  if (typeof msaidiziScheduleHandle.unref === 'function') {
    msaidiziScheduleHandle.unref();
  }
}

function createOpsTask(store, payload = {}) {
  const now = nowIso();
  const task = {
    id: id('TASK'),
    title: clean(payload.title) || 'Untitled task',
    description: clean(payload.description),
    severity: normalizeRiskSeverity(payload.severity, 'medium'),
    status: normalizeOpsTaskStatus(payload.status || 'open'),
    sourceType: clean(payload.sourceType) || 'ai',
    sourceRef: clean(payload.sourceRef),
    assignedTo: clean(payload.assignedTo),
    notes: clean(payload.notes),
    createdBy: clean(payload.createdBy) || 'system',
    createdAt: now,
    updatedAt: now,
    resolvedAt: ''
  };
  pushWithCap(store.opsTasks, task, OPS_TASK_CAP);
  return task;
}

function buildOpsTasksFromPaymentRisk(report, actor = 'system') {
  const flags = Array.isArray(report?.flags) ? report.flags : [];
  const tasks = [];
  for (const flag of flags) {
    const severity = normalizeRiskSeverity(flag.severity, 'medium');
    if (!['high', 'critical', 'medium'].includes(severity)) continue;
    tasks.push({
      title: `${severity.toUpperCase()}: ${clean(flag.title) || clean(flag.code) || 'Payment risk flag'}`,
      description: clean(flag.detail) || 'Investigate payment anomaly and reconcile before next disbursement.',
      severity,
      sourceType: 'payment-risk',
      sourceRef: clean(flag.paymentId) || clean(flag.ref),
      createdBy: actor
    });
  }
  return tasks.slice(0, 120);
}

function buildOpsTasksFromQcInsights(items, actor = 'system') {
  const rows = Array.isArray(items) ? items : [];
  const tasks = [];
  for (const insight of rows) {
    const level = normalizeRiskSeverity(insight?.riskLevel, 'low');
    if (!['high', 'medium', 'critical'].includes(level)) continue;
    tasks.push({
      title: `${level.toUpperCase()}: QC lot follow-up (${clean(insight.farmerName) || 'Unknown farmer'})`,
      description:
        clean(insight.summary) ||
        `Review lot ${clean(insight.qcRecordId)} and confirm ${clean(insight.recommendedDecision) || 'next QC decision'}.`,
      severity: level,
      sourceType: 'qc-intelligence',
      sourceRef: clean(insight.qcRecordId),
      createdBy: actor
    });
  }
  return tasks.slice(0, 120);
}

function buildAiFeedbackSummary(feedbackRows) {
  const rows = Array.isArray(feedbackRows) ? feedbackRows : [];
  const byTool = {};

  for (const row of rows) {
    const tool = clean(row?.tool) || 'unknown';
    const rating = clean(row?.rating).toLowerCase();
    if (!byTool[tool]) {
      byTool[tool] = {
        tool,
        total: 0,
        positive: 0,
        negative: 0,
        neutral: 0
      };
    }
    byTool[tool].total += 1;
    if (rating === 'up' || rating === 'positive') {
      byTool[tool].positive += 1;
    } else if (rating === 'down' || rating === 'negative') {
      byTool[tool].negative += 1;
    } else {
      byTool[tool].neutral += 1;
    }
  }

  const totals = Object.values(byTool).reduce(
    (acc, row) => {
      acc.total += row.total;
      acc.positive += row.positive;
      acc.negative += row.negative;
      acc.neutral += row.neutral;
      return acc;
    },
    { total: 0, positive: 0, negative: 0, neutral: 0 }
  );
  const positiveRate = totals.total ? Number(((totals.positive / totals.total) * 100).toFixed(1)) : 0;

  return {
    totals: {
      ...totals,
      positiveRatePct: positiveRate
    },
    byTool: Object.values(byTool)
      .map((row) => ({
        ...row,
        positiveRatePct: row.total ? Number(((row.positive / row.total) * 100).toFixed(1)) : 0
      }))
      .sort((a, b) => b.total - a.total)
  };
}

function runAiEvalSuite(store) {
  const checks = [];

  const copilotCheck = localCopilotAnswer(store, 'show farmers owed this week in Nakuru');
  checks.push({
    id: 'copilot_intent_owed',
    label: 'Copilot detects owed-farmers intent',
    pass: clean(copilotCheck.intent) === 'owed_farmers',
    details: `Intent="${clean(copilotCheck.intent)}"`
  });

  const riskReport = localPaymentRiskReport(store, { period: 'week' });
  checks.push({
    id: 'payment_risk_schema',
    label: 'Payment risk report returns summary and flags',
    pass: Boolean(riskReport && riskReport.summary && Array.isArray(riskReport.flags)),
    details: `flags=${Array.isArray(riskReport.flags) ? riskReport.flags.length : 0}`
  });

  const produceSample = Array.isArray(store.produce) ? store.produce[0] : null;
  if (produceSample) {
    const qc = localQcIntelligence(produceSample);
    checks.push({
      id: 'qc_decision_bounds',
      label: 'QC intelligence returns valid decision',
      pass: ['Accept', 'Hold', 'Reject'].includes(clean(qc.recommendedDecision)),
      details: `recommendedDecision="${clean(qc.recommendedDecision)}"`
    });
  } else {
    checks.push({
      id: 'qc_decision_bounds',
      label: 'QC intelligence returns valid decision',
      pass: true,
      details: 'Skipped (no QC rows available).'
    });
  }

  const knowledge = findKnowledgeMatches(store, 'payment reconciliation workflow', 2);
  checks.push({
    id: 'knowledge_retrieval',
    label: 'Knowledge retrieval executes without error',
    pass: Array.isArray(knowledge),
    details: `matches=${Array.isArray(knowledge) ? knowledge.length : 0}`
  });

  return checks;
}

async function generateExecutiveBrief(store, options = {}) {
  const trigger = clean(options.trigger) || 'manual';
  const requestedBy = clean(options.requestedBy) || 'system';
  const generatedAt = nowIso();

  const summary = makeSummary(store);
  const agentStats = makeAgentStats(store).slice(0, 5);
  const paymentRisk = localPaymentRiskReport(store, { period: 'week' });
  const recentQc = (store.produce || [])
    .slice()
    .sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt))
    .slice(0, 25)
    .map((row) => localQcIntelligence(row));

  const qcHighRisk = recentQc.filter((row) => row.riskLevel === 'high').length;
  const qcMediumRisk = recentQc.filter((row) => row.riskLevel === 'medium').length;
  const riskSignals = [];
  if ((summary.owedFarmers || 0) > 0) {
    riskSignals.push(`${summary.owedFarmers} farmer(s) currently owed (KES ${formatKes(summary.totalOwedKes || 0)}).`);
  }
  if ((paymentRisk.summary?.highFlags || 0) > 0) {
    riskSignals.push(`${paymentRisk.summary.highFlags} high payment-risk flag(s) in the last 7 days.`);
  }
  if (qcHighRisk > 0) {
    riskSignals.push(`${qcHighRisk} high-risk QC lot(s) in the latest inspections.`);
  }
  if (!riskSignals.length) {
    riskSignals.push('No critical operational risks detected in current snapshot.');
  }

  const actions = [];
  if ((summary.owedFarmers || 0) > 0) {
    actions.push('Open Payments > Farmers Owed and prepare settlement batch for due farmers.');
  }
  if ((paymentRisk.summary?.highFlags || 0) > 0 || (paymentRisk.summary?.mediumFlags || 0) > 0) {
    actions.push('Review Payment Risk flags and convert high-risk findings into ops tasks for follow-up.');
  }
  if (qcHighRisk > 0 || qcMediumRisk > 0) {
    actions.push('Run re-sampling on medium/high QC lots before dispatch confirmation.');
  }
  if (!actions.length) {
    actions.push('Maintain routine monitoring and continue weekly KPI review.');
  }

  let narrative =
    `Daily executive brief (${generatedAt}): ${summary.farmers || 0} farmers, ${summary.qcRecords || 0} QC records, ` +
    `${summary.purchasedRecords || 0} purchases, payment success ${summary.paymentSuccessRate || 0}%, ` +
    `SMS spend KES ${formatKes(summary.smsSpentLast24hKes || 0)} in last 24h.`;
  let source = 'local-rules';
  let model = 'local-rules';
  let warning = '';

  if (OPENAI_API_KEY) {
    const openAi = await callOpenAiJson(
      aiPromptPolicy(
        store,
        'executiveBriefSystem',
        'You are an executive operations analyst. Create concise, decision-oriented summaries. Do not claim to execute changes.'
      ),
      `Create a short executive brief using this snapshot:
- Trigger: ${trigger}
- Requested by: ${requestedBy}
- Farmers: ${summary.farmers || 0}
- QC records: ${summary.qcRecords || 0}
- Purchases: ${summary.purchasedRecords || 0}
- Owed farmers: ${summary.owedFarmers || 0}
- Total owed KES: ${summary.totalOwedKes || 0}
- Payment success: ${summary.paymentSuccessRate || 0}%
- SMS spend 24h KES: ${summary.smsSpentLast24hKes || 0}
- Payment risk summary: ${paymentRisk.narrative}
- QC risk counts: high=${qcHighRisk}, medium=${qcMediumRisk}
- Baseline narrative: ${narrative}`,
      {
        type: 'object',
        additionalProperties: false,
        required: ['narrative', 'actions'],
        properties: {
          narrative: { type: 'string' },
          actions: {
            type: 'array',
            minItems: 1,
            maxItems: 6,
            items: { type: 'string' }
          }
        }
      }
    );

    if (openAi.ok) {
      source = 'openai';
      model = OPENAI_MODEL;
      narrative = clean(openAi.data?.narrative) || narrative;
      const aiActions = Array.isArray(openAi.data?.actions)
        ? openAi.data.actions.map((row) => clean(row)).filter(Boolean).slice(0, 6)
        : [];
      if (aiActions.length) {
        actions.length = 0;
        actions.push(...aiActions);
      }
    } else {
      warning = openAi.error || 'OpenAI unavailable. Generated local executive brief.';
    }
  }

  const overallRisk = paymentRisk.summary?.overallRisk || (qcHighRisk > 0 ? 'high' : qcMediumRisk > 0 ? 'medium' : 'low');
  const brief = {
    id: id('BRIEF'),
    generatedAt,
    trigger,
    requestedBy,
    source,
    model,
    warning,
    summary,
    paymentRiskSummary: paymentRisk.summary || {},
    agentSnapshot: agentStats,
    riskSignals: riskSignals.slice(0, 10),
    actions: actions.slice(0, 8),
    narrative,
    overallRisk
  };

  if (overallRisk === 'high') {
    brief.alert = {
      id: id('AIALERT'),
      createdAt: generatedAt,
      severity: 'high',
      title: 'High operational risk detected in executive brief',
      detail: narrative
    };
  }

  return brief;
}

async function executeAiProposal(store, proposal, actor = 'system') {
  const actionType = clean(proposal?.actionType).toLowerCase();
  const payload = proposal?.payload && typeof proposal.payload === 'object' ? proposal.payload : {};

  if (!actionType) {
    throw new Error('Proposal actionType is missing.');
  }

  if (actionType === 'payment_risk_to_tasks') {
    const report = localPaymentRiskReport(store, {
      period: clean(payload?.period) || 'week',
      from: clean(payload?.from),
      to: clean(payload?.to)
    });
    const templates = buildOpsTasksFromPaymentRisk(report, actor);
    const created = templates.map((row) => createOpsTask(store, row));
    return {
      actionType,
      tasksCreated: created.length,
      taskIds: created.map((row) => row.id),
      summary: `Created ${created.length} ops task(s) from payment risk flags.`
    };
  }

  if (actionType === 'qc_risk_to_tasks') {
    const qcRecordId = clean(payload?.qcRecordId);
    const limitRaw = Number(payload?.limit);
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(120, Math.floor(limitRaw))) : 30;
    const allQc = [...(store.produce || [])].sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt));
    const selected = qcRecordId ? allQc.filter((row) => row.id === qcRecordId) : allQc.slice(0, limit);
    const insights = selected.map((row) => localQcIntelligence(row));
    const templates = buildOpsTasksFromQcInsights(insights, actor);
    const created = templates.map((row) => createOpsTask(store, row));
    return {
      actionType,
      tasksCreated: created.length,
      taskIds: created.map((row) => row.id),
      summary: `Created ${created.length} ops task(s) from QC intelligence.`
    };
  }

  if (actionType === 'create_ops_tasks') {
    const tasks = Array.isArray(payload?.tasks) ? payload.tasks : [];
    if (!tasks.length) {
      throw new Error('No tasks payload provided.');
    }
    const created = tasks.map((task) =>
      createOpsTask(store, {
        title: task?.title,
        description: task?.description,
        severity: task?.severity,
        sourceType: clean(task?.sourceType) || 'ai-proposal',
        sourceRef: clean(task?.sourceRef) || clean(proposal?.id),
        createdBy: actor
      })
    );
    return {
      actionType,
      tasksCreated: created.length,
      taskIds: created.map((row) => row.id),
      summary: `Created ${created.length} custom ops task(s).`
    };
  }

  if (actionType === 'generate_executive_brief') {
    const brief = await generateExecutiveBrief(store, {
      trigger: 'proposal-execution',
      requestedBy: actor
    });
    pushWithCap(store.aiBriefs, brief, AI_EXEC_BRIEF_MAX_HISTORY);
    if (brief.alert) {
      pushWithCap(store.aiAlerts, brief.alert, AI_EXEC_BRIEF_MAX_HISTORY);
    }
    return {
      actionType,
      tasksCreated: 0,
      taskIds: [],
      briefId: brief.id,
      summary: `Generated executive brief ${brief.id}.`
    };
  }

  throw new Error(`Unsupported proposal actionType "${actionType}".`);
}

async function runScheduledExecutiveBrief() {
  const store = await readStore();
  const brief = await generateExecutiveBrief(store, {
    trigger: 'scheduled',
    requestedBy: 'system'
  });
  pushWithCap(store.aiBriefs, brief, AI_EXEC_BRIEF_MAX_HISTORY);
  if (brief.alert) {
    pushWithCap(store.aiAlerts, brief.alert, AI_EXEC_BRIEF_MAX_HISTORY);
  }
  addActivity(store, {
    actor: 'system',
    role: 'system',
    action: 'ai.brief.scheduled',
    entity: 'ai',
    entityId: brief.id,
    details: `Generated scheduled executive brief (${brief.overallRisk} risk)`
  });
  await writeStore(store);
  return brief;
}

function scheduleExecutiveBriefs() {
  if (executiveBriefScheduleHandle) {
    clearInterval(executiveBriefScheduleHandle);
    executiveBriefScheduleHandle = null;
  }
  if (!AI_EXEC_BRIEF_ENABLED) return;

  const intervalMs = AI_EXEC_BRIEF_INTERVAL_HOURS * 60 * 60 * 1000;
  executiveBriefScheduleHandle = setInterval(() => {
    runScheduledExecutiveBrief().catch((error) => {
      console.error(`[ai] Scheduled executive brief failed: ${error.message}`);
    });
  }, intervalMs);
  if (typeof executiveBriefScheduleHandle.unref === 'function') {
    executiveBriefScheduleHandle.unref();
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

function readTextBody(req, maxBytes = 1_000_000) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk;
      if (body.length > maxBytes) reject(new Error('Request too large'));
    });
    req.on('end', () => {
      resolve(body);
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
    phone: currentUser.phone || '',
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

function normalizeLanguage(value) {
  const lower = clean(value).toLowerCase();
  if (['2', 'sw', 'kiswahili', 'swahili'].includes(lower)) return 'sw';
  if (['1', 'en', 'english'].includes(lower)) return 'en';
  return '';
}

function languageOrDefault(value) {
  return normalizeLanguage(value) || DEFAULT_LANGUAGE;
}

function parsePreferredLanguage(value) {
  const raw = clean(value);
  if (!raw) {
    return { value: DEFAULT_LANGUAGE, valid: true };
  }
  const normalized = normalizeLanguage(raw);
  if (!normalized) {
    return { value: DEFAULT_LANGUAGE, valid: false };
  }
  return { value: normalized, valid: true };
}

function languageLabel(value) {
  return languageOrDefault(value) === 'sw' ? 'Kiswahili' : 'English';
}

function inLanguage(lang, englishText, kiswahiliText) {
  return languageOrDefault(lang) === 'sw' ? kiswahiliText : englishText;
}

function normalizePhone(value) {
  const digits = clean(value).replace(/[^\d]/g, '');
  if (!digits) return '';
  if (digits.startsWith('254') && digits.length === 12) return digits;
  if (digits.startsWith('0') && digits.length === 10) return `254${digits.slice(1)}`;
  if (digits.length === 9 && digits.startsWith('7')) return `254${digits}`;
  return digits;
}

function isValidKenyaPhone(value) {
  return /^254\d{9}$/.test(clean(value));
}

function isTruthy(value) {
  if (typeof value === 'boolean') return value;
  const lower = clean(value).toLowerCase();
  return lower === 'true' || lower === '1' || lower === 'yes' || lower === 'on';
}

function defaultOnboardingSmsTemplate(lang) {
  return languageOrDefault(lang) === 'sw' ? IMPORT_ONBOARDING_SMS_DEFAULT_SW : IMPORT_ONBOARDING_SMS_DEFAULT_EN;
}

function renderImportOnboardingSms(template, farmer) {
  const lang = languageOrDefault(farmer?.preferredLanguage);
  const rawTemplate = clean(template) || defaultOnboardingSmsTemplate(lang);
  const values = {
    name: clean(farmer?.name) || 'farmer',
    phone: clean(farmer?.phone),
    nationalId: clean(farmer?.nationalId),
    location: clean(farmer?.location),
    ussd: USSD_CODE,
    portal: 'Agem Portal'
  };

  const rendered = rawTemplate.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_, key) => values[key] || '');
  return clean(rendered);
}

function normalizedUsername(value) {
  return clean(value).toLowerCase();
}

function findUserByUsername(store, username) {
  const target = normalizedUsername(username);
  return store.users.find((row) => normalizedUsername(row.username) === target);
}

function findUserByPhone(store, phone) {
  const target = normalizePhone(phone);
  if (!target) return null;
  return (
    store.users.find((row) => normalizePhone(row.phone || row.username) === target) || null
  );
}

function findFarmerUserById(store, userId) {
  const target = clean(userId);
  if (!target) return null;
  return store.users.find((row) => row.id === target && row.role === 'farmer') || null;
}

function resolveFarmerPortalUser(store, farmer) {
  if (!farmer) return null;
  const byId = findFarmerUserById(store, farmer.userId);
  if (byId) return byId;
  const byPhone = findUserByPhone(store, farmer.phone);
  if (byPhone && byPhone.role === 'farmer') return byPhone;
  return null;
}

function ensureFarmerPortalUser(store, farmer, pin) {
  const normalizedFarmerPhone = normalizePhone(farmer?.phone) || clean(farmer?.phone);
  if (!normalizedFarmerPhone) {
    return { error: 'Farmer phone number is required before setting a PIN.' };
  }

  const conflictByPhone = findUserByPhone(store, normalizedFarmerPhone);
  if (conflictByPhone && conflictByPhone.role !== 'farmer') {
    return { error: 'This phone number is already linked to a non-farmer account.' };
  }

  let user = resolveFarmerPortalUser(store, farmer);
  if (!user && conflictByPhone && conflictByPhone.role === 'farmer') {
    user = conflictByPhone;
  }
  if (user && conflictByPhone && conflictByPhone.id !== user.id) {
    return { error: 'Phone number is linked to a different farmer account.' };
  }

  const conflictByUsername = findUserByUsername(store, normalizedFarmerPhone);
  if (conflictByUsername && conflictByUsername.role !== 'farmer') {
    return { error: 'This phone number conflicts with a non-farmer username.' };
  }
  if (!user && conflictByUsername && conflictByUsername.role === 'farmer') {
    user = conflictByUsername;
  }
  if (user && conflictByUsername && conflictByUsername.id !== user.id) {
    return { error: 'Username is linked to a different farmer account.' };
  }

  const now = nowIso();
  let created = false;
  if (!user) {
    created = true;
    user = {
      id: id('USR'),
      username: normalizedFarmerPhone,
      phone: normalizedFarmerPhone,
      name: clean(farmer.name),
      role: 'farmer',
      status: 'active',
      password: hashPassword(pin),
      createdAt: now,
      updatedAt: now
    };
    store.users.push(user);
  } else {
    user.username = normalizedFarmerPhone;
    user.phone = normalizedFarmerPhone;
    user.name = clean(farmer.name);
    user.role = 'farmer';
    user.status = 'active';
    user.password = hashPassword(pin);
    user.updatedAt = now;
  }

  farmer.phone = normalizedFarmerPhone;
  farmer.userId = user.id;
  farmer.updatedAt = now;
  return { user, created };
}

function syncFarmerPortalUserIdentity(store, farmer, nextPhone, nextName) {
  const normalizedFarmerPhone = normalizePhone(nextPhone) || clean(nextPhone);
  if (!normalizedFarmerPhone) {
    return { ok: false, error: 'Farmer phone number is required.' };
  }

  const user = resolveFarmerPortalUser(store, farmer);
  if (!user) {
    return { ok: true, updated: false, user: null };
  }

  const portalConflictPhone = findUserByPhone(store, normalizedFarmerPhone);
  if (portalConflictPhone && portalConflictPhone.id !== user.id) {
    return { ok: false, error: 'Farmer phone conflicts with another portal account.' };
  }
  const portalConflictUsername = findUserByUsername(store, normalizedFarmerPhone);
  if (portalConflictUsername && portalConflictUsername.id !== user.id) {
    return { ok: false, error: 'Farmer phone conflicts with another portal username.' };
  }

  user.username = normalizedFarmerPhone;
  user.phone = normalizedFarmerPhone;
  user.name = clean(nextName);
  user.updatedAt = nowIso();
  farmer.userId = user.id;

  return { ok: true, updated: true, user };
}

function normalizedEmail(value) {
  return clean(value).toLowerCase();
}

function isValidEmail(value) {
  const email = normalizedEmail(value);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function findUserByEmail(store, email) {
  const target = normalizedEmail(email);
  if (!target) return null;
  return store.users.find((row) => normalizedEmail(row.email) === target) || null;
}

function baseAgentUsername(name) {
  const normalized = clean(name)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '.')
    .replace(/^\.+|\.+$/g, '');
  const fallback = normalized || 'agent';
  return fallback.slice(0, 24);
}

function generateAgentUsername(store, name) {
  const base = baseAgentUsername(name);
  for (let i = 0; i < 200; i += 1) {
    const suffix = String(crypto.randomInt(1000, 9999));
    let candidate = `${base}.${suffix}`;
    if (candidate.length > 32) {
      const maxBase = Math.max(3, 32 - (suffix.length + 1));
      candidate = `${base.slice(0, maxBase)}.${suffix}`;
    }
    if (!findUserByUsername(store, candidate)) {
      return candidate;
    }
  }
  throw new Error('Could not generate a unique username. Try again.');
}

function generateTemporaryPassword(length = 12) {
  const lowers = 'abcdefghijkmnopqrstuvwxyz';
  const uppers = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const digits = '23456789';
  const symbols = '!@#$%&*?';
  const all = `${lowers}${uppers}${digits}${symbols}`;

  const chars = [
    lowers[crypto.randomInt(0, lowers.length)],
    uppers[crypto.randomInt(0, uppers.length)],
    digits[crypto.randomInt(0, digits.length)],
    symbols[crypto.randomInt(0, symbols.length)]
  ];
  while (chars.length < length) {
    chars.push(all[crypto.randomInt(0, all.length)]);
  }

  for (let i = chars.length - 1; i > 0; i -= 1) {
    const j = crypto.randomInt(0, i + 1);
    const tmp = chars[i];
    chars[i] = chars[j];
    chars[j] = tmp;
  }

  return chars.join('');
}

function generateNumericPin(length = 4) {
  const max = 10 ** length;
  return String(crypto.randomInt(0, max)).padStart(length, '0');
}

function agentView(user) {
  return {
    id: user.id,
    username: user.username,
    name: user.name || '',
    email: user.email || '',
    phone: user.phone || '',
    role: user.role || 'agent',
    status: user.status || 'active',
    provisioning: user.provisioning || 'admin',
    createdAt: user.createdAt || '',
    updatedAt: user.updatedAt || user.createdAt || ''
  };
}

function parseNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : NaN;
}

function parseTrees(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return 0;
  const direct = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(direct)) return direct;
  const match = raw.match(/-?\d[\d,]*(?:\.\d+)?/);
  if (!match) return NaN;
  const num = Number(match[0].replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function parseArea(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return NaN;
  const direct = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(direct)) return direct;
  const match = raw.match(/-?\d[\d,]*(?:\.\d+)?/);
  if (!match) return NaN;
  const num = Number(match[0].replace(/,/g, ''));
  return Number.isFinite(num) ? num : NaN;
}

function parseAreaToHectares(value, defaultUnit = 'hectares') {
  const raw = clean(value);
  if (!raw) return NaN;
  const num = parseArea(raw);
  if (!Number.isFinite(num)) return NaN;
  const lower = raw.toLowerCase();
  if (/(hectare|hectares|\bha\b)/.test(lower)) return num;
  if (/(acre|acres|acreage)/.test(lower)) return num * HECTARES_PER_ACRE;
  if (/(square\s*feet|square\s*foot|sq\s*ft|sqft|ft2|ft²)/.test(lower)) return num / SQFT_PER_HECTARE;
  if (defaultUnit === 'acres') return num * HECTARES_PER_ACRE;
  if (defaultUnit === 'squareFeet') return num / SQFT_PER_HECTARE;
  return num;
}

function resolveHectaresFromValues(hectaresRaw, acresRaw, squareFeetRaw) {
  const hasHectares = clean(hectaresRaw) !== '';
  const hasAcres = clean(acresRaw) !== '';
  const hasSquareFeet = clean(squareFeetRaw) !== '';

  if (hasHectares) return parseAreaToHectares(hectaresRaw, 'hectares');
  if (hasAcres) {
    return parseAreaToHectares(acresRaw, 'acres');
  }
  if (hasSquareFeet) {
    return parseAreaToHectares(squareFeetRaw, 'squareFeet');
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

  const normalizedTarget = normalizePhone(target);
  if (normalizedTarget) {
    return (
      list.find((row) => normalizePhone(row.phone) === normalizedTarget && row.id !== excludeId) || null
    );
  }

  return list.find((row) => clean(row.phone) === target && row.id !== excludeId) || null;
}

function findFarmerByPhoneNormalized(list, phone) {
  const target = normalizePhone(phone);
  if (!target) return null;
  return list.find((row) => normalizePhone(row.phone) === target) || null;
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
  if (!parsePreferredLanguage(payload.preferredLanguage).valid) return 'preferredLanguage must be en or sw';
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

function formatAreaHint(hectares) {
  const num = Number(hectares);
  if (!Number.isFinite(num) || num <= 0) return '';
  const acres = num * ACRES_PER_HECTARE;
  return `${num.toFixed(3)} ha (~${acres.toFixed(2)} acres)`;
}

function buildFarmerImportIssue(mapped, invalid) {
  const totalHectares = resolveHectares(mapped);
  const avocadoHectares = resolveAvocadoHectares(mapped);
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
        ? `Set area under avocado to <= total farm size (${formatAreaHint(totalHectares)}), or fix units in the source row.`
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
  if (invalid === 'preferredLanguage must be en or sw') {
    issue.code = 'invalid_language';
    issue.suggestion = 'Use English/en/1 or Kiswahili/sw/2 for preferred language.';
    return issue;
  }

  return issue;
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

function joinNonEmpty(values, separator = ' ') {
  return values
    .map((value) => clean(value))
    .filter(Boolean)
    .join(separator);
}

function mapFarmerImportRecord(raw) {
  const flat = {};
  for (const [key, value] of Object.entries(raw || {})) {
    flat[normalizeImportHeader(key)] = value;
  }

  const firstName = firstNonEmpty([flat.firstname, flat.firstnameseller, flat.sellerfirstname]);
  const middleName = firstNonEmpty([flat.middlename, flat.sellermiddlename]);
  const lastName = firstNonEmpty([flat.lastname, flat.lastnameseller, flat.sellerlastname, flat.surname]);
  const composedName = joinNonEmpty([firstName, middleName, lastName]);
  const name = firstNonEmpty([flat.name, flat.farmername, flat.fullname, flat.farmerfullname, flat.growername, flat.sellername, composedName]);
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
    flat.governmentidnumber,
    flat.identificationnumber
  ]);
  const explicitLocation = firstNonEmpty([
    flat.location,
    flat.area,
    flat.village,
    flat.farmlocation,
    flat.nearestcollectionpointdepot
  ]);
  const ward = firstNonEmpty([flat.ward, flat.locationward, flat.wardname, flat['2']]);
  const county = firstNonEmpty([flat.county, flat.region, flat.locationcounty, flat.countyname, flat['3']]);
  const wardCountyLocation = joinNonEmpty([ward, county], ', ');
  const location = firstNonEmpty([explicitLocation, wardCountyLocation, ward, county, flat.locationofsocietycooperative]);
  const notes = firstNonEmpty([
    flat.notes,
    flat.note,
    flat.comments,
    flat.remarks,
    flat.description,
    flat.additionalinformationcommentsnotes
  ]);
  const treesRaw = firstNonEmpty([
    flat.trees,
    flat.treecount,
    flat.numberoftrees,
    flat.treequantity,
    flat.treenumber,
    flat.numberapproximateofhasstrees,
    flat.numberofhasstrees
  ]);
  const trees = parseTrees(treesRaw);
  const hectaresRaw = firstNonEmpty([
    flat.hectares,
    flat.hectare,
    flat.farmsizeha,
    flat.farmsizehectares,
    flat.landsizehectares,
    flat.totalfarmsizehectares
  ]);
  const acresRaw = firstNonEmpty([
    flat.acres,
    flat.acre,
    flat.acreage,
    flat.farmsizeacres,
    flat.landsizeacres,
    flat.farmsize,
    flat.totalfarmsize
  ]);
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
    flat.avocadoplotacres,
    flat.areaunderhassavocado,
    flat.areaunderavocado
  ]);
  const avocadoSquareFeetRaw = firstNonEmpty([
    flat.avocadosquarefeet,
    flat.avocadosquarefoot,
    flat.avocadosqft,
    flat.areaunderavocadosquarefeet,
    flat.areaunderavocadosqft
  ]);
  const preferredLanguageRaw = firstNonEmpty([
    flat.preferredlanguage,
    flat.language,
    flat.farmerlanguage,
    flat.smslanguage,
    flat.ussdlanguage,
    flat.lugha
  ]);
  const preferredLanguage = parsePreferredLanguage(preferredLanguageRaw).value;
  const hectares = resolveHectares({ hectares: hectaresRaw, acres: acresRaw, squareFeet: squareFeetRaw });
  const avocadoHectares = resolveAvocadoHectares({
    avocadoHectares: avocadoHectaresRaw,
    avocadoAcres: avocadoAcresRaw,
    avocadoSquareFeet: avocadoSquareFeetRaw
  });

  return {
    name,
    phone: normalizePhone(phone) || phone,
    nationalId,
    location,
    trees,
    hectares,
    avocadoHectares,
    preferredLanguage,
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

function isBillableSmsStatus(status) {
  const normalized = clean(status).toLowerCase();
  if (!normalized) return true;
  return !['failed', 'rejected', 'cancelled', 'undelivered'].includes(normalized);
}

function smsOwnerCostKes(log) {
  if (!isBillableSmsStatus(log?.status)) return 0;
  const explicitOwnerCost = parseNumber(log?.ownerCostKes);
  if (Number.isFinite(explicitOwnerCost) && explicitOwnerCost >= 0) return money(explicitOwnerCost);
  const legacyExplicitCost = parseNumber(log?.costKes);
  if (Number.isFinite(legacyExplicitCost) && legacyExplicitCost >= 0) return money(legacyExplicitCost);
  return money(SMS_OWNER_COST_PER_MESSAGE_KES);
}

function createSmsLog(store, payload) {
  const phone = clean(payload?.phone);
  const message = clean(payload?.message);
  if (!phone || !message) return null;

  const log = {
    id: id('SMS'),
    farmerId: clean(payload?.farmerId),
    farmerName: clean(payload?.farmerName),
    phone,
    message,
    provider: clean(payload?.provider) || 'AfricaTalking(Mock)',
    ownerCostKes: money(SMS_OWNER_COST_PER_MESSAGE_KES),
    status: clean(payload?.status) || 'Sent',
    createdBy: clean(payload?.createdBy) || 'system',
    createdAt: nowIso()
  };

  store.smsLogs.unshift(log);
  return log;
}

function smsCostSnapshot(smsLogs, nowMs = Date.now()) {
  const logs = Array.isArray(smsLogs) ? smsLogs : [];
  const cutoffMs = nowMs - 24 * 60 * 60 * 1000;
  let smsBillable = 0;
  let smsBillableLast24h = 0;
  let smsSpentKes = 0;
  let smsSpentLast24hKes = 0;

  for (const log of logs) {
    const costKes = smsOwnerCostKes(log);
    if (costKes > 0) {
      smsBillable += 1;
      smsSpentKes += costKes;
    }

    const createdAtMs = dateMs(log?.createdAt);
    if (createdAtMs >= cutoffMs && costKes > 0) {
      smsBillableLast24h += 1;
      smsSpentLast24hKes += costKes;
    }
  }

  return {
    smsOwnerCostPerMessageKes: money(SMS_OWNER_COST_PER_MESSAGE_KES),
    smsBillable,
    smsBillableLast24h,
    smsSpentKes: money(smsSpentKes),
    smsSpentLast24hKes: money(smsSpentLast24hKes)
  };
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

function dateShort(value) {
  if (!value) return '';
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return '';
  return date.toISOString().slice(0, 10);
}

function formatKes(value) {
  const num = Number(value || 0);
  if (!Number.isFinite(num)) return '0';
  return Math.round(num).toLocaleString('en-US');
}

function ussdReply(res, textBody, closeSession = false) {
  const prefix = closeSession ? 'END ' : 'CON ';
  text(res, 200, `${prefix}${textBody}`);
}

function ussdLanguageMenu() {
  return [
    'Agem Portal',
    'Choose language / Chagua lugha',
    '1. English',
    '2. Kiswahili'
  ].join('\n');
}

function ussdMainMenu(lang = DEFAULT_LANGUAGE) {
  if (languageOrDefault(lang) === 'sw') {
    return [
      'Agem Portal',
      '1. Hali ya Malipo',
      '2. QC ya Mazao ya Hivi Karibuni',
      '3. Wasifu Wangu',
      '4. Msaada'
    ].join('\n');
  }
  return [
    'Agem Portal',
    '1. Payment Status',
    '2. Latest Produce QC',
    '3. My Profile',
    '4. Help'
  ].join('\n');
}

function ussdUnregisteredMenu(lang = DEFAULT_LANGUAGE) {
  if (languageOrDefault(lang) === 'sw') {
    return [
      'Agem Portal',
      'Nambari hii haijasajiliwa.',
      '1. Jisajili Sasa',
      '2. Msaada'
    ].join('\n');
  }
  return [
    'Agem Portal',
    'This number is not registered.',
    '1. Register Now',
    '2. Help'
  ].join('\n');
}

function ussdHelpMessage(lang = DEFAULT_LANGUAGE) {
  return inLanguage(
    lang,
    `Agem Portal Help\nUse ${USSD_CODE} to check payment, QC status, and profile.\nSupport: ${USSD_HELP_PHONE}`,
    `Msaada wa Agem Portal\nTumia ${USSD_CODE} kuangalia malipo, QC, na wasifu.\nMsaada: ${USSD_HELP_PHONE}`
  );
}

function ussdRegistrationPrompt(lang, stepNumber) {
  if (languageOrDefault(lang) === 'sw') {
    if (stepNumber === 1) return 'Usajili (Hatua 1/7)\nWeka jina lako kamili:';
    if (stepNumber === 2) return 'Usajili (Hatua 2/7)\nWeka nambari ya kitambulisho (National ID):';
    if (stepNumber === 3) return 'Usajili (Hatua 3/7)\nWeka eneo lako (mtaa/kata):';
    if (stepNumber === 4) return 'Usajili (Hatua 4/7)\nWeka ukubwa wa shamba kwa acres:';
    if (stepNumber === 5) return 'Usajili (Hatua 5/7)\nWeka eneo la parachichi kwa acres:';
    if (stepNumber === 6) return 'Usajili (Hatua 6/7)\nWeka PIN ya namba 4:';
    return 'Usajili (Hatua 7/7)\nRudia PIN ya namba 4 kuthibitisha:';
  }

  if (stepNumber === 1) return 'Registration (Step 1/7)\nEnter your full name:';
  if (stepNumber === 2) return 'Registration (Step 2/7)\nEnter your National ID number:';
  if (stepNumber === 3) return 'Registration (Step 3/7)\nEnter your location (ward/area):';
  if (stepNumber === 4) return 'Registration (Step 4/7)\nEnter total farm size in acres:';
  if (stepNumber === 5) return 'Registration (Step 5/7)\nEnter area under avocado in acres:';
  if (stepNumber === 6) return 'Registration (Step 6/7)\nSet a 4-digit PIN:';
  return 'Registration (Step 7/7)\nConfirm your 4-digit PIN:';
}

function ussdRegistrationSmsMessage(farmer, lang = DEFAULT_LANGUAGE) {
  return inLanguage(
    lang,
    `Agem Portal: Hello ${farmer.name}. Registration complete. Dial ${USSD_CODE} for payment status, QC updates, and profile.`,
    `Agem Portal: Habari ${farmer.name}. Usajili umekamilika. Piga ${USSD_CODE} kuona malipo, taarifa za QC, na wasifu wako.`
  );
}

function ussdRegistrationCompleteMessage(farmer, lang = DEFAULT_LANGUAGE) {
  return inLanguage(
    lang,
    `Registration complete for ${farmer.name}.\nDial ${USSD_CODE} any time for payment status, QC updates, and profile.`,
    `Usajili umekamilika kwa ${farmer.name}.\nPiga ${USSD_CODE} wakati wowote kuona malipo, taarifa za QC, na wasifu.`
  );
}

function registerFarmerFromUssd(store, payload) {
  const lang = languageOrDefault(payload.language);
  const phone = normalizePhone(payload.phoneNumber) || clean(payload.phoneNumber);
  const pin = String(payload.pin || '').trim();
  const name = clean(payload.name);
  const nationalId = cleanNationalId(payload.nationalId);
  const location = clean(payload.location);
  const totalAcres = Number(payload.totalAcres);
  const avocadoAcres = Number(payload.avocadoAcres);
  const totalHectares = totalAcres * HECTARES_PER_ACRE;
  const avocadoHectares = avocadoAcres * HECTARES_PER_ACRE;

  const user = {
    id: id('USR'),
    username: phone,
    phone,
    name,
    role: 'farmer',
    status: 'active',
    password: hashPassword(pin),
    createdAt: nowIso(),
    updatedAt: nowIso()
  };
  store.users.push(user);

  const record = {
    id: id('F'),
    name,
    phone,
    nationalId,
    location,
    trees: 0,
    hectares: Number(totalHectares.toFixed(3)),
    avocadoHectares: Number(avocadoHectares.toFixed(3)),
    preferredLanguage: lang,
    notes: `Registered via USSD (${languageLabel(lang)})`,
    userId: user.id,
    createdBy: 'ussd',
    createdAt: nowIso(),
    updatedAt: nowIso()
  };
  store.farmers.unshift(record);

  addActivity(store, {
    actor: phone || 'ussd',
    role: 'farmer',
    action: 'auth.register',
    entity: 'user',
    entityId: user.id,
    details: `USSD registration created farmer account for ${phone}`
  });

  addActivity(store, {
    actor: phone || 'ussd',
    role: 'farmer',
    action: 'farmer.ussd_register',
    entity: 'farmer',
    entityId: record.id,
    details: `USSD registration completed for ${record.name}`
  });

  const smsLog = {
    id: id('SMS'),
    farmerId: record.id,
    farmerName: record.name,
    phone: record.phone,
    message: ussdRegistrationSmsMessage(record, lang),
    provider: 'AfricaTalking(Mock)',
    ownerCostKes: money(SMS_OWNER_COST_PER_MESSAGE_KES),
    status: 'Sent',
    createdBy: 'ussd',
    createdAt: nowIso()
  };

  store.smsLogs.unshift(smsLog);
  addActivity(store, {
    actor: phone || 'ussd',
    role: 'farmer',
    action: 'sms.ussd_registration',
    entity: 'sms',
    entityId: smsLog.id,
    details: `Sent USSD registration confirmation SMS to ${record.phone}`
  });

  return record;
}

function ussdPaymentMessage(store, farmer, lang = DEFAULT_LANGUAGE) {
  const purchases = store.producePurchases.filter((row) => row.farmerId === farmer.id);
  if (!purchases.length) {
    return inLanguage(
      lang,
      'No produce purchases found yet. Contact your AGEM agent for assistance.',
      'Hakuna rekodi ya manunuzi ya mazao bado. Wasiliana na afisa wa AGEM kwa msaada.'
    );
  }

  const totalPurchasedKes = purchases.reduce((sum, row) => sum + valueFromPurchase(row), 0);
  const totalPaidKes = store.payments
    .filter((row) => row.farmerId === farmer.id && normalizePaymentStatus(row.status) === 'Received')
    .reduce((sum, row) => sum + Number(row.amount || 0), 0);
  const balanceKes = Math.max(0, totalPurchasedKes - totalPaidKes);

  const lastPayment = [...store.payments]
    .filter((row) => row.farmerId === farmer.id)
    .sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt))[0];

  const lines = [
    `${inLanguage(lang, 'Farmer', 'Mkulima')}: ${farmer.name}`,
    `${inLanguage(lang, 'Total Purchased', 'Jumla Ilinunuliwa')}: KES ${formatKes(totalPurchasedKes)}`,
    `${inLanguage(lang, 'Total Paid', 'Jumla Iliyolipwa')}: KES ${formatKes(totalPaidKes)}`,
    `${inLanguage(lang, 'Balance', 'Salio')}: KES ${formatKes(balanceKes)}`
  ];

  if (lastPayment) {
    lines.push(
      `${inLanguage(lang, 'Last Payment', 'Malipo ya Mwisho')}: ${clean(lastPayment.ref) || '-'} ` +
      `(${normalizePaymentStatus(lastPayment.status)}) ${dateShort(lastPayment.createdAt)}`
    );
  }
  return lines.join('\n');
}

function ussdQcMessage(store, farmer, lang = DEFAULT_LANGUAGE) {
  const latestQc = [...store.produce]
    .filter((row) => row.farmerId === farmer.id)
    .sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt))[0];

  if (!latestQc) {
    return inLanguage(
      lang,
      'No farm-gate QC entry found yet. Your next delivery will appear here after inspection.',
      'Hakuna rekodi ya QC ya shambani bado. Uwasilishaji wako unaofuata utaonekana hapa baada ya ukaguzi.'
    );
  }

  const dryMatter = clean(latestQc.dryMatterPct) ? `${latestQc.dryMatterPct}%` : 'N/A';
  const firmness = clean(latestQc.firmnessValue) ? `${latestQc.firmnessValue} ${latestQc.firmnessUnit || ''}`.trim() : 'N/A';
  return [
    `${inLanguage(lang, 'Farmer', 'Mkulima')}: ${farmer.name}`,
    `${inLanguage(lang, 'Variety', 'Aina')}: ${clean(latestQc.variety) || 'N/A'}`,
    `${inLanguage(lang, 'Visual', 'Mwonekano')}: ${clean(latestQc.visualGrade) || 'N/A'}`,
    `${inLanguage(lang, 'Dry Matter', 'Dry Matter')}: ${dryMatter}`,
    `${inLanguage(lang, 'Firmness', 'Ugumu')}: ${firmness}`,
    `${inLanguage(lang, 'Decision', 'Uamuzi')}: ${clean(latestQc.qcDecision) || 'N/A'}`,
    `${inLanguage(lang, 'Date', 'Tarehe')}: ${dateShort(latestQc.createdAt)}`
  ].join('\n');
}

function ussdProfileMessage(farmer, lang = DEFAULT_LANGUAGE) {
  const acres = Number(farmer.hectares || 0) / HECTARES_PER_ACRE;
  const avocadoAcres = Number(farmer.avocadoHectares || 0) / HECTARES_PER_ACRE;
  return [
    `${inLanguage(lang, 'Name', 'Jina')}: ${clean(farmer.name) || '-'}`,
    `${inLanguage(lang, 'Phone', 'Simu')}: ${clean(farmer.phone) || '-'}`,
    `${inLanguage(lang, 'National ID', 'Kitambulisho')}: ${clean(farmer.nationalId) || '-'}`,
    `${inLanguage(lang, 'Location', 'Eneo')}: ${clean(farmer.location) || '-'}`,
    `${inLanguage(lang, 'Farm Size', 'Ukubwa wa Shamba')}: ${Number(farmer.hectares || 0).toFixed(3)} ha (${acres.toFixed(2)} acres)`,
    `${inLanguage(lang, 'Avocado Area', 'Eneo la Parachichi')}: ${Number(farmer.avocadoHectares || 0).toFixed(3)} ha (${avocadoAcres.toFixed(2)} acres)`,
    `${inLanguage(lang, 'Trees', 'Miti')}: ${Number(farmer.trees || 0).toFixed(0)}`
  ].join('\n');
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
    } else if (normalizedPeriod === 'week' || normalizedPeriod === 'last7') {
      to = endOfDayUtc(now);
      from = startOfDayUtc(now - 6 * 24 * 60 * 60 * 1000);
    } else if (normalizedPeriod === 'month' || normalizedPeriod === 'thismonth') {
      const nowDate = new Date(now);
      from = Date.UTC(nowDate.getUTCFullYear(), nowDate.getUTCMonth(), 1, 0, 0, 0, 0);
      to = endOfDayUtc(now);
    } else if (normalizedPeriod === 'last30') {
      to = endOfDayUtc(now);
      from = startOfDayUtc(now - 29 * 24 * 60 * 60 * 1000);
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
  const smsCosts = smsCostSnapshot(store.smsLogs);
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
    smsOwnerCostPerMessageKes: smsCosts.smsOwnerCostPerMessageKes,
    smsBillable: smsCosts.smsBillable,
    smsSentLast24h: smsCosts.smsBillableLast24h,
    smsSpentKes: smsCosts.smsSpentKes,
    smsSpentLast24hKes: smsCosts.smsSpentLast24hKes,
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

function toAreaBundle(hectaresValue) {
  const hectares = Number(hectaresValue);
  if (!Number.isFinite(hectares) || hectares <= 0) {
    return {
      hectares: '',
      acres: '',
      squareFeet: ''
    };
  }
  const acres = hectares / HECTARES_PER_ACRE;
  const squareFeet = hectares * SQFT_PER_HECTARE;
  return {
    hectares: Number(hectares.toFixed(3)),
    acres: Number(acres.toFixed(3)),
    squareFeet: Number(squareFeet.toFixed(1))
  };
}

function paymentStatsInRange(payments, fromMs, toMs) {
  const rows = (Array.isArray(payments) ? payments : []).filter((row) => {
    const at = dateMs(row?.createdAt);
    return at >= fromMs && at <= toMs;
  });
  const total = rows.length;
  const received = rows.filter((row) => normalizePaymentStatus(row?.status) === 'Received').length;
  const failed = rows.filter((row) => normalizePaymentStatus(row?.status) === 'Failed').length;
  const pending = rows.filter((row) => normalizePaymentStatus(row?.status) === 'Pending').length;
  const rate = total ? Number(((received / total) * 100).toFixed(1)) : 0;
  return { total, received, failed, pending, rate };
}

function extractCopilotPeriod(question) {
  const textValue = clean(question).toLowerCase();
  if (/\btoday\b|\b24h\b|\b24 hours?\b/.test(textValue)) return 'today';
  if (/\bweek\b|\bweekly\b|\blast 7\b|\b7 days?\b/.test(textValue)) return 'week';
  if (/\bmonth\b|\bmonthly\b|\b30 days?\b/.test(textValue)) return 'month';
  if (/\ball\b|\boverall\b/.test(textValue)) return 'all';
  return 'week';
}

function detectLocationFromQuestion(question, store) {
  const textValue = clean(question).toLowerCase();
  const uniqueLocations = Array.from(
    new Set((store.farmers || []).map((row) => clean(row.location)).filter(Boolean))
  );
  for (const location of uniqueLocations) {
    if (textValue.includes(location.toLowerCase())) return location;
  }
  return '';
}

function localCopilotAnswer(store, question) {
  const prompt = clean(question);
  const lower = prompt.toLowerCase();
  const summary = makeSummary(store);

  if (/\bowed\b|\bbalance\b|\bunpaid\b/.test(lower)) {
    const period = extractCopilotPeriod(prompt);
    const location = detectLocationFromQuestion(prompt, store);
    const owed = buildOwedRows(store, { period });
    let rows = owed.rows;
    if (location) {
      rows = rows.filter((row) => clean(row.location).toLowerCase() === location.toLowerCase());
    }
    const top = rows.slice(0, 10);
    const totalKes = money(rows.reduce((sum, row) => sum + money(row.balanceKes), 0));
    const preview = top.map((row) => ({
      farmerName: row.farmerName,
      location: row.location || '-',
      phone: row.phone || '-',
      balanceKes: money(row.balanceKes)
    }));

    return {
      intent: 'owed_farmers',
      answer:
        top.length > 0
          ? `Found ${rows.length} owed farmers${location ? ` in ${location}` : ''} for period "${period}". Total owed is KES ${totalKes}.`
          : `No owed farmers found${location ? ` in ${location}` : ''} for period "${period}".`,
      insights: preview,
      actions: [
        {
          label: 'Payments > Farmers Owed',
          description: 'Select rows and use Prepare Selected or Pay Selected Now.'
        },
        {
          label: 'Export Owed CSV',
          description: 'Download owed list for finance reconciliation.'
        }
      ]
    };
  }

  if ((/\bpayment\b/.test(lower) && /\b(success|drop|decline|failed|pending)\b/.test(lower)) || /\bwhy\b.*\bpayment\b/.test(lower)) {
    const now = new Date();
    const startToday = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 0, 0, 0, 0);
    const endToday = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 23, 59, 59, 999);
    const startYesterday = startToday - 24 * 60 * 60 * 1000;
    const endYesterday = startToday - 1;

    const today = paymentStatsInRange(store.payments, startToday, endToday);
    const yesterday = paymentStatsInRange(store.payments, startYesterday, endYesterday);
    const delta = Number((today.rate - yesterday.rate).toFixed(1));

    const trend =
      delta < 0
        ? `Payment success dropped by ${Math.abs(delta)} percentage points vs yesterday.`
        : delta > 0
          ? `Payment success improved by ${delta} percentage points vs yesterday.`
          : 'Payment success is unchanged vs yesterday.';

    return {
      intent: 'payment_success',
      answer: `${trend} Today: ${today.received}/${today.total} received (${today.rate}%). Failed: ${today.failed}, Pending: ${today.pending}.`,
      insights: [
        { period: 'today', total: today.total, received: today.received, failed: today.failed, pending: today.pending, successRatePct: today.rate },
        { period: 'yesterday', total: yesterday.total, received: yesterday.received, failed: yesterday.failed, pending: yesterday.pending, successRatePct: yesterday.rate }
      ],
      actions: [
        {
          label: 'Review Pending/Failed Payments',
          description: 'Use Payments table status buttons to resolve failed and pending transactions.'
        },
        {
          label: 'Cross-check Owed Panel',
          description: 'Ensure purchases and settlements are reconciled before disbursement.'
        }
      ]
    };
  }

  if (/\bsms\b/.test(lower) && /\b(draft|message|campaign)\b/.test(lower)) {
    return {
      intent: 'sms_drafting',
      answer:
        'Use AI SMS Draft in the SMS panel to generate bilingual messages. You can then send to selected farmers, all farmers, or one number.',
      insights: [
        {
          smsSentLast24h: summary.smsSentLast24h || 0,
          smsSpentLast24hKes: summary.smsSpentLast24hKes || 0,
          smsSpentTotalKes: summary.smsSpentKes || 0
        }
      ],
      actions: [
        { label: 'Open SMS Panel', description: 'Set audience and click Draft with AI.' }
      ]
    };
  }

  return {
    intent: 'operations_summary',
    answer:
      `Current snapshot: ${summary.farmers || 0} farmers, ${summary.qcRecords || 0} QC records, ` +
      `${summary.purchasedRecords || 0} purchase logs, ${summary.owedFarmers || 0} owed farmers, ` +
      `${summary.paymentSuccessRate || 0}% payment success, and SMS spend KES ${summary.smsSpentLast24hKes || 0} in the last 24h.`,
    insights: [
      {
        farmers: summary.farmers || 0,
        qcRecords: summary.qcRecords || 0,
        purchasedRecords: summary.purchasedRecords || 0,
        owedFarmers: summary.owedFarmers || 0,
        totalOwedKes: summary.totalOwedKes || 0,
        paymentSuccessRatePct: summary.paymentSuccessRate || 0,
        smsSpentLast24hKes: summary.smsSpentLast24hKes || 0
      }
    ],
    actions: [
      { label: 'Open Reports', description: 'Review KPI and agent performance trends.' },
      { label: 'Open Payments', description: 'Resolve owed balances and pending payouts.' }
    ]
  };
}

function normalizeFirmnessToNewton(value, unit) {
  const raw = parseNumber(value);
  if (!Number.isFinite(raw) || raw <= 0) return null;
  const normalizedUnit = normalizeFirmnessUnit(unit) || 'N';
  if (normalizedUnit === 'kgf') {
    return Number((raw * 9.80665).toFixed(2));
  }
  return Number(raw.toFixed(2));
}

function sizeCodeNumber(code) {
  const normalized = normalizeSizeCode(code);
  if (!normalized) return null;
  const parsed = Number(normalized.slice(1));
  return Number.isFinite(parsed) ? parsed : null;
}

function localQcIntelligence(record) {
  const reasons = [];
  const actions = [];
  let riskScore = 0;
  let missingSignals = 0;

  const visual = normalizeVisualGrade(record?.visualGrade || record?.quality);
  if (visual === 'Reject') {
    riskScore += 4;
    reasons.push('Visual grade is Reject.');
  } else if (visual === 'Borderline') {
    riskScore += 2;
    reasons.push('Visual grade is Borderline.');
  } else if (!visual) {
    riskScore += 1;
    missingSignals += 1;
    reasons.push('Visual grade is missing.');
  }

  const dryMatter = parseNumber(record?.dryMatterPct);
  if (!Number.isFinite(dryMatter) || dryMatter <= 0) {
    riskScore += 1;
    missingSignals += 1;
    reasons.push('Dry matter is missing.');
  } else if (dryMatter < 21) {
    riskScore += 4;
    reasons.push(`Dry matter is low (${dryMatter}%).`);
  } else if (dryMatter < 23) {
    riskScore += 2;
    reasons.push(`Dry matter is borderline (${dryMatter}%).`);
  } else if (dryMatter > 35) {
    riskScore += 1;
    reasons.push(`Dry matter looks unusually high (${dryMatter}%).`);
  }

  const firmnessN = normalizeFirmnessToNewton(record?.firmnessValue, record?.firmnessUnit);
  if (!Number.isFinite(firmnessN) || firmnessN <= 0) {
    riskScore += 1;
    missingSignals += 1;
    reasons.push('Firmness reading is missing.');
  } else if (firmnessN < 45) {
    riskScore += 4;
    reasons.push(`Firmness is very low (${firmnessN} N).`);
  } else if (firmnessN < 60) {
    riskScore += 2;
    reasons.push(`Firmness is slightly low (${firmnessN} N).`);
  } else if (firmnessN > 130) {
    riskScore += 1;
    reasons.push(`Firmness is unusually high (${firmnessN} N).`);
  }

  const avgFruitWeightG = parseNumber(record?.avgFruitWeightG);
  if (!Number.isFinite(avgFruitWeightG) || avgFruitWeightG <= 0) {
    riskScore += 1;
    missingSignals += 1;
    reasons.push('Average fruit weight is missing.');
  } else if (avgFruitWeightG < 150) {
    riskScore += 2;
    reasons.push(`Average fruit weight is low (${avgFruitWeightG} g).`);
  } else if (avgFruitWeightG > 420) {
    riskScore += 1;
    reasons.push(`Average fruit weight is high (${avgFruitWeightG} g).`);
  }

  const sampleSize = parseNumber(record?.sampleSize);
  if (Number.isFinite(sampleSize) && sampleSize > 0 && sampleSize < 10) {
    riskScore += 1;
    reasons.push(`Sample size is small (${sampleSize}).`);
  } else if (!Number.isFinite(sampleSize) || sampleSize <= 0) {
    riskScore += 1;
    missingSignals += 1;
    reasons.push('Sample size is missing.');
  }

  const sizeNum = sizeCodeNumber(record?.sizeCode);
  if (sizeNum === null) {
    riskScore += 1;
    missingSignals += 1;
    reasons.push('Size code is missing.');
  } else if (sizeNum > 26) {
    riskScore += 1;
    reasons.push(`Small-size pack code (${record.sizeCode}).`);
  }

  const currentDecision = normalizeQcDecision(record?.qcDecision) || clean(record?.qcDecision) || '';
  if (currentDecision === 'Reject') riskScore += 1;
  if (currentDecision === 'Hold') riskScore += 1;

  let riskLevel = 'low';
  let recommendedDecision = 'Accept';
  if (riskScore >= 8) {
    riskLevel = 'high';
    recommendedDecision = 'Reject';
  } else if (riskScore >= 4) {
    riskLevel = 'medium';
    recommendedDecision = 'Hold';
  }

  if (visual === 'Reject' && recommendedDecision === 'Accept') {
    recommendedDecision = 'Hold';
  }
  if (currentDecision === 'Reject') {
    recommendedDecision = 'Reject';
  }

  if (riskLevel === 'high') {
    actions.push('Hold this lot and run immediate re-sampling before dispatch.');
    actions.push('Escalate to supervisor and confirm maturity window with farm records.');
  } else if (riskLevel === 'medium') {
    actions.push('Re-test dry matter and firmness on a larger sample before final release.');
    actions.push('Tag lot for closer monitoring at loading.');
  } else {
    actions.push('Proceed with normal handling and keep standard QC trace logs.');
  }

  const confidenceBase = 0.62 + Math.min(0.3, riskScore * 0.03) - Math.min(0.2, missingSignals * 0.04);
  const confidence = Number(Math.max(0.35, Math.min(0.96, confidenceBase)).toFixed(2));
  const summary = reasons.length
    ? reasons.slice(0, 3).join(' ')
    : 'Signals are within expected farm-gate ranges.';

  return {
    qcRecordId: record.id,
    farmerId: record.farmerId || '',
    farmerName: record.farmerName || '',
    variety: record.variety || '',
    lotWeightKgs: Number(parseNumber(record.lotWeightKgs ?? record.kgs)).toFixed(1),
    createdAt: record.createdAt || '',
    currentDecision: currentDecision || '-',
    recommendedDecision,
    riskLevel,
    riskScore,
    confidence,
    summary,
    reasons,
    actions
  };
}

function median(values) {
  const cleanValues = (Array.isArray(values) ? values : [])
    .filter((value) => Number.isFinite(value))
    .sort((a, b) => a - b);
  if (!cleanValues.length) return 0;
  const mid = Math.floor(cleanValues.length / 2);
  if (cleanValues.length % 2 === 0) {
    return (cleanValues[mid - 1] + cleanValues[mid]) / 2;
  }
  return cleanValues[mid];
}

function localPaymentRiskReport(store, options = {}) {
  const range = resolveRange(options.period, options.from, options.to);
  const payments = (store.payments || []).filter((row) => isInRange(row.createdAt, range));
  const farmersById = new Map((store.farmers || []).map((row) => [clean(row.id), row]));
  const flags = [];
  const addFlag = (flag) => {
    if (!flag || !flag.code) return;
    flags.push({
      severity: flag.severity || 'medium',
      code: flag.code,
      title: flag.title || '',
      detail: flag.detail || '',
      paymentId: clean(flag.paymentId),
      farmerName: clean(flag.farmerName),
      ref: clean(flag.ref),
      amountKes: Number.isFinite(Number(flag.amountKes)) ? money(flag.amountKes) : ''
    });
  };

  const byRef = new Map();
  for (const payment of payments) {
    const ref = clean(payment.ref).toUpperCase();
    if (!ref) continue;
    const rows = byRef.get(ref) || [];
    rows.push(payment);
    byRef.set(ref, rows);
  }
  for (const [ref, rows] of byRef.entries()) {
    if (rows.length < 2) continue;
    addFlag({
      severity: 'high',
      code: 'duplicate_ref',
      title: 'Duplicate payment reference',
      detail: `Reference ${ref} appears ${rows.length} times in the selected period.`,
      paymentId: rows[0]?.id,
      farmerName: rows[0]?.farmerName,
      ref
    });
  }

  const byFarmerAmount = new Map();
  for (const payment of payments) {
    const farmerId = clean(payment.farmerId);
    const amountKes = money(payment.amount);
    const key = `${farmerId}|${amountKes}`;
    const rows = byFarmerAmount.get(key) || [];
    rows.push(payment);
    byFarmerAmount.set(key, rows);
  }

  for (const rows of byFarmerAmount.values()) {
    if (rows.length < 2) continue;
    const ordered = rows.slice().sort((a, b) => dateMs(a.createdAt) - dateMs(b.createdAt));
    for (let idx = 1; idx < ordered.length; idx += 1) {
      const prev = ordered[idx - 1];
      const next = ordered[idx];
      const gapMinutes = Math.abs(dateMs(next.createdAt) - dateMs(prev.createdAt)) / (60 * 1000);
      if (gapMinutes <= 15) {
        addFlag({
          severity: 'high',
          code: 'rapid_repeat_payment',
          title: 'Rapid repeat payment',
          detail: `Same farmer and amount posted ${Math.round(gapMinutes)} minute(s) apart.`,
          paymentId: next.id,
          farmerName: next.farmerName,
          ref: next.ref,
          amountKes: next.amount
        });
      }
    }
  }

  const amounts = payments
    .map((row) => money(row.amount))
    .filter((value) => value > 0);
  const amountMedian = median(amounts);
  const outlierThreshold = Math.max(5000, amountMedian * 3);
  if (amountMedian > 0) {
    for (const payment of payments) {
      const amountKes = money(payment.amount);
      if (amountKes >= outlierThreshold) {
        addFlag({
          severity: 'medium',
          code: 'high_amount_outlier',
          title: 'High amount outlier',
          detail: `Amount KES ${formatKes(amountKes)} is above outlier threshold KES ${formatKes(outlierThreshold)}.`,
          paymentId: payment.id,
          farmerName: payment.farmerName,
          ref: payment.ref,
          amountKes
        });
      }
    }
  }

  for (const payment of payments) {
    const farmerId = clean(payment.farmerId);
    const farmerExists = farmerId && farmersById.has(farmerId);
    if (!farmerExists) {
      addFlag({
        severity: 'medium',
        code: 'unknown_farmer_link',
        title: 'Payment linked to unknown farmer',
        detail: 'Payment record is missing a valid farmer link.',
        paymentId: payment.id,
        farmerName: payment.farmerName,
        ref: payment.ref,
        amountKes: payment.amount
      });
    }
  }

  const pendingCount = payments.filter((row) => normalizePaymentStatus(row.status) === 'Pending').length;
  const failedCount = payments.filter((row) => normalizePaymentStatus(row.status) === 'Failed').length;
  if (payments.length >= 5 && pendingCount / payments.length >= 0.35) {
    addFlag({
      severity: 'medium',
      code: 'pending_ratio_high',
      title: 'Pending ratio is high',
      detail: `${pendingCount} of ${payments.length} payments are still pending.`
    });
  }
  if (failedCount >= 3) {
    addFlag({
      severity: 'high',
      code: 'failed_payments_spike',
      title: 'Failed payment spike',
      detail: `${failedCount} failed payments detected in selected period.`
    });
  }

  const severityCounts = flags.reduce(
    (acc, flag) => {
      const key = flag.severity || 'medium';
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    },
    { high: 0, medium: 0, low: 0 }
  );
  const overallRisk = severityCounts.high > 0 ? 'high' : severityCounts.medium > 0 ? 'medium' : 'low';

  const actions = [];
  if (severityCounts.high > 0) {
    actions.push('Pause bulk disbursement and review high-severity flagged rows first.');
  }
  if (flags.some((flag) => flag.code === 'duplicate_ref' || flag.code === 'rapid_repeat_payment')) {
    actions.push('Cross-check M-PESA references for duplicate posting before settlement.');
  }
  if (flags.some((flag) => flag.code === 'high_amount_outlier')) {
    actions.push('Require supervisor approval for outlier payment amounts.');
  }
  if (pendingCount > 0) {
    actions.push('Resolve pending payments and retry failed transactions.');
  }
  if (!actions.length) {
    actions.push('No major payment risks detected. Continue normal reconciliation.');
  }

  return {
    range: {
      period: range.period || 'all',
      from: range.from !== null ? new Date(range.from).toISOString() : '',
      to: range.to !== null ? new Date(range.to).toISOString() : ''
    },
    paymentCount: payments.length,
    summary: {
      overallRisk,
      highFlags: severityCounts.high || 0,
      mediumFlags: severityCounts.medium || 0,
      lowFlags: severityCounts.low || 0,
      pendingCount,
      failedCount,
      medianAmountKes: money(amountMedian || 0),
      outlierThresholdKes: money(outlierThreshold || 0)
    },
    narrative:
      flags.length > 0
        ? `Detected ${flags.length} payment risk flag(s): ${severityCounts.high} high, ${severityCounts.medium} medium, ${severityCounts.low} low.`
        : 'No payment risk flags detected in the selected period.',
    actions,
    flags: flags.slice(0, 200)
  };
}

function smartImportAutofix(mapped, invalid) {
  const totalHectares = resolveHectares(mapped);
  const avocadoHectares = resolveAvocadoHectares(mapped);
  const patch = {};
  let reason = '';
  let confidence = 0;

  if (invalid === 'area under avocado cannot be greater than total farm size') {
    if (Number.isFinite(totalHectares) && totalHectares > 0) {
      patch.avocadoHectares = Number(totalHectares.toFixed(3));
      reason = 'Set area under avocado equal to total farm size.';
      confidence = 0.92;
    }
  } else if (invalid === 'avocadoHectares/avocadoAcres/avocadoSquareFeet is required') {
    if (Number.isFinite(totalHectares) && totalHectares > 0) {
      patch.avocadoHectares = Number(totalHectares.toFixed(3));
      reason = 'Missing avocado area; copied total farm size as a placeholder for admin review.';
      confidence = 0.58;
    }
  } else if (invalid === 'hectares/acres/square feet is required') {
    if (Number.isFinite(avocadoHectares) && avocadoHectares > 0) {
      patch.hectares = Number(avocadoHectares.toFixed(3));
      reason = 'Missing total area; copied avocado area as a minimum placeholder for review.';
      confidence = 0.55;
    }
  } else if (invalid === 'trees must be a number') {
    patch.trees = 0;
    reason = 'Tree count was not numeric; set to 0.';
    confidence = 0.82;
  }

  const hasPatch = Object.keys(patch).length > 0;
  if (!hasPatch) {
    return { hasPatch: false, reason: '', confidence: 0, corrected: null };
  }

  const corrected = {
    ...mapped,
    ...patch
  };
  return { hasPatch: true, reason, confidence, corrected };
}

function summarizeSmartImportRow(raw, rowIndex, existingByPhone, existingByNationalId) {
  const mapped = mapFarmerImportRecord(raw);
  const invalid = validateFarmer(mapped);
  const issues = [];
  let duplicate = null;

  const phoneKey = normalizePhone(mapped.phone) || clean(mapped.phone);
  const nationalIdKey = normalizedNationalId(mapped.nationalId);
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
      message: 'This row matches an existing farmer by phone or National ID.',
      suggestion: 'Use overwrite mode if new file values should replace old farmer values.'
    });
  }

  let proposedFix = null;
  let confidence = 0.9;

  if (invalid) {
    const issue = buildFarmerImportIssue(mapped, invalid);
    issues.unshift({
      code: issue.code || 'invalid_import_row',
      severity: 'error',
      message: issue.error || invalid,
      suggestion: issue.suggestion || 'Check row values before import.',
      observed: issue.observed || null
    });
    const autoFix = smartImportAutofix(mapped, invalid);
    if (autoFix.hasPatch) {
      proposedFix = {
        ...autoFix.corrected,
        totalArea: toAreaBundle(resolveHectares(autoFix.corrected)),
        avocadoArea: toAreaBundle(resolveAvocadoHectares(autoFix.corrected))
      };
      issues.push({
        code: 'auto_fix_proposed',
        severity: 'info',
        message: autoFix.reason,
        suggestion: 'Review suggested values and confirm overwrite/adjustment before import.'
      });
      confidence = autoFix.confidence;
    } else {
      confidence = 0.4;
    }
  } else {
    const hectares = resolveHectares(mapped);
    const avocadoHectares = resolveAvocadoHectares(mapped);
    const treeCount = Number(parseTrees(mapped.trees));
    const density = Number.isFinite(treeCount) && Number.isFinite(avocadoHectares) && avocadoHectares > 0
      ? treeCount / avocadoHectares
      : 0;

    if (Number.isFinite(hectares) && hectares > 50) {
      issues.push({
        code: 'large_farm_size',
        severity: 'warning',
        message: `Large farm area detected (${hectares.toFixed(2)} ha).`,
        suggestion: 'Confirm unit conversion (acres vs hectares) in source file.'
      });
      confidence = Math.min(confidence, 0.72);
    }

    if (density > 2000) {
      issues.push({
        code: 'tree_density_outlier',
        severity: 'warning',
        message: `Tree density looks high (${density.toFixed(0)} trees/ha under avocado).`,
        suggestion: 'Confirm tree count or avocado area values for this row.'
      });
      confidence = Math.min(confidence, 0.7);
    }
  }

  const status = issues.some((item) => item.severity === 'error')
    ? 'blocked'
    : issues.length
      ? 'review'
      : 'ready';
  return {
    row: rowIndex,
    status,
    confidence: Number(confidence.toFixed(2)),
    mapped: {
      ...mapped,
      totalArea: toAreaBundle(resolveHectares(mapped)),
      avocadoArea: toAreaBundle(resolveAvocadoHectares(mapped))
    },
    issues,
    proposedFix,
    duplicate
  };
}

function buildSmsDraftsLocal(payload, store) {
  const purpose = clean(payload?.purpose) || 'general update';
  const audience = clean(payload?.audience) || 'selected farmers';
  const tone = clean(payload?.tone).toLowerCase() || 'professional';
  const language = clean(payload?.language).toLowerCase() || 'bilingual';
  const maxLength = Number(payload?.maxLength);
  const summary = makeSummary(store);

  const tighten = (text) => {
    const normalized = clean(text).replace(/\s+/g, ' ');
    if (!Number.isFinite(maxLength) || maxLength < 30) return normalized;
    if (normalized.length <= maxLength) return normalized;
    return normalized.slice(0, Math.max(0, maxLength - 1)).trimEnd() + '…';
  };

  const englishCore = `Agem Portal: ${purpose}. Audience: ${audience}. Reply for support if needed.`;
  const swCore = `Agem Portal: ${purpose}. Walengwa: ${audience}. Jibu kwa msaada ukihitaji.`;
  const englishFormal = `Agem Portal update: ${purpose}. This message is for ${audience}. Contact your Agem agent for help.`;
  const swFormal = `Taarifa ya Agem Portal: ${purpose}. Ujumbe huu ni kwa ${audience}. Wasiliana na wakala wa Agem kwa msaada.`;

  const english = tone === 'friendly' ? englishCore : englishFormal;
  const swahili = tone === 'friendly' ? swCore : swFormal;

  const drafts = [];
  if (language === 'en' || language === 'english') {
    drafts.push({ language: 'English', message: tighten(english) });
  } else if (language === 'sw' || language === 'kiswahili') {
    drafts.push({ language: 'Kiswahili', message: tighten(swahili) });
  } else {
    drafts.push({ language: 'English', message: tighten(english) });
    drafts.push({ language: 'Kiswahili', message: tighten(swahili) });
  }

  return {
    drafts,
    metadata: {
      model: 'local-rules',
      smsSentLast24h: summary.smsSentLast24h || 0,
      smsSpendLast24hKes: summary.smsSpentLast24hKes || 0
    }
  };
}

function extractResponseText(payload) {
  if (typeof payload?.output_text === 'string' && clean(payload.output_text)) {
    return payload.output_text;
  }
  if (!Array.isArray(payload?.output)) return '';

  const chunks = [];
  for (const item of payload.output) {
    if (!Array.isArray(item?.content)) continue;
    for (const content of item.content) {
      if (typeof content?.text === 'string' && clean(content.text)) {
        chunks.push(content.text);
      }
    }
  }
  return chunks.join('\n').trim();
}

async function callOpenAiJson(systemPrompt, userPrompt, schema) {
  if (!OPENAI_API_KEY || typeof fetch !== 'function') {
    return { ok: false, error: 'OPENAI_API_KEY not configured.' };
  }

  const controller = new AbortController();
  const timeoutHandle = setTimeout(() => controller.abort(), OPENAI_TIMEOUT_MS);

  try {
    const response = await fetch('https://api.openai.com/v1/responses', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        input: [
          { role: 'system', content: [{ type: 'input_text', text: systemPrompt }] },
          { role: 'user', content: [{ type: 'input_text', text: userPrompt }] }
        ],
        text: {
          format: {
            type: 'json_schema',
            name: 'agem_phase1_response',
            strict: true,
            schema
          }
        }
      }),
      signal: controller.signal
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = clean(payload?.error?.message) || `OpenAI request failed (${response.status})`;
      return { ok: false, error: message };
    }

    const rawText = extractResponseText(payload);
    if (!clean(rawText)) {
      return { ok: false, error: 'OpenAI response did not include text output.' };
    }

    try {
      const parsed = JSON.parse(rawText);
      return { ok: true, data: parsed };
    } catch (error) {
      return { ok: false, error: `OpenAI output parsing failed: ${error.message}` };
    }
  } catch (error) {
    if (error?.name === 'AbortError') {
      return { ok: false, error: 'OpenAI request timed out.' };
    }
    return { ok: false, error: error.message || 'OpenAI request failed.' };
  } finally {
    clearTimeout(timeoutHandle);
  }
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
      now: nowIso(),
      storage: {
        backend: STORAGE_BACKEND,
        backupIntervalHours: BACKUP_INTERVAL_HOURS,
        backupRetentionDays: BACKUP_RETENTION_DAYS
      },
      ai: {
        model: OPENAI_MODEL,
        openAiConfigured: Boolean(OPENAI_API_KEY),
        execBriefEnabled: AI_EXEC_BRIEF_ENABLED,
        execBriefIntervalHours: AI_EXEC_BRIEF_INTERVAL_HOURS
      }
    });
    return;
  }

  if (pathname === '/api/ussd/callback' && (req.method === 'POST' || req.method === 'GET')) {
    if (!USSD_ENABLED) {
      ussdReply(res, 'USSD service is not active yet. Please try again later.', true);
      return;
    }

    try {
      let payload = {};
      if (req.method === 'GET') {
        payload = Object.fromEntries(reqUrl.searchParams.entries());
      } else {
        const contentType = clean(req.headers['content-type']).toLowerCase();
        if (contentType.includes('application/json')) {
          payload = await readBody(req, 120_000);
        } else {
          const rawBody = await readTextBody(req, 120_000);
          if (!clean(rawBody)) {
            payload = {};
          } else if (contentType.includes('application/x-www-form-urlencoded')) {
            payload = Object.fromEntries(new URLSearchParams(rawBody).entries());
          } else {
            try {
              payload = JSON.parse(rawBody);
            } catch {
              payload = Object.fromEntries(new URLSearchParams(rawBody).entries());
            }
          }
        }
      }

      if (USSD_SHARED_SECRET) {
        const providedSecret = clean(
          req.headers['x-ussd-secret'] ||
          payload.secret ||
          reqUrl.searchParams.get('secret')
        );
        if (!secureEquals(providedSecret, USSD_SHARED_SECRET)) {
          text(res, 403, 'END Unauthorized USSD request');
          return;
        }
      }

      const rawPhoneNumber = clean(payload.phoneNumber || payload.phonenumber || payload.msisdn);
      const phoneNumber = normalizePhone(rawPhoneNumber) || rawPhoneNumber;
      const inputText = clean(payload.text || payload.input || payload.ussdText);
      if (!phoneNumber) {
        ussdReply(res, 'Phone number is missing from the request.', true);
        return;
      }

      const store = await readStore();
      const parts = inputText
        .split('*')
        .map((part) => clean(part))
        .filter(Boolean);
      if (!parts.length) {
        ussdReply(res, ussdLanguageMenu(), false);
        return;
      }

      const lang = normalizeLanguage(parts[0]);
      if (!lang) {
        ussdReply(
          res,
          `Invalid option / Chaguo si sahihi.\n${ussdLanguageMenu()}`,
          false
        );
        return;
      }

      const farmer = findFarmerByPhoneNormalized(store.farmers, phoneNumber);
      const menuParts = parts.slice(1);

      if (!farmer) {
        if (!menuParts.length) {
          ussdReply(res, ussdUnregisteredMenu(lang), false);
          return;
        }

        const option = menuParts[0] || '';
        if (option === '2') {
          ussdReply(res, ussdHelpMessage(lang), true);
          return;
        }
        if (option !== '1') {
          ussdReply(
            res,
            `${inLanguage(lang, 'Invalid option.', 'Chaguo si sahihi.')}\n${ussdUnregisteredMenu(lang)}`,
            false
          );
          return;
        }

        const registrationValues = menuParts.slice(1);
        if (!registrationValues.length) {
          ussdReply(res, ussdRegistrationPrompt(lang, 1), false);
          return;
        }

        const name = clean(registrationValues[0]);
        if (name.length < 3) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'Name must be at least 3 characters. Please dial again to restart registration.',
              'Jina lazima liwe na angalau herufi 3. Tafadhali piga tena kuanza usajili upya.'
            ),
            true
          );
          return;
        }
        if (registrationValues.length === 1) {
          ussdReply(res, ussdRegistrationPrompt(lang, 2), false);
          return;
        }

        const nationalId = cleanNationalId(registrationValues[1]);
        if (nationalId.length < 4) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'National ID looks invalid. Please dial again and restart registration.',
              'Nambari ya kitambulisho si sahihi. Tafadhali piga tena na uanze usajili upya.'
            ),
            true
          );
          return;
        }
        if (registrationValues.length === 2) {
          ussdReply(res, ussdRegistrationPrompt(lang, 3), false);
          return;
        }

        const location = clean(registrationValues[2]);
        if (location.length < 2) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'Location is required. Please dial again and restart registration.',
              'Eneo linahitajika. Tafadhali piga tena na uanze usajili upya.'
            ),
            true
          );
          return;
        }
        if (registrationValues.length === 3) {
          ussdReply(res, ussdRegistrationPrompt(lang, 4), false);
          return;
        }

        const totalAcres = parseArea(registrationValues[3]);
        if (!Number.isFinite(totalAcres) || totalAcres <= 0) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'Farm size must be a number greater than 0. Please dial again and restart registration.',
              'Ukubwa wa shamba lazima uwe namba kubwa kuliko 0. Tafadhali piga tena kuanza usajili upya.'
            ),
            true
          );
          return;
        }
        if (registrationValues.length === 4) {
          ussdReply(res, ussdRegistrationPrompt(lang, 5), false);
          return;
        }

        const avocadoAcres = parseArea(registrationValues[4]);
        if (!Number.isFinite(avocadoAcres) || avocadoAcres <= 0) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'Avocado area must be a number greater than 0. Please dial again and restart registration.',
              'Eneo la parachichi lazima liwe namba kubwa kuliko 0. Tafadhali piga tena kuanza usajili upya.'
            ),
            true
          );
          return;
        }
        if (avocadoAcres > totalAcres) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'Avocado area cannot be greater than total farm size. Please dial again and restart registration.',
              'Eneo la parachichi haliwezi kuzidi ukubwa wa shamba lote. Tafadhali piga tena kuanza usajili upya.'
            ),
            true
          );
          return;
        }
        if (registrationValues.length === 5) {
          ussdReply(res, ussdRegistrationPrompt(lang, 6), false);
          return;
        }

        const pin = String(registrationValues[5] || '').trim();
        const invalidPin = validateFarmerPin(pin);
        if (invalidPin) {
          ussdReply(
            res,
            inLanguage(
              lang,
              `${invalidPin} Please dial again and restart registration.`,
              `PIN lazima iwe namba 4. Tafadhali piga tena kuanza usajili upya.`
            ),
            true
          );
          return;
        }
        if (registrationValues.length === 6) {
          ussdReply(res, ussdRegistrationPrompt(lang, 7), false);
          return;
        }

        const confirmPin = String(registrationValues[6] || '').trim();
        if (!secureEquals(pin, confirmPin)) {
          ussdReply(
            res,
            inLanguage(
              lang,
              'PIN confirmation does not match. Please dial again and restart registration.',
              'Uthibitisho wa PIN haulingani. Tafadhali piga tena kuanza usajili upya.'
            ),
            true
          );
          return;
        }

        if (findFarmerByPhoneNormalized(store.farmers, phoneNumber)) {
          ussdReply(
            res,
            inLanguage(
              lang,
              `This phone number is already registered.\nSupport: ${USSD_HELP_PHONE}`,
              `Nambari hii tayari imesajiliwa.\nMsaada: ${USSD_HELP_PHONE}`
            ),
            true
          );
          return;
        }
        if (findFarmerByNationalId(store.farmers, nationalId)) {
          ussdReply(
            res,
            inLanguage(
              lang,
              `This National ID is already registered.\nSupport: ${USSD_HELP_PHONE}`,
              `Kitambulisho hiki tayari kimesajiliwa.\nMsaada: ${USSD_HELP_PHONE}`
            ),
            true
          );
          return;
        }
        if (findUserByPhone(store, phoneNumber)) {
          ussdReply(
            res,
            inLanguage(
              lang,
              `This phone number already has an account.\nSupport: ${USSD_HELP_PHONE}`,
              `Nambari hii tayari ina akaunti.\nMsaada: ${USSD_HELP_PHONE}`
            ),
            true
          );
          return;
        }

        const registeredFarmer = registerFarmerFromUssd(store, {
          phoneNumber,
          name,
          nationalId,
          location,
          totalAcres,
          avocadoAcres,
          pin,
          language: lang
        });
        await writeStore(store);
        ussdReply(res, ussdRegistrationCompleteMessage(registeredFarmer, lang), true);
        return;
      }

      if (languageOrDefault(farmer.preferredLanguage) !== lang) {
        farmer.preferredLanguage = lang;
        farmer.updatedAt = nowIso();
        await writeStore(store);
      }

      if (!menuParts.length) {
        ussdReply(res, ussdMainMenu(lang), false);
        return;
      }

      const option = menuParts[0] || '';
      if (option === '1') {
        ussdReply(res, ussdPaymentMessage(store, farmer, lang), true);
        return;
      }
      if (option === '2') {
        ussdReply(res, ussdQcMessage(store, farmer, lang), true);
        return;
      }
      if (option === '3') {
        ussdReply(res, ussdProfileMessage(farmer, lang), true);
        return;
      }
      if (option === '4') {
        ussdReply(res, ussdHelpMessage(lang), true);
        return;
      }

      ussdReply(res, `${inLanguage(lang, 'Invalid option.', 'Chaguo si sahihi.')}\n${ussdMainMenu(lang)}`, false);
    } catch (error) {
      ussdReply(res, `USSD processing failed: ${error.message}`, true);
    }
    return;
  }

  if (pathname === '/api/ussd/events' && (req.method === 'POST' || req.method === 'GET')) {
    try {
      let payload = {};
      if (req.method === 'GET') {
        payload = Object.fromEntries(reqUrl.searchParams.entries());
      } else {
        const contentType = clean(req.headers['content-type']).toLowerCase();
        if (contentType.includes('application/json')) {
          payload = await readBody(req, 120_000);
        } else {
          const rawBody = await readTextBody(req, 120_000);
          payload = Object.fromEntries(new URLSearchParams(rawBody).entries());
        }
      }

      if (USSD_SHARED_SECRET) {
        const providedSecret = clean(
          req.headers['x-ussd-secret'] ||
          payload.secret ||
          reqUrl.searchParams.get('secret')
        );
        if (!secureEquals(providedSecret, USSD_SHARED_SECRET)) {
          json(res, 403, { error: 'Unauthorized USSD events request' });
          return;
        }
      }

      const store = await readStore();
      addActivity(store, {
        actor: 'ussd-gateway',
        role: 'system',
        action: 'ussd.event',
        entity: 'ussd',
        details: `USSD event received: ${clean(payload.status || payload.event || payload.sessionStatus || 'unknown')}`
      });
      await writeStore(store);

      json(res, 200, { data: { ok: true } });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/auth/login' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      const identityRaw = clean(payload.username || payload.phone);
      const secret = String(payload.password || payload.pin || '');
      const username = normalizedUsername(identityRaw);
      const phone = normalizePhone(identityRaw);
      const store = await readStore();

      if (activeUserCount(store) === 0) {
        json(res, 503, {
          error:
            'No active staff accounts configured. Set ADMIN_USERNAME and ADMIN_PASSWORD in environment variables, then redeploy.'
        });
        return;
      }
      if (!identityRaw || !secret) {
        json(res, 422, { error: 'Username/phone and password/PIN are required.' });
        return;
      }

      const user = findUserByUsername(store, username) || findUserByPhone(store, phone);
      if (!user || user.status !== 'active' || !verifyPassword(secret, user.password)) {
        json(res, 401, { error: 'Invalid username/phone or password/PIN' });
        return;
      }

      const token = randomToken();
      const createdAt = nowIso();
      const expiresAt = new Date(Date.now() + SESSION_TTL_MS).toISOString();

      store.sessions[token] = {
        userId: user.id,
        username: user.username,
        phone: user.phone || '',
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

      await writeStore(store);
      json(res, 200, {
        data: {
          token,
          expiresAt,
          user: {
            id: user.id,
            username: user.username,
            phone: user.phone || '',
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
      const phone = normalizePhone(payload.phone);
      const nationalId = cleanNationalId(payload.nationalId);
      const location = clean(payload.location);
      const preferredLanguageInput = parsePreferredLanguage(payload.preferredLanguage);
      const preferredLanguage = preferredLanguageInput.value;
      const pin = String(payload.pin || '');
      const confirmPin = payload.confirmPin === undefined ? pin : String(payload.confirmPin || '');
      const notes = clean(payload.notes);
      const treesValue = payload.trees === undefined ? 0 : parseTrees(payload.trees);
      const trees = Number.isFinite(treesValue) ? Number(treesValue.toFixed(2)) : NaN;
      const hectares = resolveHectares(payload);
      const avocadoHectares = resolveAvocadoHectares(payload);

      if (!name || !phone || !nationalId || !location || !pin || !confirmPin) {
        json(res, 422, {
          error: 'name, phone, nationalId, location, total area, area under avocado, pin, and confirmPin are required.'
        });
        return;
      }
      if (!preferredLanguageInput.valid) {
        json(res, 422, { error: 'preferredLanguage must be en or sw.' });
        return;
      }
      if (!secureEquals(pin, confirmPin)) {
        json(res, 422, { error: 'PIN confirmation does not match.' });
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

      const invalidPin = validateFarmerPin(pin);
      if (invalidPin) {
        json(res, 422, { error: invalidPin });
        return;
      }

      const store = await readStore();
      if (findFarmerByPhone(store.farmers, phone)) {
        json(res, 409, { error: 'A farmer with this phone number already exists.' });
        return;
      }
      if (findUserByPhone(store, phone)) {
        json(res, 409, { error: 'An account with this phone number already exists.' });
        return;
      }
      if (findFarmerByNationalId(store.farmers, nationalId)) {
        json(res, 409, { error: 'A farmer with this National ID already exists.' });
        return;
      }

      const username = phone;
      const user = {
        id: id('USR'),
        username,
        phone,
        name,
        role: 'farmer',
        status: 'active',
        password: hashPassword(pin),
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
        preferredLanguage,
        notes,
        userId: user.id,
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
        phone: user.phone || '',
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

      await writeStore(store);
      json(res, 201, {
        data: {
          token,
          expiresAt,
          user: {
            id: user.id,
            username: user.username,
            phone: user.phone || '',
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
      const store = await readStore();
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

      const user = store.users.find((row) => row.id === auth.session.userId && row.status === 'active');
      if (!user) {
        json(res, 401, { error: 'Session is no longer valid. Please sign in again.' });
        return;
      }
      const updatingPin = user.role === 'farmer';

      if (!currentPassword || !newPassword) {
        json(res, 422, {
          error: updatingPin
            ? 'Current and new PIN are required.'
            : 'Current and new password are required.'
        });
        return;
      }
      if (!secureEquals(newPassword, confirmPassword)) {
        json(res, 422, {
          error: updatingPin
            ? 'New PIN and confirmation do not match.'
            : 'New password and confirmation do not match.'
        });
        return;
      }
      if (secureEquals(currentPassword, newPassword)) {
        json(res, 422, {
          error: updatingPin
            ? 'New PIN must be different from current PIN.'
            : 'New password must be different from current password.'
        });
        return;
      }

      const invalidCredential = updatingPin
        ? validateFarmerPin(newPassword)
        : validateNewPassword(newPassword);
      if (invalidCredential) {
        json(res, 422, { error: invalidCredential });
        return;
      }

      if (!verifyPassword(currentPassword, user.password)) {
        json(res, 401, {
          error: updatingPin ? 'Current PIN is incorrect.' : 'Current password is incorrect.'
        });
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
        details: updatingPin
          ? 'PIN changed by signed-in farmer.'
          : 'Password changed by signed-in user.'
      });

      await writeStore(store);
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
      const recoveryCode = String(payload.recoveryCode || '').trim();

      if (!role || !['admin', 'agent'].includes(role)) {
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

      const store = await readStore();
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
      await writeStore(store);

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
      const recoveryCode = String(payload.recoveryCode || '').trim();
      const newPassword = String(payload.newPassword || '');
      const confirmPassword = payload.confirmPassword === undefined
        ? newPassword
        : String(payload.confirmPassword || '');

      if (!role || !['admin', 'agent'].includes(role)) {
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

      const store = await readStore();
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
      await writeStore(store);

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
    const store = await readStore();
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
          phone: auth.session.phone || '',
          role: auth.session.role,
          name: auth.session.name
        },
        expiresAt: auth.session.expiresAt
      }
    });
    return;
  }

  if (pathname === '/api/auth/logout' && req.method === 'POST') {
    const store = await readStore();
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
      await writeStore(store);
    }

    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/agents' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const includeDisabled = reqUrl.searchParams.get('includeDisabled') === 'true';
    const query = clean(reqUrl.searchParams.get('q')).toLowerCase();
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 200, 2000);
    const offset = safeQueryOffset(reqUrl.searchParams, 'offset', 0, 2000000);

    let rows = store.users.filter((row) => {
      if (!row || row.role !== 'agent') return false;
      if (!includeDisabled && row.status !== 'active') return false;
      return true;
    });

    if (query) {
      rows = rows.filter((row) =>
        [row.username, row.name, row.email, row.phone, row.status, row.provisioning].some((field) =>
          clean(field).toLowerCase().includes(query)
        )
      );
    }

    rows = rows.sort((a, b) => {
      const left = clean(a.updatedAt || a.createdAt);
      const right = clean(b.updatedAt || b.createdAt);
      if (left === right) {
        return clean(a.username).localeCompare(clean(b.username));
      }
      return left > right ? -1 : 1;
    });

    const total = rows.length;
    const paged = rows.slice(offset, offset + limit).map(agentView);

    json(res, 200, {
      data: paged,
      meta: {
        total,
        limit,
        offset,
        includeDisabled
      }
    });
    return;
  }

  if (pathname === '/api/agents' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const name = clean(payload.name);
      const email = normalizedEmail(payload.email);
      const phone = normalizePhone(payload.phone) || clean(payload.phone);

      if (!name || !email || !phone) {
        json(res, 422, { error: 'name, email, and phone are required.' });
        return;
      }
      if (!isValidEmail(email)) {
        json(res, 422, { error: 'Enter a valid email address.' });
        return;
      }
      if (!isValidKenyaPhone(phone)) {
        json(res, 422, { error: 'Enter a valid Kenyan phone number (e.g. 2547XXXXXXXX).' });
        return;
      }
      if (findUserByEmail(store, email)) {
        json(res, 409, { error: 'This email is already linked to an existing user.' });
        return;
      }
      if (findUserByPhone(store, phone)) {
        json(res, 409, { error: 'This phone number is already linked to an existing user.' });
        return;
      }
      const username = generateAgentUsername(store, name);
      const temporaryPassword = generateTemporaryPassword();

      const agent = {
        id: id('USR'),
        username,
        name,
        email,
        phone,
        role: 'agent',
        status: 'active',
        password: hashPassword(temporaryPassword),
        provisioning: 'admin',
        createdAt: nowIso(),
        updatedAt: nowIso()
      };
      store.users.push(agent);

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'agent.create',
        entity: 'user',
        entityId: agent.id,
        details: `Created agent account ${agent.username} (${agent.email}, ${agent.phone})`
      });

      await writeStore(store);
      json(res, 201, {
        data: {
          ...agentView(agent),
          temporaryPassword
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/farmers' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.farmers];
    rows = filterByQuery(rows, reqUrl.searchParams, ['name', 'phone', 'nationalId', 'location', 'preferredLanguage', 'notes']);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 200, 2000);
    rows = rows.slice(0, limit).map((farmer) => {
      const user = resolveFarmerPortalUser(store, farmer);
      return {
        ...farmer,
        hasPortalAccess: Boolean(user && user.status === 'active'),
        portalUsername: user ? user.username : ''
      };
    });
    json(res, 200, { data: rows });
    return;
  }

  if (pathname === '/api/farmers' && req.method === 'POST') {
    try {
      const store = await readStore();
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
      const preferredLanguage = parsePreferredLanguage(payload.preferredLanguage).value;
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
        phone: normalizePhone(payload.phone) || clean(payload.phone),
        nationalId: cleanNationalId(payload.nationalId),
        location: clean(payload.location),
        trees: Number(parseTrees(payload.trees).toFixed(2)),
        hectares: Number(hectares.toFixed(3)),
        avocadoHectares: Number(avocadoHectares.toFixed(3)),
        preferredLanguage,
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
      await writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/farmers/import' && req.method === 'POST') {
    try {
      const store = await readStore();
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
      const sendOnboardingSms = isTruthy(payload.sendOnboardingSms);
      const onboardingSmsTemplate = clean(payload.onboardingSmsTemplate);
      if (sendOnboardingSms && onboardingSmsTemplate.length > IMPORT_ONBOARDING_SMS_MAX_LENGTH) {
        json(res, 422, {
          error: `onboardingSmsTemplate must be ${IMPORT_ONBOARDING_SMS_MAX_LENGTH} characters or fewer`
        });
        return;
      }
      const farmersByPhone = new Map(
        store.farmers
          .map((row) => [normalizePhone(row.phone) || clean(row.phone), row])
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
          const issue = buildFarmerImportIssue(mapped, invalid);
          errors.push({ row: idx + 1, ...issue });
          return;
        }

        const phone = normalizePhone(mapped.phone) || clean(mapped.phone);
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
                error: 'Conflicting duplicate match: phone and National ID belong to different farmers',
                code: 'duplicate_conflict',
                suggestion: 'Review this row manually: phone and National ID currently belong to different existing farmers.'
              });
              return;
            }

            const nextName = clean(mapped.name);
            const syncResult = syncFarmerPortalUserIdentity(store, existing, phone, nextName);
            if (!syncResult.ok) {
              errors.push({
                row: idx + 1,
                error: syncResult.error,
                code: 'identity_sync_conflict',
                suggestion: 'Check phone and name values, then retry overwrite for this specific row.'
              });
              return;
            }

            const oldPhoneKey = normalizePhone(existing.phone) || clean(existing.phone);
            const oldNationalIdKey = normalizedNationalId(existing.nationalId);
            if (oldPhoneKey && farmersByPhone.get(oldPhoneKey)?.id === existing.id) {
              farmersByPhone.delete(oldPhoneKey);
            }
            if (oldNationalIdKey && farmersByNationalId.get(oldNationalIdKey)?.id === existing.id) {
              farmersByNationalId.delete(oldNationalIdKey);
            }

            existing.name = nextName;
            existing.phone = phone;
            existing.nationalId = nationalId;
            existing.location = clean(mapped.location);
            existing.trees = Number(parseTrees(mapped.trees).toFixed(2));
            existing.hectares = Number(resolveHectares(mapped).toFixed(3));
            existing.avocadoHectares = Number(resolveAvocadoHectares(mapped).toFixed(3));
            existing.preferredLanguage = parsePreferredLanguage(mapped.preferredLanguage).value;
            existing.notes = clean(mapped.notes);
            existing.updatedAt = nowIso();

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
          id: id('F'),
          name: clean(mapped.name),
          phone,
          nationalId,
          location: clean(mapped.location),
          trees: Number(parseTrees(mapped.trees).toFixed(2)),
          hectares: Number(resolveHectares(mapped).toFixed(3)),
          avocadoHectares: Number(resolveAvocadoHectares(mapped).toFixed(3)),
          preferredLanguage: parsePreferredLanguage(mapped.preferredLanguage).value,
          notes: clean(mapped.notes),
          createdBy: auth.session.username,
          createdAt: nowIso(),
          updatedAt: nowIso()
        };

        farmersByPhone.set(phone, record);
        farmersByNationalId.set(nationalIdKey, record);
        created.push(record);
      });

      const importSmsLogs = [];
      if (sendOnboardingSms && (created.length || updated.length)) {
        const byPhone = new Map();
        for (const farmer of created.concat(updated)) {
          const phone = clean(farmer.phone);
          if (!phone || byPhone.has(phone)) continue;
          byPhone.set(phone, farmer);
        }

        const createdAt = nowIso();
        for (const [phone, farmer] of byPhone.entries()) {
          const message = renderImportOnboardingSms(onboardingSmsTemplate, farmer);
          if (!message) continue;
          importSmsLogs.push({
            id: id('SMS'),
            farmerId: farmer.id || '',
            farmerName: farmer.name || '',
            phone,
            message,
            provider: 'AfricaTalking(Mock)',
            ownerCostKes: money(SMS_OWNER_COST_PER_MESSAGE_KES),
            status: 'Sent',
            createdBy: auth.session.username,
            createdAt
          });
        }
      }

      if (created.length || updated.length || importSmsLogs.length) {
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
        if (importSmsLogs.length) {
          store.smsLogs = importSmsLogs.reverse().concat(store.smsLogs);
          addActivity(store, {
            actor: auth.session.username,
            role: auth.session.role,
            action: 'sms.import_onboarding',
            entity: 'sms',
            details: `Sent onboarding SMS to ${importSmsLogs.length} imported farmer mobile numbers`
          });
        }
        await writeStore(store);
      }

      json(res, 200, {
        data: {
          totalRows: records.length,
          imported: created.length,
          updated: updated.length,
          skipped: records.length - created.length - updated.length,
          duplicateMode: onDuplicate,
          sendOnboardingSms,
          smsSent: importSmsLogs.length,
          anomalyCount: errors.filter((entry) => clean(entry.suggestion)).length,
          errors: errors.slice(0, 250)
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  const farmerPinMatch = pathname.match(/^\/api\/farmers\/([^/]+)\/reset-pin$/);
  if (farmerPinMatch && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const farmerId = decodeURIComponent(farmerPinMatch[1]);
      const farmer = findById(store.farmers, farmerId);
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found' });
        return;
      }

      const payload = await readBody(req);
      const generate = isTruthy(payload.generate);
      const pin = String(payload.pin || '').trim();
      const confirmPin = payload.confirmPin === undefined ? pin : String(payload.confirmPin || '').trim();

      let nextPin = pin;
      if (generate || !nextPin) {
        nextPin = generateNumericPin(4);
      } else {
        const invalidPin = validateFarmerPin(nextPin);
        if (invalidPin) {
          json(res, 422, { error: invalidPin });
          return;
        }
        if (!secureEquals(nextPin, confirmPin)) {
          json(res, 422, { error: 'PIN confirmation does not match.' });
          return;
        }
      }

      const invalidGeneratedPin = validateFarmerPin(nextPin);
      if (invalidGeneratedPin) {
        json(res, 422, { error: invalidGeneratedPin });
        return;
      }

      const upsertResult = ensureFarmerPortalUser(store, farmer, nextPin);
      if (upsertResult.error) {
        json(res, 409, { error: upsertResult.error });
        return;
      }

      const user = upsertResult.user;
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'farmer.pin_reset',
        entity: 'user',
        entityId: user.id,
        details: `${upsertResult.created ? 'Created' : 'Updated'} portal access for farmer ${farmer.name}`
      });

      await writeStore(store);
      json(res, 200, {
        data: {
          farmerId: farmer.id,
          farmerName: farmer.name,
          phone: farmer.phone,
          username: user.username,
          accountCreated: upsertResult.created,
          generatedPin: generate ? nextPin : ''
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/farmers/') && req.method === 'GET') {
    const store = await readStore();
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
      const store = await readStore();
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
      const nextPhone =
        payload.phone !== undefined
          ? normalizePhone(payload.phone) || clean(payload.phone)
          : clean(farmer.phone);
      const nextNationalId =
        payload.nationalId !== undefined ? cleanNationalId(payload.nationalId) : cleanNationalId(farmer.nationalId);
      const nextLocation = payload.location !== undefined ? clean(payload.location) : clean(farmer.location);
      const nextNotes = payload.notes !== undefined ? clean(payload.notes) : clean(farmer.notes);
      const preferredLanguageInput = parsePreferredLanguage(
        payload.preferredLanguage !== undefined ? payload.preferredLanguage : farmer.preferredLanguage
      );
      const nextPreferredLanguage = preferredLanguageInput.value;

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
      if (!preferredLanguageInput.valid) {
        json(res, 422, { error: 'preferredLanguage must be en or sw' });
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
      const linkedFarmerUser = resolveFarmerPortalUser(store, farmer);
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
      farmer.preferredLanguage = nextPreferredLanguage;
      farmer.notes = nextNotes;
      farmer.updatedAt = nowIso();
      if (linkedFarmerUser) {
        const syncResult = syncFarmerPortalUserIdentity(store, farmer, nextPhone, nextName);
        if (!syncResult.ok) {
          json(res, 409, { error: syncResult.error });
          return;
        }
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'farmer.update',
        entity: 'farmer',
        entityId: farmer.id,
        details: `Updated farmer ${farmer.name}`
      });
      await writeStore(store);
      json(res, 200, { data: farmer });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/farmers/') && req.method === 'DELETE') {
    const store = await readStore();
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
    const removedPortalUser = resolveFarmerPortalUser(store, removed);
    if (removedPortalUser) {
      store.users = store.users.filter((row) => row.id !== removedPortalUser.id);
      invalidateUserSessions(store, removedPortalUser.id);
    }
    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'farmer.delete',
      entity: 'farmer',
      entityId: removed.id,
      details: `Deleted farmer ${removed.name}`
    });
    if (removedPortalUser) {
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'auth.user_delete',
        entity: 'user',
        entityId: removedPortalUser.id,
        details: `Deleted linked farmer portal account for ${removed.name}`
      });
    }
    await writeStore(store);
    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/produce' && req.method === 'GET') {
    const store = await readStore();
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
      const store = await readStore();
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
      await writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/produce/') && req.method === 'DELETE') {
    const store = await readStore();
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
    await writeStore(store);
    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/produce-purchases' && req.method === 'GET') {
    const store = await readStore();
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
      const store = await readStore();
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
      await writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/produce-purchases/') && req.method === 'DELETE') {
    const store = await readStore();
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
    await writeStore(store);
    json(res, 200, { data: { success: true } });
    return;
  }

  if (pathname === '/api/payments/owed' && req.method === 'GET') {
    const store = await readStore();
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

  if (pathname === '/api/payments/recommendations' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const statusFilter = clean(reqUrl.searchParams.get('status')).toLowerCase();
    const q = clean(reqUrl.searchParams.get('q'));
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 200, 2000);
    const offset = safeQueryOffset(reqUrl.searchParams, 'offset', 0, 2000000);

    let rows = [...store.paymentRecommendations];
    if (auth.session.role === 'agent') {
      rows = rows.filter((row) => clean(row.createdBy) === clean(auth.session.username));
    }

    if (statusFilter && statusFilter !== 'all') {
      rows = rows.filter((row) => normalizePaymentRecommendationStatus(row.status) === statusFilter);
    }
    rows = filterByTextQuery(rows, q, [
      'id',
      'farmerName',
      'farmerPhone',
      'reason',
      'createdBy',
      'status',
      'rejectionReason'
    ]);

    rows = rows.sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt));
    const total = rows.length;
    const counts = rows.reduce(
      (acc, row) => {
        const key = normalizePaymentRecommendationStatus(row.status);
        if (key in acc) acc[key] += 1;
        return acc;
      },
      { pending: 0, approved: 0, rejected: 0 }
    );

    json(res, 200, {
      data: rows.slice(offset, offset + limit),
      meta: {
        total,
        limit,
        offset,
        q,
        status: statusFilter || 'all',
        counts
      }
    });
    return;
  }

  if (pathname === '/api/payments/recommendations' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req);
      const farmerId = clean(payload.farmerId);
      const amountRaw = parseNumber(payload.amount);
      const reason = clean(payload.reason);
      if (!farmerId || !Number.isFinite(amountRaw) || amountRaw <= 0 || !reason) {
        json(res, 422, { error: 'farmerId, amount (> 0), and reason are required.' });
        return;
      }

      const farmer = findById(store.farmers, farmerId);
      if (!farmer) {
        json(res, 404, { error: 'Farmer not found.' });
        return;
      }

      const owedSnapshot = buildOwedRows(store, { period: 'all' });
      const owedRow = owedSnapshot.rows.find((row) => row.farmerId === farmerId);
      const outstanding = money(owedRow?.balanceKes);
      if (!(outstanding > 0)) {
        json(res, 422, { error: 'This farmer currently has no outstanding balance.' });
        return;
      }

      const amount = money(amountRaw);
      if (amount > outstanding) {
        json(res, 422, {
          error: `Recommended amount (KES ${formatKes(amount)}) exceeds outstanding balance (KES ${formatKes(outstanding)}).`
        });
        return;
      }

      const recommendation = {
        id: id('PREQ'),
        farmerId: farmer.id,
        farmerName: clean(farmer.name),
        farmerPhone: clean(farmer.phone),
        amount,
        requestedOwedKes: outstanding,
        status: 'pending',
        reason,
        decisionNote: '',
        createdBy: auth.session.username,
        createdAt: nowIso(),
        updatedAt: nowIso(),
        approvedBy: '',
        approvedAt: '',
        rejectedBy: '',
        rejectedAt: '',
        rejectionReason: '',
        paymentId: ''
      };

      pushWithCap(store.paymentRecommendations, recommendation, PAYMENT_RECOMMENDATION_CAP);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'payment.recommend.create',
        entity: 'paymentRecommendation',
        entityId: recommendation.id,
        details: `Recommended payment of KES ${formatKes(amount)} for farmer ${recommendation.farmerName}`
      });
      await writeStore(store);

      const pendingCount = store.paymentRecommendations.filter((row) => row.status === 'pending').length;
      json(res, 201, {
        data: recommendation,
        meta: {
          pendingForAdmin: pendingCount
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (
    pathname.startsWith('/api/payments/recommendations/') &&
    pathname.endsWith('/approve') &&
    req.method === 'POST'
  ) {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const recommendationId = routeParamWithSuffix(
        pathname,
        '/api/payments/recommendations/',
        '/approve'
      );
      const recommendation = findById(store.paymentRecommendations, recommendationId);
      if (!recommendation) {
        json(res, 404, { error: 'Payment recommendation not found.' });
        return;
      }
      if (normalizePaymentRecommendationStatus(recommendation.status) !== 'pending') {
        json(res, 409, { error: 'Only pending recommendations can be approved.' });
        return;
      }

      const payload = await readBody(req);
      const decisionNote = clean(payload.decisionNote);
      const farmer = findById(store.farmers, recommendation.farmerId);
      if (!farmer) {
        json(res, 404, { error: 'Farmer linked to recommendation was not found.' });
        return;
      }

      const owedSnapshot = buildOwedRows(store, { period: 'all' });
      const owedRow = owedSnapshot.rows.find((row) => row.farmerId === recommendation.farmerId);
      const outstanding = money(owedRow?.balanceKes);
      if (!(outstanding > 0)) {
        json(res, 409, { error: 'Farmer has no outstanding balance. Reject recommendation or review purchases.' });
        return;
      }

      const amount = money(Math.min(money(recommendation.amount), outstanding));
      const ref = `APR${Date.now().toString().slice(-8)}${crypto.randomInt(10, 99)}`;
      const paymentRecord = {
        id: id('TX'),
        farmerId: farmer.id,
        farmerName: farmer.name,
        amount,
        ref,
        status: 'Received',
        method: 'M-PESA(Mock)',
        notes: clean(
          `Approved from recommendation ${recommendation.id}${decisionNote ? `: ${decisionNote}` : ''}`
        ),
        source: 'agent-recommendation',
        recommendationId: recommendation.id,
        createdBy: auth.session.username,
        createdAt: nowIso(),
        updatedAt: nowIso()
      };
      store.payments.unshift(paymentRecord);

      recommendation.status = 'approved';
      recommendation.approvedBy = auth.session.username;
      recommendation.approvedAt = nowIso();
      recommendation.updatedAt = recommendation.approvedAt;
      recommendation.paymentId = paymentRecord.id;
      recommendation.decisionNote = decisionNote;
      recommendation.rejectedBy = '';
      recommendation.rejectedAt = '';
      recommendation.rejectionReason = '';

      const agentUser = findUserByUsername(store, recommendation.createdBy);
      const agentPhone = clean(agentUser?.phone);
      const farmerPhone = clean(farmer.phone || recommendation.farmerPhone);

      const farmerSms = createSmsLog(store, {
        farmerId: farmer.id,
        farmerName: farmer.name,
        phone: farmerPhone,
        message: `Agem Portal: Payment of KES ${formatKes(amount)} has been approved and processed. Ref ${ref}.`,
        createdBy: auth.session.username
      });
      const agentSms = createSmsLog(store, {
        farmerId: farmer.id,
        farmerName: farmer.name,
        phone: agentPhone,
        message: `Agem Portal: Your recommendation ${recommendation.id} for ${farmer.name} was APPROVED. Payment ref ${ref}.`,
        createdBy: auth.session.username
      });

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'payment.recommend.approve',
        entity: 'paymentRecommendation',
        entityId: recommendation.id,
        details: `Approved recommendation and posted payment ${paymentRecord.ref} (KES ${formatKes(amount)})`
      });
      await writeStore(store);

      json(res, 200, {
        data: recommendation,
        payment: paymentRecord,
        meta: {
          smsNotified: {
            farmer: Boolean(farmerSms),
            agent: Boolean(agentSms)
          }
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (
    pathname.startsWith('/api/payments/recommendations/') &&
    pathname.endsWith('/reject') &&
    req.method === 'POST'
  ) {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const recommendationId = routeParamWithSuffix(
        pathname,
        '/api/payments/recommendations/',
        '/reject'
      );
      const recommendation = findById(store.paymentRecommendations, recommendationId);
      if (!recommendation) {
        json(res, 404, { error: 'Payment recommendation not found.' });
        return;
      }
      if (normalizePaymentRecommendationStatus(recommendation.status) !== 'pending') {
        json(res, 409, { error: 'Only pending recommendations can be rejected.' });
        return;
      }

      const payload = await readBody(req);
      const reason = clean(payload.reason);
      const farmer = findById(store.farmers, recommendation.farmerId);
      const farmerName = clean(farmer?.name || recommendation.farmerName);
      const farmerPhone = clean(farmer?.phone || recommendation.farmerPhone);

      recommendation.status = 'rejected';
      recommendation.rejectedBy = auth.session.username;
      recommendation.rejectedAt = nowIso();
      recommendation.updatedAt = recommendation.rejectedAt;
      recommendation.rejectionReason = reason;
      recommendation.decisionNote = reason;
      recommendation.approvedBy = '';
      recommendation.approvedAt = '';
      recommendation.paymentId = '';

      const agentUser = findUserByUsername(store, recommendation.createdBy);
      const agentPhone = clean(agentUser?.phone);

      const farmerSms = createSmsLog(store, {
        farmerId: recommendation.farmerId,
        farmerName,
        phone: farmerPhone,
        message: `Agem Portal: Payment recommendation for your account was not approved${reason ? ` (${reason})` : ''}. Contact your AGEM agent for details.`,
        createdBy: auth.session.username
      });
      const agentSms = createSmsLog(store, {
        farmerId: recommendation.farmerId,
        farmerName,
        phone: agentPhone,
        message: `Agem Portal: Recommendation ${recommendation.id} for ${farmerName} was REJECTED${reason ? ` (${reason})` : ''}.`,
        createdBy: auth.session.username
      });

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'payment.recommend.reject',
        entity: 'paymentRecommendation',
        entityId: recommendation.id,
        details: `Rejected recommendation ${recommendation.id}${reason ? ` (${reason.slice(0, 120)})` : ''}`
      });
      await writeStore(store);

      json(res, 200, {
        data: recommendation,
        meta: {
          smsNotified: {
            farmer: Boolean(farmerSms),
            agent: Boolean(agentSms)
          }
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/payments/settle' && req.method === 'POST') {
    try {
      const store = await readStore();
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
      await writeStore(store);

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
    const store = await readStore();
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
      const store = await readStore();
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
      await writeStore(store);
      json(res, 201, { data: record });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/payments/') && pathname.endsWith('/status') && req.method === 'PATCH') {
    try {
      const store = await readStore();
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
      await writeStore(store);
      json(res, 200, { data: payment });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/sms' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.smsLogs];
    rows = filterByQuery(rows, reqUrl.searchParams, ['phone', 'message', 'createdBy', 'farmerName']);

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 300, 2000);
    json(res, 200, {
      data: rows.slice(0, limit),
      meta: smsCostSnapshot(store.smsLogs)
    });
    return;
  }

  if (pathname === '/api/sms/recipients' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const q = clean(reqUrl.searchParams.get('q'));
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 50, 200);
    const offset = safeQueryOffset(reqUrl.searchParams, 'offset', 0, 2_000_000);

    let recipients = store.farmers.filter((row) => clean(row.phone));
    recipients = filterByTextQuery(recipients, q, ['id', 'name', 'phone', 'nationalId', 'location', 'preferredLanguage', 'notes']);

    const total = recipients.length;
    const rows = recipients.slice(offset, offset + limit).map((row) => ({
      id: row.id,
      name: row.name,
      phone: row.phone,
      nationalId: row.nationalId,
      location: row.location,
      preferredLanguage: languageOrDefault(row.preferredLanguage)
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
      const store = await readStore();
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
        ownerCostKes: money(SMS_OWNER_COST_PER_MESSAGE_KES),
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
      await writeStore(store);
      json(res, 201, { data: log });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/sms/send-bulk' && req.method === 'POST') {
    try {
      const store = await readStore();
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
          ownerCostKes: money(SMS_OWNER_COST_PER_MESSAGE_KES),
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

      await writeStore(store);
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

  if (pathname === '/api/ai/msaidizi/status' && req.method === 'GET') {
    const store = await readStore();
    const session = currentSession(req, store);
    const role = clean(session?.role).toLowerCase() || 'guest';
    const syncResult = syncMsaidiziInStore(store, {
      reason: 'status',
      actor: session?.username || 'guest',
      force: false
    });
    if (syncResult.updated) {
      await writeStore(store);
    }

    json(res, 200, {
      data: {
        enabled: AI_MSAIDIZI_SYNC_ENABLED,
        intervalHours: AI_MSAIDIZI_SYNC_INTERVAL_HOURS,
        status: clean(store.aiMsaidizi?.status) || 'ready',
        lastSyncAt: clean(store.aiMsaidizi?.lastSyncAt),
        sourceSignature: clean(store.aiMsaidizi?.sourceSignature),
        moduleCount: msaidiziAccessibleModules(store, role).length
      }
    });
    return;
  }

  if (pathname === '/api/ai/msaidizi/sync' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }
      const payload = await readBody(req, 80_000);
      const force = Boolean(payload?.force);
      const syncResult = syncMsaidiziInStore(store, {
        reason: clean(payload?.reason) || 'manual',
        actor: auth.session.username,
        force
      });

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.msaidizi.sync',
        entity: 'ai',
        details: syncResult.updated
          ? `Msaidizi synced and regenerated (${syncResult.modules.length} modules).`
          : 'Msaidizi sync checked. No changes detected.'
      });
      await writeStore(store);
      json(res, 200, {
        data: {
          updated: syncResult.updated,
          status: clean(store.aiMsaidizi?.status) || 'ready',
          lastSyncAt: clean(store.aiMsaidizi?.lastSyncAt),
          sourceSignature: clean(store.aiMsaidizi?.sourceSignature),
          moduleCount: (store.aiMsaidizi?.modules || []).length
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/msaidizi/context' && req.method === 'GET') {
    const store = await readStore();
    const session = currentSession(req, store);
    const role = clean(session?.role).toLowerCase() || clean(reqUrl.searchParams.get('role')).toLowerCase() || 'guest';
    const pane = clean(reqUrl.searchParams.get('pane')).toLowerCase() || 'auth';
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 4, 10);

    const syncResult = syncMsaidiziInStore(store, {
      reason: 'context',
      actor: session?.username || role || 'guest',
      force: false
    });
    if (syncResult.updated) {
      await writeStore(store);
    }

    const modules = findMsaidiziModules(store, '', pane, role, limit).map((module) => ({
      id: module.id,
      title: module.title,
      short: module.short,
      panes: module.panes || [],
      roles: module.roles || [],
      steps: Array.isArray(module.steps) ? module.steps.slice(0, 5) : [],
      doNow: Array.isArray(module.doNow) ? module.doNow.slice(0, 4) : [],
      avoid: Array.isArray(module.avoid) ? module.avoid.slice(0, 4) : [],
      waitFor: Array.isArray(module.waitFor) ? module.waitFor.slice(0, 4) : []
    }));

    json(res, 200, {
      data: {
        pane,
        role,
        modules,
        status: clean(store.aiMsaidizi?.status) || 'ready',
        lastSyncAt: clean(store.aiMsaidizi?.lastSyncAt),
        sourceSignature: clean(store.aiMsaidizi?.sourceSignature)
      }
    });
    return;
  }

  if (pathname === '/api/ai/msaidizi/ask' && req.method === 'POST') {
    try {
      const store = await readStore();
      const session = currentSession(req, store);
      const role =
        clean(session?.role).toLowerCase() ||
        clean(req.headers['x-role']).toLowerCase() ||
        clean(reqUrl.searchParams.get('role')).toLowerCase() ||
        'guest';
      const payload = await readBody(req, 500_000);
      const question = clean(payload.question);
      const pane = clean(payload.pane).toLowerCase() || 'auth';
      const mode = normalizeMsaidiziMode(payload.mode);
      const responseId = id('MSA');
      const syncResult = syncMsaidiziInStore(store, {
        reason: 'ask',
        actor: session?.username || role || 'guest',
        force: false
      });

      const modules = findMsaidiziModules(store, question, pane, role, 4);
      let local = buildMsaidiziLocalAnswer(modules, mode);
      let source = 'local-rules';
      let model = 'local-rules';
      let warning = '';

      if (OPENAI_API_KEY) {
        const moduleContext = modules
          .map((module, index) => {
            return `[MODULE-${index + 1}] ${module.title}
Short: ${module.short}
Steps: ${(module.steps || []).join(' | ')}
Troubleshooting: ${(module.troubleshooting || []).join(' | ')}
Do now: ${(module.doNow || []).join(' | ')}
Avoid: ${(module.avoid || []).join(' | ')}
Wait for: ${(module.waitFor || []).join(' | ')}`;
          })
          .join('\n\n');

        const openAi = await callOpenAiJson(
          'You are Msaidizi, the AGEM portal instruction assistant. Explain only from provided module context. Be concise, practical, and step-oriented. If mode is simple, use very plain wording. If mode is step_by_step, provide full ordered steps.',
          `Mode: ${mode}
Role: ${role}
Current page: ${pane}
Question: ${question || '(no question provided - return the best quick guide for current page)'}

Instruction modules:
${moduleContext || 'No modules matched.'}`,
          {
            type: 'object',
            additionalProperties: false,
            required: ['answerShort', 'steps', 'troubleshoot', 'doNow', 'avoid', 'waitFor', 'relatedModules'],
            properties: {
              answerShort: { type: 'string' },
              steps: { type: 'array', maxItems: 10, items: { type: 'string' } },
              troubleshoot: { type: 'array', maxItems: 8, items: { type: 'string' } },
              doNow: { type: 'array', maxItems: 8, items: { type: 'string' } },
              avoid: { type: 'array', maxItems: 8, items: { type: 'string' } },
              waitFor: { type: 'array', maxItems: 8, items: { type: 'string' } },
              relatedModules: { type: 'array', maxItems: 6, items: { type: 'string' } }
            }
          }
        );

        if (openAi.ok) {
          source = 'openai';
          model = OPENAI_MODEL;
          local = {
            answerShort: clean(openAi.data?.answerShort) || local.answerShort,
            steps: Array.isArray(openAi.data?.steps)
              ? openAi.data.steps.map((row) => clean(row)).filter(Boolean).slice(0, 10)
              : local.steps,
            troubleshoot: Array.isArray(openAi.data?.troubleshoot)
              ? openAi.data.troubleshoot.map((row) => clean(row)).filter(Boolean).slice(0, 8)
              : local.troubleshoot,
            doNow: Array.isArray(openAi.data?.doNow)
              ? openAi.data.doNow.map((row) => clean(row)).filter(Boolean).slice(0, 8)
              : local.doNow,
            avoid: Array.isArray(openAi.data?.avoid)
              ? openAi.data.avoid.map((row) => clean(row)).filter(Boolean).slice(0, 8)
              : local.avoid,
            waitFor: Array.isArray(openAi.data?.waitFor)
              ? openAi.data.waitFor.map((row) => clean(row)).filter(Boolean).slice(0, 8)
              : local.waitFor,
            relatedModules: Array.isArray(openAi.data?.relatedModules)
              ? openAi.data.relatedModules.map((row) => clean(row)).filter(Boolean).slice(0, 6)
              : local.relatedModules
          };
        } else {
          warning = openAi.error || 'OpenAI unavailable. Showing local instructions.';
        }
      }

      addActivity(store, {
        actor: session?.username || role || 'guest',
        role: session?.role || role || 'guest',
        action: 'ai.msaidizi.ask',
        entity: 'ai',
        entityId: responseId,
        details: `Msaidizi asked from pane "${pane}" (${mode} mode): "${question.slice(0, 120)}"`
      });

      if (syncResult.updated) {
        addActivity(store, {
          actor: session?.username || role || 'guest',
          role: session?.role || role || 'guest',
          action: 'ai.msaidizi.sync',
          entity: 'ai',
          details: `Msaidizi docs auto-updated before response (signature ${syncResult.sourceSignature}).`
        });
      }

      await writeStore(store);

      json(res, 200, {
        data: {
          responseId,
          mode,
          pane,
          role,
          source,
          model,
          warning,
          answer: local,
          modules: modules.map((module) => ({
            id: module.id,
            title: module.title,
            short: module.short,
            panes: module.panes || []
          })),
          status: clean(store.aiMsaidizi?.status) || 'ready',
          lastSyncAt: clean(store.aiMsaidizi?.lastSyncAt),
          sourceSignature: clean(store.aiMsaidizi?.sourceSignature)
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/copilot' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 400_000);
      const question = clean(payload.question);
      if (!question) {
        json(res, 422, { error: 'question is required' });
        return;
      }
      const responseId = id('AIR');
      const maxContextDocsRaw = Number(store?.aiPromptConfig?.maxContextDocs);
      const maxContextDocs = Number.isFinite(maxContextDocsRaw)
        ? Math.max(1, Math.min(8, Math.floor(maxContextDocsRaw)))
        : 4;
      const citations = findKnowledgeMatches(store, question, maxContextDocs);

      const local = localCopilotAnswer(store, question);
      let source = 'local-rules';
      let model = 'local-rules';
      let warning = '';
      let answer = local.answer;
      let insights = (local.insights || []).map((item) =>
        typeof item === 'string'
          ? item
          : Object.entries(item || {})
            .map(([key, value]) => `${key}: ${value}`)
            .join(', ')
      );
      let actions = (local.actions || []).map((item) =>
        typeof item === 'string' ? item : `${item.label}: ${item.description}`
      );

      if (OPENAI_API_KEY) {
        const summary = makeSummary(store);
        const contextBlock = citations.length
          ? citations
            .map((doc, index) => {
              const tags = Array.isArray(doc.tags) && doc.tags.length ? ` | tags: ${doc.tags.join(', ')}` : '';
              return `[DOC-${index + 1}] ${doc.title}${doc.source ? ` (${doc.source})` : ''}${tags}\n${doc.snippet}`;
            })
            .join('\n\n')
          : 'No knowledge documents matched this question.';
        const openAi = await callOpenAiJson(
          aiPromptPolicy(
            store,
            'copilotSystem',
            'You are an operations copilot for an avocado platform. Give concise, actionable, factual guidance. Do not claim to execute payments or system changes.'
          ),
          `Question: ${question}

Platform snapshot:
- Farmers: ${summary.farmers || 0}
- QC records: ${summary.qcRecords || 0}
- Purchased records: ${summary.purchasedRecords || 0}
- Owed farmers: ${summary.owedFarmers || 0}
- Total owed KES: ${summary.totalOwedKes || 0}
- Payments received KES: ${summary.paymentsReceived || 0}
- Payment success %: ${summary.paymentSuccessRate || 0}
- SMS spend 24h KES: ${summary.smsSpentLast24hKes || 0}

Knowledge context (use only if relevant):
${contextBlock}

Local baseline answer:
${local.answer}`,
          {
            type: 'object',
            additionalProperties: false,
            required: ['answer', 'insights', 'actions'],
            properties: {
              answer: { type: 'string' },
              insights: {
                type: 'array',
                maxItems: 6,
                items: { type: 'string' }
              },
              actions: {
                type: 'array',
                maxItems: 6,
                items: { type: 'string' }
              }
            }
          }
        );

        if (openAi.ok) {
          source = 'openai';
          model = OPENAI_MODEL;
          answer = clean(openAi.data?.answer) || answer;
          insights = Array.isArray(openAi.data?.insights) && openAi.data.insights.length
            ? openAi.data.insights.map((item) => clean(item)).filter(Boolean)
            : insights;
          actions = Array.isArray(openAi.data?.actions) && openAi.data.actions.length
            ? openAi.data.actions.map((item) => clean(item)).filter(Boolean)
            : actions;
        } else {
          warning = openAi.error || 'OpenAI unavailable. Showing local copilot answer.';
        }
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.copilot.ask',
        entity: 'ai',
        entityId: responseId,
        details: `Admin Copilot asked: "${question.slice(0, 140)}"`
      });
      await writeStore(store);

      json(res, 200, {
        data: {
          responseId,
          question,
          intent: local.intent || 'operations_summary',
          answer,
          insights,
          actions,
          citations,
          source,
          model,
          warning
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/import-qa' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 12_000_000);
      const records = Array.isArray(payload.records) ? payload.records : [];
      if (!records.length) {
        json(res, 422, { error: 'records array is required' });
        return;
      }
      if (records.length > 10000) {
        json(res, 422, { error: 'For Smart Import QA, upload up to 10,000 rows per run.' });
        return;
      }

      const existingByPhone = new Map(
        store.farmers
          .map((row) => [normalizePhone(row.phone) || clean(row.phone), row])
          .filter(([phone]) => Boolean(phone))
      );
      const existingByNationalId = new Map(
        store.farmers
          .map((row) => [normalizedNationalId(row.nationalId), row])
          .filter(([nationalId]) => Boolean(nationalId))
      );

      const rows = records.map((raw, idx) => {
        if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
          return {
            row: idx + 1,
            status: 'blocked',
            confidence: 0.2,
            mapped: {},
            issues: [
              {
                code: 'invalid_row_shape',
                severity: 'error',
                message: 'Row is not a valid object.',
                suggestion: 'Ensure the sheet has a header row and structured columns.'
              }
            ],
            proposedFix: null,
            duplicate: null
          };
        }
        return summarizeSmartImportRow(raw, idx + 1, existingByPhone, existingByNationalId);
      });

      const statusCounts = rows.reduce(
        (acc, row) => {
          acc[row.status] = (acc[row.status] || 0) + 1;
          return acc;
        },
        { ready: 0, review: 0, blocked: 0 }
      );

      const topIssues = {};
      rows.forEach((row) => {
        row.issues.forEach((issue) => {
          topIssues[issue.code] = (topIssues[issue.code] || 0) + 1;
        });
      });
      const issueSummary = Object.entries(topIssues)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8)
        .map(([code, count]) => ({ code, count }));

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.import.qa',
        entity: 'ai',
        details: `Smart Import QA run on ${records.length} rows`
      });
      await writeStore(store);

      json(res, 200, {
        data: {
          totalRows: records.length,
          readyRows: statusCounts.ready || 0,
          reviewRows: statusCounts.review || 0,
          blockedRows: statusCounts.blocked || 0,
          autoFixableRows: rows.filter((row) => Boolean(row.proposedFix)).length,
          issueSummary,
          rows: rows.slice(0, 500)
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/sms-draft' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 300_000);
      const responseId = id('AIR');
      const purpose = clean(payload.purpose || payload.goal);
      if (!purpose) {
        json(res, 422, { error: 'purpose is required' });
        return;
      }

      const audience = clean(payload.audience) || 'selected farmers';
      const tone = clean(payload.tone).toLowerCase() || 'professional';
      const language = clean(payload.language).toLowerCase() || 'bilingual';
      const maxLengthRaw = Number(payload.maxLength);
      const maxLength = Number.isFinite(maxLengthRaw) ? Math.max(30, Math.min(500, Math.floor(maxLengthRaw))) : null;

      const localDraft = buildSmsDraftsLocal(
        {
          purpose,
          audience,
          tone,
          language,
          maxLength
        },
        store
      );

      let drafts = localDraft.drafts;
      let source = 'local-rules';
      let model = 'local-rules';
      let warning = '';

      if (OPENAI_API_KEY) {
        const openAi = await callOpenAiJson(
          aiPromptPolicy(
            store,
            'smsSystem',
            'Draft concise operational SMS messages for Kenyan farmers. Keep wording simple. Respect requested language and max length. Return only final draft text.'
          ),
          `Draft SMS for Agem Portal.
Purpose: ${purpose}
Audience: ${audience}
Tone: ${tone}
Language request: ${language}
Max length: ${maxLength || 'none'}`,
          {
            type: 'object',
            additionalProperties: false,
            required: ['drafts'],
            properties: {
              drafts: {
                type: 'array',
                minItems: 1,
                maxItems: 3,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['language', 'message'],
                  properties: {
                    language: { type: 'string' },
                    message: { type: 'string' }
                  }
                }
              }
            }
          }
        );

        if (openAi.ok && Array.isArray(openAi.data?.drafts) && openAi.data.drafts.length) {
          drafts = openAi.data.drafts
            .map((item) => ({
              language: clean(item.language) || 'English',
              message: clean(item.message)
            }))
            .filter((item) => item.message);
          if (!drafts.length) drafts = localDraft.drafts;
          source = 'openai';
          model = OPENAI_MODEL;
        } else if (!openAi.ok) {
          warning = openAi.error || 'OpenAI unavailable. Showing local SMS drafts.';
        }
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.sms.draft',
        entity: 'ai',
        entityId: responseId,
        details: `AI SMS draft generated for purpose "${purpose.slice(0, 120)}"`
      });
      await writeStore(store);

      json(res, 200, {
        data: {
          responseId,
          purpose,
          audience,
          tone,
          language,
          maxLength: maxLength || '',
          drafts,
          source,
          model,
          warning,
          metadata: localDraft.metadata
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/qc-intelligence' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 300_000);
      const responseId = id('AIR');
      const qcRecordId = clean(payload.qcRecordId);
      const limitRaw = parseNumber(payload.limit);
      const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(100, Math.floor(limitRaw))) : 30;

      const allQc = [...store.produce].sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt));
      if (!allQc.length) {
        json(res, 422, { error: 'No QC records found yet.' });
        return;
      }

      let selected = [];
      if (qcRecordId) {
        const match = allQc.find((row) => row.id === qcRecordId);
        if (!match) {
          json(res, 404, { error: 'QC record not found.' });
          return;
        }
        selected = [match];
      } else {
        selected = allQc.slice(0, limit);
      }

      let source = 'local-rules';
      let model = 'local-rules';
      let warning = '';

      let items = selected.map((record) => localQcIntelligence(record));
      const summary = {
        totalAnalyzed: items.length,
        highRisk: items.filter((item) => item.riskLevel === 'high').length,
        mediumRisk: items.filter((item) => item.riskLevel === 'medium').length,
        lowRisk: items.filter((item) => item.riskLevel === 'low').length,
        flagged:
          items.filter((item) => item.recommendedDecision !== 'Accept' || item.riskLevel !== 'low').length
      };

      if (OPENAI_API_KEY && items.length === 1) {
        const item = items[0];
        const openAi = await callOpenAiJson(
          aiPromptPolicy(
            store,
            'qcSystem',
            'You are a farm-gate avocado QC assistant. Improve explanation quality but keep decisions conservative and practical.'
          ),
          `QC Snapshot:
- Variety: ${item.variety}
- Lot weight kg: ${item.lotWeightKgs}
- Current decision: ${item.currentDecision}
- Suggested decision (local): ${item.recommendedDecision}
- Risk level: ${item.riskLevel}
- Risk score: ${item.riskScore}
- Reasons: ${(item.reasons || []).join(' | ')}`,
          {
            type: 'object',
            additionalProperties: false,
            required: ['summary', 'reasons', 'actions'],
            properties: {
              summary: { type: 'string' },
              reasons: {
                type: 'array',
                minItems: 1,
                maxItems: 5,
                items: { type: 'string' }
              },
              actions: {
                type: 'array',
                minItems: 1,
                maxItems: 5,
                items: { type: 'string' }
              }
            }
          }
        );

        if (openAi.ok) {
          source = 'openai';
          model = OPENAI_MODEL;
          item.summary = clean(openAi.data?.summary) || item.summary;
          if (Array.isArray(openAi.data?.reasons) && openAi.data.reasons.length) {
            item.reasons = openAi.data.reasons.map((row) => clean(row)).filter(Boolean).slice(0, 5);
          }
          if (Array.isArray(openAi.data?.actions) && openAi.data.actions.length) {
            item.actions = openAi.data.actions.map((row) => clean(row)).filter(Boolean).slice(0, 5);
          }
          items = [item];
        } else {
          warning = openAi.error || 'OpenAI unavailable. Showing local QC intelligence.';
        }
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.qc.intelligence',
        entity: 'ai',
        entityId: responseId,
        details: `QC intelligence run on ${items.length} lot(s)`
      });
      await writeStore(store);

      json(res, 200, {
        data: {
          responseId,
          qcRecordId: qcRecordId || '',
          limit,
          summary,
          items,
          source,
          model,
          warning
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/payment-risk' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 200_000);
      const responseId = id('AIR');
      const period = clean(payload.period) || 'week';
      const from = clean(payload.from);
      const to = clean(payload.to);

      let source = 'local-rules';
      let model = 'local-rules';
      let warning = '';

      const report = localPaymentRiskReport(store, { period, from, to });

      if (OPENAI_API_KEY) {
        const openAi = await callOpenAiJson(
          aiPromptPolicy(
            store,
            'paymentRiskSystem',
            'You are a payments risk assistant. Summarize risk and provide concise operational actions. Do not claim to execute payments.'
          ),
          `Payment risk report:
- Period: ${report.range.period}
- Payments analyzed: ${report.paymentCount}
- Overall risk: ${report.summary.overallRisk}
- High flags: ${report.summary.highFlags}
- Medium flags: ${report.summary.mediumFlags}
- Low flags: ${report.summary.lowFlags}
- Pending count: ${report.summary.pendingCount}
- Failed count: ${report.summary.failedCount}
- Local narrative: ${report.narrative}
- Top flags: ${report.flags.slice(0, 5).map((row) => `${row.code}: ${row.detail}`).join(' | ')}`,
          {
            type: 'object',
            additionalProperties: false,
            required: ['narrative', 'actions'],
            properties: {
              narrative: { type: 'string' },
              actions: {
                type: 'array',
                minItems: 1,
                maxItems: 6,
                items: { type: 'string' }
              }
            }
          }
        );

        if (openAi.ok) {
          source = 'openai';
          model = OPENAI_MODEL;
          report.narrative = clean(openAi.data?.narrative) || report.narrative;
          if (Array.isArray(openAi.data?.actions) && openAi.data.actions.length) {
            report.actions = openAi.data.actions.map((row) => clean(row)).filter(Boolean).slice(0, 6);
          }
        } else {
          warning = openAi.error || 'OpenAI unavailable. Showing local payment risk report.';
        }
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.payment.risk',
        entity: 'ai',
        entityId: responseId,
        details: `Payment risk check run for period "${report.range.period}" (${report.paymentCount} payment(s))`
      });
      await writeStore(store);

      json(res, 200, {
        data: {
          responseId,
          ...report,
          source,
          model,
          warning
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/proposals' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.aiProposals];
    rows = filterByQuery(rows, reqUrl.searchParams, ['id', 'status', 'actionType', 'title', 'description', 'source', 'createdBy']);
    const statusFilter = normalizeAiProposalStatus(reqUrl.searchParams.get('status'));
    if (clean(reqUrl.searchParams.get('status'))) {
      rows = rows.filter((row) => normalizeAiProposalStatus(row.status) === statusFilter);
    }

    const actionType = clean(reqUrl.searchParams.get('actionType')).toLowerCase();
    if (actionType) {
      rows = rows.filter((row) => clean(row.actionType).toLowerCase() === actionType);
    }

    rows.sort((a, b) => dateMs(b.updatedAt || b.createdAt) - dateMs(a.updatedAt || a.createdAt));
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 100, 500);
    const offset = safeQueryInt(reqUrl.searchParams, 'offset', 0, 5000);
    const total = rows.length;
    const paged = rows.slice(offset, offset + limit);
    const counts = rows.reduce(
      (acc, row) => {
        const key = normalizeAiProposalStatus(row.status);
        acc[key] = (acc[key] || 0) + 1;
        return acc;
      },
      { pending: 0, approved: 0, rejected: 0, executed: 0 }
    );

    json(res, 200, {
      data: paged,
      meta: {
        total,
        limit,
        offset,
        counts
      }
    });
    return;
  }

  if (pathname === '/api/ai/proposals' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 600_000);
      const actionType = clean(payload.actionType).toLowerCase();
      const allowedActionTypes = new Set([
        'payment_risk_to_tasks',
        'qc_risk_to_tasks',
        'create_ops_tasks',
        'generate_executive_brief'
      ]);
      if (!actionType || !allowedActionTypes.has(actionType)) {
        json(res, 422, {
          error:
            'actionType is required and must be one of: payment_risk_to_tasks, qc_risk_to_tasks, create_ops_tasks, generate_executive_brief'
        });
        return;
      }

      const now = nowIso();
      const confidenceRaw = Number(payload.confidence);
      const confidence = Number.isFinite(confidenceRaw)
        ? Number(Math.max(0, Math.min(1, confidenceRaw)).toFixed(2))
        : '';
      const title =
        clean(payload.title) ||
        (
          actionType === 'payment_risk_to_tasks'
            ? 'Create Ops Tasks from Payment Risk'
            : actionType === 'qc_risk_to_tasks'
              ? 'Create Ops Tasks from QC Risk'
              : actionType === 'create_ops_tasks'
                ? 'Create Custom Ops Tasks'
                : 'Generate Executive Brief'
        );
      const proposal = {
        id: id('AIP'),
        status: 'pending',
        actionType,
        title,
        description: clean(payload.description),
        confidence,
        source: clean(payload.source) || 'admin',
        payload: payload.payload && typeof payload.payload === 'object' && !Array.isArray(payload.payload) ? payload.payload : {},
        createdBy: auth.session.username,
        createdAt: now,
        updatedAt: now,
        approvedBy: '',
        approvedAt: '',
        rejectedBy: '',
        rejectedAt: '',
        rejectionReason: '',
        executedBy: '',
        executedAt: '',
        executionSummary: ''
      };

      pushWithCap(store.aiProposals, proposal, AI_PROPOSAL_CAP);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.proposal.create',
        entity: 'ai',
        entityId: proposal.id,
        details: `Created proposal "${proposal.title}" (${proposal.actionType})`
      });
      await writeStore(store);

      json(res, 201, { data: proposal });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/ai/proposals/') && pathname.endsWith('/approve') && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const proposalId = routeParamWithSuffix(pathname, '/api/ai/proposals/', '/approve');
      const proposal = findById(store.aiProposals, proposalId);
      if (!proposal) {
        json(res, 404, { error: 'Proposal not found' });
        return;
      }
      if (normalizeAiProposalStatus(proposal.status) !== 'pending') {
        json(res, 409, { error: 'Only pending proposals can be approved.' });
        return;
      }

      proposal.status = 'approved';
      proposal.approvedBy = auth.session.username;
      proposal.approvedAt = nowIso();
      proposal.updatedAt = proposal.approvedAt;
      proposal.rejectedBy = '';
      proposal.rejectedAt = '';
      proposal.rejectionReason = '';

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.proposal.approve',
        entity: 'ai',
        entityId: proposal.id,
        details: `Approved proposal "${proposal.title}"`
      });
      await writeStore(store);
      json(res, 200, { data: proposal });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/ai/proposals/') && pathname.endsWith('/reject') && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const proposalId = routeParamWithSuffix(pathname, '/api/ai/proposals/', '/reject');
      const proposal = findById(store.aiProposals, proposalId);
      if (!proposal) {
        json(res, 404, { error: 'Proposal not found' });
        return;
      }
      const currentStatus = normalizeAiProposalStatus(proposal.status);
      if (!['pending', 'approved'].includes(currentStatus)) {
        json(res, 409, { error: 'Only pending or approved proposals can be rejected.' });
        return;
      }

      const payload = await readBody(req, 100_000);
      proposal.status = 'rejected';
      proposal.rejectedBy = auth.session.username;
      proposal.rejectedAt = nowIso();
      proposal.rejectionReason = clean(payload.reason);
      proposal.updatedAt = proposal.rejectedAt;

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.proposal.reject',
        entity: 'ai',
        entityId: proposal.id,
        details: `Rejected proposal "${proposal.title}"${proposal.rejectionReason ? ` (${proposal.rejectionReason.slice(0, 120)})` : ''}`
      });
      await writeStore(store);
      json(res, 200, { data: proposal });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/ai/proposals/') && pathname.endsWith('/execute') && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const proposalId = routeParamWithSuffix(pathname, '/api/ai/proposals/', '/execute');
      const proposal = findById(store.aiProposals, proposalId);
      if (!proposal) {
        json(res, 404, { error: 'Proposal not found' });
        return;
      }
      if (normalizeAiProposalStatus(proposal.status) !== 'approved') {
        json(res, 409, { error: 'Only approved proposals can be executed.' });
        return;
      }

      const result = await executeAiProposal(store, proposal, auth.session.username);
      proposal.status = 'executed';
      proposal.executedBy = auth.session.username;
      proposal.executedAt = nowIso();
      proposal.updatedAt = proposal.executedAt;
      proposal.executionSummary = clean(result?.summary);

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.proposal.execute',
        entity: 'ai',
        entityId: proposal.id,
        details: proposal.executionSummary || `Executed proposal "${proposal.title}"`
      });
      await writeStore(store);
      json(res, 200, {
        data: proposal,
        result
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/briefs/run-now' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const brief = await generateExecutiveBrief(store, {
        trigger: 'manual',
        requestedBy: auth.session.username
      });
      pushWithCap(store.aiBriefs, brief, AI_EXEC_BRIEF_MAX_HISTORY);
      if (brief.alert) {
        pushWithCap(store.aiAlerts, brief.alert, AI_EXEC_BRIEF_MAX_HISTORY);
      }
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.brief.run_now',
        entity: 'ai',
        entityId: brief.id,
        details: `Generated executive brief (${brief.overallRisk} risk)`
      });
      await writeStore(store);
      json(res, 201, { data: brief });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/briefs' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 20, 300);
    const includeAlerts = reqUrl.searchParams.get('includeAlerts') !== 'false';
    const briefs = [...store.aiBriefs]
      .sort((a, b) => dateMs(b.generatedAt || b.createdAt) - dateMs(a.generatedAt || a.createdAt))
      .slice(0, limit);
    const alerts = includeAlerts
      ? [...store.aiAlerts]
        .sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt))
        .slice(0, limit)
      : [];

    json(res, 200, {
      data: {
        briefs,
        alerts
      }
    });
    return;
  }

  if (pathname === '/api/ai/alerts' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const severity = clean(reqUrl.searchParams.get('severity')).toLowerCase();
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 50, 500);
    let rows = [...store.aiAlerts];
    if (severity) {
      rows = rows.filter((row) => normalizeRiskSeverity(row.severity, '') === severity);
    }
    rows.sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt));
    json(res, 200, { data: rows.slice(0, limit) });
    return;
  }

  if (pathname === '/api/ai/feedback' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 120_000);
      const tool = clean(payload.tool).toLowerCase();
      if (!tool) {
        json(res, 422, { error: 'tool is required' });
        return;
      }

      const rawRating = clean(payload.rating).toLowerCase();
      const rating =
        rawRating === 'up' || rawRating === 'positive'
          ? 'up'
          : rawRating === 'down' || rawRating === 'negative'
            ? 'down'
            : 'neutral';

      const entry = {
        id: id('AIFB'),
        tool,
        rating,
        note: clean(payload.note).slice(0, 2000),
        responseId: clean(payload.responseId),
        prompt: clean(payload.prompt).slice(0, 3000),
        actor: auth.session.username,
        role: auth.session.role,
        createdAt: nowIso()
      };
      pushWithCap(store.aiFeedback, entry, AI_FEEDBACK_CAP);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.feedback.submit',
        entity: 'ai',
        entityId: entry.id,
        details: `Feedback submitted for ${entry.tool} (${entry.rating})`
      });
      await writeStore(store);
      json(res, 201, { data: entry });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/feedback/summary' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 40, 500);
    const summary = buildAiFeedbackSummary(store.aiFeedback);
    const recent = [...store.aiFeedback]
      .sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt))
      .slice(0, limit);
    json(res, 200, { data: { ...summary, recent } });
    return;
  }

  if (pathname === '/api/ai/evals/run' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const checks = runAiEvalSuite(store);
      const passCount = checks.filter((row) => Boolean(row.pass)).length;
      const failCount = checks.length - passCount;
      const run = {
        id: id('AIEVAL'),
        createdAt: nowIso(),
        actor: auth.session.username,
        totalChecks: checks.length,
        passCount,
        failCount,
        scorePct: checks.length ? Number(((passCount / checks.length) * 100).toFixed(1)) : 0,
        checks
      };
      pushWithCap(store.aiEvalRuns, run, AI_EVAL_CAP);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.eval.run',
        entity: 'ai',
        entityId: run.id,
        details: `AI eval run completed: ${run.passCount}/${run.totalChecks} checks passed`
      });
      await writeStore(store);
      json(res, 201, { data: run });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/evals' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 20, 200);
    const rows = [...store.aiEvalRuns]
      .sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt))
      .slice(0, limit);
    json(res, 200, { data: rows });
    return;
  }

  if (pathname === '/api/ai/prompt-config' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }
    json(res, 200, { data: store.aiPromptConfig || {} });
    return;
  }

  if (pathname === '/api/ai/prompt-config' && req.method === 'PATCH') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 400_000);
      const allowedStringKeys = [
        'copilotSystem',
        'smsSystem',
        'qcSystem',
        'paymentRiskSystem',
        'executiveBriefSystem'
      ];

      for (const key of allowedStringKeys) {
        if (payload[key] === undefined) continue;
        const next = clean(payload[key]);
        if (!next) {
          delete store.aiPromptConfig[key];
          continue;
        }
        store.aiPromptConfig[key] = next.slice(0, 4000);
      }
      if (payload.maxContextDocs !== undefined) {
        const parsed = Number(payload.maxContextDocs);
        if (!Number.isFinite(parsed)) {
          json(res, 422, { error: 'maxContextDocs must be numeric.' });
          return;
        }
        store.aiPromptConfig.maxContextDocs = Math.max(1, Math.min(8, Math.floor(parsed)));
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.prompt_config.update',
        entity: 'ai',
        details: 'Updated AI prompt policy configuration'
      });
      await writeStore(store);
      json(res, 200, { data: store.aiPromptConfig || {} });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ai/knowledge' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const q = clean(reqUrl.searchParams.get('q'));
    const includeContent = reqUrl.searchParams.get('includeContent') === 'true';
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 100, 1000);
    let docs = [...store.aiKnowledgeDocs]
      .sort((a, b) => dateMs(b.updatedAt || b.createdAt) - dateMs(a.updatedAt || a.createdAt));
    if (q) {
      const tokens = tokenizeSearchText(q);
      docs = docs.filter((doc) => {
        const haystack = [
          clean(doc.title).toLowerCase(),
          clean(doc.source).toLowerCase(),
          Array.isArray(doc.tags) ? doc.tags.map((row) => clean(row).toLowerCase()).join(' ') : '',
          clean(doc.content).toLowerCase()
        ].join(' ');
        return tokens.every((token) => haystack.includes(token));
      });
    }
    docs = docs.slice(0, limit).map((doc) => ({
      ...doc,
      snippet: buildKnowledgeSnippet(doc.content, tokenizeSearchText(q)),
      content: includeContent ? clean(doc.content) : ''
    }));

    json(res, 200, { data: docs });
    return;
  }

  if (pathname === '/api/ai/knowledge' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 800_000);
      const title = clean(payload.title);
      const content = clean(payload.content);
      if (!title || !content) {
        json(res, 422, { error: 'title and content are required.' });
        return;
      }

      const tagsRaw = Array.isArray(payload.tags)
        ? payload.tags
        : clean(payload.tags)
          ? clean(payload.tags).split(',')
          : [];
      const tags = tagsRaw
        .map((row) => clean(row).toLowerCase())
        .filter(Boolean)
        .slice(0, 20);
      const now = nowIso();
      const doc = {
        id: id('KDOC'),
        title: title.slice(0, 220),
        source: clean(payload.source).slice(0, 220),
        tags,
        content: content.slice(0, 20000),
        createdBy: auth.session.username,
        createdAt: now,
        updatedAt: now
      };
      pushWithCap(store.aiKnowledgeDocs, doc, AI_KNOWLEDGE_CAP);
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ai.knowledge.add',
        entity: 'ai',
        entityId: doc.id,
        details: `Added knowledge document "${doc.title}"`
      });
      await writeStore(store);
      json(res, 201, { data: doc });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname.startsWith('/api/ai/knowledge/') && req.method === 'DELETE') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const docId = routeParam(pathname, '/api/ai/knowledge/');
    const before = store.aiKnowledgeDocs.length;
    const removed = store.aiKnowledgeDocs.find((row) => row.id === docId);
    store.aiKnowledgeDocs = store.aiKnowledgeDocs.filter((row) => row.id !== docId);
    if (before === store.aiKnowledgeDocs.length) {
      json(res, 404, { error: 'Knowledge document not found.' });
      return;
    }

    addActivity(store, {
      actor: auth.session.username,
      role: auth.session.role,
      action: 'ai.knowledge.delete',
      entity: 'ai',
      entityId: docId,
      details: `Deleted knowledge document "${clean(removed?.title) || docId}"`
    });
    await writeStore(store);
    json(res, 200, { ok: true });
    return;
  }

  if (pathname === '/api/ops/tasks/from-ai' && req.method === 'POST') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const payload = await readBody(req, 800_000);
      const source = clean(payload.source).toLowerCase();
      const actor = auth.session.username;

      let templates = [];
      let sourceSummary = '';
      if (source === 'payment-risk') {
        const report = localPaymentRiskReport(store, {
          period: clean(payload.period) || 'week',
          from: clean(payload.from),
          to: clean(payload.to)
        });
        templates = buildOpsTasksFromPaymentRisk(report, actor);
        sourceSummary = `payment-risk (${report.summary?.overallRisk || 'unknown'} risk)`;
      } else if (source === 'qc-intelligence') {
        const qcRecordId = clean(payload.qcRecordId);
        const limitRaw = Number(payload.limit);
        const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(120, Math.floor(limitRaw))) : 30;
        const allQc = [...(store.produce || [])].sort((a, b) => dateMs(b.createdAt) - dateMs(a.createdAt));
        const selected = qcRecordId ? allQc.filter((row) => row.id === qcRecordId) : allQc.slice(0, limit);
        const insights = selected.map((row) => localQcIntelligence(row));
        templates = buildOpsTasksFromQcInsights(insights, actor);
        sourceSummary = `qc-intelligence (${insights.length} lot(s))`;
      } else if (source === 'custom') {
        const tasks = Array.isArray(payload.tasks) ? payload.tasks : [];
        templates = tasks.map((task) => ({
          title: task?.title,
          description: task?.description,
          severity: task?.severity,
          sourceType: clean(task?.sourceType) || 'ai-custom',
          sourceRef: clean(task?.sourceRef),
          assignedTo: clean(task?.assignedTo),
          notes: clean(task?.notes),
          createdBy: actor
        }));
        sourceSummary = 'custom';
      } else {
        json(res, 422, { error: 'source is required and must be one of: payment-risk, qc-intelligence, custom' });
        return;
      }

      const created = templates.map((row) => createOpsTask(store, row));
      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ops.tasks.create_from_ai',
        entity: 'ai',
        details: `Created ${created.length} ops task(s) from ${sourceSummary}`
      });
      await writeStore(store);
      json(res, 201, {
        data: created,
        meta: {
          source,
          createdCount: created.length
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/ops/tasks' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    let rows = [...store.opsTasks];
    rows = filterByQuery(rows, reqUrl.searchParams, ['title', 'description', 'sourceType', 'sourceRef', 'assignedTo', 'status']);
    const status = clean(reqUrl.searchParams.get('status')).toLowerCase();
    const severity = clean(reqUrl.searchParams.get('severity')).toLowerCase();
    if (status) rows = rows.filter((row) => normalizeOpsTaskStatus(row.status) === normalizeOpsTaskStatus(status));
    if (severity) rows = rows.filter((row) => normalizeRiskSeverity(row.severity, '') === normalizeRiskSeverity(severity, ''));
    rows.sort((a, b) => dateMs(b.updatedAt || b.createdAt) - dateMs(a.updatedAt || a.createdAt));
    const limit = safeQueryInt(reqUrl.searchParams, 'limit', 200, 2000);
    const offset = safeQueryInt(reqUrl.searchParams, 'offset', 0, 10000);
    const total = rows.length;
    const counts = rows.reduce(
      (acc, row) => {
        const key = normalizeOpsTaskStatus(row.status);
        acc[key] = (acc[key] || 0) + 1;
        return acc;
      },
      { open: 0, in_progress: 0, resolved: 0, closed: 0 }
    );
    json(res, 200, {
      data: rows.slice(offset, offset + limit),
      meta: {
        total,
        limit,
        offset,
        counts
      }
    });
    return;
  }

  if (pathname.startsWith('/api/ops/tasks/') && req.method === 'PATCH') {
    try {
      const store = await readStore();
      const auth = requireSession(req, store, ['admin', 'agent']);
      if (!auth.ok) {
        json(res, auth.status, { error: auth.error });
        return;
      }

      const taskId = routeParam(pathname, '/api/ops/tasks/');
      const task = findById(store.opsTasks, taskId);
      if (!task) {
        json(res, 404, { error: 'Task not found.' });
        return;
      }

      const payload = await readBody(req, 120_000);
      if (payload.title !== undefined) task.title = clean(payload.title) || task.title;
      if (payload.description !== undefined) task.description = clean(payload.description);
      if (payload.severity !== undefined) task.severity = normalizeRiskSeverity(payload.severity, task.severity || 'medium');
      if (payload.assignedTo !== undefined) task.assignedTo = clean(payload.assignedTo);
      if (payload.notes !== undefined) task.notes = clean(payload.notes);
      if (payload.status !== undefined) {
        task.status = normalizeOpsTaskStatus(payload.status);
      }
      task.updatedAt = nowIso();
      if (task.status === 'resolved' || task.status === 'closed') {
        task.resolvedAt = task.resolvedAt || task.updatedAt;
      } else {
        task.resolvedAt = '';
      }

      addActivity(store, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'ops.task.update',
        entity: 'ai',
        entityId: task.id,
        details: `Updated ops task "${task.title}" (${task.status})`
      });
      await writeStore(store);
      json(res, 200, { data: task });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/reports/summary' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    json(res, 200, { data: makeSummary(store) });
    return;
  }

  if (pathname === '/api/reports/agents' && req.method === 'GET') {
    const store = await readStore();
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
      const store = await readStore();
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
      await writeStore(store);

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
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const range = resolveRange(
      reqUrl.searchParams.get('period'),
      reqUrl.searchParams.get('from'),
      reqUrl.searchParams.get('to')
    );

    const farmerRows = store.farmers.filter((row) => isInRange(row.createdAt, range)).map((row) => {
      const hectares = Number(row.hectares || 0);
      const avocadoHectares = Number(row.avocadoHectares || 0);
      const acres = Number.isFinite(hectares) ? hectares / HECTARES_PER_ACRE : 0;
      const squareFeet = Number.isFinite(hectares) ? hectares * SQFT_PER_HECTARE : 0;
      const avocadoAcres = Number.isFinite(avocadoHectares) ? avocadoHectares / HECTARES_PER_ACRE : 0;
      const avocadoSquareFeet = Number.isFinite(avocadoHectares) ? avocadoHectares * SQFT_PER_HECTARE : 0;

      return {
        ...row,
        preferredLanguage: languageOrDefault(row.preferredLanguage),
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
      { key: 'preferredLanguage', label: 'preferredLanguage' },
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
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const range = resolveRange(
      reqUrl.searchParams.get('period'),
      reqUrl.searchParams.get('from'),
      reqUrl.searchParams.get('to')
    );

    const produceRows = store.produce.filter((row) => isInRange(row.createdAt, range)).map((row) => ({
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
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const range = resolveRange(
      reqUrl.searchParams.get('period'),
      reqUrl.searchParams.get('from'),
      reqUrl.searchParams.get('to')
    );

    const rows = store.producePurchases.filter((row) => isInRange(row.createdAt, range));

    const csvData = toCsv(rows, [
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
    const store = await readStore();
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
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const range = resolveRange(
      reqUrl.searchParams.get('period'),
      reqUrl.searchParams.get('from'),
      reqUrl.searchParams.get('to')
    );

    const rows = store.payments.filter((row) => isInRange(row.createdAt, range));

    const csvData = toCsv(rows, [
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

  if (pathname === '/api/exports/sms.csv' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin', 'agent']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const range = resolveRange(
      reqUrl.searchParams.get('period'),
      reqUrl.searchParams.get('from'),
      reqUrl.searchParams.get('to')
    );

    let smsRows = [...store.smsLogs];
    smsRows = smsRows.filter((row) => isInRange(row.createdAt, range));

    smsRows = smsRows.map((row) => ({
      ...row,
      ownerCostKes: smsOwnerCostKes(row),
      billable: smsOwnerCostKes(row) > 0 ? 'yes' : 'no'
    }));

    const csvData = toCsv(smsRows, [
      { key: 'id', label: 'id' },
      { key: 'farmerId', label: 'farmerId' },
      { key: 'farmerName', label: 'farmerName' },
      { key: 'phone', label: 'phone' },
      { key: 'message', label: 'message' },
      { key: 'provider', label: 'provider' },
      { key: 'status', label: 'status' },
      { key: 'ownerCostKes', label: 'ownerCostKes' },
      { key: 'billable', label: 'billable' },
      { key: 'createdBy', label: 'createdBy' },
      { key: 'createdAt', label: 'createdAt' }
    ]);
    csv(res, 'sms.csv', csvData);
    return;
  }

  if (pathname === '/api/exports/activity.csv' && req.method === 'GET') {
    const store = await readStore();
    const auth = requireSession(req, store, ['admin']);
    if (!auth.ok) {
      json(res, auth.status, { error: auth.error });
      return;
    }

    const range = resolveRange(
      reqUrl.searchParams.get('period'),
      reqUrl.searchParams.get('from'),
      reqUrl.searchParams.get('to')
    );

    const rows = store.activityLogs.filter((row) => isInRange(row.at, range));

    const csvData = toCsv(rows, [
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
    const store = await readStore();
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
    await writeStore(store);

    json(res, 201, {
      data: {
        filename: snapshot.filename,
        createdAt: snapshot.createdAt
      }
    });
    return;
  }

  if (pathname === '/api/admin/backups' && req.method === 'GET') {
    const store = await readStore();
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

    json(res, 200, {
      data: files,
      meta: {
        storageBackend: STORAGE_BACKEND,
        backupIntervalHours: BACKUP_INTERVAL_HOURS,
        backupRetentionDays: BACKUP_RETENTION_DAYS
      }
    });
    return;
  }

  if (pathname === '/api/admin/restore' && req.method === 'POST') {
    try {
      const store = await readStore();
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

      const filename = path.basename(clean(payload.filename || ''));
      if (!filename || !filename.endsWith('.json')) {
        json(res, 422, { error: 'filename is required and must end with .json' });
        return;
      }

      const filepath = path.join(BACKUP_DIR, filename);
      if (!fs.existsSync(filepath)) {
        json(res, 404, { error: 'Backup file not found' });
        return;
      }

      const raw = fs.readFileSync(filepath, 'utf-8');
      const restored = normalizeStore(JSON.parse(raw));
      const preRestore = createBackupSnapshot(store, 'pre-restore');

      const currentToken = auth.session.token;
      if (currentToken && store.sessions[currentToken]) {
        restored.sessions[currentToken] = store.sessions[currentToken];
      }

      addActivity(restored, {
        actor: auth.session.username,
        role: auth.session.role,
        action: 'admin.restore',
        entity: 'backup',
        entityId: filename,
        details: `Restored dataset from ${filename}. Pre-restore snapshot: ${preRestore.filename}`
      });
      await writeStore(restored);

      json(res, 200, {
        data: {
          restored: true,
          fromBackup: filename,
          preRestoreBackup: preRestore.filename,
          farmers: restored.farmers.length,
          produce: restored.produce.length,
          producePurchases: restored.producePurchases.length,
          payments: restored.payments.length,
          smsLogs: restored.smsLogs.length
        }
      });
    } catch (error) {
      json(res, 400, { error: error.message });
    }
    return;
  }

  if (pathname === '/api/admin/reset' && req.method === 'POST') {
    try {
      const store = await readStore();
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
      await writeStore(store);

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
    const store = await readStore();
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
    await writeStore(store);

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

async function shutdownStorage() {
  if (!sqlStoreAdapter) return;
  try {
    await sqlStoreAdapter.close();
  } catch (error) {
    console.error(`[storage] Failed to close ${sqlStoreAdapter.backend} adapter: ${error.message}`);
  } finally {
    sqlStoreAdapter = null;
  }
}

let shuttingDown = false;
async function handleShutdown(signal) {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log(`[system] ${signal} received, shutting down gracefully...`);
  try {
    if (backupScheduleHandle) {
      clearInterval(backupScheduleHandle);
      backupScheduleHandle = null;
    }
    if (executiveBriefScheduleHandle) {
      clearInterval(executiveBriefScheduleHandle);
      executiveBriefScheduleHandle = null;
    }
    if (msaidiziScheduleHandle) {
      clearInterval(msaidiziScheduleHandle);
      msaidiziScheduleHandle = null;
    }
    await shutdownStorage();
  } finally {
    process.exit(0);
  }
}

process.on('SIGINT', () => {
  handleShutdown('SIGINT');
});
process.on('SIGTERM', () => {
  handleShutdown('SIGTERM');
});

async function startServer() {
  try {
    await ensureStore();
    pruneOldBackups();
    scheduleAutomaticBackups();
    scheduleExecutiveBriefs();
    scheduleMsaidiziSync();
    await runScheduledMsaidiziSync();

    server.listen(PORT, HOST, () => {
      console.log(`Agem MVP server running on http://${HOST}:${PORT}`);
      console.log(
        `[storage] backend=${STORAGE_BACKEND}, backupIntervalHours=${BACKUP_INTERVAL_HOURS}, backupRetentionDays=${BACKUP_RETENTION_DAYS}, aiExecBriefEnabled=${AI_EXEC_BRIEF_ENABLED}, aiExecBriefIntervalHours=${AI_EXEC_BRIEF_INTERVAL_HOURS}, msaidiziAutoSync=${AI_MSAIDIZI_SYNC_ENABLED}, msaidiziSyncHours=${AI_MSAIDIZI_SYNC_INTERVAL_HOURS}`
      );
    });
  } catch (error) {
    console.error(`[startup] Failed to initialize storage: ${error.message}`);
    process.exit(1);
  }
}

startServer();
