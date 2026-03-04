# Storage Hardening Guide

## 1) Pick storage backend

The API now supports three storage backends:

- `postgres` (recommended for production)
- `mysql`
- `json` (local/dev fallback only)

Set with:

```bash
export STORAGE_BACKEND="postgres"
export DATABASE_URL="postgres://user:password@host:5432/dbname"
```

For MySQL:

```bash
export STORAGE_BACKEND="mysql"
export DATABASE_URL="mysql://user:password@host:3306/dbname"
```

If `STORAGE_BACKEND` is not set, the server auto-detects from `DATABASE_URL`; otherwise it falls back to `json`.

## 2) Migration behavior

On first SQL startup, if no SQL store row exists:

1. The server attempts to seed from `backend/data/store.json`.
2. If no JSON file exists, a fresh empty store is created.
3. A `migration-seed` backup snapshot is written in `backend/data/backups/`.

## 3) Backup schedule and retention

Configure:

```bash
export BACKUP_INTERVAL_HOURS="24"
export BACKUP_RETENTION_DAYS="30"
```

Behavior:

- Scheduled snapshots are written to `backend/data/backups/`.
- Backups older than retention days are pruned automatically.
- Manual backups are available via `POST /api/admin/backup`.

## 4) Restore flow

API endpoint (admin only):

```http
POST /api/admin/restore
Content-Type: application/json
Authorization: Bearer <token>

{
  "filename": "2026-03-03T23-37-52-097Z-manual.json",
  "confirm": true
}
```

Safety behavior:

- A `pre-restore` snapshot is automatically created before restoring.
- Current admin session is preserved.

## 5) Restore verification test

Run:

```bash
npm run test:restore
```

This performs:

1. Create farmer
2. Create backup
3. Reset data
4. Restore backup
5. Verify farmer returns

## 6) Operational checklist

- Keep `STORAGE_BACKEND=postgres` (or `mysql`) in production.
- Keep `ALLOW_DEMO_USERS=false` in production.
- Review backups regularly from `/api/admin/backups`.
- Run `npm run test:restore` after major release changes.
