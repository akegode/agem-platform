# Agem Portal (Complete MVP)

This project is now a full working web application with authentication, role-based permissions, operational forms, reports, exports, and admin tools.

## What is implemented
- Secure sign-in sessions with private staff accounts (admin, agent roles)
- Farmer self-registration from the sign-in card (creates farmer account + farmer profile)
- Admin-only agent account management (enter name + email + phone, auto-generate temporary username/password, and list multiple agent logins)
- Signed-in password change flow
- Recovery endpoints for forgotten username/password for admin and agent accounts via role recovery codes
- Farmer management (create, edit, view, delete with safeguards) with mandatory National ID, total farm size + area under avocado, and auto-conversion across hectares, acres, and square feet
- Bulk farmer import (CSV/Excel) for admin; required fields (including National ID, total area, and area under avocado) are mapped, extra columns ignored, acreage/square-feet values are converted to hectares, and optional onboarding SMS can be sent to imported farmers
- Farmer preferred language support (English/Kiswahili) for onboarding and USSD engagement
- Duplicate protection for admin/agent/farmer accounts and farmer phone records (no double registration)
- Import duplicate handling: choose to keep existing farmer data or overwrite with incoming file data
- Farm-gate avocado QC capture (Hass/Fuerte): visual grade, dry matter, firmness, fruit weight, size code, lot weight, and QC decision
- Separate produce purchasing log linked to farmer (and optional QC lot): purchased kg, price/kg, value, and buyer
- Purchase-to-payment workflow: owed balances, day/week/month owed filtering, bulk settlement for selected farmers, and owed CSV export
- Payment logging and payment status management
- M-PESA disbursement simulation endpoint
- SMS sending simulation + message logs
- Bulk SMS for admin (all farmers or selected recipients) with searchable recipient picker
- SMS cost tracking for code-owner spend (total + rolling 24-hour spend in dashboard/reports)
- AI Phase 2A: QC Intelligence assistant (lot risk scoring + pass/hold/reject guidance) and Payment Risk Check assistant (duplicate ref/outlier/pending risk detection)
- Safaricom-ready USSD callback endpoint (`/api/ussd/callback`) with language selection (English/Kiswahili), self-registration for new phone numbers, and menu for payment status, latest QC, profile, and help (via gateway such as Africa's Talking)
- Operational KPI dashboard and agent performance report
- CSV exports (farmers, produce, payments, activity)
- Backup snapshots + backup listing + backup restore endpoint
- SQL-backed persistence (PostgreSQL or MySQL) with automatic migration from existing `store.json`
- Scheduled automatic backups with retention policy
- Admin reset workflow with pre-reset backup
- Offline fallback mode when backend is unavailable

## Project structure
- Frontend:
  - `frontend/index.html`
  - `frontend/styles.css`
  - `frontend/app.js`
- Backend:
  - `backend/server.js`
  - `backend/data/store.json` (legacy/local JSON mode only)
  - `backend/data/backups/`
- Tests:
  - `scripts/test.js`
  - `scripts/restore-test.js`

## Run the site (simple)
1. Open terminal in this folder.
2. Set private staff credentials (required):
   ```bash
   export ADMIN_USERNAME="your-admin-username"
   export ADMIN_PASSWORD="your-strong-admin-password"
   export AGENT_USERNAME="your-bootstrap-agent-username"
   export AGENT_PASSWORD="your-bootstrap-agent-password"
   export ADMIN_RECOVERY_CODE="your-admin-recovery-code"
   export AGENT_RECOVERY_CODE="your-agent-recovery-code"
   export ALLOW_DEMO_USERS="false"
   export ALLOW_FARMER_REGISTRATION="true"
   export USSD_ENABLED="true"
   export USSD_CODE="*483#"
   export USSD_HELP_PHONE="+254700000000"
   export USSD_SHARED_SECRET="set-a-secret-used-by-your-ussd-gateway"
   export SMS_OWNER_COST_PER_MESSAGE_KES="0.25"
   export STORAGE_BACKEND="postgres"   # postgres | mysql | json
   export DATABASE_URL="postgres://user:password@host:5432/dbname"
   export BACKUP_INTERVAL_HOURS="24"
   export BACKUP_RETENTION_DAYS="30"
   ```
3. Run:
   ```bash
   npm run dev
   ```
4. Open:
   [http://127.0.0.1:8080](http://127.0.0.1:8080)
5. Sign in with the private usernames/passwords you set above.
6. New farmers can register from the left account panel using **New farmer? Register here**.

## One-command verification
Run:
```bash
npm test
```

This runs end-to-end checks for auth, permissions, CRUD, reports, exports, backup, and reset.

Run storage restore verification:
```bash
npm run test:restore
```

This verifies backup -> reset -> restore end-to-end using the currently configured storage backend.

## Deploy to your website (Render + subdomain)
This is the easiest path if you are new to deployment.

1. Push this folder to a GitHub repository.
2. In Render, create a **Blueprint** service from that GitHub repo.
3. Render will read `render.yaml` and deploy automatically.
4. In Render service settings, set environment variables:
   - `STORAGE_BACKEND=postgres`
   - `DATABASE_URL=<from Render PostgreSQL connection string>`
   - `BACKUP_INTERVAL_HOURS=24`
   - `BACKUP_RETENTION_DAYS=30`
   - `ALLOW_DEMO_USERS=false`
   - `ALLOW_FARMER_REGISTRATION=true` (or `false` to disable public farmer signup)
   - `ADMIN_USERNAME` + `ADMIN_PASSWORD`
   - `AGENT_USERNAME` + `AGENT_PASSWORD`
   - `ADMIN_RECOVERY_CODE` + `AGENT_RECOVERY_CODE`
   - `USSD_ENABLED=true`
   - `USSD_CODE=*483#`
   - `USSD_HELP_PHONE=+2547xxxxxxxx`
   - `USSD_SHARED_SECRET=<strong-random-secret>`
   - `SMS_OWNER_COST_PER_MESSAGE_KES=0.25`
5. In Render service settings, add a custom domain like `portal.yourdomain.com`.
6. In your DNS provider, add the DNS record Render asks for.
7. In your existing website, add a button/link to `https://portal.yourdomain.com`.

Your current website stays as-is, and this app runs as a portal under a subdomain.

## Important notes
- Data can be stored in PostgreSQL/MySQL (`STORAGE_BACKEND=postgres|mysql`) or JSON (`STORAGE_BACKEND=json` for local fallback).
- If SQL storage is enabled and a legacy `backend/data/store.json` exists, the server seeds SQL from that file on first run.
- Backups are saved in `backend/data/backups/` and can be restored with `POST /api/admin/restore`.
- Automatic backup schedule is controlled by `BACKUP_INTERVAL_HOURS` and backup cleanup by `BACKUP_RETENTION_DAYS`.
- SQL backends use `pg` and `mysql2`.
- Excel import parsing is handled in-browser via the `xlsx` CDN script; CSV import works without it.
- Default demo credentials are disabled when `ALLOW_DEMO_USERS=false` (recommended for production).
- Farmer self-registration can be turned off by setting `ALLOW_FARMER_REGISTRATION=false`.
- Admins can create additional agent accounts from the in-app **Agents** section by entering agent name and email. The portal auto-generates temporary login credentials, and env `AGENT_*` remains the initial bootstrap agent.
- For live Safaricom USSD, use a gateway provider (e.g., Africa's Talking), point callback URL to `https://portal.agemlimited.com/api/ussd/callback?secret=<same-secret>` and events URL to `https://portal.agemlimited.com/api/ussd/events?secret=<same-secret>`.
- USSD self-registration collects full name, National ID, location, total farm size (acres), and area under avocado (acres), then creates a farmer profile and sends a confirmation SMS in the selected language.
- SMS spend metrics treat successful sends as billable and compute owner spend using `SMS_OWNER_COST_PER_MESSAGE_KES` (default `0.25`).
- If you later install PHP/Composer/MySQL/Redis, this can be migrated to Laravel + React/Inertia exactly as planned.
