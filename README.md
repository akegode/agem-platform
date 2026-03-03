# Agem Portal (Complete MVP)

This project is now a full working web application with authentication, role-based permissions, operational forms, reports, exports, and admin tools.

## What is implemented
- Secure sign-in sessions with private staff accounts (admin, agent roles)
- Farmer self-registration from the sign-in card (creates farmer account + farmer profile)
- Admin-only agent account management (enter name + email, auto-generate temporary username/password, and list multiple agent logins)
- Signed-in password change flow
- Recovery endpoints for forgotten username/password via role recovery codes
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
- Safaricom-ready USSD callback endpoint (`/api/ussd/callback`) with language selection (English/Kiswahili), self-registration for new phone numbers, and menu for payment status, latest QC, profile, and help (via gateway such as Africa's Talking)
- Operational KPI dashboard and agent performance report
- CSV exports (farmers, produce, payments, activity)
- Backup snapshots + backup listing
- Admin reset workflow with pre-reset backup
- Offline fallback mode when backend is unavailable

## Project structure
- Frontend:
  - `frontend/index.html`
  - `frontend/styles.css`
  - `frontend/app.js`
- Backend:
  - `backend/server.js`
  - `backend/data/store.json`
  - `backend/data/backups/`
- Tests:
  - `scripts/test.js`

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

## Deploy to your website (Render + subdomain)
This is the easiest path if you are new to deployment.

1. Push this folder to a GitHub repository.
2. In Render, create a **Blueprint** service from that GitHub repo.
3. Render will read `render.yaml` and deploy automatically.
4. In Render service settings, set environment variables:
   - `ALLOW_DEMO_USERS=false`
   - `ALLOW_FARMER_REGISTRATION=true` (or `false` to disable public farmer signup)
   - `ADMIN_USERNAME` + `ADMIN_PASSWORD`
   - `AGENT_USERNAME` + `AGENT_PASSWORD`
   - `ADMIN_RECOVERY_CODE` + `AGENT_RECOVERY_CODE`
   - `USSD_ENABLED=true`
   - `USSD_CODE=*483#`
   - `USSD_HELP_PHONE=+2547xxxxxxxx`
   - `USSD_SHARED_SECRET=<strong-random-secret>`
5. In Render service settings, add a custom domain like `portal.yourdomain.com`.
6. In your DNS provider, add the DNS record Render asks for.
7. In your existing website, add a button/link to `https://portal.yourdomain.com`.

Your current website stays as-is, and this app runs as a portal under a subdomain.

## Important notes
- Data is stored locally in `backend/data/store.json`.
- Backups are saved in `backend/data/backups/`.
- This is implemented with Node.js built-ins only (no external packages).
- Excel import parsing is handled in-browser via the `xlsx` CDN script; CSV import works without it.
- Default demo credentials are disabled when `ALLOW_DEMO_USERS=false` (recommended for production).
- Farmer self-registration can be turned off by setting `ALLOW_FARMER_REGISTRATION=false`.
- Admins can create additional agent accounts from the in-app **Agents** section by entering agent name and email. The portal auto-generates temporary login credentials, and env `AGENT_*` remains the initial bootstrap agent.
- For live Safaricom USSD, use a gateway provider (e.g., Africa's Talking), point callback URL to `https://portal.agemlimited.com/api/ussd/callback?secret=<same-secret>` and events URL to `https://portal.agemlimited.com/api/ussd/events?secret=<same-secret>`.
- USSD self-registration collects full name, National ID, location, total farm size (acres), and area under avocado (acres), then creates a farmer profile and sends a confirmation SMS in the selected language.
- If you later install PHP/Composer/MySQL/Redis, this can be migrated to Laravel + React/Inertia exactly as planned.
