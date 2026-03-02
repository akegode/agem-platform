# Agem Avocado Platform (Complete MVP)

This project is now a full working web application with authentication, role-based permissions, operational forms, reports, exports, and admin tools.

## What is implemented
- Secure sign-in sessions with private staff accounts (admin, agent roles)
- Farmer management (create, edit, view, delete with safeguards)
- Produce collection tracking
- Payment logging and payment status management
- M-PESA disbursement simulation endpoint
- SMS sending simulation + message logs
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
   export AGENT_USERNAME="your-agent-username"
   export AGENT_PASSWORD="your-strong-agent-password"
   export ALLOW_DEMO_USERS="false"
   ```
3. Run:
   ```bash
   npm run dev
   ```
4. Open:
   [http://127.0.0.1:8080](http://127.0.0.1:8080)
5. Sign in with the private usernames/passwords you set above.

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
   - `ADMIN_USERNAME` + `ADMIN_PASSWORD`
   - `AGENT_USERNAME` + `AGENT_PASSWORD`
5. In Render service settings, add a custom domain like `portal.yourdomain.com`.
6. In your DNS provider, add the DNS record Render asks for.
7. In your existing website, add a button/link to `https://portal.yourdomain.com`.

Your current website stays as-is, and this app runs as a portal under a subdomain.

## Important notes
- Data is stored locally in `backend/data/store.json`.
- Backups are saved in `backend/data/backups/`.
- This is implemented with Node.js built-ins only (no external packages).
- Default demo credentials are disabled when `ALLOW_DEMO_USERS=false` (recommended for production).
- If you later install PHP/Composer/MySQL/Redis, this can be migrated to Laravel + React/Inertia exactly as planned.
