# Agem Avocado Platform (Complete MVP)

This project is now a full working web application with authentication, role-based permissions, operational forms, reports, exports, and admin tools.

## What is implemented
- Secure sign-in sessions (admin, agent, farmer roles)
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
2. Run:
   ```bash
   npm run dev
   ```
3. Open:
   [http://127.0.0.1:8080](http://127.0.0.1:8080)
4. Sign in with one of these demo accounts:
   - `admin / admin123`
   - `agent / agent123`
   - `farmer / farmer123`

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
4. In Render service settings, add a custom domain like `portal.yourdomain.com`.
5. In your DNS provider, add the DNS record Render asks for.
6. In your existing website, add a button/link to `https://portal.yourdomain.com`.

Your current website stays as-is, and this app runs as a portal under a subdomain.

## Important notes
- Data is stored locally in `backend/data/store.json`.
- Backups are saved in `backend/data/backups/`.
- This is implemented with Node.js built-ins only (no external packages).
- If you later install PHP/Composer/MySQL/Redis, this can be migrated to Laravel + React/Inertia exactly as planned.
