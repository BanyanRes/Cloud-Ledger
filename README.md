# CloudLedger - Multi-Entity Cloud Accounting

A full-stack accounting application supporting 40+ entities with double-entry bookkeeping, financial reporting, and bank reconciliation. Built for small teams who need real multi-user access.

## Features

- **Multi-entity**: Independent chart of accounts, journal entries, and reports per entity
- **Multi-user**: Role-based access (Admin, Accountant, Viewer) with JWT authentication
- **Double-entry bookkeeping**: Balanced journal entries enforced at entry time
- **Financial reports**: Trial Balance, Balance Sheet, Income Statement
- **General Ledger**: Running balance per account with transaction history
- **Bank reconciliation**: Statement matching, cleared/outstanding tracking, reconciliation history
- **Bulk import**: Add dozens of entities at once via CSV format

## Tech Stack

- **Backend**: Node.js + Express + SQLite (via better-sqlite3)
- **Frontend**: React 18 + Vite
- **Auth**: bcrypt + JSON Web Tokens
- **Database**: SQLite with WAL mode (handles concurrent reads well)

## Quick Start (Local Development)

```bash
# 1. Clone and install
git clone <your-repo-url>
cd cloudledger
npm run setup

# 2. Create .env file
cp .env.example .env
# Edit .env and set a real JWT_SECRET

# 3. Start in development mode
npm run dev

# Frontend: http://localhost:5173
# Backend:  http://localhost:3000
```

Default login: `admin@company.com` / `admin`

---

## Deploy to the Cloud (Recommended Options)

### Option A: Railway (Easiest - ~2 minutes)

1. Push this project to a **GitHub repo**
2. Go to [railway.app](https://railway.app) and sign in with GitHub
3. Click **"New Project"** → **"Deploy from GitHub Repo"**
4. Select your repo
5. Add these **environment variables** in Railway dashboard:
   - `JWT_SECRET` = (generate a random string, e.g. `openssl rand -hex 32`)
   - `PORT` = `3000`
   - `NODE_ENV` = `production`
6. Railway auto-detects the Dockerfile and deploys
7. Click **"Generate Domain"** to get your public URL

**Important for Railway**: SQLite data lives on the container filesystem. To persist data across deploys, add a **Railway Volume**:
- Go to your service → Settings → Volumes
- Mount path: `/app/data`
- This keeps your database safe across deploys and restarts

### Option B: Render

1. Push to GitHub
2. Go to [render.com](https://render.com) → New → **Web Service**
3. Connect your repo
4. Settings:
   - **Runtime**: Docker
   - **Instance Type**: Starter ($7/mo) or Free
5. Add environment variables (same as Railway above)
6. Add a **Disk** mounted at `/app/data` for SQLite persistence

### Option C: Any VPS (DigitalOcean, Linode, AWS Lightsail)

```bash
# On your server:
git clone <your-repo> cloudledger
cd cloudledger
cp .env.example .env
nano .env   # Set JWT_SECRET and NODE_ENV=production

npm run setup
npm start

# Runs on port 3000. Put behind nginx or caddy for HTTPS.
```

Recommended: use **pm2** to keep it running:
```bash
npm install -g pm2
pm2 start server/index.js --name cloudledger
pm2 save
pm2 startup
```

---

## Storage & Memory Considerations

SQLite is more than sufficient for your use case:

| Scenario | Estimated DB Size |
|----------|------------------|
| 40 entities × 100 JEs each | ~2 MB |
| 40 entities × 1,000 JEs each | ~15 MB |
| 40 entities × 10,000 JEs each | ~120 MB |
| 40 entities × 50,000 JEs each | ~500 MB |

SQLite handles databases up to **281 TB** and can process thousands of transactions per second. For a team of 5-20 users doing normal accounting work, it will never be the bottleneck.

**Backup**: The entire database is one file (`data/cloudledger.db`). Copy it to back up everything.

### When to consider PostgreSQL

If you eventually need:
- 50+ concurrent users writing at the same time
- Multiple application servers (horizontal scaling)
- Point-in-time recovery / replication

The migration path is straightforward - the Express API stays the same, only the database driver changes.

---

## Project Structure

```
cloudledger/
├── server/
│   └── index.js          # Express API + SQLite database
├── client/
│   ├── src/
│   │   ├── App.jsx       # Full React frontend
│   │   ├── api.js        # API client
│   │   └── main.jsx      # Entry point
│   ├── index.html
│   └── vite.config.js
├── data/                  # SQLite database (auto-created)
├── Dockerfile
├── railway.json
├── package.json
└── .env.example
```

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | /api/auth/login | Login |
| POST | /api/auth/signup | Create account |
| GET | /api/entities | List all entities |
| POST | /api/entities | Create entity |
| POST | /api/entities/bulk | Bulk create entities |
| GET | /api/entities/:id/accounts | Get chart of accounts |
| POST | /api/entities/:id/accounts | Add account |
| GET | /api/entities/:id/entries | Get journal entries |
| POST | /api/entities/:id/entries | Post journal entry |
| GET | /api/entities/:id/balances | Get account balances |
| GET | /api/entities/:id/reconciliations | Get reconciliation history |
| POST | /api/entities/:id/reconciliations | Finalize reconciliation |
| GET | /api/summary | Org-wide entity summary |

## Security Notes

- Change `JWT_SECRET` in production (use `openssl rand -hex 32`)
- Passwords are hashed with bcrypt (10 rounds)
- CORS and Helmet enabled
- Consider adding HTTPS via reverse proxy (nginx/caddy) or your hosting platform
