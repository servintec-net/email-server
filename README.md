# Server

Express + WebSocket backend for the email application. Handles Microsoft OAuth, mail sync, categorization (OpenAI), and real-time updates via WebSocket.

## Prerequisites

- Node.js (v18+)
- MySQL
- `.env` configured (see below)

## Environment

Create a `.env` file in the `server` directory. Required variables:

| Variable | Description |
|----------|-------------|
| `PORT` | HTTP/WS port (default: `4000`) |
| `APP_JWT_SECRET` | Secret for JWT signing (required) |
| `JWT_EXPIRES_IN` | JWT expiry (default: `7d`) |
| `CLIENT_ID` | Microsoft OAuth app client ID |
| `CLIENT_SECRET` | Microsoft OAuth app client secret |
| `AUTHORITY` | Microsoft authority URL (default: consumers) |
| `REDIRECT_URI` | OAuth redirect URI |
| `FRONTEND` | Frontend origin URL |
| `MYSQL_HOST` | MySQL host |
| `MYSQL_USER` | MySQL user |
| `MYSQL_PASS` | MySQL password |
| `MYSQL_DB` | MySQL database name |
| `OPENAI_API_KEY` | OpenAI API key (for categorization) |

Optional: `LOG_THROTTLE=0` to disable throttle logging.

## Development

```bash
npm install
npm start
```

Server runs on `http://localhost:4000` (or your `PORT`). WebSocket path: `/ws`.

---

## Production with PM2

[PM2](https://pm2.keymetrics.io/) keeps the Node process running, restarts on crash, and can manage logs and clustering.

### Install PM2

```bash
npm install -g pm2
```

### Run with PM2

**Single instance (simple):**

```bash
pm2 start server.js --name "email-server"
```

**With ecosystem file (recommended):**

Create `ecosystem.config.cjs` in the `server` directory:

```javascript
module.exports = {
  apps: [
    {
      name: "email-server",
      script: "server.js",
      cwd: __dirname,
      instances: 1,
      exec_mode: "fork",
      env: { NODE_ENV: "development" },
      env_production: { NODE_ENV: "production" },
      max_memory_restart: "500M",
      error_file: "logs/err.log",
      out_file: "logs/out.log",
      merge_logs: true,
      time: true,
    },
  ],
};
```

Then:

```bash
# Create logs dir if needed
mkdir -p logs

# Start in production
pm2 start ecosystem.config.cjs --env production

# Or start and save process list so it restarts on reboot
pm2 start ecosystem.config.cjs --env production
pm2 save
pm2 startup
```

### PM2 commands

| Command | Description |
|---------|-------------|
| `pm2 list` | List processes |
| `pm2 logs email-server` | Stream logs |
| `pm2 restart email-server` | Restart app |
| `pm2 stop email-server` | Stop app |
| `pm2 delete email-server` | Remove from PM2 |
| `pm2 monit` | Real-time monitor |
| `pm2 save` | Save process list |
| `pm2 startup` | Generate startup script (run after `pm2 save`) |

### Clustering (optional)

For multiple CPU cores, set `instances` to a number or `"max"` and `exec_mode: "cluster"`. Note: WebSocket state is in-memory per process; for shared WS state you’d need a store (e.g. Redis) or keep `instances: 1` with WebSockets.

---

## API / WebSocket

- **HTTP**: REST endpoints for auth, mail, folders, etc.
- **WebSocket** (`/ws`): Connect with `?` or send `{ type: "auth", token: "<jwt>" }` after connect for real-time push.
