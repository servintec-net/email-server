require("dotenv").config();
const http = require("http");
const express = require("express");
const { WebSocketServer } = require("ws");
const axios = require("axios");
const cors = require("cors");
const cron = require("node-cron");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");

const {
    SCOPES,
    CLIENT_ID,
    CLIENT_SECRET,
    AUTHORITY,
    REDIRECT_URI,
    FRONTEND,
    POOL,
    APP_JWT_SECRET,
    JWT_EXPIRES_IN
} = require("./constants");

const JWT_SECRET = APP_JWT_SECRET;
const { runCategorizer } = require("./categorizer");
const {
    collapseToLatestPerConversation,
    htmlToText,
    getFirstNameFromSender,
    pickGreeting,
    pickSignoff,
} = require("./helper");
const OpenAI = require("openai");
const { getValidAccessToken, getMailboxRowOrThrow } = require("./token");

const app = express();
app.use(cors());
app.use(express.json());

const server = http.createServer(app);

// WebSocket: push folder-count and folder-update events (move to inbox, new mail via periodic invalidation)
const wsClientsByUserId = new Map();

const wss = new WebSocketServer({ server, path: "/ws" });
wss.on("connection", (ws) => {
    let userId = null;
    ws.on("message", (raw) => {
        try {
            const msg = JSON.parse(String(raw));
            if (msg.type === "auth" && msg.token) {
                try {
                    const payload = jwt.verify(msg.token, JWT_SECRET);
                    userId = payload.uid;
                    if (!wsClientsByUserId.has(userId)) wsClientsByUserId.set(userId, new Set());
                    wsClientsByUserId.get(userId).add(ws);
                    ws.send(JSON.stringify({ type: "auth", ok: true }));
                } catch {
                    ws.send(JSON.stringify({ type: "auth", ok: false }));
                }
            }
        } catch (_) { }
    });
    ws.on("close", () => {
        if (userId != null && wsClientsByUserId.has(userId)) {
            wsClientsByUserId.get(userId).delete(ws);
            if (wsClientsByUserId.get(userId).size === 0) wsClientsByUserId.delete(userId);
        }
    });
});

function broadcastToUser(userId, payload) {
    const clients = wsClientsByUserId.get(userId);
    if (!clients) return;
    const str = typeof payload === "string" ? payload : JSON.stringify(payload);
    for (const ws of clients) {
        if (ws.readyState === 1) try { ws.send(str); } catch (_) { }
    }
}

// Cache key format: `${userId}::${mailboxId}::${folderId}` to ensure uniqueness across users and mailboxes
const folderCountCache = new Map();
const FOLDER_COUNT_CACHE_TTL = 5 * 60 * 1000; // 5 minutes (increased from 30 seconds)

// Request queue to prevent concurrent folder count requests for the same mailbox
const folderCountRequestQueue = new Map(); // key: `${userId}::${mailboxId}` -> Promise
const folderIdCache = new Map();
const FOLDER_ID_CACHE_TTL = 10 * 60 * 60 * 1000;

// Concurrency limiter for Microsoft Graph API
// Microsoft Graph API has a HARD LIMIT of 4 concurrent requests per mailbox
// This limit cannot be increased and applies to all Outlook/mail operations
// See: https://learn.microsoft.com/en-us/graph/throttling-limits
// When exceeded, you get: "Application is over its MailboxConcurrency limit."
const GRAPH_CONCURRENCY_LIMIT = 4;
const graphRequestQueues = new Map(); // key: `${userId}::${mailboxId}` -> Array of pending requests

// Helper to limit concurrent Graph API requests per mailbox
async function executeWithConcurrencyLimit(userId, mailboxId, fn) {
    const key = `${userId}::${mailboxId}`;

    if (!graphRequestQueues.has(key)) {
        graphRequestQueues.set(key, { running: 0, queue: [] });
    }

    const queue = graphRequestQueues.get(key);

    return new Promise((resolve, reject) => {
        queue.queue.push({ fn, resolve, reject });
        processGraphQueue(key);
    });
}

async function processGraphQueue(key) {
    const queue = graphRequestQueues.get(key);

    if (queue.running >= GRAPH_CONCURRENCY_LIMIT || queue.queue.length === 0) {
        return;
    }

    const { fn, resolve, reject } = queue.queue.shift();
    queue.running++;

    try {
        const result = await fn();
        resolve(result);
    } catch (error) {
        reject(error);
    } finally {
        queue.running--;
        processGraphQueue(key);
    }
}

function signAppToken(user) {
    return jwt.sign(
        { uid: user.id, username: user.username, email: user.email },
        JWT_SECRET,
        { expiresIn: JWT_EXPIRES_IN }
    );
}

function requireAuth(req, res, next) {
    const header = req.headers.authorization || "";
    const token = header.startsWith("Bearer ") ? header.slice(7) : null;
    if (!token) return res.status(401).json({ error: "Missing auth token" });

    try {
        const payload = jwt.verify(token, JWT_SECRET);
        req.user = { id: payload.uid, username: payload.username, email: payload.email };
        return next();
    } catch (e) {
        return res.status(401).json({ error: "Invalid/expired auth token" });
    }
}

app.post("/auth/signup", async (req, res) => {
    const { username, email, password } = req.body || {};
    if (!username || !email || !password) return res.status(400).json({ error: "username, email, password required" });

    try {
        const password_hash = await bcrypt.hash(String(password), 12);

        await POOL.query(
            `INSERT INTO users (username, email, password_hash)
       VALUES (?, ?, ?)`,
            [String(username).trim(), String(email).trim().toLowerCase(), password_hash]
        );

        const [rows] = await POOL.query(`SELECT id, username, email FROM users WHERE email = ? LIMIT 1`, [
            String(email).trim().toLowerCase(),
        ]);
        const user = rows[0];
        const token = signAppToken(user);
        res.json({ token, user });
    } catch (err) {
        // duplicate username/email
        if (err?.code === "ER_DUP_ENTRY") {
            return res.status(409).json({ error: "Username or email already exists" });
        }
        console.error("signup error:", err);
        res.status(500).json({ error: "Signup failed" });
    }
});

app.post("/auth/login", async (req, res) => {
    const { emailOrUsername, password } = req.body || {};
    if (!emailOrUsername || !password) return res.status(400).json({ error: "emailOrUsername and password required" });

    try {
        const key = String(emailOrUsername).trim();
        const [rows] = await POOL.query(
            `SELECT id, username, email, password_hash
       FROM users
       WHERE email = ? OR username = ?
       LIMIT 1`,
            [key.toLowerCase(), key]
        );

        const user = rows?.[0];
        if (!user) return res.status(401).json({ error: "Invalid credentials" });

        const ok = await bcrypt.compare(String(password), user.password_hash);
        if (!ok) return res.status(401).json({ error: "Invalid credentials" });

        const token = signAppToken(user);
        res.json({ token, user: { id: user.id, username: user.username, email: user.email } });
    } catch (err) {
        console.error("login error:", err);
        res.status(500).json({ error: "Login failed" });
    }
});

app.get("/me", requireAuth, async (req, res) => {
    try {
        const [[row]] = await POOL.query(
            `SELECT id, username, email FROM users WHERE id = ? LIMIT 1`,
            [req.user.id]
        );
        if (!row) return res.status(404).json({ error: "User not found" });
        res.json({
            id: row.id,
            username: row.username,
            email: row.email,
        });
    } catch (err) {
        console.error("GET /me error:", err);
        res.status(500).json({ error: "Failed to load user" });
    }
});

app.patch("/me/password", requireAuth, async (req, res) => {
    const { currentPassword, newPassword } = req.body || {};
    if (!currentPassword || !newPassword) {
        return res.status(400).json({ error: "currentPassword and newPassword are required" });
    }
    if (String(newPassword).length < 8) {
        return res.status(400).json({ error: "New password must be at least 8 characters" });
    }
    try {
        const [rows] = await POOL.query(
            `SELECT password_hash FROM users WHERE id = ? LIMIT 1`,
            [req.user.id]
        );
        const user = rows?.[0];
        if (!user) return res.status(404).json({ error: "User not found" });

        const ok = await bcrypt.compare(String(currentPassword), user.password_hash);
        if (!ok) return res.status(401).json({ error: "Current password is incorrect" });

        const password_hash = await bcrypt.hash(String(newPassword), 12);
        await POOL.query(`UPDATE users SET password_hash = ? WHERE id = ?`, [password_hash, req.user.id]);
        res.json({ success: true });
    } catch (err) {
        console.error("Change password error:", err);
        res.status(500).json({ error: "Failed to change password" });
    }
});

const DEFAULT_GPT_PROMPT = `Write a crisp, professional email reply. Use precise, business-appropriate language, but keep it natural, direct, and clearly human (not templated). Mirror the sender's tone (more formal vs. casual) without overdoing it. If the sender uses bullets or line-by-line requirements (e.g., lines starting with "-"), you may respond in the same format; otherwise, write in short paragraphs.

Formatting rules (important):
- Use natural email spacing.
- Prefer 1–3 short paragraphs.
- Insert a blank line between paragraphs.
- Keep each paragraph to 1–2 sentences when possible.
- Always separate greeting, body, and sign-off into distinct blocks (with blank lines).

Style rules:
- Be concise: aim for 3–7 short sentences unless the email requires more.
- Vary sentence structure and wording to avoid generic phrasing; let a realistic professional voice come through.
- Avoid filler ("Hope you're doing well", "Good morning/afternoon", excessive enthusiasm).
- Don't invent facts. If something is missing, ask one clear question.

Content rules:
- Include 1–2 specific details from the email (e.g., names, role, req ID, dates/times, location, next step).
- If scheduling is involved: propose 2 concrete time windows with timezone OR ask for their preferred windows + timezone (choose the option that best fits the thread).
- If the email asks multiple questions: answer them in the same order.
- Keep commitments clear (what you will do, when you'll follow up).

Output rules:
- Output ONLY the email reply (no explanations).
- No bullet points unless the sender used bullets.
- Do not compress the body into a single line; use paragraph breaks where they read naturally.`;

app.get("/me/settings/gpt-prompt", requireAuth, async (req, res) => {
    try {
        const [[row]] = await POOL.query(
            `SELECT gpt_prompt FROM users WHERE id = ? LIMIT 1`,
            [req.user.id]
        );
        const prompt = row?.gpt_prompt != null && String(row.gpt_prompt).trim() !== ""
            ? String(row.gpt_prompt).trim()
            : DEFAULT_GPT_PROMPT;
        res.json({ prompt });
    } catch (err) {
        console.error("GET gpt-prompt:", err);
        res.status(500).json({ error: "Failed to load GPT prompt" });
    }
});

app.put("/me/settings/gpt-prompt", requireAuth, async (req, res) => {
    const prompt = req.body?.prompt != null ? String(req.body.prompt).trim() : "";
    try {
        await POOL.query(
            `UPDATE users SET gpt_prompt = ? WHERE id = ?`,
            [prompt || null, req.user.id]
        );
        res.json({ prompt: prompt || DEFAULT_GPT_PROMPT });
    } catch (err) {
        console.error("PUT gpt-prompt:", err);
        res.status(500).json({ error: "Failed to save GPT prompt" });
    }
});

// Generate AI reply draft: user's GPT prompt + mailbox name (myName), sender name, subject, content (body only, HTML stripped to plain text, 2000 chars)

app.post("/me/ai-draft", requireAuth, async (req, res) => {
    const subject = req.body?.subject != null ? String(req.body.subject).trim() : "";
    const myName = req.body?.myName != null ? String(req.body.myName).trim() : "";
    const senderName = req.body?.senderName != null ? String(req.body.senderName).trim() : "";
    let content = req.body?.content != null ? String(req.body.content) : "";
    if (typeof content !== "string") content = "";
    const plainContent = htmlToText(content);
    const contentSliced = plainContent.slice(0, 2000);

    const emailContext = [
        myName ? `You are replying as: ${myName}` : null,
        senderName ? `Reply to: ${senderName}` : null,
        subject ? `Subject: ${subject}` : null,
        `Content:\n${contentSliced}`,
    ].filter(Boolean).join("\n\n");

    try {
        const [[row]] = await POOL.query(
            `SELECT gpt_prompt FROM users WHERE id = ? LIMIT 1`,
            [req.user.id]
        );
        const systemPrompt = row?.gpt_prompt != null && String(row.gpt_prompt).trim() !== ""
            ? String(row.gpt_prompt).trim()
            : DEFAULT_GPT_PROMPT;

        const firstName = getFirstNameFromSender(senderName)
        const greeting = pickGreeting(firstName)
        const signoff = pickSignoff()

        const apiKey = process.env.OPENAI_API_KEY;
        if (!apiKey) {
            return res.status(503).json({ error: "AI draft is not configured (missing OPENAI_API_KEY)." });
        }
        const openai = new OpenAI({ apiKey });
        const temperature = 0.5 + Math.random() * 0.3;
        const resp = await openai.chat.completions.create({
            model: "gpt-4.1-mini",
            temperature,
            messages: [
                { role: "system", content: systemPrompt },
                {
                    role: "user", content: `
                    Use this greeting EXACTLY as the first line:
${greeting}
Use this sign-off EXACTLY:
${signoff}
                    \n\nReply to this email. Write only the reply body (no subject, no headers).\n\n${emailContext}`
                },
            ],
        });
        const draft = resp.choices?.[0]?.message?.content?.trim() || "";
        res.json({ draft });
    } catch (err) {
        console.error("POST ai-draft:", err?.response?.data || err);
        const status = err?.response?.status || err?.status || 500;
        const msg = err?.response?.data?.error?.message || err?.message || "Failed to generate draft.";
        res.status(status).json({ error: msg });
    }
});

app.get("/me/mailboxes", requireAuth, async (req, res) => {
    const limit = Math.min(Math.max(Number(req.query.limit) || 20, 1), 100);
    const offset = Math.max(Number(req.query.offset) || 0, 0);
    const userId = req.user.id;

    const [[{ total }]] = await POOL.query(
        `SELECT COUNT(*) as total FROM user_mailboxes WHERE user_id = ?`,
        [userId]
    );

    const [rows] = await POOL.query(
        `SELECT id, mailbox_email, tenant_id, graph_user_id, 
                (refresh_token IS NOT NULL) as is_connected, 
                created_at, updated_at
         FROM user_mailboxes
         WHERE user_id = ?
         ORDER BY created_at DESC
         LIMIT ? OFFSET ?`,
        [userId, limit, offset]
    );

    const totalCount = Number(total) || 0;
    const hasMore = offset + rows.length < totalCount;
    const nextOffset = offset + rows.length;

    res.json({
        value: rows,
        total: totalCount,
        hasMore,
        nextOffset: hasMore ? nextOffset : null,
    });
});

// Cache: userId::mailboxId -> displayName. Filled by "fetch all at first" when client loads mailbox list.
const mailboxDisplayNameCache = new Map();

// Fetch all connected mailbox display names from Graph, cache, and return. Called "at first" when loading mailboxes.
app.get("/me/mailbox-display-names", requireAuth, async (req, res) => {
    const userId = req.user.id;
    try {
        const [rows] = await POOL.query(
            `SELECT id, mailbox_email FROM user_mailboxes WHERE user_id = ? AND refresh_token IS NOT NULL`,
            [userId]
        );
        const out = {};
        for (const row of rows || []) {
            const mid = row.id;
            const key = `${userId}::${mid}`;
            const cached = mailboxDisplayNameCache.get(key);
            if (cached) {
                out[mid] = cached;
                continue;
            }
            try {
                const accessToken = await getValidAccessToken({ userId, mailboxId: mid });
                const meResp = await axios.get("https://graph.microsoft.com/v1.0/me?$select=displayName", {
                    headers: { Authorization: `Bearer ${accessToken}` },
                });
                const name = (meResp.data.displayName || "").trim();
                if (name) {
                    mailboxDisplayNameCache.set(key, name);
                    out[mid] = name;
                }
            } catch (e) {
                console.warn(`Mailbox ${mid} display name fetch failed:`, e?.response?.data || e?.message);
            }
        }
        res.json(out);
    } catch (err) {
        console.error("GET /me/mailbox-display-names error:", err?.message || err);
        res.status(500).json({ error: "Failed to load mailbox display names" });
    }
});

// Get mailbox profile (display name) from Graph for top bar
app.get("/me/mailboxes/:mailboxId/profile", requireAuth, async (req, res) => {
    const mailboxId = req.params.mailboxId;
    const userId = req.user.id;
    try {
        await getMailboxRowOrThrow({ userId, mailboxId });
        const accessToken = await getValidAccessToken({ userId, mailboxId });
        const meResp = await axios.get("https://graph.microsoft.com/v1.0/me?$select=displayName", {
            headers: { Authorization: `Bearer ${accessToken}` },
        });
        res.json({ displayName: meResp.data.displayName || null });
    } catch (err) {
        console.error("GET /me/mailboxes/:mailboxId/profile error:", err?.response?.data || err);
        res.status(err?.response?.status === 401 ? 401 : 500).json({ error: "Failed to load mailbox profile" });
    }
});

// Disconnect/delete a mailbox
app.delete("/me/mailboxes/:mailboxId", requireAuth, async (req, res) => {
    const { mailboxId } = req.params;
    const userId = req.user.id;

    try {
        // Verify mailbox belongs to user
        const [rows] = await POOL.query(
            `SELECT id FROM user_mailboxes WHERE id = ? AND user_id = ?`,
            [mailboxId, userId]
        );

        if (rows.length === 0) {
            return res.status(404).json({ error: "Mailbox not found" });
        }

        mailboxDisplayNameCache.delete(`${userId}::${mailboxId}`);
        await POOL.query(
            `DELETE FROM user_mailboxes WHERE id = ? AND user_id = ?`,
            [mailboxId, userId]
        );

        res.json({ success: true, message: "Mailbox disconnected successfully" });
    } catch (e) {
        console.error("Error disconnecting mailbox:", e);
        res.status(500).json({ error: "Failed to disconnect mailbox" });
    }
});

// Endpoint to get OAuth URL (for frontend to fetch with auth header)
app.get("/ms/login-url", requireAuth, (req, res) => {
    const state = jwt.sign({ uid: req.user.id, t: Date.now() }, JWT_SECRET, { expiresIn: "10m" });

    const url =
        `${AUTHORITY}/oauth2/v2.0/authorize` +
        `?client_id=${CLIENT_ID}` +
        `&response_type=code` +
        `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
        `&response_mode=query` +
        `&scope=${encodeURIComponent(SCOPES)}` +
        `&prompt=select_account` +
        `&state=${encodeURIComponent(state)}`;

    res.json({ url });
});

// Legacy redirect endpoint (kept for backward compatibility)
app.get("/ms/login", requireAuth, (req, res) => {
    const state = jwt.sign({ uid: req.user.id, t: Date.now() }, JWT_SECRET, { expiresIn: "10m" });

    const url =
        `${AUTHORITY}/oauth2/v2.0/authorize` +
        `?client_id=${CLIENT_ID}` +
        `&response_type=code` +
        `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
        `&response_mode=query` +
        `&scope=${encodeURIComponent(SCOPES)}` +
        `&prompt=select_account` +
        `&state=${encodeURIComponent(state)}`;

    res.redirect(url);
});

// Function to process all mailboxes (used by cron and immediate triggers)
async function processAllMailboxes() {
    console.log("⏰ Processing all mailboxes", new Date().toISOString());
    try {
        const [mailboxes] = await POOL.query(
            `SELECT id, user_id, mailbox_email
       FROM user_mailboxes
       WHERE refresh_token IS NOT NULL
       ORDER BY id`
        );

        console.log(`📧 Found ${mailboxes.length} mailbox(es) to process`);

        for (const mb of mailboxes) {
            try {
                console.log(`🔄 Processing mailbox: ${mb.mailbox_email} (user_id: ${mb.user_id}, mailbox_id: ${mb.id})`);
                const accessToken = await getValidAccessToken({ userId: mb.user_id, mailboxId: mb.id });

                await runCategorizer({
                    userId: mb.user_id,
                    mailboxId: mb.id,
                    mailboxEmail: mb.mailbox_email,
                    accessToken
                });
            } catch (e) {
                console.error(`❌ Mailbox processing error for ${mb.mailbox_email} (id: ${mb.id}):`, e?.response?.data || e?.message || e);
            }
        }
    } catch (e) {
        console.error("❌ Error processing mailboxes:", e);
    }
}

// Function to process a specific mailbox immediately
async function processMailboxImmediately(userId, mailboxId, mailboxEmail) {
    try {
        console.log(`🚀 Immediate categorization triggered for ${mailboxEmail} (user_id: ${userId}, mailbox_id: ${mailboxId})`);
        const accessToken = await getValidAccessToken({ userId, mailboxId });
        await runCategorizer({
            userId,
            mailboxId,
            mailboxEmail,
            accessToken
        });
    } catch (e) {
        console.error(`❌ Immediate processing error for ${mailboxEmail}:`, e?.response?.data || e?.message || e);
    }
}

app.get("/auth/callback", async (req, res) => {
    // keep your path name if you want; it’s fine.
    const code = req.query.code;
    const state = req.query.state;

    try {
        const payload = jwt.verify(String(state || ""), JWT_SECRET);
        const userId = payload.uid;

        const tokenResp = await axios.post(
            `${AUTHORITY}/oauth2/v2.0/token`,
            new URLSearchParams({
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET,
                grant_type: "authorization_code",
                code,
                redirect_uri: REDIRECT_URI,
                scope: SCOPES,
            }),
            { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
        );

        const { access_token, refresh_token, expires_in } = tokenResp.data;

        const meResp = await axios.get("https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail", {
            headers: { Authorization: `Bearer ${access_token}` },
        });

        const mailbox_email = (meResp.data.mail || meResp.data.userPrincipalName || "").toLowerCase();
        const graph_user_id = meResp.data.id;
        const tenant_id = ""; // optionally parse from id_token if you request it; leaving empty is OK for now
        const expiresAt = new Date(Date.now() + Number(expires_in || 3600) * 1000);

        // Upsert mailbox by (user_id, mailbox_email) — your requested uniqueness
        const [insertResult] = await POOL.query(
            `INSERT INTO user_mailboxes
        (user_id, mailbox_email, tenant_id, graph_user_id, access_token, refresh_token, expires_at, scopes)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE
         tenant_id=VALUES(tenant_id),
         graph_user_id=VALUES(graph_user_id),
         access_token=VALUES(access_token),
         refresh_token=VALUES(refresh_token),
         expires_at=VALUES(expires_at),
         scopes=VALUES(scopes),
         updated_at=CURRENT_TIMESTAMP`,
            [
                userId,
                mailbox_email,
                tenant_id,
                graph_user_id,
                access_token,
                refresh_token,
                expiresAt,
                SCOPES,
            ]
        );

        // Get the mailbox ID (either from insert or existing record)
        let mailboxId = insertResult.insertId;
        if (!mailboxId) {
            // If it was an update (duplicate key), fetch the existing ID
            const [rows] = await POOL.query(
                `SELECT id FROM user_mailboxes WHERE user_id = ? AND mailbox_email = ?`,
                [userId, mailbox_email]
            );
            mailboxId = rows[0]?.id;
        }

        // Trigger immediate categorization for this mailbox
        if (mailboxId) {
            processMailboxImmediately(userId, mailboxId, mailbox_email).catch((e) => {
                console.error("Failed to trigger immediate categorization:", e);
            });
        }

        // redirect back to frontend (no more ?email=... needed)
        res.redirect(`${FRONTEND}/?connected=1`);
    } catch (err) {
        console.error("MS CALLBACK ERROR:", err.response?.data || err);
        res.status(500).send("Mailbox connect failed");
    }
});

async function getOrCreateChildFolder(accessToken, parentId, displayName) {
    const name = String(displayName || "").trim().replace(/,.*$/, "").trim() || String(displayName).trim();
    const res = await axios.get(
        `https://graph.microsoft.com/v1.0/me/mailFolders/${parentId}/childFolders?$top=200&$select=id,displayName`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const list = res.data.value || [];
    const existing = list.find((f) => (f.displayName || "").toLowerCase() === name.toLowerCase());
    if (existing) return existing.id;

    try {
        const created = await axios.post(
            `https://graph.microsoft.com/v1.0/me/mailFolders/${parentId}/childFolders`,
            { displayName: name },
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );
        return created.data.id;
    } catch (err) {
        const code = err?.response?.data?.error?.code;
        if (code === "ErrorFolderExists" || (err?.response?.status === 400 && String(err?.response?.data?.error?.message || "").includes("already exists"))) {
            const retry = await axios.get(
                `https://graph.microsoft.com/v1.0/me/mailFolders/${parentId}/childFolders?$top=200&$select=id,displayName`,
                { headers: { Authorization: `Bearer ${accessToken}` } }
            );
            const found = (retry.data.value || []).find((f) => (f.displayName || "").toLowerCase() === name.toLowerCase());
            if (found) return found.id;
        }
        throw err;
    }
}

async function resolveFolderIdFromPath(accessToken, mailboxId, folderPath) {
    const raw = String(folderPath || "").trim();
    if (!raw) throw new Error("folderPath is required");

    const normalized = raw
        .split(">")
        .map((p) => p.trim())
        .filter(Boolean)
        .join(" > ");

    const key = `${mailboxId}::${normalized}`;
    const cached = folderIdCache.get(key);
    if (cached && Date.now() - cached.ts < FOLDER_ID_CACHE_TTL) return cached.id;

    let parts = normalized.split(">").map((p) => p.trim()).filter(Boolean);

    // Find longest cached prefix so we only resolve the uncached suffix
    let startIdx = 0;
    let parentId = null;
    let pathSoFar = "";
    for (let i = 1; i <= parts.length; i++) {
        const prefix = parts.slice(0, i).join(" > ");
        const prefixCached = folderIdCache.get(`${mailboxId}::${prefix}`);
        if (prefixCached && Date.now() - prefixCached.ts < FOLDER_ID_CACHE_TTL) {
            startIdx = i;
            parentId = prefixCached.id;
            pathSoFar = prefix;
        }
    }

    const first = (parts[0] || "").toLowerCase();
    if (first === "inbox") {
        if (parts.length === 1) {
            folderIdCache.set(key, { id: "inbox", ts: Date.now() });
            return "inbox";
        }
        if (parentId === null) {
            parentId = "inbox";
            pathSoFar = "Inbox";
            folderIdCache.set(`${mailboxId}::Inbox`, { id: "inbox", ts: Date.now() });
            startIdx = 1;
        }
        parts = parts.slice(1);
        const resolvedCount = pathSoFar ? pathSoFar.split(" > ").length - 1 : 0;
        for (let i = resolvedCount; i < parts.length; i++) {
            const name = parts[i];
            parentId = await getOrCreateChildFolder(accessToken, parentId, name);
            pathSoFar = pathSoFar ? pathSoFar + " > " + name : name;
            folderIdCache.set(`${mailboxId}::${pathSoFar}`, { id: parentId, ts: Date.now() });
        }
        return parentId;
    }

    if (parentId === null) {
        parentId = "inbox";
        startIdx = 0;
    }
    for (let i = startIdx; i < parts.length; i++) {
        const name = parts[i];
        parentId = await getOrCreateChildFolder(accessToken, parentId, name);
        pathSoFar = pathSoFar ? pathSoFar + " > " + name : name;
        folderIdCache.set(`${mailboxId}::${pathSoFar}`, { id: parentId, ts: Date.now() });
    }
    return parentId;
}

app.get("/emails", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const folderPath = req.query.folderPath;
    const category = req.query.category ? decodeURIComponent(String(req.query.category)) : null;
    const top = Number(req.query.top || 50);
    const skip = Number(req.query.skip || 0);

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        // verify ownership + get token
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const selectFields =
            "id,conversationId,subject,from,bodyPreview,isRead,receivedDateTime,hasAttachments,inferenceClassification,categories";

        // Fetch by label/category (Graph filter on categories)
        if (category && category.trim()) {
            const escaped = String(category).replace(/'/g, "''");
            let url =
                `https://graph.microsoft.com/v1.0/me/messages` +
                `?$filter=categories/any(c: c eq '${escaped}')` +
                `&$select=${selectFields}` +
                "&$orderby=receivedDateTime desc" +
                `&$top=${top}`;
            if (skip > 0) url += `&$skip=${skip}`;

            const { data } = await axios.get(url, {
                headers: { Authorization: `Bearer ${accessToken}` },
            });

            const msgs = (data.value || []).map((m) => ({
                ...m,
                categories: m.categories || [],
            }));

            const collapsed = collapseToLatestPerConversation(msgs);
            const hasMore = (data.value || []).length === top;

            return res.json({
                value: collapsed,
                hasMore,
                skip: skip + collapsed.length,
            });
        }

        const folderPaths = folderPath == null ? [] : (Array.isArray(folderPath) ? folderPath : [folderPath]);

        if (folderPaths.length === 1) {
            const singlePath = folderPaths[0];
            const folderId = await resolveFolderIdFromPath(accessToken, mailboxId, singlePath);

            let url =
                `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages` +
                `?$select=${selectFields}` +
                "&$orderby=receivedDateTime desc" +
                `&$top=${top}`;

            if (skip > 0) {
                url += `&$skip=${skip}`;
            }

            const { data } = await axios.get(url, {
                headers: { Authorization: `Bearer ${accessToken}` },
            });

            const msgs = (data.value || []).map((m) => ({
                ...m,
                categories: m.categories || [],
                folderPath: singlePath,
            }));

            const collapsed = collapseToLatestPerConversation(msgs);
            const hasMore = (data.value || []).length === top;

            return res.json({
                value: collapsed,
                hasMore: hasMore,
                skip: skip + collapsed.length
            });
        }

        if (folderPaths.length > 1) {
            const perFolderTop = Math.min(200, skip + top);
            const allMsgs = [];
            for (const fp of folderPaths) {
                try {
                    const folderId = await resolveFolderIdFromPath(accessToken, mailboxId, fp);
                    const url =
                        `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages` +
                        `?$select=${selectFields}` +
                        "&$orderby=receivedDateTime desc" +
                        `&$top=${perFolderTop}`;
                    const { data } = await axios.get(url, {
                        headers: { Authorization: `Bearer ${accessToken}` },
                    });
                    const list = (data.value || []).map((m) => ({
                        ...m,
                        categories: m.categories || [],
                        folderPath: fp,
                    }));
                    allMsgs.push(...list);
                } catch (err) {
                    console.warn(`Failed to fetch folder ${fp}:`, err?.response?.status || err.message);
                }
            }
            const byId = new Map();
            for (const m of allMsgs) {
                const existing = byId.get(m.id);
                if (!existing || new Date(m.receivedDateTime) > new Date(existing.receivedDateTime)) {
                    byId.set(m.id, m);
                }
            }
            const merged = Array.from(byId.values()).sort(
                (a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime)
            );
            const collapsed = collapseToLatestPerConversation(merged);
            const slice = collapsed.slice(skip, skip + top);
            const hasMore = collapsed.length > skip + top;

            return res.json({
                value: slice,
                hasMore,
                skip: skip + slice.length,
            });
        }

        // fallback: inbox + junk
        let inboxURL =
            "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages" +
            `?$select=${selectFields}` +
            "&$orderby=receivedDateTime desc" +
            `&$top=${top}`;

        let junkURL =
            "https://graph.microsoft.com/v1.0/me/mailFolders/JunkEmail/messages" +
            `?$select=${selectFields}` +
            "&$orderby=receivedDateTime desc" +
            `&$top=${top}`;

        if (skip > 0) {
            inboxURL += `&$skip=${skip}`;
            junkURL += `&$skip=${skip}`;
        }

        const [inboxRes, junkRes] = await Promise.all([
            axios.get(inboxURL, { headers: { Authorization: `Bearer ${accessToken}` } }),
            axios.get(junkURL, { headers: { Authorization: `Bearer ${accessToken}` } }),
        ]);

        const inbox = (inboxRes.data.value || []).map((m) => ({ ...m, folder: "Inbox", categories: m.categories || [] }));
        const junk = (junkRes.data.value || []).map((m) => ({ ...m, folder: "Junk", categories: m.categories || [] }));

        const combined = collapseToLatestPerConversation([...inbox, ...junk]);
        const hasMore = (inboxRes.data.value || []).length === top || (junkRes.data.value || []).length === top;

        res.json({
            value: combined,
            hasMore: hasMore,
            skip: skip + combined.length
        });
    } catch (err) {
        console.error("EMAIL LIST ERROR:", err.response?.data || err);
        res.status(err.status || 500).send("Failed to load email list");
    }
});

app.get("/email/:id", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const id = req.params.id;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const detailRes = await axios.get(
            `https://graph.microsoft.com/v1.0/me/messages/${id}?$select=id,subject,body,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,isRead,hasAttachments`,
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        res.json(detailRes.data);
    } catch (err) {
        console.error("DETAIL ERROR:", err.response?.data || err);
        res.status(err.status || 500).send("Failed to load email detail");
    }
});

app.get("/email/:id/attachments", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const id = req.params.id;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const attachmentsRes = await axios.get(
            `https://graph.microsoft.com/v1.0/me/messages/${id}/attachments?$select=id,name,size,contentType,isInline,contentBytes`,
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        const attachments = (attachmentsRes.data.value || []).map((a) => ({
            id: a.id,
            name: a.name,
            size: a.size,
            type: a["@odata.type"],
            contentType: a.contentType,
            isInline: !!a.isInline,
            preview:
                a["@odata.type"] === "#microsoft.graph.fileAttachment" &&
                    a.contentBytes &&
                    (a.contentType || "").startsWith("image/")
                    ? `data:${a.contentType};base64,${a.contentBytes}`
                    : null,
        }));

        res.json(attachments);
    } catch (err) {
        console.error("ATTACHMENTS ERROR:", err.response?.data || err);
        res.status(err.status || 500).send("Failed to load attachments");
    }
});

app.get("/email/:id/attachment/:attId", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const { disposition = "attachment" } = req.query; // "inline" | "attachment"
    const { id, attId } = req.params;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const attRes = await axios.get(
            `https://graph.microsoft.com/v1.0/me/messages/${id}/attachments/${attId}`,
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        const att = attRes.data;

        if (att["@odata.type"] === "#microsoft.graph.fileAttachment") {
            const safeName = (att.name || "attachment").replace(/"/g, "");
            const disp = disposition === "inline" ? "inline" : "attachment";

            res.setHeader("Content-Type", att.contentType || "application/octet-stream");
            res.setHeader("Content-Disposition", `${disp}; filename="${safeName}"`);

            // fileAttachment has contentBytes base64
            res.send(Buffer.from(att.contentBytes || "", "base64"));
            return;
        }

        if (att["@odata.type"] === "#microsoft.graph.referenceAttachment") {
            return res.redirect(att.sourceUrl);
        }

        res.json(att);
    } catch (err) {
        console.error("DOWNLOAD/PREVIEW ERROR:", err.response?.data || err);
        res.status(err.status || 500).send("Failed to fetch attachment");
    }
});

// Helper function to invalidate folder count cache for a mailbox and push to WebSocket clients
function invalidateFolderCountCache(userId, mailboxId) {
    const prefix = `${userId}::${mailboxId}::`;
    for (const [key] of folderCountCache.entries()) {
        if (key.startsWith(prefix)) {
            folderCountCache.delete(key);
        }
    }
    broadcastToUser(userId, { type: "folderCountsInvalidated", mailboxId });
}

app.patch("/email/:id/read", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const id = req.params.id;
    const isRead = req.query.isRead !== "false"; // Default to true, set to false to mark as unread

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        await axios.patch(
            `https://graph.microsoft.com/v1.0/me/messages/${id}`,
            { isRead: isRead },
            { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } }
        );

        // Invalidate folder count cache for this mailbox
        invalidateFolderCountCache(req.user.id, mailboxId);

        res.json({ success: true });
    } catch (err) {
        console.error("PATCH ERROR:", err.response?.data || err);
        res.status(err.status || 500).send(`Failed to mark as ${isRead ? "read" : "unread"}`);
    }
});

// Batch endpoint to mark multiple emails as read/unread
app.post("/emails/mark-read", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const { emailIds, isRead = true, folderPath } = req.body;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });
    if (!Array.isArray(emailIds) || emailIds.length === 0) {
        return res.status(400).json({ error: "emailIds array is required" });
    }

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        // Use Microsoft Graph batch API (max 20 requests per batch)
        const BATCH_SIZE = 20;
        const batches = [];
        for (let i = 0; i < emailIds.length; i += BATCH_SIZE) {
            const batch = emailIds.slice(i, i + BATCH_SIZE);
            batches.push(batch);
        }

        let successCount = 0;
        let errorCount = 0;

        for (const batch of batches) {
            const requests = batch.map((emailId, idx) => ({
                id: String(idx + 1),
                method: "PATCH",
                url: `/me/messages/${emailId}`,
                body: { isRead: isRead },
                headers: { "Content-Type": "application/json" },
            }));

            try {
                const batchResp = await executeWithConcurrencyLimit(req.user.id, mailboxId, async () => {
                    return await axios.post(
                        "https://graph.microsoft.com/v1.0/$batch",
                        { requests },
                        { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } }
                    );
                });

                const responses = batchResp.data?.responses || [];
                for (const r of responses) {
                    if (r.status >= 200 && r.status < 300) {
                        successCount++;
                    } else {
                        errorCount++;
                        console.error(`Failed to mark email in batch:`, r.status, r.body);
                    }
                }

                // Small delay between batches to respect rate limits
                if (batches.length > 1) {
                    await new Promise((r) => setTimeout(r, 100));
                }
            } catch (err) {
                console.error("Batch mark read error:", err.response?.data || err);
                errorCount += batch.length;
            }
        }

        // Invalidate folder count cache for this mailbox
        invalidateFolderCountCache(req.user.id, mailboxId);

        res.json({
            success: true,
            successCount,
            errorCount,
            total: emailIds.length,
        });
    } catch (err) {
        console.error("BATCH MARK READ ERROR:", err.response?.data || err);
        res.status(err.status || 500).json({ error: "Failed to mark emails" });
    }
});

// Delete an email
app.delete("/email/:id", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const id = req.params.id;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        await axios.delete(
            `https://graph.microsoft.com/v1.0/me/messages/${id}`,
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        // Invalidate folder count cache for this mailbox
        invalidateFolderCountCache(req.user.id, mailboxId);

        res.json({ success: true });
    } catch (err) {
        console.error("DELETE ERROR:", err.response?.data || err);
        res.status(err.status || 500).json({ error: "Failed to delete email" });
    }
});

// Move an email to a folder (e.g., Inbox)
app.post("/email/:id/move", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const id = req.params.id;
    const { destinationId = "inbox" } = req.body; // Default to inbox

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const moveResp = await axios.post(
            `https://graph.microsoft.com/v1.0/me/messages/${id}/move`,
            { destinationId },
            { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } }
        );

        // Invalidate folder count cache for this mailbox
        invalidateFolderCountCache(req.user.id, mailboxId);
        // Notify client so it can refetch the folder's email list (e.g. Inbox)
        const folderPath = destinationId === "inbox" ? "Inbox" : null;
        broadcastToUser(req.user.id, { type: "folderUpdated", mailboxId, folderPath });

        res.json({ success: true, newId: moveResp.data?.id });
    } catch (err) {
        console.error("MOVE ERROR:", err.response?.data || err);
        res.status(err.status || 500).json({ error: "Failed to move email" });
    }
});

// Reply in same thread (Microsoft Graph message reply — keeps conversation thread)
app.post("/me/messages/:messageId/reply", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const messageId = req.params.messageId;
    const { comment } = req.body;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });
    if (!messageId) return res.status(400).json({ error: "messageId is required" });

    const commentStr = (comment != null && comment !== "") ? String(comment).trim() : " ";

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        await axios.post(
            `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/reply`,
            { comment: commentStr },
            { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } }
        );

        res.status(202).json({ success: true });
    } catch (err) {
        console.error("REPLY ERROR:", err.response?.data || err);
        const status = err.response?.status || err.status || 500;
        const msg = err.response?.data?.error?.message || "Failed to send reply";
        if (err.response?.status === 403 || err.response?.data?.error?.code === "ErrorAccessDenied") {
            return res.status(403).json({
                error: "This mailbox doesn't have permission to send mail. Disconnect and reconnect it in Manage Mailboxes to grant send permission.",
                code: "SendPermissionRequired",
            });
        }
        res.status(status).json({ error: msg });
    }
});

// Send mail (reply/compose) via Microsoft Graph
app.post("/me/send-mail", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const { to, subject, body } = req.body;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });
    if (!to || typeof to !== "string" || !to.trim()) {
        return res.status(400).json({ error: "to (recipient email) is required" });
    }

    const toAddress = to.trim();
    const subjectStr = (subject != null && subject !== "") ? String(subject).trim() : "";
    const bodyStr = (body != null && body !== "") ? String(body).trim() : "";

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const payload = {
            message: {
                subject: subjectStr,
                body: {
                    contentType: "Text",
                    content: bodyStr || " ",
                },
                toRecipients: [
                    { emailAddress: { address: toAddress } },
                ],
            },
            saveToSentItems: true,
        };

        await axios.post(
            "https://graph.microsoft.com/v1.0/me/sendMail",
            payload,
            { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } }
        );

        res.status(202).json({ success: true });
    } catch (err) {
        console.error("SEND MAIL ERROR:", err.response?.data || err);
        const status = err.response?.status || err.status || 500;
        const graphError = err.response?.data?.error;
        const code = graphError?.code || "";
        const msg = graphError?.message || "Failed to send email";

        // Access denied usually means mailbox was connected before Mail.Send was added — user must reconnect
        if (status === 403 || code === "ErrorAccessDenied" || (typeof msg === "string" && msg.toLowerCase().includes("access is denied"))) {
            return res.status(403).json({
                error: "This mailbox doesn't have permission to send mail. Disconnect and reconnect it in Manage Mailboxes to grant send permission.",
                code: "SendPermissionRequired",
            });
        }

        res.status(status).json({ error: msg });
    }
});

app.get("/thread/:conversationId", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const { conversationId } = req.params;
    const MAX_MESSAGES = 10;

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });

    try {
        await getMailboxRowOrThrow({ userId: req.user.id, mailboxId });
        const accessToken = await getValidAccessToken({ userId: req.user.id, mailboxId });

        const select =
            "id,conversationId,subject,from,receivedDateTime,isRead,hasAttachments,body,bodyPreview,isDraft";

        let url =
            "https://graph.microsoft.com/v1.0/me/messages" +
            `?$filter=conversationId eq '${conversationId}'` +
            `&$select=${select}` +
            `&$top=${MAX_MESSAGES}`;

        const all = [];
        while (url && all.length < MAX_MESSAGES) {
            const { data } = await axios.get(url, {
                headers: { Authorization: `Bearer ${accessToken}` },
            });
            all.push(...(data.value || []));
            url = data["@odata.nextLink"] || null;
        }

        const messages = all
            .slice(0, MAX_MESSAGES)
            .sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));

        const withAttachments = await Promise.all(
            messages.map(async (m) => {
                if (!m.hasAttachments) return { ...m, attachments: [] };

                const attRes = await axios.get(
                    `https://graph.microsoft.com/v1.0/me/messages/${m.id}/attachments?$select=id,name,contentType,size,isInline`,
                    { headers: { Authorization: `Bearer ${accessToken}` } }
                );

                return { ...m, attachments: attRes.data.value || [] };
            })
        );

        res.json({ value: withAttachments });
    } catch (err) {
        console.error("THREAD ERROR:", err?.response?.data || err);
        res.status(err.status || 500).send("Failed to load thread");
    }
});

app.post("/folderCounts", requireAuth, async (req, res) => {
    const mailboxId = Number(req.query.mailboxId);
    const paths = Array.isArray(req.body?.paths) ? req.body.paths : [];

    if (!mailboxId) return res.status(400).json({ error: "mailboxId is required" });
    if (!paths.length) return res.json({ counts: {} });

    const userId = req.user.id;
    const queueKey = `${userId}::${mailboxId}`;

    // Queue requests to prevent concurrent requests for the same mailbox
    let queuePromise = folderCountRequestQueue.get(queueKey);
    if (queuePromise) {
        // Wait for existing request to complete
        try {
            await queuePromise;
        } catch (e) {
            // Ignore errors from previous request
        }
    }

    // Create new request promise
    const requestPromise = (async () => {
        try {
            await getMailboxRowOrThrow({ userId, mailboxId });
            const accessToken = await getValidAccessToken({ userId, mailboxId });

            const uniquePaths = Array.from(new Set(paths.map((p) => String(p || "").trim()).filter(Boolean)))
                .sort((a, b) => (a.split(">").length - b.split(">").length)); // parents first for better cache hits

            // Resolve folder IDs with concurrency limit (max 4 concurrent requests per mailbox)
            const pathToId = new Map();
            // Process in chunks of 4 to respect Microsoft Graph API concurrency limit
            const CONCURRENCY_CHUNK = 4;
            for (let i = 0; i < uniquePaths.length; i += CONCURRENCY_CHUNK) {
                const chunk = uniquePaths.slice(i, i + CONCURRENCY_CHUNK);
                await Promise.all(
                    chunk.map(async (p) => {
                        const folderId = await executeWithConcurrencyLimit(userId, mailboxId, async () => {
                            return await resolveFolderIdFromPath(accessToken, mailboxId, p);
                        });
                        pathToId.set(p, folderId);
                    })
                );
            }

            const counts = {};
            const batchTargets = [];

            // Use userId in cache key to prevent cross-user collisions
            for (const [path, folderId] of pathToId.entries()) {
                const ck = `${userId}::${mailboxId}::${folderId}`;
                const cached = folderCountCache.get(ck);
                if (cached && Date.now() - cached.ts < FOLDER_COUNT_CACHE_TTL) {
                    counts[path] = { unread: cached.unread, total: cached.total };
                } else {
                    batchTargets.push({ path, folderId });
                }
            }

            const chunk = (arr, size) => {
                const out = [];
                for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
                return out;
            };

            // Retry helper with exponential backoff for throttling
            const retryWithBackoff = async (fn, maxRetries = 3) => {
                for (let attempt = 0; attempt < maxRetries; attempt++) {
                    try {
                        return await fn();
                    } catch (e) {
                        const isThrottled = e?.response?.data?.error?.code === 'ApplicationThrottled' ||
                            e?.response?.data?.error?.code === 'ThrottledRequest' ||
                            e?.response?.status === 429;

                        if (isThrottled && attempt < maxRetries - 1) {
                            const delay = Math.min(1000 * Math.pow(2, attempt), 10000); // Exponential backoff, max 10s
                            console.log(`⏳ Throttled, retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
                            await new Promise(r => setTimeout(r, delay));
                            continue;
                        }
                        throw e;
                    }
                }
            };

            for (const group of chunk(batchTargets, 20)) {
                const requests = group.map((t, idx) => ({
                    id: String(idx + 1),
                    method: "GET",
                    url: `/me/mailFolders/${t.folderId}?$select=id,displayName,unreadItemCount,totalItemCount`,
                }));

                try {
                    const batchResp = await executeWithConcurrencyLimit(userId, mailboxId, async () => {
                        return await retryWithBackoff(async () => {
                            return await axios.post(
                                "https://graph.microsoft.com/v1.0/$batch",
                                { requests },
                                { headers: { Authorization: `Bearer ${accessToken}` } }
                            );
                        });
                    });

                    const responses = batchResp.data?.responses || [];
                    for (let i = 0; i < group.length; i++) {
                        const t = group[i];
                        const r = responses.find((x) => String(x.id) === String(i + 1));

                        // Handle missing response
                        if (!r) {
                            console.warn(`⚠️ No response for folder ${t.path}, using cached or default values`);
                            const cached = folderCountCache.get(`${userId}::${mailboxId}::${t.folderId}`);
                            if (cached) {
                                counts[t.path] = { unread: cached.unread, total: cached.total };
                            } else {
                                // Default to 0 if no cache available
                                counts[t.path] = { unread: 0, total: 0 };
                            }
                            continue;
                        }

                        const body = r?.body || {};

                        // Check for throttling in individual responses (Graph returns 429/ApplicationThrottled when rate limit is hit; we fall back to cache or 0)
                        if (r?.status === 429 || body?.error?.code === 'ApplicationThrottled' || body?.error?.code === 'ThrottledRequest') {
                            if (process.env.LOG_THROTTLE !== '0') console.warn(`⚠️ Throttled for folder ${t.path}, using cached or default values`);
                            // Use cached value if available, otherwise skip
                            const cached = folderCountCache.get(`${userId}::${mailboxId}::${t.folderId}`);
                            if (cached) {
                                counts[t.path] = { unread: cached.unread, total: cached.total };
                            } else {
                                // Default to 0 if no cache available
                                counts[t.path] = { unread: 0, total: 0 };
                            }
                            continue;
                        }

                        // Check for other errors
                        if (r?.status >= 400 || body?.error) {
                            console.warn(`⚠️ Error for folder ${t.path}:`, r?.status, body?.error);
                            const cached = folderCountCache.get(`${userId}::${mailboxId}::${t.folderId}`);
                            if (cached) {
                                counts[t.path] = { unread: cached.unread, total: cached.total };
                            } else {
                                counts[t.path] = { unread: 0, total: 0 };
                            }
                            continue;
                        }

                        // Extract unread and total counts - ensure they are numbers
                        const unread = Number(body.unreadItemCount ?? 0) || 0;
                        const total = Number(body.totalItemCount ?? 0) || 0;

                        // Always set counts - Microsoft Graph API should always return these fields
                        // Log if we're missing unreadItemCount to help debug
                        if (typeof body.unreadItemCount === 'undefined') {
                            console.warn(`⚠️ Missing unreadItemCount for folder ${t.path} (${t.folderId}), using 0`);
                        }
                        if (typeof body.totalItemCount === 'undefined') {
                            console.warn(`⚠️ Missing totalItemCount for folder ${t.path} (${t.folderId}), using 0`);
                        }

                        counts[t.path] = { unread, total };
                        // Include userId in cache key to prevent cross-user/mailbox collisions
                        folderCountCache.set(`${userId}::${mailboxId}::${t.folderId}`, { unread, total, ts: Date.now() });
                    }
                } catch (err) {
                    // If entire batch fails, try to use cached values
                    const isThrottled = err?.response?.data?.error?.code === 'ApplicationThrottled' ||
                        err?.response?.data?.error?.code === 'ThrottledRequest' ||
                        err?.response?.status === 429;

                    if (isThrottled) {
                        if (process.env.LOG_THROTTLE !== '0') console.warn(`⚠️ Batch throttled, using cached values where available`);
                        for (const t of group) {
                            const cached = folderCountCache.get(`${userId}::${mailboxId}::${t.folderId}`);
                            if (cached && Date.now() - cached.ts < FOLDER_COUNT_CACHE_TTL * 2) {
                                // Use cached value even if slightly stale
                                counts[t.path] = { unread: cached.unread, total: cached.total };
                            } else if (!counts[t.path]) {
                                // Ensure we set a default value if no cache and not already set
                                counts[t.path] = { unread: 0, total: 0 };
                            }
                        }
                    } else {
                        console.error("FOLDER COUNTS BATCH ERROR:", err.response?.data || err);
                        // Set default values for folders that failed
                        for (const t of group) {
                            if (!counts[t.path]) {
                                const cached = folderCountCache.get(`${userId}::${mailboxId}::${t.folderId}`);
                                if (cached) {
                                    counts[t.path] = { unread: cached.unread, total: cached.total };
                                } else {
                                    counts[t.path] = { unread: 0, total: 0 };
                                }
                            }
                        }
                    }
                }

                // Brief delay between batches to avoid throttling (reduced from 300ms)
                if (batchTargets.length > 20) {
                    await new Promise(r => setTimeout(r, 80));
                }
            }

            // Ensure all requested paths have counts (set to 0 if missing)
            for (const path of uniquePaths) {
                if (!counts[path]) {
                    const folderId = pathToId.get(path);
                    if (folderId) {
                        const cached = folderCountCache.get(`${userId}::${mailboxId}::${folderId}`);
                        if (cached) {
                            counts[path] = { unread: cached.unread, total: cached.total };
                        } else {
                            counts[path] = { unread: 0, total: 0 };
                        }
                    } else {
                        counts[path] = { unread: 0, total: 0 };
                    }
                }
            }

            res.json({ counts, ts: Date.now() });
        } catch (err) {
            console.error("FOLDER COUNTS ERROR:", err.response?.data || err);
            res.status(err.status || 500).send("Failed to load folder counts");
        } finally {
            // Remove from queue when done
            folderCountRequestQueue.delete(queueKey);
        }
    })();

    folderCountRequestQueue.set(queueKey, requestPromise);
    await requestPromise;
});

server.listen(4000, () => {
    console.log("🚀 Server running at http://localhost:4000 (WebSocket /ws)");
    // Trigger categorization immediately on server start
    setTimeout(() => {
        console.log("🚀 Triggering initial categorization on server start...");
        processAllMailboxes();
    }, 5000); // Wait 5 seconds for server to fully initialize
});

// Schedule cron job to run every 10 minutes
cron.schedule("*/10 * * * *", processAllMailboxes);

// Every 2 minutes: invalidate folder counts for all mailboxes so clients get fresh counts (new mail in Inbox etc.)
cron.schedule("*/2 * * * *", async () => {
    try {
        const [rows] = await POOL.query("SELECT user_id, id FROM user_mailboxes WHERE refresh_token IS NOT NULL");
        for (const row of rows || []) {
            invalidateFolderCountCache(row.user_id, row.id);
        }
    } catch (e) {
        console.error("Folder count invalidation cron:", e);
    }
});
