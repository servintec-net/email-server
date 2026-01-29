const axios = require("axios");
const {
    SCOPES,
    CLIENT_ID,
    CLIENT_SECRET,
    AUTHORITY,
    POOL,
} = require("./constants");

const EXPIRY_SKEW_SECONDS = 60;

function isExpired(expiresAt) {
    if (!expiresAt) return true;
    const ms = new Date(expiresAt).getTime() - Date.now();
    return ms <= EXPIRY_SKEW_SECONDS * 1000;
}
async function getMailboxRowOrThrow({ userId, mailboxId }) {
    const [rows] = await POOL.query(
        `SELECT *
     FROM user_mailboxes
     WHERE id = ? AND user_id = ? AND refresh_token IS NOT NULL
     LIMIT 1`,
        [mailboxId, userId]
    );

    const row = rows?.[0];
    if (!row) {
        const err = new Error("Mailbox not found or not owned by user.");
        err.status = 404;
        throw err;
    }
    return row;
}

async function refreshAccessToken({ refreshToken }) {
    const tokenResp = await axios.post(
        `${AUTHORITY}/oauth2/v2.0/token`,
        new URLSearchParams({
            client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET,
            grant_type: "refresh_token",
            refresh_token: refreshToken,
            // scope is optional on refresh for v2.0; keep omitted to avoid mismatch issues
        }),
        { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    return tokenResp.data; // access_token, refresh_token (maybe), expires_in
}

/**
 * Get a valid access token for a specific mailbox owned by a specific app user.
 */
async function getValidAccessToken({ userId, mailboxId }) {
    const mb = await getMailboxRowOrThrow({ userId, mailboxId });

    if (mb.access_token && !isExpired(mb.expires_at)) {
        return mb.access_token;
    }

    if (!mb.refresh_token) {
        const err = new Error("Mailbox missing refresh token (reconnect required).");
        err.status = 401;
        throw err;
    }

    const data = await refreshAccessToken({ refreshToken: mb.refresh_token });

    const access_token = data.access_token;
    const refresh_token = data.refresh_token || mb.refresh_token;
    const expires_in = Number(data.expires_in || 3600);
    const expiresAt = new Date(Date.now() + expires_in * 1000);

    await POOL.query(
        `UPDATE user_mailboxes
     SET access_token = ?, refresh_token = ?, expires_at = ?, updated_at = CURRENT_TIMESTAMP
     WHERE id = ? AND user_id = ?`,
        [access_token, refresh_token, expiresAt, mailboxId, userId]
    );

    return access_token;
}

module.exports = {
    getValidAccessToken,
    getMailboxRowOrThrow,
};