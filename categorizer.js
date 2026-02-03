// categorizer.js
const axios = require("axios");
const OpenAI = require("openai");
const { performance } = require("perf_hooks");

const {
    CLASSIFY_SYSTEM_PROMPT,
    CLASSIFY_USER_PROMPT,
    FOLDER_MAP_TEMPLATE,
} = require("./constants");

const { htmlToText } = require("./helper");
const { getValidAccessToken } = require("./token");

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

function now() {
    return performance.now();
}

function ms(start) {
    return (performance.now() - start).toFixed(1);
}

function normPath(p) {
    return String(p || "")
        .split(">")
        .map((s) => s.trim())
        .filter(Boolean)
        .join(" > ");
}

/**
 * ✅ Per-mailbox locking (no global isRunning)
 */
const running = new Set(); // key: `${userId}::${mailboxId}`

/**
 * ✅ Per-mailbox folder map cache (no global folderMap)
 * folderMapByMailbox.get(key) -> { "Applications > Job Alerts": "<folderId>", ... }
 */
const folderMapByMailbox = new Map(); // key -> map
const folderMapTs = new Map(); // key -> ts
const FOLDERMAP_TTL = 6 * 60 * 60 * 1000; // 6 hours

function mailboxKey({ userId, mailboxId }) {
    return `${userId}::${mailboxId}`;
}

async function classifyEmail({ subject, from, snippet }) {
    const userPrompt = CLASSIFY_USER_PROMPT
        // support either placeholder variant (you had mismatch earlier)
        .replace("{{sender}}", from || "")
        .replace("{{subject}}", subject || "")
        .replace("{{body}}", String(snippet || "").slice(0, 2000));

    const resp = await openai.chat.completions.create({
        model: "gpt-4.1-mini",
        temperature: 0,
        messages: [
            { role: "system", content: CLASSIFY_SYSTEM_PROMPT },
            { role: "user", content: userPrompt },
        ],
    });

    return JSON.parse(resp.choices[0].message.content);
}

// Batch classification: process multiple emails in one API call
async function classifyEmailsBatch(emails) {
    if (emails.length === 0) return [];
    if (emails.length === 1) {
        // Single email - use existing function
        const result = await classifyEmail(emails[0]);
        return [result];
    }

    // Build batch prompt with numbered emails
    const emailData = emails.map((email, idx) => {
        const snippet = String(email.snippet || "").slice(0, 2000); // Reduced per email for batch
        return `EMAIL ${idx + 1}:
From: ${email.from || ""}
Subject: ${email.subject || ""}
Body: ${snippet}`;
    }).join("\n\n---\n\n");

    // Modified system prompt for batch processing
    const batchSystemPrompt = CLASSIFY_SYSTEM_PROMPT.replace(
        "Return exactly ONE JSON object",
        `Return a JSON ARRAY of ${emails.length} objects`
    ).replace(
        "Return RAW JSON ONLY.",
        "Return RAW JSON ARRAY ONLY. Each element must be a JSON object with keys: coreFolder, jobBoard, role."
    );

    const batchPrompt = `Classify the following ${emails.length} emails. Return a JSON array with ${emails.length} elements, one classification object per email in order.

${emailData}

Return ONLY a JSON array, no markdown, no explanations. Example format:
[{"coreFolder": "Applications > Job Alerts", "jobBoard": "LinkedIn", "role": "Software Engineer"}, {"coreFolder": "System Noise", "jobBoard": null, "role": null}, ...]`;

    const resp = await openai.chat.completions.create({
        model: "gpt-4.1-mini",
        temperature: 0,
        messages: [
            { role: "system", content: batchSystemPrompt },
            { role: "user", content: batchPrompt },
        ],
    });

    let results;
    try {
        results = JSON.parse(resp.choices[0].message.content);
    } catch (e) {
        console.error("❌ Failed to parse batch classification result:", e);
        // Fallback to individual classification
        return await Promise.all(emails.map(e => classifyEmail(e)));
    }

    // Ensure we return an array with the same length as input
    if (!Array.isArray(results)) {
        console.warn("⚠️ Batch classification returned non-array, falling back to single classification");
        return await Promise.all(emails.map(e => classifyEmail(e)));
    }

    // Pad or truncate to match input length
    while (results.length < emails.length) {
        results.push({ coreFolder: null, jobBoard: null, role: null });
    }

    return results.slice(0, emails.length);
}

async function moveJunkToInbox(accessToken) {
    const start = now();
    console.log("📦 Moving Junk emails to Inbox...");

    const junk = await axios.get(
        "https://graph.microsoft.com/v1.0/me/mailFolders/JunkEmail/messages?$top=50&$select=id",
        { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const msgs = junk.data.value || [];
    console.log(`📨 Junk messages found: ${msgs.length}`);

    for (const msg of msgs) {
        try {
            await axios.post(
                `https://graph.microsoft.com/v1.0/me/messages/${msg.id}/move`,
                { destinationId: "inbox" }, // ✅ use well-known id
                { headers: { Authorization: `Bearer ${accessToken}` } }
            );
        } catch (e) {
            console.warn("⚠️ Failed moving junk:", msg.id, e?.response?.data || e?.message);
        }
    }

    console.log(`✅ Junk → Inbox completed in ${ms(start)} ms`);
    return msgs.length;
}

async function getOrCreateChildFolder(accessToken, parentId, displayName) {
    const name = String(displayName || "").trim().replace(/,.*$/, "").trim() || String(displayName || "").trim();
    // Never create a child folder named "Inbox" under the real Inbox (avoids Inbox/Inbox and "Inbox,Inbox" recurrence)
    if (parentId === "inbox" && (name || "").toLowerCase() === "inbox") return "inbox";
    const res = await axios.get(
        `https://graph.microsoft.com/v1.0/me/mailFolders/${parentId}/childFolders?$top=200&$select=id,displayName`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
    );

    const list = res.data.value || [];
    const existing = list.find((f) => String(f.displayName || "").toLowerCase() === name.toLowerCase());
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
            const found = (retry.data.value || []).find((f) => String(f.displayName || "").toLowerCase() === name.toLowerCase());
            if (found) return found.id;
        }
        throw err;
    }
}

async function ensureFolderPath(accessToken, parts) {
    let parentId = "inbox";
    for (const name of parts) {
        parentId = await getOrCreateChildFolder(accessToken, parentId, name);
    }
    return parentId;
}

async function ensureAllFoldersForMailbox({ accessToken, key }) {
    const ts = folderMapTs.get(key);
    const cached = folderMapByMailbox.get(key);

    if (cached && ts && Date.now() - ts < FOLDERMAP_TTL) {
        return cached;
    }

    const hydrated = {};
    const paths = Object.keys(FOLDER_MAP_TEMPLATE).map(normPath).filter(Boolean);

    for (const path of paths) {
        const parts = path.split(">").map((p) => p.trim()).filter(Boolean);
        const leafId = await ensureFolderPath(accessToken, parts);
        hydrated[path] = leafId; // ✅ map FULL PATH -> folder id (no leaf collisions)
    }

    folderMapByMailbox.set(key, hydrated);
    folderMapTs.set(key, Date.now());
    return hydrated;
}

/* ------------------------------------------------------------------ */
/* MAIN JOB */
/* ------------------------------------------------------------------ */
async function runCategorizer({ userId, mailboxId, mailboxEmail, accessToken: accessTokenMaybe }) {
    if (!userId || !mailboxId) throw new Error("runCategorizer requires { userId, mailboxId }");

    const key = mailboxKey({ userId, mailboxId });
    if (running.has(key)) {
        console.log(`⏳ Categorizer already running for mailbox ${key}, skipping`);
        return;
    }

    running.add(key);
    const jobStart = now();

    try {
        let accessToken =
            accessTokenMaybe || (await getValidAccessToken({ userId, mailboxId }));

        // Helper function to refresh token if needed (on 401 errors)
        const refreshTokenIfNeeded = async (error) => {
            if (error?.response?.status === 401 || error?.response?.data?.error?.code === 'InvalidAuthenticationToken') {
                console.log(`🔄 Access token expired, refreshing...`);
                accessToken = await getValidAccessToken({ userId, mailboxId });
                return true;
            }
            return false;
        };

        console.log(`🟢 Starting categorization for ${mailboxEmail || key} (user_id: ${userId}, mailbox_id: ${mailboxId})`);

        // STEP 1: Junk -> Inbox (optional)
        const junkStart = now();
        const junkMoved = await moveJunkToInbox(accessToken);
        console.log(`📦 Junk → Inbox: ${junkMoved} messages in ${ms(junkStart)} ms`);

        // STEP 2: Ensure folder tree (cached per mailbox)
        const folderMap = await ensureAllFoldersForMailbox({ accessToken, key });

        // STEP 3: Fetch Inbox messages
        const inboxFetchStart = now();
        const PAGE_SIZE = 50;
        const MAX_TO_PROCESS = 500;

        const select =
            "id,subject,from,body,bodyPreview,categories,receivedDateTime,hasAttachments";

        let url =
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages" +
            `?$top=${PAGE_SIZE}` +
            `&$select=${select}` +
            "&$orderby=receivedDateTime desc";

        const messages = [];
        while (url && messages.length < MAX_TO_PROCESS) {
            try {
                const { data } = await axios.get(url, {
                    headers: { Authorization: `Bearer ${accessToken}` },
                });
                messages.push(...(data.value || []));
                url = data["@odata.nextLink"] || null;
            } catch (e) {
                if (await refreshTokenIfNeeded(e)) {
                    // Retry with new token
                    continue;
                }
                throw e;
            }
        }

        const batch = messages.slice(0, MAX_TO_PROCESS);

        console.log(`📨 Inbox fetched: ${batch.length} messages in ${ms(inboxFetchStart)} ms`);

        // STEP 4: Classify & move emails (only unlabeled)
        let processed = 0;

        // Filter unlabeled emails
        const unlabeledEmails = batch.filter((msg) => !(Array.isArray(msg.categories) && msg.categories.length > 0));

        // Check if batch processing is enabled (default: true)
        const USE_BATCH_CLASSIFICATION = process.env.USE_BATCH_CLASSIFICATION !== 'false';
        const BATCH_SIZE = Number(process.env.CLASSIFICATION_BATCH_SIZE) || 50;

        if (!USE_BATCH_CLASSIFICATION) {
            // Single email processing (original method - one API call per email)
            console.log(`📧 Processing ${unlabeledEmails.length} emails individually (batch mode disabled)`);

            for (const msg of unlabeledEmails) {
                const emailStart = now();

                try {
                    const contentType = msg.body?.contentType || "text";
                    const rawBody = msg.body?.content || msg.bodyPreview || "";
                    const textBody =
                        String(contentType).toLowerCase() === "html"
                            ? htmlToText(rawBody)
                            : String(rawBody || "").trim();

                    // CLASSIFY (single email)
                    const classifyStart = now();
                    const result = await classifyEmail({
                        subject: msg.subject || "",
                        from: msg.from?.emailAddress?.address || "",
                        snippet: textBody || "",
                    });
                    const classifyTime = ms(classifyStart);

                    const coreRaw = normPath(result.coreFolder || "");
                    if (!coreRaw) continue;

                    // DEST FOLDER
                    const destId = folderMap[coreRaw];
                    if (!destId) {
                        console.warn("⚠️ No folder for coreFolder:", coreRaw, "| result:", result);
                        continue;
                    }

                    // MOVE (Graph returns NEW message id)
                    const moveStart = now();
                    let moveResp;
                    try {
                        moveResp = await axios.post(
                            `https://graph.microsoft.com/v1.0/me/messages/${msg.id}/move`,
                            { destinationId: destId },
                            { headers: { Authorization: `Bearer ${accessToken}` } }
                        );
                    } catch (e) {
                        if (await refreshTokenIfNeeded(e)) {
                            // Retry with new token
                            moveResp = await axios.post(
                                `https://graph.microsoft.com/v1.0/me/messages/${msg.id}/move`,
                                { destinationId: destId },
                                { headers: { Authorization: `Bearer ${accessToken}` } }
                            );
                        } else {
                            throw e;
                        }
                    }
                    const movedMsgId = moveResp?.data?.id || msg.id;
                    const moveTime = ms(moveStart);

                    // LABEL (patch categories on moved message id)
                    let labelTime = "0.0";
                    const coreForNoiseCheck = String(result.coreFolder || "").trim();

                    if (coreForNoiseCheck !== "System Noise") {
                        const categories = [];
                        if (result.jobBoard) categories.push(`JobBoard: ${result.jobBoard}`);
                        if (result.role) categories.push(`Role: ${result.role}`);

                        if (categories.length) {
                            const labelStart = now();
                            try {
                                await axios.patch(
                                    `https://graph.microsoft.com/v1.0/me/messages/${movedMsgId}`,
                                    { categories },
                                    { headers: { Authorization: `Bearer ${accessToken}` } }
                                );
                            } catch (e) {
                                if (await refreshTokenIfNeeded(e)) {
                                    // Retry with new token
                                    await axios.patch(
                                        `https://graph.microsoft.com/v1.0/me/messages/${movedMsgId}`,
                                        { categories },
                                        { headers: { Authorization: `Bearer ${accessToken}` } }
                                    );
                                } else {
                                    throw e;
                                }
                            }
                            labelTime = ms(labelStart);
                        }
                    }

                    processed++;

                    console.log(
                        `✉️ ${msg.id.slice(0, 8)} | classify ${classifyTime} ms | move ${moveTime} ms | label ${labelTime} ms | total ${ms(
                            emailStart
                        )} ms`
                    );

                    // Rate limit safety: delay between emails to avoid Microsoft Graph API throttling
                    const rateLimitDelay = Number(process.env.RATE_LIMIT_DELAY_MS) || 700;
                    await new Promise((r) => setTimeout(r, rateLimitDelay));
                } catch (err) {
                    console.error("❌ Email processing error:", err?.response?.data || err?.message || err);
                }
            }
        } else {
            // Batch processing (multiple emails per API call - faster but uses more tokens)
            console.log(`📧 Processing ${unlabeledEmails.length} emails in batches of ${BATCH_SIZE} (batch mode enabled)`);

            for (let i = 0; i < unlabeledEmails.length; i += BATCH_SIZE) {
                const emailBatch = unlabeledEmails.slice(i, Math.min(i + BATCH_SIZE, unlabeledEmails.length));
                const batchStart = now();

                // Prepare email data for batch classification
                const emailData = emailBatch.map((msg) => {
                    const contentType = msg.body?.contentType || "text";
                    const rawBody = msg.body?.content || msg.bodyPreview || "";
                    const textBody =
                        String(contentType).toLowerCase() === "html"
                            ? htmlToText(rawBody)
                            : String(rawBody || "").trim();

                    return {
                        subject: msg.subject || "",
                        from: msg.from?.emailAddress?.address || "",
                        snippet: textBody || "",
                        msg: msg, // Keep reference to original message
                    };
                });

                // Batch classify
                const classifyStart = now();
                let classificationResults;
                try {
                    classificationResults = await classifyEmailsBatch(emailData.map(({ msg, ...rest }) => rest));
                } catch (e) {
                    console.error(`❌ Batch classification error (batch ${Math.floor(i / BATCH_SIZE) + 1}):`, e?.message || e);
                    // Fallback to individual classification for this batch
                    classificationResults = await Promise.all(
                        emailData.map(({ msg, ...rest }) => classifyEmail(rest))
                    );
                }
                const classifyTime = ms(classifyStart);
                console.log(`📦 Classified batch ${Math.floor(i / BATCH_SIZE) + 1} (${emailBatch.length} emails) in ${classifyTime} ms`);

                // Process each email with its classification result
                for (let j = 0; j < emailBatch.length; j++) {
                    const msg = emailBatch[j];
                    const result = classificationResults[j];
                    const emailStart = now();

                    try {
                        const coreRaw = normPath(result.coreFolder || "");
                        if (!coreRaw) continue;

                        // DEST FOLDER
                        const destId = folderMap[coreRaw];
                        if (!destId) {
                            console.warn("⚠️ No folder for coreFolder:", coreRaw, "| result:", result);
                            continue;
                        }

                        // MOVE (Graph returns NEW message id)
                        const moveStart = now();
                        let moveResp;
                        try {
                            moveResp = await axios.post(
                                `https://graph.microsoft.com/v1.0/me/messages/${msg.id}/move`,
                                { destinationId: destId },
                                { headers: { Authorization: `Bearer ${accessToken}` } }
                            );
                        } catch (e) {
                            if (await refreshTokenIfNeeded(e)) {
                                // Retry with new token
                                moveResp = await axios.post(
                                    `https://graph.microsoft.com/v1.0/me/messages/${msg.id}/move`,
                                    { destinationId: destId },
                                    { headers: { Authorization: `Bearer ${accessToken}` } }
                                );
                            } else {
                                throw e;
                            }
                        }
                        const movedMsgId = moveResp?.data?.id || msg.id;
                        const moveTime = ms(moveStart);

                        // LABEL (patch categories on moved message id)
                        let labelTime = "0.0";
                        const coreForNoiseCheck = String(result.coreFolder || "").trim();

                        if (coreForNoiseCheck !== "System Noise") {
                            const categories = [];
                            if (result.jobBoard) categories.push(`JobBoard: ${result.jobBoard}`);
                            if (result.role) categories.push(`Role: ${result.role}`);

                            if (categories.length) {
                                const labelStart = now();
                                try {
                                    await axios.patch(
                                        `https://graph.microsoft.com/v1.0/me/messages/${movedMsgId}`,
                                        { categories },
                                        { headers: { Authorization: `Bearer ${accessToken}` } }
                                    );
                                } catch (e) {
                                    if (await refreshTokenIfNeeded(e)) {
                                        // Retry with new token
                                        await axios.patch(
                                            `https://graph.microsoft.com/v1.0/me/messages/${movedMsgId}`,
                                            { categories },
                                            { headers: { Authorization: `Bearer ${accessToken}` } }
                                        );
                                    } else {
                                        throw e;
                                    }
                                }
                                labelTime = ms(labelStart);
                            }
                        }

                        processed++;

                        console.log(
                            `✉️ ${msg.id.slice(0, 8)} | batch classify | move ${moveTime} ms | label ${labelTime} ms | total ${ms(
                                emailStart
                            )} ms`
                        );

                        // Rate limit safety: delay between emails to avoid Microsoft Graph API throttling
                        const rateLimitDelay = Number(process.env.RATE_LIMIT_DELAY_MS) || 700;
                        await new Promise((r) => setTimeout(r, rateLimitDelay));
                    } catch (err) {
                        console.error("❌ Email processing error:", err?.response?.data || err?.message || err);
                    }
                }

                // Small delay between batches to avoid overwhelming the API
                if (i + BATCH_SIZE < unlabeledEmails.length) {
                    await new Promise((r) => setTimeout(r, 200));
                }
            }
        }

        console.log(`✅ Processed ${processed} emails`);
    } catch (err) {
        console.error("🔥 Categorizer failed:", err?.response?.data || err?.message || err);
    } finally {
        console.log(`🏁 Categorization finished in ${ms(jobStart)} ms`);
        running.delete(key);
    }
}

module.exports = { runCategorizer };
