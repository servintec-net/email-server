const htmlToText = (html) => {
    return String(html || "")
        // remove script/style blocks
        .replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<style[\s\S]*?<\/style>/gi, " ")
        // strip all tags
        .replace(/<[^>]+>/g, " ")
        // decode a few common entities (optional but helpful)
        .replace(/&nbsp;/gi, " ")
        .replace(/&amp;/gi, "&")
        .replace(/&lt;/gi, "<")
        .replace(/&gt;/gi, ">")
        .replace(/&quot;/gi, '"')
        .replace(/&#39;/gi, "'")
        // normalize whitespace
        .replace(/\s+/g, " ")
        .trim();
};

const collapseToLatestPerConversation = (messages) => {
    const map = new Map();

    for (const msg of messages) {
        const existing = map.get(msg.conversationId);
        if (
            !existing ||
            new Date(msg.receivedDateTime) > new Date(existing.receivedDateTime)
        ) {
            map.set(msg.conversationId, msg);
        }
    }

    return Array.from(map.values()).sort(
        (a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime)
    );
}

const getFirstNameFromSender = (senderName, fallback = "there") => {
    const toTitleCase = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : "");

    const badLocalParts = new Set([
        "noreply", "no-reply", "donotreply", "do-not-reply",
        "support", "help", "sales", "info", "hello",
        "team", "hr", "recruiting", "talent", "jobs", "careers"
    ]);

    const prefixes = new Set(["mr", "mrs", "ms", "miss", "dr", "prof", "sir", "madam"]);

    const raw = senderName == null ? "" : String(senderName).trim();
    if (!raw || raw.toLowerCase() === "unknown") return fallback;

    // 1) If it's an email, use local-part
    if (raw.includes("@")) {
        const local = (raw.split("@")[0] || "").trim();
        const token = local
            .replace(/[^a-zA-Z0-9._-]/g, "")
            .split(/[._-]+/)
            .filter(Boolean)[0];

        if (!token) return fallback;
        if (badLocalParts.has(token.toLowerCase())) return fallback;
        return toTitleCase(token);
    }

    // 2) Otherwise treat as a display name
    const cleaned = raw
        .replace(/[<>()"]/g, " ")
        .replace(/\s+/g, " ")
        .trim();

    const parts = cleaned.split(" ").filter(Boolean);
    if (!parts.length) return fallback;

    // Skip prefix like "Dr." / "Mr."
    const first0 = parts[0].replace(/\./g, "").toLowerCase();
    const candidate = (prefixes.has(first0) && parts.length > 1) ? parts[1] : parts[0];

    const name = candidate.replace(/[^a-zA-Z'-]/g, "");
    if (!name) return fallback;

    // If the "name" is obviously a generic mailbox/team label, fall back
    if (badLocalParts.has(name.toLowerCase())) return fallback;

    return toTitleCase(name);
}

const pickGreeting = (firstName) => {
    const name = (firstName || "").trim();
    const hasName = name && name.toLowerCase() !== "there";

    const list = hasName
        ? ["Hi {FirstName},", "Hi {FirstName},", "Hi {FirstName},", "Hi {FirstName},", "Hello {FirstName},", "Hello {FirstName},", "{FirstName},", "Hi there {FirstName},", "Hi,", "Hi,", "Hello,", "Hi there,"]
        : ["Hi there,", "Hi there,", "Hello there,", "Hello there,", "Hi,", "Hi,", "Hi,", "Hello,"];

    const g = list[(Math.random() * list.length) | 0];
    return hasName ? g.replace("{FirstName}", name) : g;
}

const pickSignoff = () => {
    const list = ["Best,", "Best,", "Best,", "Thanks,", "Regards,", "Regards,", "Regards,", "Sincerely,", "Sincerely,", "Warmly,", "Best Regards,", "Best Regards,", "Warm Regards,", "Kind Regards,"];
    return list[(Math.random() * list.length) | 0];
}

module.exports = {
    htmlToText,
    collapseToLatestPerConversation,
    getFirstNameFromSender,
    pickGreeting,
    pickSignoff,
};
