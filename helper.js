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

module.exports = {
    htmlToText,
    collapseToLatestPerConversation,
};
