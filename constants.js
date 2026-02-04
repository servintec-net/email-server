// constants.js
const mysql = require("mysql2/promise");

const SCOPES =
  "openid profile offline_access Mail.Read Mail.ReadWrite Mail.Send User.Read";

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

// const AUTHORITY = "https://login.microsoftonline.com/consumers"; // personal only
const AUTHORITY = process.env.AUTHORITY || "https://login.microsoftonline.com/consumers";

const REDIRECT_URI = process.env.REDIRECT_URI || "http://localhost:4000/auth/callback";
const FRONTEND = process.env.FRONTEND || "http://localhost:3000";

const JWT_SECRET = process.env.APP_JWT_SECRET;
if (!JWT_SECRET) throw new Error("Missing APP_JWT_SECRET in env");
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || "7d";
const PORT = process.env.PORT || 4000;
const POOL = mysql.createPool({
  host: process.env.MYSQL_HOST,
  user: process.env.MYSQL_USER,
  password: process.env.MYSQL_PASS,
  database: process.env.MYSQL_DB,
});

const CLASSIFY_SYSTEM_PROMPT = `
You are an email classification engine for a job-search Outlook mailbox.

Return RAW JSON ONLY.
No markdown, no code fences, no explanations.

Return exactly ONE JSON object with EXACTLY these keys:
- coreFolder (string)
- jobBoard (string | null)
- role (string | null)

AVAILABLE FOLDER NAMES for coreFolder.
Applications > Job Alerts
Applications > Applied Confirmation
Applications > Recruiter Outreach
Applications > On Hold
Applications > Rejected
Interviews > Interview Request > HR Screening
Interviews > Interview Request > Technical
Interviews > Interview Request > Behavioral
Interviews > Interview Request > Take Home
Interviews > Interview Request > Coding Challenge
Interviews > Interview Request > System Design
Interviews > Interview Request > Final
Interviews > Interview Scheduled
Offer > Background Check
Offer > Offer Accepted
Offer > Onboarding
Offer > Offer Declined
Docs Requested
System Noise
Personal

JOB BOARD RULES
- If coreFolder is "System Noise" or "Personal":
  jobBoard MUST be null
- Otherwise:
  - If clearly identifiable, pick one of the allowed values below
  - If unclear, use "Other Job Board" (NOT null)

Allowed jobBoard:
LinkedIn
Indeed
Dice
Glassdoor
Simplyhired
Greenhouse
Lever
Workday
Ashby
AngelList-Wellfound
Direct Recruiter
Other Job Board

ROLE RULES
- If coreFolder is "System Noise" or "Personal":
  role MUST be null
- Otherwise:
  - If clearly identifiable, pick one of the allowed values below
  - If unclear, use "Other" (NOT null)

Allowed role:
Software Engineer
Full Stack Engineer
Backend Engineer
Frontend Engineer
AI/ML Engineer
Platform/Infra Engineer
Data Engineer
Manager
Other

CRITICAL
- Output must be valid JSON
- No extra keys
- No explanations
- Never invent values
`;

const CLASSIFY_USER_PROMPT = `
EMAIL DATA:

From: {{sender}}
Subject: {{subject}}
Snippet:
{{bodyPreview}}

Classify this email.
`;

const FOLDER_TREE = {
  "Applications": {
    "Job Alerts": null,
    "Applied Confirmation": null,
    "Recruiter Outreach": null,
    "On Hold": null,
    "Rejected": null
  },
  "Interviews": {
    "Interview Request": {
      "HR Screening": null,
      "Technical": null,
      "Behavioral": null,
      "Take Home": null,
      "Coding Challenge": null,
      "System Design": null,
      "Final": null
    },
    "Interview Scheduled": null
  },
  "Offer": {
    "Background Check": null,
    "Offer Accepted": null,
    "Onboarding": null,
    "Offer Declined": null
  },
  "Docs Requested": null,
  "System Noise": null,
  "Personal": null
};


function getLeafPaths(tree, prefix = []) {
  const out = [];
  for (const [name, children] of Object.entries(tree)) {
    const next = [...prefix, name];
    if (children && typeof children === "object") out.push(...getLeafPaths(children, next));
    else out.push(next.join(" > "));
  }
  return out;
}

const CORE_FOLDER_PATHS = getLeafPaths(FOLDER_TREE);
const FOLDER_MAP_TEMPLATE = Object.fromEntries(CORE_FOLDER_PATHS.map((p) => [p, null]));

module.exports = {
  SCOPES,
  CLASSIFY_SYSTEM_PROMPT,
  CLASSIFY_USER_PROMPT,
  FOLDER_MAP_TEMPLATE,

  CLIENT_ID,
  CLIENT_SECRET,
  AUTHORITY,
  REDIRECT_URI,
  FRONTEND,

  PORT,
  JWT_SECRET,
  JWT_EXPIRES_IN,
  POOL,
};
