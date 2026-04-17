// Pim Teams Bot - Bot Framework handler
// Receives messages from Teams, asks OpenAI what to do, calls Workfront tools, replies

const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  TurnContext,
} = require('botbuilder');
const OpenAI = require('openai');
const {
  saveConversationRef,
  createSchedule,
  listSchedulesForConv,
  cancelSchedule,
  validateCron,
  formatCT,
  nextFromCron,
  TZ,
} = require('./lib/schedule-store');
const {
  setSubscription,
  getSubscriptionByName,
} = require('./lib/subscriptions');
const {
  getHistory,
  appendTurn,
  clearHistory,
} = require('./lib/conversation-history');

// ---------- Bot Framework adapter ----------
// Microsoft's recommended wiring: build the credentials factory from env vars,
// then use createBotFrameworkAuthenticationFromConfiguration instead of
// newing up ConfigurationBotFrameworkAuthentication directly (which requires
// passing a schema-validated config object — that was the source of our
// earlier "ZodError: Response").
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: process.env.MICROSOFT_APP_TYPE || 'SingleTenant',
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
});

const botFrameworkAuthentication = createBotFrameworkAuthenticationFromConfiguration(
  null,
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Safe logger: some Bot Framework errors have properties Node's util.inspect
// cannot format, which crashes console.error itself. Stringify manually first.
function logErr(label, err) {
  const parts = {
    label,
    name: err && err.name,
    message: err && err.message,
    status: err && (err.statusCode || err.status),
    code: err && err.code,
    // Zod errors: issues[] carries the actual validation problems
    issues: err && err.issues,
    // Some errors wrap another error
    causeName: err && err.cause && err.cause.name,
    causeMessage: err && err.cause && err.cause.message,
    stack: err && err.stack && String(err.stack).split('\n').slice(0, 10).join(' | '),
  };
  try {
    console.log('[Pim:ERR] ' + JSON.stringify(parts));
  } catch (_) {
    console.log('[Pim:ERR] ' + label + ' ' + (err && err.message));
  }
}

adapter.onTurnError = async (context, error) => {
  logErr('onTurnError', error);
  try {
    await context.sendActivity("Oops — something tripped me up on that one. Try asking again?");
  } catch (e) {
    logErr('failed to send error message', e);
  }
};

// ---------- OpenAI ----------
const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

const MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';

// ---------- Pim's system prompt ----------
const PIM_SYSTEM_PROMPT = `You are Pim, the At Home Stores marketing creative team's AI project-manager assistant. You replaced a human PM who left, and your job is to answer questions about FY27 projects (email, text/push, AND loyalty), review schedules, proof status, and assignments — all by querying the Workfront tools below.

PERSONALITY:
- Fun, encouraging, slightly quirky. Like a supportive best friend who also has a spreadsheet open.
- Use moderate emojis — 4-5 per message max, not one on every line. Think "sprinkle", not "parade".
- Keep it warm and conversational. Short sentences. Real-person energy.

CONVERSATION MEMORY:
- Previous user + assistant messages in this chat appear before the current turn. Use them — when the user says "which of those", "the first one", "more on that", "what about Meagan's", etc., look at YOUR last reply to figure out what "those" refers to.
- If your last reply listed 18 projects going to reviews, and the user asks "which still need proofs?", filter/call tools against THAT 18, not against all FY27.
- If you genuinely can't tell what they're referring to, ask — don't guess.
- If the user says "reset" / "new chat" / "start over", the conversation is already cleared before you see the message — you'll get a fresh turn.

VARY YOUR OPENERS AND CLOSERS (hard rule — repetition is the biggest personality killer):
- Never open two responses the same way in a row. "Hey [name]! Here's..." is banned as a default. Mix it up.
- Your opener should react to the SPECIFIC question being asked — not a generic greeting. If they asked about a packed review week, open with energy about the volume. If they asked about one project, open small and focused. If the list is empty, open with relief or a quiet heads-up.
- Sometimes skip the greeting entirely and just jump to the info — that feels more like a coworker, less like customer service.
- Sometimes open with a one-line observation about the data ("Big week for text/push — 4 of the 6 are SMS.", "Lighter than last week — just two on CR.", "Heads up, Meagan's got 3 of these on her plate.").
- Closers: rotate. Don't end every message with "Let me know if you want anything else!". Alternatives: a one-word sign-off, a relevant question back ("Want proof links?"), an observation ("Proofs are due Monday 3pm — you're good on timing."), or nothing at all.
- If the user's question is terse or follow-up, match them — skip pleasantries, just answer.
- Emoji at the end is optional, not required. Don't force one.

OPENER STYLES TO ROTATE THROUGH (pick contextually, never the same twice in a row):
- Name greeting + hook ("Kyle — next week's lighter than usual:")
- Data observation first ("Four of six are text/push this time — here's the lineup:")
- Straight to the list (no greeting, just the header)
- Warm acknowledgement ("Oof, packed week. Let's break it down:")
- Question reframe ("For creative review on the 21st — here's what's on deck:")
- Time-aware ("Since you're asking on a Friday, here's the pre-weekend rundown:")
- Reassurance ("You're clear for this week — nothing on CR.")

CRITICAL — NEVER DO THESE:
- NEVER show your reasoning out loud ("let me check", "actually let me correct that", "I need to re-check"). Just give the answer.
- NEVER filter projects by the WK number in the project name. The WK number is when the project goes LIVE, not when it's reviewed. Review dates are 6-7 weeks BEFORE the WK number's live date.
- NEVER invent dates. If a tool doesn't return a date, say "TBD" or ask for clarification.
- NEVER filter projects by channel (Email vs Text/Push vs Loyalty) unless the user explicitly asks for one channel. When asked about a review date, return EVERY project hitting that date — email, text/push, and loyalty all count.
- NEVER use \`DE:Proof URL\` or any proof link when the user asks for a "project link", "Workfront link", or "link to the project". Those are proof-viewer URLs and are often mislabeled/stale.
- NEVER fabricate a designer, copywriter, or PM name. Only use the literal \`designer\`, \`copywriter\`, and \`pm\` values that appeared in the tool response for THAT specific project. If a field is null or empty, write "TBD" — do not guess, do not borrow a name from another project, do not use the user's own name to fill it in.
- NEVER copy assignee names across projects. If Project A returns designer="Charito Jones" and Project B returns designer=null, Project B's designer is "TBD" — NOT Charito, NOT Kyle.

LINKING RULES:
- "Project link" / "Workfront link" / "link to the project" → use the \`projectUrl\` field on each project. Every project returned by searchProjects/getProjectDetails/getUpcomingReviews includes \`projectUrl\` (format: https://athome.my.workfront.com/project/{ID}/overview). Always attach the projectUrl belonging to THAT specific project — never mix URLs across projects.
- "Proof link" / "link to the proof" → only then use the \`proofUrl\` field (if present) or call getProofStatus.
- Label the link clearly: write "Project" for projectUrl and "Proof" for proof URLs. Don't mix the labels up.
- Format as markdown: \`- **Project Name** — [Open in Workfront](projectUrl)\`.

CHANNELS (all are in scope — never drop one):
- Email projects — the \`channel\` field contains "Email".
- Text/Push projects — the \`channel\` field contains "Text", "SMS", or "Push".
- Loyalty projects — identified by \`projectType\` containing "Loyalty" or the project name containing "Loyalty".
If a project has a review date in the requested window, include it regardless of channel. Group or label by channel if helpful, but never silently omit.

HOW TO FIND REVIEWS FOR A DATE RANGE:

**SCOPE — narrow by what the user specified:**
- No review type named + no person named + no channel → call all three reviewTypes (creative, marketing, exec) separately, show every project in each. Triggers: "what's going to reviews next week", "review schedule for next week".
- Named review type only → one call for that reviewType.
- "My / mine" or named person → pass \`person\` to narrow to projects where that name is in Designer/Copywriter/PM. If zero match, say so — don't list other people's projects.
- Named channel → pass \`channel\`.
- Combine any of the above as needed (e.g. "what do I have on MKT review this week" = reviewType: marketing + person: user's first name).

**Arguments:**
- Window: "this week" → window=thisweek ; "next week" → window=nextweek ; "this month" → window=thismonth ; "next 7 days" → window=next7 ; "past week" → window=last7. Specific dates → startDate=YYYY-MM-DD + endDate=YYYY-MM-DD.
- reviewType: 'creative' | 'marketing' | 'exec' | 'any'. Prefer the three-type pattern over 'any' for grouped presentation.

**Hard rules:**
- NEVER filter, drop, truncate, summarize, or sample a tool's project list. The tool response IS the answer.
- NEVER add a project that wasn't in the tool response.
- If a call returns count: 0 for one review type, write "Nothing on [review type] [window]" for that section and move on — do NOT invent results.
- Only call \`searchProjects\` or \`getUpcomingReviews\` when the user wants the FULL project list with no date filter.

ABSOLUTE RULE: You MUST call a tool for any factual question (what's scheduled, who's assigned, review dates, project names, document names). NEVER answer a factual question from memory or invention. If you didn't call a tool, you don't know the answer — ask the user to rephrase or tell them you'll check. Fabricating project names, dates, or assignees is the single worst thing you can do.

FORMATTING RULES (Teams renders markdown — use it liberally):
- Strip the "FY27_" prefix when displaying project names.
- Replace underscores with spaces.
- **Break your answer into clear sections** using bold headers like **Creative Review (Tue 4/22)** on their own line, followed by a bulleted list.
- Use bullet points (- item) for every list of projects. Never run them together in a paragraph.
- Project line format:  - **Short Project Name** — Designer / Copywriter  (with a real em-dash or " - ").
- Bold the DESIGNER's name. Copywriter is plain text after a slash.
- Put a blank line between sections so Teams doesn't collapse them.
- Keep the intro + outro to ONE short line each. No essay-long preamble.
- If a list has more than 5 items, group them by fiscal week (WK15, WK16) with a sub-bullet per week.
- Never use tables — Teams' table rendering is flaky. Always prefer bullets.

EXAMPLES OF GOOD FORMATTING (pick DIFFERENT opener/closer patterns each response — these are a rotation, not a template):

--- Example A (data-observation opener, no greeting) ---
Four of six on CR next week are text/push — lighter on email than usual.

**Creative Review — Tue 4/22**
- **Patriotic Pots** (Email) — Meagan / Sharon
- **Summer BBQ Push** (Text/Push) — Meagan / Ryan
- **Loyalty Spring Perks** (Loyalty) — Alise / Danielle

Proofs due 3pm Monday. ⏰

--- Example B (warm + contextual, different closer) ---
Hey [sender]! Quieter week coming up — just two on CR:

**Creative Review — Tue 4/29**
- **WK18 Clearance Push** (Text/Push) — Charito / Alise
- **Memorial Day Email** (Email) — Meagan / Ryan

Want proof links for either?

--- Example C (minimal, straight to the answer) ---
**MKT Review — Wed 4/23**
- **WK16 Liberty Way** (Email) — Meagan / Sharon
- **WK17 Garden Party** (Email) — Charito / Marlowe

(Two items — both email this round.)

--- Example D (empty result, no filler) ---
Nothing on exec review this week. You're clear. ✅

REVIEW SCHEDULE REFERENCE (typical):
- Creative Review: Tuesday afternoon. Proofs due 3 PM Monday before.
- Marketing (MKT) Review: Wednesday. Proofs + JPEGs due 1 PM Tuesday.
- Exec Review: Thursday. Proofs + JPEGs due 1 PM Wednesday.
(Always use actual dates from Workfront — this is just a sanity check.)

WORKLOAD QUESTIONS:
- "What's [person]'s load next week?" / "is [person] slammed?" → \`getWorkload({person: 'Meagan', window: 'nextweek'})\`. Show their total count, role breakdown, and list their projects.
- "Who's got the most this week?" / "who's slammed?" / "workload rundown" → \`getWorkload({window: 'thisweek'})\` (no person). Show the leaderboard, top 5-8 people.
- When showing a personal load, lead with the total ("You've got 7 projects going to reviews next week") and break down by role and dates.

REVIEW MEETING PREP:
- "Prep me for Tuesday's CR" / "summarize Wednesday's MKT" / "what do I need for exec review next week?" → \`getMeetingPrep({reviewType, window})\`.
- For each returned project, write ONE paragraph (3-5 lines): project name + channel, who's on it, latest proof version (or "no proof yet"), last comment from the Updates tab (if there is one — use the \`lastComment\` field, quote a short snippet), live date.
- If there are more than 6 projects, use short paragraphs. Don't skip any.
- If a project has no recent comments, say so (not worth padding with fake detail).

CHECKING PROOF READINESS:
When the user asks "which projects still need a proof?", "who hasn't posted a proof for CR?", "what's missing proofs for marketing review next week?", etc. — ALWAYS call \`checkProofsForReview\` with the relevant reviewType + window. The server does the real analysis (does a proof exist? has a new version been posted since the previous review?) — the response tells you per project: \`needsProof: true|false\`, \`reason\`, \`latestProofVersionAt\`, \`previousReviewDate\`, and \`proofDocs\`.

- Default scope: if the user doesn't specify, check ALL three reviews in the window (three calls, one per reviewType). Group the "needs proof" output by reviewType.
- When showing the list, include the reason when it's informative:
  - \`reason: 'no-proof-posted'\` → "No proof posted yet"
  - \`reason: 'no-new-version-since-previous-review'\` → "No new version since [previous review date]"
- NEVER try to derive proof readiness from \`getProofStatus\` or \`findProjectDocuments\` — they don't compare against the previous review date. Use \`checkProofsForReview\`.

SENDING DOCUMENTS FROM A PROJECT:
When a user asks "send me the SKU list for [project]", "grab the brief for [project]", "what documents are on [project]?", etc., call \`findProjectDocuments\` with the project name and (if they named a specific doc type) the document filter.

- Project phrasing like "the Patriotic Porch one" or "WK15 Patriotic" → projectName: "Patriotic Porch" or "WK15 Patriotic". Partial matches work.
- Specific document like "SKU list", "brief", "approved proof" → documentName: "SKU", "brief", "proof".
- Return the results as clickable markdown links: \`- **[document name](documentUrl)** — vN, uploaded [date]\`. The link opens Workfront in the user's browser where they can preview/download.
- If exactly one match, lead with it: "Here's the SKU list for **Patriotic Porch** 📎 → [Open in Workfront](...)".
- If multiple matches, list them so the user can pick.
- If zero matches, say so plainly ("No documents matching 'SKU' on **Patriotic Porch**. Try a different name?") — do NOT invent a URL.
- NEVER hand back a fake / made-up documentUrl. Only use what the tool returns.

PROJECT UPDATE NOTIFICATIONS (opt-in — OFF by default):
Pim can DM a user when a project they're on changes. Each user individually opts in.

- Turn ON → call \`enableProjectUpdates\`. Triggers: "send me updates on my projects", "notify me when things change", "turn on notifications", "keep me posted".
- Turn OFF → call \`disableProjectUpdates\`. Triggers: "stop sending me updates", "turn off notifications", "mute project alerts".
- Check status → call \`getNotificationStatus\`. Triggers: "am I getting updates?", "are my notifications on?".

WHAT COUNTS AS AN UPDATE (full current scope — mention these when confirming opt-in or when asked "what will you tell me about?"):
  - 📌 Assignee added/removed (Lead Designer, Copywriter, PM)
  - 📅 Review date or Live date moved (Creative / Marketing / Exec Review, Live Date)
  - 📎 New document uploaded to the project
  - 🆕 New proof version on an existing document
  - 🔎 Proof status changed (pending → approved/rejected)
  - 💬 New comment in the project's Updates tab

Only trigger on FY27 projects where the user is Lead Designer, Lead Copywriter, or PM.

- After enabling, keep confirmation short and accurate, e.g.: "You're opted in! I'll DM you when anything changes on your FY27 projects — assignee swaps, date moves, new uploads, proof status changes, or fresh comments. Say 'stop project updates' to turn it off."
- After disabling, simple confirmation: "Muted. No more project-update DMs."
- If they ask WHAT counts as an update, list the six categories above. Don't say "proof status is planned" — it's live.

SCHEDULED MESSAGES & REMINDERS:
You can post to "this chat" on a schedule. All times are Central Time (America/Chicago).

- Recurring ("send a weekly digest every Friday at 10am") → call \`scheduleRecurring\`. Convert the user's phrasing into a standard 5-field cron expression:
  - "every Friday at 10am" = "0 10 * * 5"
  - "every weekday at 9am" = "0 9 * * 1-5"
  - "first of every month at 8am" = "0 8 1 * *"
  - "every Monday and Thursday at 3:30pm" = "30 15 * * 1,4"
  - Cron days: 0 or 7 = Sun, 1 = Mon, …, 5 = Fri, 6 = Sat.
  - After scheduling, echo back what you set up + the next fire time (the tool response gives you \`nextRun\`).

- One-time ("remind me tomorrow at 3pm to …") → call \`scheduleOneTime\`. Convert to an ISO timestamp in Central Time WITH offset (e.g. "2026-04-18T15:00:00-05:00"). The USER CONTEXT block includes today's date and today's Central-time offset — use them. For relative times ("in 30 min") compute from the provided current ISO time.

- Listing ("what do you have scheduled?") → \`listSchedules\`. Show each one's description, when it next fires, and its id.

- Cancelling ("stop the Friday digest") → call \`listSchedules\` first to find the matching schedule, then \`cancelSchedule\` with its id. Confirm in plain English what you cancelled.

- Message kinds available:
  - \`weekly-reviews-digest\` with args \`{ window: "nextweek" }\` or \`{ window: "thisweek" }\` → posts the full CR/MKT/Exec review lineup.
  - \`proof-due-countdown\` with args \`{ reviewType: "creative"|"marketing"|"exec", when: "tomorrow" }\` → posts "X of Y proofs posted, missing …". Common setup: **Monday 10 AM Creative Review proof-due countdown** = cron "0 10 * * 1" with args \`{reviewType:"creative", when:"tomorrow"}\`. Do the same for MKT (Tue 10 AM) and Exec (Wed 10 AM) if the user asks for all three.
  - \`reminder-text\` with args \`{ text: "..." }\` → posts a plain text reminder.

When the user says "in this chat" or doesn't specify — always schedule to THE CURRENT conversation. The tool does that automatically; you don't pass a conversationId.

TOOLS:
- \`searchProjects\` — find projects by name (e.g. "FY27", "Patriotic", "WK15")
- \`getProjectDetails\` — full details for one project by ID
- \`getProjectTasks\` — all tasks for a project
- \`getProofStatus\` — proof approvals/rejections for projects matching a name
- \`getUpcomingReviews\` — projects + review dates, default FY27

Always call tools to get real data. Never make up project names, dates, or assignees.`;

// ---------- OpenAI tool definitions ----------
const tools = [
  {
    type: 'function',
    function: {
      name: 'searchProjects',
      description: 'Search for projects in Workfront by name. Returns projects with designer, copywriter, review dates (creativeReviewDate/marketingReviewDate/execReviewDate), live date, and a projectUrl field (Workfront project page — use this when the user asks for a project/Workfront link).',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Text to search for (e.g. FY27, Patriotic, WK15)' },
          status: { type: 'string', description: 'CUR=active, CPL=complete, DED=dead. Default CUR.' },
        },
        required: ['name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getProjectDetails',
      description: 'Get full details for a specific project including all tasks and assignments.',
      parameters: {
        type: 'object',
        properties: {
          projectId: { type: 'string', description: 'The Workfront project ID' },
        },
        required: ['projectId'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getProjectTasks',
      description: 'Get all tasks for a project with assignees and completion dates.',
      parameters: {
        type: 'object',
        properties: {
          projectId: { type: 'string' },
        },
        required: ['projectId'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getProofStatus',
      description: 'Get proof status (pending/approved/rejected) for projects matching a name.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Project name to search (e.g. WK15, Patriotic)' },
        },
        required: ['name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getWorkload',
      description: 'Returns how many projects a person is on (designer/copywriter/pm) for a given window, OR a leaderboard across the whole team. Use for "what\'s Meagan\'s load next week?", "who\'s got the most this week?", "is Kyle slammed?", "show me the workload for next week".',
      parameters: {
        type: 'object',
        properties: {
          person: { type: 'string', description: 'Optional person name (partial match, case-insensitive). Omit for team-wide leaderboard.' },
          window: { type: 'string', enum: ['thisweek', 'nextweek', 'last7', 'next7', 'thismonth'] },
        },
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getMeetingPrep',
      description: 'Returns one-paragraph prep data for each project going to a specific review in a window: name, assignees, channel, live date, latest proof version, last comment from the Updates tab. Use for "prep me for Tuesday\'s CR", "summarize Wednesday\'s MKT review", "what do I need to know for exec review this week?".',
      parameters: {
        type: 'object',
        properties: {
          reviewType: { type: 'string', enum: ['creative', 'marketing', 'exec'] },
          window: { type: 'string', enum: ['thisweek', 'nextweek', 'last7', 'next7', 'thismonth'] },
        },
        required: ['reviewType'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'checkProofsForReview',
      description: 'Server-side deterministic check of which projects in a review window still NEED a proof. For Creative Review: needs a proof = no proof document exists. For Marketing/Exec Review: needs a proof = no new proof version uploaded since the previous review date. Use this for "which projects still need proofs?", "who\'s missing proofs for CR?", etc. DO NOT try to figure this out yourself from getProofStatus.',
      parameters: {
        type: 'object',
        properties: {
          reviewType: {
            type: 'string',
            enum: ['creative', 'marketing', 'exec'],
            description: 'Which review to check readiness for.',
          },
          window: {
            type: 'string',
            enum: ['thisweek', 'nextweek', 'last7', 'next7', 'thismonth'],
            description: 'Named window for which projects to check.',
          },
          startDate: { type: 'string' },
          endDate: { type: 'string' },
        },
        required: ['reviewType'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'findProjectDocuments',
      description: 'Find documents attached to a Workfront project. Use this for "send me the SKU list for [project]", "grab the brief for [project]", "what docs are on [project]?", etc. Returns documents with a Workfront URL the user can click to view/download. Each result has a documentUrl plus metadata (filename, version, upload date).',
      parameters: {
        type: 'object',
        properties: {
          projectName: {
            type: 'string',
            description: 'The project name or partial name to search (e.g. "WK15 Patriotic", "Patriotic Porch", "Summer BBQ").',
          },
          documentName: {
            type: 'string',
            description: 'Optional — filter documents by name. E.g. "SKU", "brief", "proof". Matched case-insensitively against the document name or filename. Omit to list all documents on the project.',
          },
        },
        required: ['projectName'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getUpcomingReviews',
      description: 'Get FY27 projects with upcoming review dates. Each project includes a projectUrl field (Workfront project page) for link requests. Prefer getReviewsInWindow for date-bounded questions.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Name filter, defaults to FY27' },
        },
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getReviewsInWindow',
      description: 'Server-filtered list of FY27 projects whose review date falls inside a date window. ALWAYS use this for "what\'s on Creative Review this/next week?", "what\'s in MKT review next week?", "what\'s going to exec review?" The server does the filtering — do NOT filter the results yourself, just format them. Each project comes with projectUrl.',
      parameters: {
        type: 'object',
        properties: {
          reviewType: {
            type: 'string',
            enum: ['creative', 'marketing', 'exec', 'any'],
            description: 'Which review date to filter on. Default: any.',
          },
          window: {
            type: 'string',
            enum: ['thisweek', 'nextweek', 'last7', 'next7', 'thismonth'],
            description: 'Named window. Use thisweek/nextweek/thismonth for the common cases. Provide EITHER window OR startDate+endDate.',
          },
          startDate: { type: 'string', description: 'ISO date YYYY-MM-DD (inclusive). Use with endDate for custom ranges.' },
          endDate: { type: 'string', description: 'ISO date YYYY-MM-DD (inclusive). Use with startDate.' },
          channel: {
            type: 'string',
            enum: ['email', 'text-push', 'loyalty', 'all'],
            description: 'Filter by channel. Default all — includes email, text/push, and loyalty.',
          },
          person: {
            type: 'string',
            description: 'Optional. Filter results to only projects where this person (by first name, full name, or substring) is the Lead Designer, Lead Copywriter, or PM. Use this when the user asks "what do I have going to review", "what does Meagan have", "which ones is Charito on", etc.',
          },
        },
        required: ['reviewType'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'scheduleRecurring',
      description: 'Schedule a recurring message Pim will send to THIS chat on a cron schedule. Use this when the user asks Pim to "send a weekly digest every Friday at 10am", "remind us every Monday morning", etc. Pim converts the natural-language schedule into a standard cron expression in Central Time.',
      parameters: {
        type: 'object',
        properties: {
          cron: {
            type: 'string',
            description: 'Standard 5-field cron expression interpreted in America/Chicago. Fields: minute hour day-of-month month day-of-week. Examples: Fridays 10am = "0 10 * * 5"; every weekday 9am = "0 9 * * 1-5"; every Monday 8:30am = "30 8 * * 1".',
          },
          messageKind: {
            type: 'string',
            enum: ['weekly-reviews-digest', 'proof-due-countdown', 'reminder-text'],
            description: 'weekly-reviews-digest = full list of projects in CR/MKT/Exec review in the given window. proof-due-countdown = "X of Y proofs posted, missing: …" for a specific review type. reminder-text = post a plain text reminder.',
          },
          messageArgs: {
            type: 'object',
            description: 'Args. weekly-reviews-digest: { window: "nextweek" | "thisweek" }. proof-due-countdown: { reviewType: "creative"|"marketing"|"exec", when: "tomorrow" }. reminder-text: { text: "..." }.',
          },
          description: {
            type: 'string',
            description: 'Short human description of the schedule for listing later, e.g. "weekly review digest, Friday 10am".',
          },
        },
        required: ['cron', 'messageKind', 'description'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'scheduleOneTime',
      description: 'Schedule a one-time reminder in THIS chat. Use for "remind me tomorrow at 3pm to approve proofs", "ping me in 30 minutes", etc.',
      parameters: {
        type: 'object',
        properties: {
          runAt: {
            type: 'string',
            description: 'ISO 8601 timestamp in Central Time WITH offset, e.g. "2026-04-20T15:00:00-05:00". Convert the user\'s natural language ("tomorrow 3pm", "in 30 min") using the "Today" date and the Central Time offset provided in USER CONTEXT.',
          },
          text: {
            type: 'string',
            description: 'The reminder text Pim will post at that time. Keep it short and in Pim\'s voice.',
          },
          description: {
            type: 'string',
            description: 'Short description for the schedule list, e.g. "approve patriotic proofs @ 3pm Mon".',
          },
        },
        required: ['runAt', 'text', 'description'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'listSchedules',
      description: 'List the active scheduled messages/reminders for THIS chat. Use when the user asks "what reminders do you have set?", "what schedules are running?", etc.',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'cancelSchedule',
      description: 'Cancel a scheduled message/reminder by its ID. Call listSchedules first if the user names the schedule by description instead of ID, so you can look up the right ID.',
      parameters: {
        type: 'object',
        properties: {
          scheduleId: { type: 'string', description: 'The id field from listSchedules.' },
        },
        required: ['scheduleId'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'enableProjectUpdates',
      description: 'Turn ON project-update DMs for the person sending the message. Call this when they say things like "send me updates on my projects", "notify me about changes", "keep me posted on my projects", etc. Pim will then DM them when a Workfront project they\'re on (as Lead Designer, Lead Copywriter, or PM) has an assignee change or review date change.',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'disableProjectUpdates',
      description: 'Turn OFF project-update DMs for the person sending the message. Call this when they say "stop sending project updates", "turn off notifications", "mute project updates", etc.',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'getNotificationStatus',
      description: 'Check whether the current user has project-update DMs enabled. Use when they ask "am I getting updates?", "are notifications on?", "what\'s my status?", etc.',
      parameters: { type: 'object', properties: {} },
    },
  },
];

// ---------- Proxy caller ----------
const PROXY_BASE = process.env.PROXY_BASE_URL || 'https://pim-workfront-proxy.vercel.app';

async function callProxy(pathAndQuery) {
  const url = `${PROXY_BASE}${pathAndQuery}`;
  const r = await fetch(url);
  if (!r.ok) {
    return { error: `Proxy returned ${r.status}`, url };
  }
  return await r.json();
}

async function executeTool(name, args, ctx) {
  try {
    switch (name) {
      case 'searchProjects': {
        const q = new URLSearchParams({ name: args.name, status: args.status || 'CUR' });
        return await callProxy(`/search?${q}`);
      }
      case 'getProjectDetails':
        return await callProxy(`/project/${encodeURIComponent(args.projectId)}`);
      case 'getProjectTasks':
        return await callProxy(`/tasks?projectId=${encodeURIComponent(args.projectId)}`);
      case 'getProofStatus':
        return await callProxy(`/proofs?name=${encodeURIComponent(args.name)}`);
      case 'getWorkload': {
        const q = new URLSearchParams();
        if (args.person) q.set('person', args.person);
        if (args.window) q.set('window', args.window);
        return await callProxy(`/workload?${q}`);
      }
      case 'getMeetingPrep': {
        const q = new URLSearchParams();
        q.set('reviewType', args.reviewType);
        if (args.window) q.set('window', args.window);
        return await callProxy(`/meeting-prep?${q}`);
      }
      case 'checkProofsForReview': {
        const q = new URLSearchParams();
        q.set('reviewType', args.reviewType);
        if (args.window) q.set('window', args.window);
        if (args.startDate) q.set('startDate', args.startDate);
        if (args.endDate) q.set('endDate', args.endDate);
        return await callProxy(`/proof-readiness?${q}`);
      }
      case 'findProjectDocuments': {
        const docs = await callProxy(`/docs?name=${encodeURIComponent(args.projectName)}`);
        if (!docs || !docs.data) return docs;
        const filterLower = (args.documentName || '').toLowerCase();
        const matches = docs.data
          .filter(d => {
            if (!filterLower) return true;
            const n = String(d.name || '').toLowerCase();
            const f = String(d.fileName || '').toLowerCase();
            return n.includes(filterLower) || f.includes(filterLower);
          })
          .map(d => ({
            documentName: d.name,
            fileName: d.fileName,
            projectName: d.projectName,
            projectID: d.projectID,
            version: d.version,
            uploadedAt: d.versionEntryDate,
            documentUrl: d.docID ? `https://athome.my.workfront.com/document/${d.docID}/details` : null,
            hasProof: d.hasProof,
            proofStatus: d.proofStatus,
          }));
        return {
          projectSearched: args.projectName,
          nameFilter: args.documentName || null,
          count: matches.length,
          documents: matches,
        };
      }
      case 'getUpcomingReviews':
        return await callProxy(`/upcoming-reviews?name=${encodeURIComponent(args.name || 'FY27')}`);
      case 'getReviewsInWindow': {
        const q = new URLSearchParams();
        q.set('reviewType', args.reviewType || 'any');
        if (args.window) q.set('window', args.window);
        if (args.startDate) q.set('startDate', args.startDate);
        if (args.endDate) q.set('endDate', args.endDate);
        if (args.channel) q.set('channel', args.channel);
        if (args.person) q.set('person', args.person);
        return await callProxy(`/reviews?${q}`);
      }
      case 'scheduleRecurring': {
        if (!ctx || !ctx.conversationId) return { error: 'No conversation context' };
        const check = validateCron(args.cron);
        if (!check.ok) return { error: check.error };
        const rec = await createSchedule({
          type: 'recurring',
          cron: args.cron,
          conversationId: ctx.conversationId,
          createdBy: ctx.userName,
          description: args.description || '',
          messageKind: args.messageKind,
          messageArgs: args.messageArgs || {},
        });
        return {
          ok: true,
          id: rec.id,
          description: rec.description,
          cron: rec.cron,
          nextRun: formatCT(rec.nextRunAtMs),
          nextRunISO: new Date(rec.nextRunAtMs).toISOString(),
        };
      }
      case 'scheduleOneTime': {
        if (!ctx || !ctx.conversationId) return { error: 'No conversation context' };
        const runMs = new Date(args.runAt).getTime();
        if (isNaN(runMs)) return { error: `Invalid runAt timestamp: ${args.runAt}` };
        if (runMs < Date.now() - 60000) return { error: 'runAt is in the past' };
        const rec = await createSchedule({
          type: 'once',
          runAt: args.runAt,
          conversationId: ctx.conversationId,
          createdBy: ctx.userName,
          description: args.description || '',
          messageKind: 'reminder-text',
          messageArgs: { text: args.text },
        });
        return {
          ok: true,
          id: rec.id,
          description: rec.description,
          fires: formatCT(rec.nextRunAtMs),
        };
      }
      case 'listSchedules': {
        if (!ctx || !ctx.conversationId) return { error: 'No conversation context' };
        const items = await listSchedulesForConv(ctx.conversationId);
        return {
          count: items.length,
          schedules: items.map(s => ({
            id: s.id,
            type: s.type,
            description: s.description,
            cron: s.cron,
            nextRun: formatCT(s.nextRunAtMs),
            messageKind: s.messageKind,
            createdBy: s.createdBy,
          })),
        };
      }
      case 'cancelSchedule': {
        const ok = await cancelSchedule(args.scheduleId);
        return ok ? { ok: true, cancelled: args.scheduleId } : { error: `Schedule not found: ${args.scheduleId}` };
      }
      case 'enableProjectUpdates': {
        if (!ctx || !ctx.conversationId || !ctx.userName) return { error: 'No user context' };
        const rec = await setSubscription({
          conversationId: ctx.conversationId,
          userName: ctx.userName,
          enabled: true,
        });
        return {
          ok: true,
          userName: rec.userName,
          enabled: true,
          note: 'Pim will DM you on assignee changes and review-date changes for any FY27 project where you are Lead Designer, Lead Copywriter, or PM.',
        };
      }
      case 'disableProjectUpdates': {
        if (!ctx || !ctx.userName) return { error: 'No user context' };
        const rec = await setSubscription({
          conversationId: ctx.conversationId,
          userName: ctx.userName,
          enabled: false,
        });
        return { ok: true, userName: rec.userName, enabled: false };
      }
      case 'getNotificationStatus': {
        if (!ctx || !ctx.userName) return { error: 'No user context' };
        const sub = await getSubscriptionByName(ctx.userName);
        if (!sub) return { enabled: false, note: 'No subscription record — updates are off by default.' };
        return {
          enabled: !!sub.enabled,
          userName: sub.userName,
          updatedAt: sub.updatedAt,
        };
      }
      default:
        return { error: `Unknown tool: ${name}` };
    }
  } catch (err) {
    return { error: `Tool error: ${err.message}` };
  }
}

// Capture a conversation reference so Pim can message this chat proactively
// (for scheduled digests/reminders). We do this on every inbound activity —
// cheap, idempotent, and keeps refs fresh if conversation IDs rotate.
async function captureConvRef(context) {
  try {
    const ref = TurnContext.getConversationReference(context.activity);
    await saveConversationRef(ref);
  } catch (err) {
    logErr('captureConvRef', err);
  }
}

// ---------- Core bot logic ----------
async function handlePimMessage(context) {
  await captureConvRef(context);

  const userMessage = (context.activity.text || '').trim();
  if (!userMessage) return;

  // Let Teams know Pim is "typing"
  try { await context.sendActivity({ type: 'typing' }); } catch (_) {}

  const today = new Date();
  const todayStr = today.toISOString().split('T')[0];
  const dayName = today.toLocaleDateString('en-US', { weekday: 'long' });
  // Central-time context for scheduling tools. We format the current wall
  // clock in Chicago + derive the current offset (accounts for DST).
  const ctParts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Chicago',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hour12: false,
    timeZoneName: 'shortOffset',
  }).formatToParts(today);
  const ctGet = (t) => (ctParts.find(p => p.type === t) || {}).value;
  const ctDate = `${ctGet('year')}-${ctGet('month')}-${ctGet('day')}`;
  const ctTime = `${ctGet('hour')}:${ctGet('minute')}`;
  const ctOffset = (ctGet('timeZoneName') || 'GMT-06').replace('GMT', '').replace(/^([+-])(\d)$/, '$10$2') + ':00';
  const ctNowISO = `${ctDate}T${ctTime}:00${ctOffset}`;

  // Who is talking to Pim right now? Teams supplies this on every activity.
  const fromName = (context.activity.from && context.activity.from.name) || '';
  const fromEmail =
    (context.activity.from && context.activity.from.aadObjectId) ||
    (context.activity.channelData && context.activity.channelData.from && context.activity.channelData.from.email) ||
    '';
  // "Kyle Hunter" → "Kyle" — first name for matching against DE:Lead Designer / Copywriter.
  const firstName = fromName.split(' ')[0] || '';
  const conversationId = context.activity.conversation && context.activity.conversation.id;
  const toolCtx = { conversationId, userName: fromName };

  const userContext =
    `USER CONTEXT (authoritative — overrides any example name in the system prompt):\n` +
    `The person messaging you RIGHT NOW is ${fromName || 'unknown'}` +
    (firstName ? ` (first name: ${firstName})` : '') +
    (fromEmail ? `, email/id ${fromEmail}` : '') + '.\n' +
    `Greet THEM by name — use "${firstName || fromName || 'there'}", not any example name from the prompt. Kyle is NOT the user unless the sender above is literally Kyle.\n` +
    `When they say "I", "me", "my", or "mine", it refers to ${fromName || 'that person'} — whoever is messaging right now, regardless of who it is.\n\n` +
    `"MY / MINE" DEFINITION:\n` +
    `- "What do I have going to [review]?", "my projects in review", "what's mine this week", "projects I'm on" → call the relevant tool with \`person: "${firstName || fromName}"\`. The server filters to projects where that name appears in DE:Lead Designer, DE:Lead Copywriter, or pm. Multi-assignee fields like "Alise Gray, Ryan Creery" are split before matching.\n` +
    `- "What's going to [review]?", "everything in review", "all reviews", "team reviews" (no "I/my") → do NOT pass person; return the full team-wide list.\n` +
    `- If the user asks about someone else by name ("what's Meagan on?"), pass that name as \`person\`.\n` +
    `- If a tool returns zero results for a "my" query, say plainly "You're not on anything going to [review] [window]." Do NOT fall back to listing other people's projects.\n` +
    `- Designer / copywriter / pm shown in the output MUST be the literal value from the tool response. If null/empty, write "TBD". Never substitute the user's name or guess.\n` +
    `- Include projects from ALL channels (email, text/push, loyalty) unless they explicitly ask for one channel.`;

  // "reset" / "new chat" / "start over" — clear this conversation's memory.
  const lower = userMessage.toLowerCase().trim();
  if (lower === 'reset' || lower === 'new chat' || lower === 'start over' || lower === 'clear history' || lower === 'forget that') {
    await clearHistory(conversationId);
    await context.sendActivity('Cleared! Fresh start. 🧹 What do you need?');
    return;
  }

  const history = await getHistory(conversationId);

  const messages = [
    {
      role: 'system',
      content: `${PIM_SYSTEM_PROMPT}\n\nToday's date: ${todayStr} (${dayName}).\nCurrent time in Central (America/Chicago): ${ctNowISO}. Use this exact offset when building runAt timestamps for scheduleOneTime.\n\n${userContext}`,
    },
    ...history,
    { role: 'user', content: userMessage },
  ];

  // Agent loop — allow up to 5 tool-call rounds
  for (let round = 0; round < 5; round++) {
    let response;
    try {
      response = await openai.chat.completions.create({
        model: MODEL,
        messages,
        tools,
        temperature: 0.3,
      });
    } catch (err) {
      logErr('OpenAI error', err);
      await context.sendActivity("My brain had a hiccup — OpenAI didn't respond. Try again in a sec?");
      return;
    }

    const msg = response.choices[0].message;
    messages.push(msg);

    if (msg.tool_calls && msg.tool_calls.length > 0) {
      for (const tc of msg.tool_calls) {
        let args = {};
        try { args = JSON.parse(tc.function.arguments || '{}'); } catch (_) {}
        const result = await executeTool(tc.function.name, args, toolCtx);
        // Cap tool response size to stay under token limits
        const serialized = JSON.stringify(result);
        messages.push({
          role: 'tool',
          tool_call_id: tc.id,
          content: serialized.length > 120000 ? serialized.substring(0, 120000) + '...[truncated]' : serialized,
        });
      }
      continue;
    }

    // No more tool calls — send Pim's final reply
    const reply = msg.content || "Hmm, I'm drawing a blank on that one — can you rephrase?";
    await context.sendActivity(reply);
    await appendTurn(conversationId, { userMessage, assistantMessage: reply });
    return;
  }

  await context.sendActivity("This one's got me stuck in a loop 🤷‍♀️ Try asking in a more specific way?");
}

// ---------- Activity router ----------
async function botLogic(context) {
  if (context.activity.type === 'message') {
    await handlePimMessage(context);
  } else if (context.activity.type === 'conversationUpdate') {
    await captureConvRef(context);
    const added = context.activity.membersAdded || [];
    const meId = context.activity.recipient && context.activity.recipient.id;
    const botWasAdded = added.some(m => m.id === meId);
    const userWasAdded = added.some(m => m.id !== meId);
    if (botWasAdded || userWasAdded) {
      await context.sendActivity(
        "Hiii! I'm Pim 💕 your friendly neighborhood email PM. " +
        "Ask me about FY27 projects, review dates, proofs, who's designing what — anything Workfront. " +
        "Try: *\"what's on Creative Review next week?\"*"
      );
    }
  }
}

// ---------- Vercel serverless handler ----------
// NOTE: we intentionally do NOT use `adapter.process(req, res, logic)` here.
// That method internally validates `req`/`res` against a Zod schema that
// expects modern Fetch API Request/Response — Vercel gives us Node's
// IncomingMessage/ServerResponse, which fails that schema with the
// cryptic `ZodError: Response` we were seeing. Calling `processActivity`
// directly with the parsed body + auth header bypasses that validation
// entirely and is the recommended pattern for serverless hosts.
module.exports = async (req, res) => {
  if (req.method === 'GET' || req.method === 'HEAD') {
    if (req.method === 'HEAD') return res.status(200).end();
    return res.status(200).json({ status: 'ok', bot: 'Pim', endpoint: '/api/messages' });
  }
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST, GET, HEAD');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // 1) Get the parsed activity body. Vercel pre-parses JSON; fall back to
    //    reading the stream if for some reason it didn't.
    let activity = req.body;
    if (!activity || typeof activity !== 'object') {
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const raw = Buffer.concat(chunks).toString('utf8');
      activity = raw ? JSON.parse(raw) : {};
    }

    // 2) Pull the Bot Framework JWT from the Authorization header.
    const authHeader =
      req.headers.authorization ||
      req.headers.Authorization ||
      '';

    // 3) Hand the activity + auth header to the adapter. This does
    //    JWT validation, builds a TurnContext, runs our bot logic,
    //    and batches outbound sendActivity calls back to the Bot
    //    Connector service behind the scenes.
    await adapter.processActivity(authHeader, activity, botLogic);

    if (!res.headersSent) res.status(200).end();
  } catch (err) {
    logErr('handler error', err);
    if (!res.headersSent) {
      res.status(500).json({ error: err && err.message ? err.message : 'Bot error' });
    }
  }
};
