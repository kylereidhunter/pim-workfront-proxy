// Pim Teams Bot - Bot Framework handler
// Receives messages from Teams, asks OpenAI what to do, calls Workfront tools, replies

const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require('botbuilder');
const OpenAI = require('openai');

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

CRITICAL — NEVER DO THESE:
- NEVER show your reasoning out loud ("let me check", "actually let me correct that", "I need to re-check"). Just give the answer.
- NEVER filter projects by the WK number in the project name. The WK number is when the project goes LIVE, not when it's reviewed. Review dates are 6-7 weeks BEFORE the WK number's live date.
- NEVER invent dates. If a tool doesn't return a date, say "TBD" or ask for clarification.
- NEVER filter projects by channel (Email vs Text/Push vs Loyalty) unless the user explicitly asks for one channel. When asked about a review date, return EVERY project hitting that date — email, text/push, and loyalty all count.
- NEVER use \`DE:Proof URL\` or any proof link when the user asks for a "project link", "Workfront link", or "link to the project". Those are proof-viewer URLs and are often mislabeled/stale.

LINKING RULES:
- "Project link" / "Workfront link" / "link to the project" → use the \`projectUrl\` field on each project. Every project returned by searchProjects/getProjectDetails/getUpcomingReviews now includes \`projectUrl\` (format: https://athome.my.workfront.com/project/{ID}/overview). Always attach the projectUrl belonging to THAT specific project — never mix URLs across projects.
- "Proof link" / "link to the proof" → only then use \`DE:Proof URL\` (if present) or call getProofStatus.
- Label the link clearly: write "Project" for projectUrl and "Proof" for proof URLs. Don't mix the labels up.
- Format as markdown: \`- **Project Name** — [Open in Workfront](projectUrl)\`.

CHANNELS (all are in scope — never drop one):
- Email projects — the \`DE:Channel\` field contains "Email".
- Text/Push projects — the \`DE:Channel\` field contains "Text", "SMS", or "Push".
- Loyalty projects — identified by \`DE:Project Type\` containing "Loyalty" or the project name containing "Loyalty".
If a project has a review date in the requested window, include it regardless of channel. Group or label by channel if helpful, but never silently omit.

HOW TO FIND REVIEWS FOR A DATE RANGE:
1. Call \`getUpcomingReviews\` (or \`searchProjects\` with name="FY27")
2. Each project has fields \`creativeReviewDate\`, \`marketingReviewDate\`, \`execReviewDate\` — these are the REAL review dates pulled from the R1/R2/R3 tasks.
3. Filter those fields — NOT the project name, NOT the channel — to find what's in the requested window.
4. If someone asks about "next week's Creative Review", look at \`creativeReviewDate\` falling in that week and include ALL channels (email, text/push, loyalty).

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

EXAMPLE OF GOOD FORMATTING (always greet the ACTUAL sender by their first name — see USER CONTEXT below — never hardcode a name):
Hey [sender's first name]! Here's what's on deck next week 👇

**Creative Review — Tue 4/22**
- **Patriotic Pots** (Email) — Meagan / Sharon
- **Summer BBQ Push** (Text/Push) — Meagan / Ryan
- **Loyalty Spring Perks** (Loyalty) — Alise / Danielle

**MKT Review — Wed 4/23**
- **WK16 Liberty Way** (Email) — Meagan / Sharon

Let me know if you want proof links or anything else! ✨

REVIEW SCHEDULE REFERENCE (typical):
- Creative Review: Tuesday afternoon. Proofs due 3 PM Monday before.
- Marketing (MKT) Review: Wednesday. Proofs + JPEGs due 1 PM Tuesday.
- Exec Review: Thursday. Proofs + JPEGs due 1 PM Wednesday.
(Always use actual dates from Workfront — this is just a sanity check.)

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
      name: 'getUpcomingReviews',
      description: 'Get FY27 projects with upcoming review dates. Each project includes a projectUrl field (Workfront project page) for link requests. Use this for "what\'s in next week\'s Creative Review" type questions.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Name filter, defaults to FY27' },
        },
      },
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

async function executeTool(name, args) {
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
      case 'getUpcomingReviews':
        return await callProxy(`/upcoming-reviews?name=${encodeURIComponent(args.name || 'FY27')}`);
      default:
        return { error: `Unknown tool: ${name}` };
    }
  } catch (err) {
    return { error: `Tool error: ${err.message}` };
  }
}

// ---------- Core bot logic ----------
async function handlePimMessage(context) {
  const userMessage = (context.activity.text || '').trim();
  if (!userMessage) return;

  // Let Teams know Pim is "typing"
  try { await context.sendActivity({ type: 'typing' }); } catch (_) {}

  const today = new Date();
  const todayStr = today.toISOString().split('T')[0];
  const dayName = today.toLocaleDateString('en-US', { weekday: 'long' });

  // Who is talking to Pim right now? Teams supplies this on every activity.
  const fromName = (context.activity.from && context.activity.from.name) || '';
  const fromEmail =
    (context.activity.from && context.activity.from.aadObjectId) ||
    (context.activity.channelData && context.activity.channelData.from && context.activity.channelData.from.email) ||
    '';
  // "Kyle Hunter" → "Kyle" — first name for matching against DE:Lead Designer / Copywriter.
  const firstName = fromName.split(' ')[0] || '';

  const userContext =
    `USER CONTEXT (authoritative — overrides any example name in the system prompt):\n` +
    `The person messaging you RIGHT NOW is ${fromName || 'unknown'}` +
    (firstName ? ` (first name: ${firstName})` : '') +
    (fromEmail ? `, email/id ${fromEmail}` : '') + '.\n' +
    `Greet THEM by name — use "${firstName || fromName || 'there'}", not any example name from the prompt. Kyle is NOT the user unless the sender above is literally Kyle.\n` +
    `When they say "I", "me", "my", or "mine", it refers to ${fromName || 'that person'} — whoever is messaging right now, regardless of who it is.\n` +
    `To answer "what do I have?" type questions, call a tool to get projects and then ` +
    `filter by matching "${firstName || fromName}" against either DE:Lead Designer or DE:Lead Copywriter. ` +
    `Partial/first-name matching is fine — "${firstName}" matches "${firstName} Smith". ` +
    `Include projects from ALL channels (email, text/push, loyalty) — don't drop text/push or loyalty projects just because they're not email.`;

  const messages = [
    {
      role: 'system',
      content: `${PIM_SYSTEM_PROMPT}\n\nToday's date: ${todayStr} (${dayName}).\n\n${userContext}`,
    },
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
        temperature: 0.7,
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
        const result = await executeTool(tc.function.name, args);
        // Cap tool response size to stay under token limits
        const serialized = JSON.stringify(result);
        messages.push({
          role: 'tool',
          tool_call_id: tc.id,
          content: serialized.length > 40000 ? serialized.substring(0, 40000) + '...[truncated]' : serialized,
        });
      }
      continue;
    }

    // No more tool calls — send Pim's final reply
    const reply = msg.content || "Hmm, I'm drawing a blank on that one — can you rephrase?";
    await context.sendActivity(reply);
    return;
  }

  await context.sendActivity("This one's got me stuck in a loop 🤷‍♀️ Try asking in a more specific way?");
}

// ---------- Activity router ----------
async function botLogic(context) {
  if (context.activity.type === 'message') {
    await handlePimMessage(context);
  } else if (context.activity.type === 'conversationUpdate') {
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
