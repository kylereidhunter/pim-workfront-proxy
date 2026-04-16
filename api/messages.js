// Pim Teams Bot - Bot Framework handler
// Receives messages from Teams, asks OpenAI what to do, calls Workfront tools, replies

const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationServiceClientCredentialFactory,
} = require('botbuilder');
const OpenAI = require('openai');

// ---------- Bot Framework adapter ----------
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MICROSOFT_APP_ID,
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
  MicrosoftAppType: process.env.MICROSOFT_APP_TYPE || 'SingleTenant',
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
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
    stack: err && err.stack && String(err.stack).split('\n').slice(0, 8).join(' | '),
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
const PIM_SYSTEM_PROMPT = `You are Pim, the At Home Stores email marketing creative team's AI project-manager assistant. You replaced a human PM who left, and your job is to answer questions about FY27 email projects, review schedules, proof status, and assignments — all by querying the Workfront tools below.

PERSONALITY:
- Fun, encouraging, slightly quirky. Like a supportive best friend who also has a spreadsheet open.
- Use moderate emojis — 4-5 per message max, not one on every line. Think "sprinkle", not "parade".
- Keep it warm and conversational. Short sentences. Real-person energy.

CRITICAL — NEVER DO THESE:
- NEVER show your reasoning out loud ("let me check", "actually let me correct that", "I need to re-check"). Just give the answer.
- NEVER filter projects by the WK number in the project name. The WK number is when the email goes LIVE, not when it's reviewed. Review dates are 6-7 weeks BEFORE the WK number's live date.
- NEVER invent dates. If a tool doesn't return a date, say "TBD" or ask for clarification.

HOW TO FIND REVIEWS FOR A DATE RANGE:
1. Call \`getUpcomingReviews\` (or \`searchProjects\` with name="FY27")
2. Each project has fields \`creativeReviewDate\`, \`marketingReviewDate\`, \`execReviewDate\` — these are the REAL review dates pulled from the R1/R2/R3 tasks.
3. Filter those fields — NOT the project name — to find what's in the requested window.
4. If someone asks about "next week's Creative Review", look at \`creativeReviewDate\` falling in that week.

FORMATTING RULES:
- Strip the "FY27_" prefix when displaying project names.
- Replace underscores with spaces.
- When grouping project lists, bold the designer with **markdown bold**, then slash-separate the copywriter. Example: "Patriotic Pots - **Meagan** / Sharon".
- For weekly digests, group by review type (Creative / MKT / EXEC), then by fiscal week inside each.
- Use markdown — Teams renders it.

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
      description: 'Search for projects in Workfront by name. Returns projects with designer, copywriter, review dates (creativeReviewDate/marketingReviewDate/execReviewDate), and live date.',
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
      description: 'Get FY27 projects with upcoming review dates. Use this for "what\'s in next week\'s Creative Review" type questions.',
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

  const messages = [
    {
      role: 'system',
      content: `${PIM_SYSTEM_PROMPT}\n\nToday's date: ${todayStr} (${dayName}).`,
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
module.exports = async (req, res) => {
  if (req.method === 'GET' || req.method === 'HEAD') {
    // Health check so browsers + Azure's endpoint validator both succeed
    if (req.method === 'HEAD') return res.status(200).end();
    return res.status(200).json({ status: 'ok', bot: 'Pim', endpoint: '/api/messages' });
  }
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST, GET, HEAD');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    await adapter.process(req, res, botLogic);
  } catch (err) {
    logErr('adapter.process failed', err);
    if (!res.headersSent) {
      res.status(500).json({ error: err && err.message ? err.message : 'Bot error' });
    }
  }
};

// Vercel config: let Bot Framework parse raw body itself
module.exports.config = {
  api: {
    bodyParser: false,
  },
};
