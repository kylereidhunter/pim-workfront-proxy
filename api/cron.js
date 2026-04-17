// Vercel Cron endpoint — fires every N minutes (see vercel.json).
// Pulls due schedules from KV, sends each proactively, updates nextRunAt.

const {
  getDueSchedules,
  getConversationRef,
  markFired,
} = require('./lib/schedule-store');
const { sendProactive } = require('./lib/proactive');
const { buildMessage } = require('./lib/message-builder');

module.exports = async (req, res) => {
  // Vercel Cron invokes with a known User-Agent + an Authorization header
  // (CRON_SECRET). We accept either auth form so manual testing works too.
  const auth = req.headers.authorization || '';
  if (process.env.CRON_SECRET && !auth.includes(process.env.CRON_SECRET)) {
    // Allow unauthenticated GET so Vercel's cron system can hit the endpoint;
    // Vercel Cron uses a Bearer token with CRON_SECRET when the env var is set.
    // If CRON_SECRET is set in env and the request doesn't present it, reject.
    return res.status(401).json({ error: 'unauthorized' });
  }

  try {
    const due = await getDueSchedules();
    if (due.length === 0) {
      return res.status(200).json({ ok: true, fired: 0 });
    }
    const results = [];
    for (const sched of due) {
      try {
        const ref = await getConversationRef(sched.conversationId);
        if (!ref) {
          results.push({ id: sched.id, status: 'no-conv-ref' });
          continue;
        }
        const text = await buildMessage(sched.messageKind, sched.messageArgs);
        await sendProactive(ref, text);
        await markFired(sched.id);
        results.push({ id: sched.id, status: 'sent' });
      } catch (err) {
        results.push({ id: sched.id, status: 'error', error: err.message });
      }
    }
    return res.status(200).json({ ok: true, fired: results.length, results });
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
};
