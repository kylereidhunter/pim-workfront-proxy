// Per-conversation message history so Pim can understand follow-ups
// ("which of those...", "more on the first one", etc.).
//
// Stored in Redis, TTL 24h (refreshed on every write). We only persist
// user + assistant final messages — not tool_calls or tool results —
// to keep token usage bounded. If a follow-up needs fresh data, Pim
// will call the relevant tool again.

const { Redis } = require('@upstash/redis');

const MAX_MESSAGES = 20;            // 10 exchanges
const TTL_SECONDS = 24 * 60 * 60;   // 24h — conversations reset daily
const MAX_CONTENT_CHARS = 8000;     // trim individual messages that balloon

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) throw new Error('KV not configured');
  return new Redis({ url, token });
}

async function getHistory(conversationId) {
  if (!conversationId) return [];
  try {
    const r = kv();
    const raw = await r.get(`hist:conv:${conversationId}`);
    if (!raw) return [];
    const arr = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return Array.isArray(arr) ? arr : [];
  } catch (_) {
    return [];
  }
}

function clip(content) {
  if (!content) return '';
  const s = String(content);
  return s.length > MAX_CONTENT_CHARS ? s.slice(0, MAX_CONTENT_CHARS) + '…' : s;
}

async function appendTurn(conversationId, { userMessage, assistantMessage }) {
  if (!conversationId) return;
  try {
    const r = kv();
    const current = await getHistory(conversationId);
    const next = current.slice();
    if (userMessage) next.push({ role: 'user', content: clip(userMessage) });
    if (assistantMessage) next.push({ role: 'assistant', content: clip(assistantMessage) });
    const trimmed = next.slice(-MAX_MESSAGES);
    await r.set(`hist:conv:${conversationId}`, JSON.stringify(trimmed), { ex: TTL_SECONDS });
  } catch (_) { /* don't break the bot if KV hiccups */ }
}

async function clearHistory(conversationId) {
  if (!conversationId) return;
  try {
    const r = kv();
    await r.del(`hist:conv:${conversationId}`);
  } catch (_) {}
}

module.exports = { getHistory, appendTurn, clearHistory };
