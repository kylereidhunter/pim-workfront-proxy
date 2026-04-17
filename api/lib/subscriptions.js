// Per-user opt-in subscription store. A user is "subscribed" when they've
// explicitly told Pim to start sending them project updates. Pim will never
// DM someone who isn't subscribed.
//
// Indexed by normalized (lowercased, trimmed) user display name so we can
// look up "Charito Jones" from a Workfront project field and find their
// Teams conversation ref.

const { Redis } = require('@upstash/redis');

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) throw new Error('KV not configured');
  return new Redis({ url, token });
}

function normName(s) {
  return String(s || '').trim().toLowerCase();
}

async function setSubscription({ conversationId, userName, enabled }) {
  if (!userName) throw new Error('userName required');
  const key = normName(userName);
  const r = kv();
  const record = {
    userName,
    userNameLower: key,
    conversationId: conversationId || null,
    enabled: !!enabled,
    updatedAt: Date.now(),
  };
  await r.set(`sub:user:${key}`, JSON.stringify(record));
  if (enabled) {
    await r.sadd('subs:enabled', key);
  } else {
    await r.srem('subs:enabled', key);
  }
  return record;
}

async function getSubscriptionByName(userName) {
  const r = kv();
  const raw = await r.get(`sub:user:${normName(userName)}`);
  if (!raw) return null;
  return typeof raw === 'string' ? JSON.parse(raw) : raw;
}

async function getSubscriptionByConvId(conversationId) {
  const r = kv();
  const names = await r.smembers('subs:enabled');
  for (const n of names || []) {
    const sub = await r.get(`sub:user:${n}`);
    const parsed = typeof sub === 'string' ? JSON.parse(sub) : sub;
    if (parsed && parsed.conversationId === conversationId) return parsed;
  }
  return null;
}

async function listEnabled() {
  const r = kv();
  const names = await r.smembers('subs:enabled');
  if (!names || names.length === 0) return [];
  const records = await Promise.all(names.map(n => r.get(`sub:user:${n}`)));
  return records
    .map(raw => (typeof raw === 'string' ? JSON.parse(raw) : raw))
    .filter(Boolean);
}

module.exports = {
  setSubscription,
  getSubscriptionByName,
  getSubscriptionByConvId,
  listEnabled,
  normName,
};
