// KV-backed store for Pim's conversation refs + schedules.
// Uses Upstash Redis (Vercel KV). Required env vars:
//   KV_REST_API_URL   (or UPSTASH_REDIS_REST_URL)
//   KV_REST_API_TOKEN (or UPSTASH_REDIS_REST_TOKEN)

const { Redis } = require('@upstash/redis');
const cronParser = require('cron-parser');
const crypto = require('crypto');

const TZ = 'America/Chicago';

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) throw new Error('KV not configured: set KV_REST_API_URL and KV_REST_API_TOKEN');
  return new Redis({ url, token });
}

// ---------- Conversation references ----------
// A conversation reference lets us message a Teams user/chat later without
// them initiating the conversation first.
async function saveConversationRef(ref) {
  if (!ref || !ref.conversation || !ref.conversation.id) return;
  const r = kv();
  const id = ref.conversation.id;
  await r.set(`convref:${id}`, JSON.stringify(ref));
  // Index by serviceUrl+channel for diagnostics
  await r.sadd('convrefs:all', id);
}

async function getConversationRef(conversationId) {
  const r = kv();
  const raw = await r.get(`convref:${conversationId}`);
  if (!raw) return null;
  return typeof raw === 'string' ? JSON.parse(raw) : raw;
}

// ---------- Schedules ----------
// A schedule is either:
//   { type: 'recurring', cron: '0 10 * * 5', ... }
//   { type: 'once', runAt: <ISO timestamp in UTC>, ... }
// Both carry: id, conversationId, createdBy, createdAt, description,
// messageKind, messageArgs, active, nextRunAtMs.

function nextFromCron(cronExpr, fromDate) {
  const it = cronParser.parseExpression(cronExpr, {
    tz: TZ,
    currentDate: fromDate || new Date(),
  });
  return it.next().getTime();
}

function newId() {
  return crypto.randomBytes(6).toString('hex');
}

async function createSchedule(input) {
  const r = kv();
  const now = Date.now();
  const id = newId();
  let nextRunAtMs;
  if (input.type === 'recurring') {
    nextRunAtMs = nextFromCron(input.cron, new Date());
  } else if (input.type === 'once') {
    nextRunAtMs = new Date(input.runAt).getTime();
    if (isNaN(nextRunAtMs)) throw new Error(`Invalid runAt: ${input.runAt}`);
  } else {
    throw new Error(`Unknown schedule type: ${input.type}`);
  }
  const record = {
    id,
    type: input.type,
    cron: input.cron || null,
    runAt: input.runAt || null,
    conversationId: input.conversationId,
    createdBy: input.createdBy || null,
    createdAt: now,
    description: input.description || '',
    messageKind: input.messageKind,
    messageArgs: input.messageArgs || {},
    active: true,
    nextRunAtMs,
  };
  await r.set(`schedule:${id}`, JSON.stringify(record));
  await r.sadd(`schedules:conv:${input.conversationId}`, id);
  await r.zadd('schedules:due', { score: nextRunAtMs, member: id });
  return record;
}

async function getSchedule(id) {
  const r = kv();
  const raw = await r.get(`schedule:${id}`);
  if (!raw) return null;
  return typeof raw === 'string' ? JSON.parse(raw) : raw;
}

async function listSchedulesForConv(conversationId) {
  const r = kv();
  const ids = await r.smembers(`schedules:conv:${conversationId}`);
  if (!ids || ids.length === 0) return [];
  const records = await Promise.all(ids.map(id => getSchedule(id)));
  return records.filter(Boolean).filter(s => s.active);
}

async function cancelSchedule(id) {
  const r = kv();
  const existing = await getSchedule(id);
  if (!existing) return false;
  existing.active = false;
  await r.set(`schedule:${id}`, JSON.stringify(existing));
  await r.zrem('schedules:due', id);
  await r.srem(`schedules:conv:${existing.conversationId}`, id);
  return true;
}

async function markFired(id) {
  const r = kv();
  const s = await getSchedule(id);
  if (!s) return;
  if (s.type === 'once') {
    s.active = false;
    await r.zrem('schedules:due', id);
    await r.srem(`schedules:conv:${s.conversationId}`, id);
  } else if (s.type === 'recurring' && s.cron) {
    s.nextRunAtMs = nextFromCron(s.cron, new Date());
    await r.zadd('schedules:due', { score: s.nextRunAtMs, member: id });
  }
  s.lastFiredAt = Date.now();
  await r.set(`schedule:${id}`, JSON.stringify(s));
}

async function getDueSchedules(nowMs = Date.now()) {
  const r = kv();
  const ids = await r.zrange('schedules:due', 0, nowMs, { byScore: true });
  if (!ids || ids.length === 0) return [];
  const records = await Promise.all(ids.map(id => getSchedule(id)));
  return records.filter(s => s && s.active);
}

// ---------- Natural-language time parsing helpers ----------
// Pim passes us a cron expression or an ISO timestamp directly (GPT does the
// parsing). We validate here and surface useful errors.
function validateCron(cron) {
  try {
    cronParser.parseExpression(cron, { tz: TZ });
    return { ok: true };
  } catch (err) {
    return { ok: false, error: `Invalid cron "${cron}": ${err.message}` };
  }
}

function formatCT(ms) {
  try {
    return new Date(ms).toLocaleString('en-US', {
      timeZone: TZ,
      weekday: 'short',
      month: 'short',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
    });
  } catch {
    return new Date(ms).toISOString();
  }
}

module.exports = {
  TZ,
  saveConversationRef,
  getConversationRef,
  createSchedule,
  getSchedule,
  listSchedulesForConv,
  cancelSchedule,
  markFired,
  getDueSchedules,
  validateCron,
  nextFromCron,
  formatCT,
};
