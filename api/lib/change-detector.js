// Polls the Workfront proxy, snapshots each project's key fields in Redis,
// diffs against the last snapshot, and emits change events.
//
// Tracked fields per project:
//   designer, copywriter, pm
//   creativeReviewDate, marketingReviewDate, execReviewDate
//
// The first time we poll (no prior snapshot), we write the baseline and
// emit ZERO events — we don't want to spam everyone with "you were just
// assigned to 194 projects".

const { Redis } = require('@upstash/redis');
const { listEnabled, normName } = require('./subscriptions');
const { sendProactive } = require('./proactive');

const PROXY_BASE = process.env.PROXY_BASE_URL || 'https://pim-workfront-proxy.vercel.app';
const TRACKED_FIELDS = [
  'designer',
  'copywriter',
  'pm',
  'creativeReviewDate',
  'marketingReviewDate',
  'execReviewDate',
];

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) throw new Error('KV not configured');
  return new Redis({ url, token });
}

async function fetchCurrentProjects() {
  const r = await fetch(`${PROXY_BASE}/upcoming-reviews?name=FY27`);
  if (!r.ok) throw new Error(`proxy /upcoming-reviews -> ${r.status}`);
  const body = await r.json();
  return body.data || [];
}

function snapshotFor(proj) {
  const out = {};
  for (const f of TRACKED_FIELDS) out[f] = proj[f] == null ? null : proj[f];
  return out;
}

function fieldLabel(field) {
  return {
    designer: 'Lead Designer',
    copywriter: 'Lead Copywriter',
    pm: 'PM',
    creativeReviewDate: 'Creative Review',
    marketingReviewDate: 'Marketing Review',
    execReviewDate: 'Exec Review',
  }[field] || field;
}

function isAssigneeField(f) {
  return f === 'designer' || f === 'copywriter' || f === 'pm';
}

function formatDate(d) {
  if (!d) return 'unset';
  const fixed = String(d).replace(/(\d{2}):(\d{3})/, '$1.$2');
  const dt = new Date(fixed);
  if (isNaN(dt.getTime())) return String(d);
  return dt.toLocaleDateString('en-US', {
    timeZone: 'America/Chicago',
    weekday: 'short',
    month: 'numeric',
    day: 'numeric',
  });
}

function shortName(name) {
  return String(name || '').replace(/^FY27_/, '').replace(/_/g, ' ').trim();
}

// Compare prev vs curr snapshots, return a list of {field, from, to}.
function diff(prev, curr) {
  const changes = [];
  for (const f of TRACKED_FIELDS) {
    const a = prev[f] == null ? null : prev[f];
    const b = curr[f] == null ? null : curr[f];
    if (a !== b) changes.push({ field: f, from: a, to: b });
  }
  return changes;
}

// For an assignee change, affected = new assignee + old assignee (they should
// know they were removed). For a date change, affected = current designer,
// copywriter, and pm.
function affectedNamesFor(change, currProj) {
  const set = new Set();
  if (isAssigneeField(change.field)) {
    if (change.to) set.add(change.to);
    if (change.from) set.add(change.from);
  } else {
    if (currProj.designer) set.add(currProj.designer);
    if (currProj.copywriter) set.add(currProj.copywriter);
    if (currProj.pm) set.add(currProj.pm);
  }
  return [...set];
}

function buildNotificationText(change, currProj, recipientName) {
  const proj = shortName(currProj.name);
  const url = currProj.projectUrl;
  const openLink = url ? ` — [Open in Workfront](${url})` : '';
  if (isAssigneeField(change.field)) {
    const role = fieldLabel(change.field);
    if (change.to && normName(change.to) === normName(recipientName)) {
      return `📌 You were just assigned as **${role}** on **${proj}**${openLink}`;
    }
    if (change.from && normName(change.from) === normName(recipientName)) {
      const replaced = change.to ? ` — now ${change.to}` : '';
      return `🔄 You were removed as **${role}** on **${proj}**${replaced}${openLink}`;
    }
    // Someone else's role changed — recipient is a related assignee
    const who = change.to || 'TBD';
    return `🔁 **${role}** on **${proj}** changed to **${who}** (prev: ${change.from || 'unset'})${openLink}`;
  }
  // Date change
  const label = fieldLabel(change.field);
  return `📅 **${label}** for **${proj}** moved from _${formatDate(change.from)}_ to _${formatDate(change.to)}_${openLink}`;
}

// Main polling cycle. Called by /cron. Returns a summary.
async function detectAndNotify() {
  const r = kv();
  const projects = await fetchCurrentProjects();
  const snapKeys = projects.map(p => `snap:project:${p.ID}`);

  // Load prior snapshots in one batch
  const priorRaws = snapKeys.length ? await r.mget(...snapKeys) : [];
  const priorMap = new Map();
  projects.forEach((p, i) => {
    const raw = priorRaws[i];
    const parsed = raw == null ? null : (typeof raw === 'string' ? JSON.parse(raw) : raw);
    priorMap.set(p.ID, parsed);
  });

  const isFirstRun = priorRaws.every(x => x == null);

  // Load subscribers once
  const subs = await listEnabled();
  const subsByName = new Map(subs.map(s => [s.userNameLower, s]));

  const events = [];
  const notifications = []; // { userName, text, conversationId }

  for (const proj of projects) {
    const curr = snapshotFor(proj);
    const prev = priorMap.get(proj.ID);
    // Always update the snapshot after processing
    if (!prev) {
      await r.set(`snap:project:${proj.ID}`, JSON.stringify(curr));
      continue;
    }
    const changes = diff(prev, curr);
    if (changes.length === 0) continue;
    for (const ch of changes) {
      events.push({ projectId: proj.ID, projectName: proj.name, ...ch });
      const names = affectedNamesFor(ch, proj);
      for (const name of names) {
        const sub = subsByName.get(normName(name));
        if (!sub || !sub.enabled || !sub.conversationId) continue;
        const text = buildNotificationText(ch, proj, sub.userName);
        notifications.push({ userName: sub.userName, text, conversationId: sub.conversationId });
      }
    }
    await r.set(`snap:project:${proj.ID}`, JSON.stringify(curr));
  }

  // Dispatch notifications. Dedupe identical (user, text) within this tick.
  const seen = new Set();
  const sendResults = [];
  const { getConversationRef } = require('./schedule-store');
  for (const n of notifications) {
    const key = `${n.conversationId}|${n.text}`;
    if (seen.has(key)) continue;
    seen.add(key);
    try {
      const ref = await getConversationRef(n.conversationId);
      if (!ref) { sendResults.push({ user: n.userName, status: 'no-ref' }); continue; }
      await sendProactive(ref, n.text);
      sendResults.push({ user: n.userName, status: 'sent' });
    } catch (err) {
      sendResults.push({ user: n.userName, status: 'error', error: err.message });
    }
  }

  return {
    projectsPolled: projects.length,
    isFirstRun,
    changes: events.length,
    notificationsQueued: notifications.length,
    notificationsSent: sendResults.filter(r => r.status === 'sent').length,
    sendResults,
  };
}

module.exports = { detectAndNotify };
