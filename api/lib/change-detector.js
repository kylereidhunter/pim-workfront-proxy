// Polls the Workfront proxy, diffs against the last snapshot, and emits
// change events. Subscribed users affected by each event get a DM.
//
// Events detected:
//   - Assignee changes    (designer / copywriter / pm)
//   - Review date changes (creative / marketing / exec)
//   - Live date changes
//   - Document uploads    (new docID appears on a project)
//   - New proof versions  (version bump on existing doc)
//   - Proof status changes (pending -> approved / rejected / etc.)
//   - New project comments (Updates tab notes)
//
// First run after deploy: builds baseline snapshots, fires ZERO events.

const { Redis } = require('@upstash/redis');
const { listEnabled, normName } = require('./subscriptions');
const { sendProactive } = require('./proactive');

const PROXY_BASE = process.env.PROXY_BASE_URL || 'https://pim-workfront-proxy.vercel.app';
const TRACKED_PROJECT_FIELDS = [
  'designer',
  'copywriter',
  'pm',
  'creativeReviewDate',
  'marketingReviewDate',
  'execReviewDate',
  'liveDate',
];

function kv() {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) throw new Error('KV not configured');
  return new Redis({ url, token });
}

async function fetchJSON(path) {
  const r = await fetch(`${PROXY_BASE}${path}`);
  if (!r.ok) throw new Error(`proxy ${path} -> ${r.status}`);
  return r.json();
}

function fieldLabel(field) {
  return {
    designer: 'Lead Designer',
    copywriter: 'Lead Copywriter',
    pm: 'PM',
    creativeReviewDate: 'Creative Review',
    marketingReviewDate: 'Marketing Review',
    execReviewDate: 'Exec Review',
    liveDate: 'Live Date',
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

function projectSnapshotFrom(proj) {
  const out = {};
  for (const f of TRACKED_PROJECT_FIELDS) out[f] = proj[f] == null ? null : proj[f];
  return out;
}

function diffProject(prev, curr) {
  const changes = [];
  for (const f of TRACKED_PROJECT_FIELDS) {
    const a = prev[f] == null ? null : prev[f];
    const b = curr[f] == null ? null : curr[f];
    if (a !== b) changes.push({ field: f, from: a, to: b });
  }
  return changes;
}

// Assignees on a project after a change, used to figure out who gets DMed.
function currentAssigneeNames(proj) {
  const set = new Set();
  if (proj.designer) set.add(proj.designer);
  if (proj.copywriter) set.add(proj.copywriter);
  if (proj.pm) set.add(proj.pm);
  return [...set];
}

function affectedForChange(change, currProj) {
  if (isAssigneeField(change.field)) {
    const s = new Set();
    if (change.to) s.add(change.to);
    if (change.from) s.add(change.from);
    // PM should also know about designer/copywriter swaps
    if (currProj.pm) s.add(currProj.pm);
    return [...s];
  }
  return currentAssigneeNames(currProj);
}

function buildProjectChangeText(change, currProj, recipientName) {
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
    const who = change.to || 'TBD';
    return `🔁 **${role}** on **${proj}** changed to **${who}** (prev: ${change.from || 'unset'})${openLink}`;
  }
  const label = fieldLabel(change.field);
  return `📅 **${label}** for **${proj}** moved from _${formatDate(change.from)}_ to _${formatDate(change.to)}_${openLink}`;
}

function buildDocText(event, currProj) {
  const proj = shortName(currProj.name);
  const url = currProj.projectUrl;
  const openLink = url ? ` — [Open in Workfront](${url})` : '';
  switch (event.type) {
    case 'doc-uploaded':
      return `📎 New document **${event.fileName || event.docName}** uploaded to **${proj}**${openLink}`;
    case 'proof-version-bump':
      return `🆕 New proof version (v${event.version}) on **${event.docName}** (project **${proj}**)${openLink}`;
    case 'proof-status-change':
      return `🔎 Proof status on **${event.docName}** (project **${proj}**) changed: _${event.from || 'pending'}_ → **${event.to || 'updated'}**${openLink}`;
    default:
      return `Document update on **${proj}**${openLink}`;
  }
}

function buildCommentText(note, currProj) {
  const proj = shortName(currProj.name);
  const url = currProj.projectUrl;
  const openLink = url ? ` — [Open in Workfront](${url})` : '';
  const author = note.ownerName || 'someone';
  const snippet = String(note.text || '').replace(/\s+/g, ' ').slice(0, 200);
  const suffix = snippet.length < String(note.text || '').length ? '…' : '';
  return `💬 **${author}** commented on **${proj}**: _"${snippet}${suffix}"_${openLink}`;
}

async function detectAndNotify() {
  const r = kv();
  const summary = {
    projectsPolled: 0,
    isFirstRun: false,
    changes: 0,
    notificationsQueued: 0,
    notificationsSent: 0,
    errors: [],
  };

  // --- Subscriber map ---
  const subs = await listEnabled();
  const subsByName = new Map(subs.map(s => [s.userNameLower, s]));
  const { getConversationRef } = require('./schedule-store');

  // --- Fetch current state ---
  const [projectsRes, docsRes, notesRes] = await Promise.all([
    fetchJSON('/upcoming-reviews?name=FY27').catch(e => ({ error: e.message })),
    fetchJSON('/docs?name=FY27').catch(e => ({ error: e.message })),
    fetchJSON('/updates?name=FY27&sinceHours=24').catch(e => ({ error: e.message })),
  ]);
  if (projectsRes.error) summary.errors.push(`projects: ${projectsRes.error}`);
  if (docsRes.error) summary.errors.push(`docs: ${docsRes.error}`);
  if (notesRes.error) summary.errors.push(`notes: ${notesRes.error}`);

  const projects = (projectsRes.data || []);
  const docs = (docsRes.data || []);
  const notes = (notesRes.data || []);
  summary.projectsPolled = projects.length;

  const projectById = new Map(projects.map(p => [p.ID, p]));
  const events = []; // { text, affectedNames, projectID }

  // --- Project-level diff ---
  if (projects.length) {
    const keys = projects.map(p => `snap:project:${p.ID}`);
    const prior = await r.mget(...keys);
    summary.isFirstRun = prior.every(x => x == null);
    for (let i = 0; i < projects.length; i++) {
      const proj = projects[i];
      const prev = prior[i] == null ? null : (typeof prior[i] === 'string' ? JSON.parse(prior[i]) : prior[i]);
      const curr = projectSnapshotFrom(proj);
      if (!prev) {
        await r.set(`snap:project:${proj.ID}`, JSON.stringify(curr));
        continue;
      }
      const changes = diffProject(prev, curr);
      for (const ch of changes) {
        const affected = affectedForChange(ch, proj);
        events.push({
          projectID: proj.ID,
          change: ch,
          affected,
          formatFor: (name) => buildProjectChangeText(ch, proj, name),
        });
      }
      if (changes.length) await r.set(`snap:project:${proj.ID}`, JSON.stringify(curr));
    }
  }

  // --- Document / proof diff ---
  // Snapshot shape per doc: { version, proofStatus }
  if (docs.length) {
    const docKeys = docs.map(d => `snap:doc:${d.docID}`);
    const prior = await r.mget(...docKeys);
    const firstDocRun = prior.every(x => x == null);
    for (let i = 0; i < docs.length; i++) {
      const d = docs[i];
      const proj = projectById.get(d.projectID);
      if (!proj) continue; // project dropped off
      const prev = prior[i] == null ? null : (typeof prior[i] === 'string' ? JSON.parse(prior[i]) : prior[i]);
      const curr = {
        version: d.version,
        proofStatus: d.proofStatus || null,
        proofDecision: d.proofDecision || null,
      };
      if (!prev) {
        // First time seeing this doc — fire upload event (but only if not first run overall)
        if (!firstDocRun) {
          const affected = currentAssigneeNames(proj);
          const ev = { type: 'doc-uploaded', docName: d.name, fileName: d.fileName };
          events.push({
            projectID: proj.ID,
            affected,
            formatFor: () => buildDocText(ev, proj),
          });
        }
        await r.set(`snap:doc:${d.docID}`, JSON.stringify(curr));
        continue;
      }
      // Version bump?
      if (d.version != null && prev.version != null && d.version > prev.version) {
        const affected = currentAssigneeNames(proj);
        const ev = { type: 'proof-version-bump', docName: d.name, version: d.version };
        events.push({
          projectID: proj.ID,
          affected,
          formatFor: () => buildDocText(ev, proj),
        });
      }
      // Proof status change?
      if (d.hasProof && d.proofStatus !== prev.proofStatus) {
        const affected = currentAssigneeNames(proj);
        const ev = {
          type: 'proof-status-change',
          docName: d.name,
          from: prev.proofStatus,
          to: d.proofStatus,
        };
        events.push({
          projectID: proj.ID,
          affected,
          formatFor: () => buildDocText(ev, proj),
        });
      }
      await r.set(`snap:doc:${d.docID}`, JSON.stringify(curr));
    }
  }

  // --- Comments / Updates tab ---
  // Use last-seen note ID per project. Store: snap:notes:<projectID> = lastEntryDate ISO string.
  if (notes.length) {
    const notesByProject = new Map();
    for (const n of notes) {
      if (!n.projectID) continue;
      const arr = notesByProject.get(n.projectID) || [];
      arr.push(n);
      notesByProject.set(n.projectID, arr);
    }
    for (const [projID, projNotes] of notesByProject) {
      const proj = projectById.get(projID);
      if (!proj) continue;
      const prevLastRaw = await r.get(`snap:notes:${projID}`);
      const prevLast = prevLastRaw ? new Date(prevLastRaw).getTime() : null;
      // Find max entryDate in this batch
      const sorted = projNotes
        .map(n => ({ n, ts: n.entryDate ? new Date(String(n.entryDate).replace(/(\d{2}):(\d{3})/, '$1.$2')).getTime() : 0 }))
        .sort((a, b) => b.ts - a.ts);
      const maxTs = sorted[0].ts;
      if (prevLast == null) {
        // First time tracking this project's comments — set baseline, no events
        await r.set(`snap:notes:${projID}`, new Date(maxTs).toISOString());
        continue;
      }
      const fresh = sorted.filter(x => x.ts > prevLast).map(x => x.n);
      for (const note of fresh) {
        const affected = currentAssigneeNames(proj);
        events.push({
          projectID: proj.ID,
          affected,
          formatFor: () => buildCommentText(note, proj),
        });
      }
      if (fresh.length) await r.set(`snap:notes:${projID}`, new Date(maxTs).toISOString());
    }
  }

  summary.changes = events.length;

  // --- Dispatch DMs ---
  const seen = new Set();
  const sendResults = [];
  for (const ev of events) {
    for (const name of ev.affected) {
      const sub = subsByName.get(normName(name));
      if (!sub || !sub.enabled || !sub.conversationId) continue;
      const text = ev.formatFor(sub.userName);
      const key = `${sub.conversationId}|${text}`;
      if (seen.has(key)) continue;
      seen.add(key);
      summary.notificationsQueued++;
      try {
        const ref = await getConversationRef(sub.conversationId);
        if (!ref) { sendResults.push({ user: sub.userName, status: 'no-ref' }); continue; }
        await sendProactive(ref, text);
        summary.notificationsSent++;
        sendResults.push({ user: sub.userName, status: 'sent' });
      } catch (err) {
        sendResults.push({ user: sub.userName, status: 'error', error: err.message });
      }
    }
  }
  summary.sendResults = sendResults;
  return summary;
}

module.exports = { detectAndNotify };
