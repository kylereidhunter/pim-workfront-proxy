// Builds the text Pim proactively sends when a schedule fires.
// Kept dumb/deterministic — no LLM in the proactive path so we don't
// burn tokens on every scheduled fire and we don't risk hallucination.

const PROXY_BASE = process.env.PROXY_BASE_URL || 'https://pim-workfront-proxy.vercel.app';

async function fetchJSON(path) {
  const r = await fetch(`${PROXY_BASE}${path}`);
  if (!r.ok) throw new Error(`Proxy ${path} -> ${r.status}`);
  return r.json();
}

function shortName(name) {
  return (name || '').replace(/^FY27_/, '').replace(/_/g, ' ').trim();
}

function fmtMD(dateStr) {
  if (!dateStr) return '';
  const fixed = dateStr.replace(/(\d{2}):(\d{3})/, '$1.$2');
  const d = new Date(fixed);
  if (isNaN(d.getTime())) return '';
  const wd = d.toLocaleDateString('en-US', { weekday: 'short', timeZone: 'America/Chicago' });
  const md = `${d.getMonth() + 1}/${d.getDate()}`;
  return `${wd} ${md}`;
}

function labelChannel(proj) {
  const chRaw = proj.channel;
  const ch = String(chRaw || '').toLowerCase();
  const type = String(proj.projectType || '').toLowerCase();
  const name = String(proj.name || '').toLowerCase();
  if (type.includes('loyalty') || name.includes('loyalty') || ch.includes('insider')) return 'Loyalty';
  // Multi-channel projects ("Email/SMS/Push" etc) get every matching token
  // listed — Text/Push projects shouldn't be silently labeled "Email".
  const parts = [];
  if (ch.includes('email')) parts.push('Email');
  if (ch.includes('text') || ch.includes('sms') || ch.includes('push')) parts.push('Text/Push');
  if (ch.includes('direct mail')) parts.push('Direct Mail');
  if (ch.includes('paid media')) parts.push('Paid Media');
  if (ch.includes('organic social') || ch.includes('social')) parts.push('Social');
  if (ch.includes('website') || ch.includes('web')) parts.push('Web');
  if (ch.includes('in store') || ch.includes('in-store')) parts.push('In-Store');
  if (parts.length) return parts.join(' + ');
  if (Array.isArray(chRaw) && chRaw.length) return String(chRaw[0]);
  if (chRaw) return String(chRaw);
  return '';
}

function renderProjectLine(p) {
  const label = labelChannel(p);
  const chBit = label ? ` (${label})` : '';
  const designer = p.designer || 'TBD';
  const copy = p.copywriter || 'TBD';
  const url = p.projectUrl ? ` — [Workfront](${p.projectUrl})` : '';
  return `- **${shortName(p.name)}**${chBit} — ${designer} / ${copy}${url}`;
}

function parseWFDateMs(s) {
  if (!s) return 0;
  const fixed = String(s).replace(/(\d{2}):(\d{3})/, '$1.$2');
  const d = new Date(fixed);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function groupByDay(projects, dateKey) {
  // Sort projects by their date first so Map insertion order is chronological.
  const sorted = projects
    .slice()
    .sort((a, b) => parseWFDateMs(a[dateKey]) - parseWFDateMs(b[dateKey]));
  const map = new Map();
  for (const p of sorted) {
    const dateStr = fmtMD(p[dateKey]);
    if (!dateStr) continue;
    if (!map.has(dateStr)) map.set(dateStr, []);
    map.get(dateStr).push(p);
  }
  return map;
}

// Scoped digest builder. Accepts:
//   window     - 'thisweek' | 'nextweek' | ...
//   reviewType - 'creative' | 'marketing' | 'exec' | undefined (all three)
//   person     - optional person name filter
async function buildWeeklyReviewsDigest({ window = 'nextweek', reviewType, person } = {}) {
  const wantTypes = reviewType
    ? [reviewType]
    : ['creative', 'marketing', 'exec'];

  const results = await Promise.all(
    wantTypes.map(t => {
      const q = new URLSearchParams({ reviewType: t, window });
      if (person) q.set('person', person);
      return fetchJSON(`/reviews?${q}`);
    })
  );
  const byType = Object.fromEntries(wantTypes.map((t, i) => [t, results[i]]));

  const windowLabel = window === 'nextweek' ? 'next week' : window === 'thisweek' ? 'this week' : window;
  const total = results.reduce((s, r) => s + (r.count || 0), 0);
  const personBit = person ? ` for ${person}` : '';
  const scopeBit = reviewType
    ? (reviewType === 'creative' ? 'Creative Review' : reviewType === 'marketing' ? 'Marketing Review' : 'Exec Review')
    : 'Reviews';

  let header;
  if (total === 0) {
    header = `**${scopeBit} ${windowLabel}${personBit}** — all clear. ✅\n`;
    return header;
  }
  header = reviewType
    ? `**${scopeBit} — ${windowLabel}${personBit}** (${total} project${total === 1 ? '' : 's'})\n`
    : `**Pim's Weekly Digest — Reviews for ${windowLabel}${personBit}** (${total} project${total === 1 ? '' : 's'})\n`;

  const sections = [];
  const dateKeyFor = {
    creative: 'creativeReviewDate',
    marketing: 'marketingReviewDate',
    exec: 'execReviewDate',
  };
  const titleFor = {
    creative: 'Creative Review',
    marketing: 'Marketing Review',
    exec: 'Exec Review',
  };

  for (const t of wantTypes) {
    const endpoint = byType[t];
    if (!endpoint.projects || endpoint.projects.length === 0) {
      // Only add an empty-section note in the multi-type view.
      if (!reviewType) sections.push(`\n**${titleFor[t]}** — nothing scheduled`);
      continue;
    }
    const byDay = groupByDay(endpoint.projects, dateKeyFor[t]);
    const lines = reviewType
      ? []
      : [`\n**${titleFor[t]}** (${endpoint.projects.length})`];
    if (byDay.size === 0) {
      endpoint.projects.forEach(p => lines.push(renderProjectLine(p)));
    } else {
      for (const [day, items] of byDay) {
        lines.push(`\n_${day}_`);
        items.forEach(p => lines.push(renderProjectLine(p)));
      }
    }
    sections.push(lines.join('\n'));
  }

  const footer = (!reviewType && total > 0)
    ? `\n\n_Proofs: CR due 3pm Mon · MKT due 1pm Tue · Exec due 1pm Wed._`
    : '';

  return header + sections.join('\n') + footer;
}

async function buildProofDueCountdown({ reviewType = 'creative', when = 'tomorrow' } = {}) {
  // Fetch proof readiness for the given review window.
  // Default: CR proofs due today, review tomorrow.
  const windowParam = when === 'today' ? 'thisweek' : 'thisweek';
  const r = await fetchJSON(`/proof-readiness?reviewType=${reviewType}&window=${windowParam}`);
  if (!r.total) {
    return `☑️ No ${reviewType === 'creative' ? 'Creative' : reviewType === 'marketing' ? 'Marketing' : 'Exec'} Review scheduled for ${when}. Enjoy the breather!`;
  }
  const reviewLabel = { creative: 'Creative Review', marketing: 'MKT Review', exec: 'Exec Review' }[reviewType] || 'Review';
  const dueBy = { creative: 'Mon 3pm', marketing: 'Tue 1pm', exec: 'Wed 1pm' }[reviewType] || 'soon';
  const missing = (r.projects || []).filter(p => p.needsProof);
  const posted = r.total - missing.length;
  let msg = `⏰ **${reviewLabel} proofs due ${dueBy}** — ${posted} of ${r.total} posted.\n`;
  if (missing.length === 0) {
    msg += `All caught up! 🎉`;
  } else {
    msg += `\nStill missing:\n`;
    missing.forEach(p => {
      const sn = shortName(p.projectName);
      const who = p.designer || 'TBD';
      const url = p.projectUrl ? ` — [Open](${p.projectUrl})` : '';
      msg += `- **${sn}** — ${who}${url}\n`;
    });
  }
  return msg;
}

async function buildMessage(kind, args) {
  switch (kind) {
    case 'weekly-reviews-digest':
      return buildWeeklyReviewsDigest(args || {});
    case 'proof-due-countdown':
      return buildProofDueCountdown(args || {});
    case 'reminder-text':
      return args && args.text ? args.text : '⏰ Reminder!';
    default:
      return `⏰ Scheduled message fired (unknown kind: ${kind})`;
  }
}

module.exports = { buildMessage, buildWeeklyReviewsDigest };
