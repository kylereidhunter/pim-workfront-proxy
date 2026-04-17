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
  if (type.includes('loyalty') || name.includes('loyalty')) return 'Loyalty';
  // Email is the primary channel when present, even if it's bundled with text/push.
  if (ch.includes('email')) return 'Email';
  if (ch.includes('text') || ch.includes('sms') || ch.includes('push')) return 'Text/Push';
  if (ch.includes('direct mail')) return 'Direct Mail';
  if (ch.includes('paid media')) return 'Paid Media';
  if (ch.includes('organic social')) return 'Social';
  if (ch.includes('website')) return 'Web';
  if (ch.includes('insider')) return 'Loyalty';
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

function groupByDay(projects, dateKey) {
  const map = new Map();
  for (const p of projects) {
    const dateStr = fmtMD(p[dateKey]);
    if (!dateStr) continue;
    if (!map.has(dateStr)) map.set(dateStr, []);
    map.get(dateStr).push(p);
  }
  return map;
}

async function buildWeeklyReviewsDigest({ window = 'nextweek' } = {}) {
  const [cr, mkt, execR] = await Promise.all([
    fetchJSON(`/reviews?reviewType=creative&window=${window}`),
    fetchJSON(`/reviews?reviewType=marketing&window=${window}`),
    fetchJSON(`/reviews?reviewType=exec&window=${window}`),
  ]);

  const total = cr.count + mkt.count + execR.count;
  const windowLabel = window === 'nextweek' ? 'next week' : window === 'thisweek' ? 'this week' : window;
  let header;
  if (total === 0) {
    header = `**Reviews for ${windowLabel}** — all clear. Nothing on the calendar. ✅\n`;
  } else {
    header = `**Pim's Weekly Digest — Reviews for ${windowLabel}** (${total} project${total === 1 ? '' : 's'})\n`;
  }

  const sections = [];

  function section(title, endpoint, dateKey) {
    if (!endpoint.projects || endpoint.projects.length === 0) return;
    const byDay = groupByDay(endpoint.projects, dateKey);
    const lines = [`\n**${title}** (${endpoint.projects.length})`];
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

  section('Creative Review', cr, 'creativeReviewDate');
  section('Marketing Review', mkt, 'marketingReviewDate');
  section('Exec Review', execR, 'execReviewDate');

  const footer = total > 0
    ? `\n\n_Proofs: CR due 3pm Mon · MKT due 1pm Tue · Exec due 1pm Wed._`
    : '';

  return header + sections.join('\n') + footer;
}

async function buildMessage(kind, args) {
  switch (kind) {
    case 'weekly-reviews-digest':
      return buildWeeklyReviewsDigest(args || {});
    case 'reminder-text':
      return args && args.text ? args.text : '⏰ Reminder!';
    default:
      return `⏰ Scheduled message fired (unknown kind: ${kind})`;
  }
}

module.exports = { buildMessage, buildWeeklyReviewsDigest };
