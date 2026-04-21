// Adaptive Card builders for Teams. All return Bot Framework attachment
// objects ready to drop into `context.sendActivity({attachments: [...]})`.
//
// Kept to Adaptive Card v1.4 for broad Teams compatibility.

const CARD_CONTENT_TYPE = 'application/vnd.microsoft.card.adaptive';
const CARD_VERSION = '1.4';

function attach(card) {
  return { contentType: CARD_CONTENT_TYPE, content: card };
}

function baseCard(body, actions) {
  const card = {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: CARD_VERSION,
    body,
  };
  if (actions && actions.length) card.actions = actions;
  return card;
}

function shortName(name) {
  return String(name || '').replace(/^FY27_/, '').replace(/_/g, ' ').trim();
}

function formatMD(dateStr) {
  if (!dateStr) return '';
  const fixed = String(dateStr).replace(/(\d{2}):(\d{3})/, '$1.$2');
  const d = new Date(fixed);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('en-US', {
    timeZone: 'America/Chicago',
    weekday: 'short',
    month: 'numeric',
    day: 'numeric',
  });
}

function channelLabel(proj) {
  const ch = String(proj.channel || '').toLowerCase();
  const type = String(proj.projectType || '').toLowerCase();
  const name = String(proj.name || '').toLowerCase();
  if (type.includes('loyalty') || name.includes('loyalty')) return 'Loyalty';
  if (ch.includes('email')) return 'Email';
  if (ch.includes('text') || ch.includes('sms') || ch.includes('push')) return 'Text/Push';
  if (ch.includes('direct mail')) return 'Direct Mail';
  if (ch.includes('paid media')) return 'Paid Media';
  if (ch.includes('organic social')) return 'Social';
  if (ch.includes('website')) return 'Web';
  if (Array.isArray(proj.channel) && proj.channel.length) return String(proj.channel[0]);
  return proj.channel ? String(proj.channel) : '';
}

function projectRow(p, opts = {}) {
  const ch = channelLabel(p);
  const chBit = ch ? ` · ${ch}` : '';
  const dateLabel = opts.dateLabel || '';
  const statusEmoji = opts.statusEmoji ? `${opts.statusEmoji} ` : '';
  const leftCol = {
    type: 'Column',
    width: 'stretch',
    items: [
      {
        type: 'TextBlock',
        text: `${statusEmoji}**${shortName(p.name || p.projectName)}**${chBit}`,
        wrap: true,
        size: 'Default',
      },
      {
        type: 'TextBlock',
        text: `${p.designer || 'TBD'} / ${p.copywriter || 'TBD'}${dateLabel ? ` · ${dateLabel}` : ''}`,
        spacing: 'None',
        isSubtle: true,
        size: 'Small',
        wrap: true,
      },
    ],
  };
  if (opts.subline) {
    leftCol.items.push({
      type: 'TextBlock',
      text: opts.subline,
      spacing: 'None',
      isSubtle: true,
      size: 'Small',
      wrap: true,
    });
  }
  const rightCol = p.projectUrl
    ? {
        type: 'Column',
        width: 'auto',
        verticalContentAlignment: 'Center',
        items: [
          {
            type: 'ActionSet',
            actions: [{ type: 'Action.OpenUrl', title: 'Open', url: p.projectUrl }],
          },
        ],
      }
    : { type: 'Column', width: 'auto', items: [] };
  return {
    type: 'ColumnSet',
    spacing: 'Small',
    columns: [leftCol, rightCol],
  };
}

// ---------- Agenda card (weekly review digest) ----------
function buildAgendaCard({ window, sections, personLabel }) {
  const windowLabel = window === 'nextweek' ? 'Next Week' : window === 'thisweek' ? 'This Week' : window;
  const total = sections.reduce((s, sec) => s + (sec.projects ? sec.projects.length : 0), 0);
  const body = [
    {
      type: 'TextBlock',
      size: 'Large',
      weight: 'Bolder',
      text: `📋 Review Agenda — ${windowLabel}${personLabel ? ` · ${personLabel}` : ''}`,
      wrap: true,
    },
    {
      type: 'TextBlock',
      text: total === 0 ? 'Nothing scheduled. You\'re clear. ✅' : `${total} project${total === 1 ? '' : 's'}`,
      spacing: 'Small',
      isSubtle: true,
    },
  ];
  for (const section of sections) {
    const projs = section.projects || [];
    body.push({
      type: 'TextBlock',
      text: `${section.title}${projs.length ? ` (${projs.length})` : ''}`,
      weight: 'Bolder',
      size: 'Medium',
      spacing: 'Medium',
      separator: true,
      wrap: true,
    });
    if (!projs.length) {
      body.push({
        type: 'TextBlock',
        text: '_Nothing scheduled_',
        spacing: 'Small',
        isSubtle: true,
      });
      continue;
    }
    for (const p of projs) {
      body.push(projectRow(p, { dateLabel: p.dateLabel || formatMD(p.reviewDate) }));
    }
  }
  return attach(baseCard(body));
}

// ---------- Notification card (per change event) ----------
function buildNotificationCard({ emoji, headline, subline, projectName, url, extraActions }) {
  const body = [
    {
      type: 'TextBlock',
      text: `${emoji || '🔔'} ${headline}`,
      weight: 'Bolder',
      size: 'Medium',
      wrap: true,
    },
  ];
  if (projectName) {
    body.push({
      type: 'TextBlock',
      text: shortName(projectName),
      isSubtle: true,
      spacing: 'None',
      wrap: true,
    });
  }
  if (subline) {
    body.push({
      type: 'TextBlock',
      text: subline,
      spacing: 'Small',
      wrap: true,
    });
  }
  const actions = [];
  if (url) actions.push({ type: 'Action.OpenUrl', title: 'Open in Workfront', url });
  if (extraActions) actions.push(...extraActions);
  return attach(baseCard(body, actions));
}

// ---------- Proof readiness card ----------
function buildProofReadinessCard({ reviewType, windowLabel, projects, personLabel }) {
  const typeLabel = {
    creative: 'Creative Review',
    marketing: 'Marketing Review',
    exec: 'Exec Review',
  }[reviewType] || 'Review';
  const needs = (projects || []).filter(p => p.needsProof);
  const body = [
    {
      type: 'TextBlock',
      size: 'Large',
      weight: 'Bolder',
      text: `✅ Proof Readiness — ${typeLabel} ${windowLabel || ''}${personLabel ? ` · ${personLabel}` : ''}`,
      wrap: true,
    },
    {
      type: 'TextBlock',
      text: needs.length === 0
        ? '🎉 Everyone has a proof posted — you\'re all clear!'
        : `${needs.length} of ${projects.length} project${projects.length === 1 ? '' : 's'} still need${needs.length === 1 ? 's' : ''} a proof`,
      spacing: 'Small',
      color: needs.length ? 'Warning' : 'Good',
    },
  ];
  for (const p of needs) {
    const emoji = p.reason === 'no-proof-posted' ? '🔴' : '🟡';
    const subline = p.reason === 'no-proof-posted'
      ? 'No proof posted yet'
      : `No new version since previous review${p.latestProofVersionAt ? ` (last: ${formatMD(p.latestProofVersionAt)})` : ''}`;
    body.push({
      type: 'Container',
      spacing: 'Small',
      separator: true,
      items: [
        projectRow(
          { ...p, name: p.projectName, projectUrl: p.projectUrl },
          { statusEmoji: emoji, subline, dateLabel: formatMD(p.reviewDate) }
        ),
      ],
    });
  }
  return attach(baseCard(body));
}

// ---------- Workload card ----------
function buildWorkloadCard({ windowLabel, leaderboard, person, personTotal, personBreakdown }) {
  const body = [
    {
      type: 'TextBlock',
      size: 'Large',
      weight: 'Bolder',
      text: person
        ? `👤 ${person}'s Load — ${windowLabel || ''}`
        : `👥 Team Workload — ${windowLabel || ''}`,
      wrap: true,
    },
  ];
  if (person) {
    body.push({
      type: 'TextBlock',
      text: `${personTotal || 0} project${personTotal === 1 ? '' : 's'}`,
      spacing: 'Small',
      isSubtle: true,
    });
    if (personBreakdown) {
      body.push({
        type: 'FactSet',
        spacing: 'Small',
        facts: [
          { title: 'Designer', value: String(personBreakdown.designer || 0) },
          { title: 'Copywriter', value: String(personBreakdown.copywriter || 0) },
          { title: 'PM', value: String(personBreakdown.pm || 0) },
        ],
      });
    }
    return attach(baseCard(body));
  }
  const rows = (leaderboard || []).slice(0, 10);
  if (!rows.length) {
    body.push({ type: 'TextBlock', text: 'Nobody has projects in this window.', spacing: 'Medium' });
    return attach(baseCard(body));
  }
  const max = Math.max(...rows.map(r => r.total || 1), 1);
  const BAR_MAX = 14;
  for (const r of rows) {
    const bars = Math.max(1, Math.round((r.total / max) * BAR_MAX));
    const bar = '█'.repeat(bars) + '░'.repeat(BAR_MAX - bars);
    body.push({
      type: 'ColumnSet',
      spacing: 'Small',
      columns: [
        {
          type: 'Column', width: 110,
          items: [{ type: 'TextBlock', text: r.name, wrap: true, size: 'Small' }],
        },
        {
          type: 'Column', width: 'stretch',
          items: [{ type: 'TextBlock', text: bar, fontType: 'Monospace', spacing: 'None', size: 'Small' }],
        },
        {
          type: 'Column', width: 40,
          items: [{ type: 'TextBlock', text: String(r.total), horizontalAlignment: 'Right', weight: 'Bolder', size: 'Small' }],
        },
      ],
    });
  }
  return attach(baseCard(body));
}

module.exports = {
  buildAgendaCard,
  buildNotificationCard,
  buildProofReadinessCard,
  buildWorkloadCard,
  shortName,
  formatMD,
  channelLabel,
};
