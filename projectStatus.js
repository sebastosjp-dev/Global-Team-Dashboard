/**
 * projectStatus.js — "Project Status" view for the PROJECTS tab.
 *
 * A management view of all non-POC projects (Installation, Renewal, Maintenance,
 * Pipeline): score cards that filter, a frozen table on the left and a Gantt
 * timeline on the right, with expandable per-project steps from the STEPS tab.
 *
 * Sheet contract
 *   PROJECTS: one row per update. The row with Latest? = TRUE is the current status.
 *             Optional helper columns "% Done" and "Short Note".
 *   STEPS:    Project | Step | Status | Start | End or Target | Owner | Working with | Note
 * @module projectStatus
 */

/* ═══════════════════════════════════════════════════════════════
   Constants & helpers
   ═══════════════════════════════════════════════════════════════ */

const DAY = 86400000;

const HEALTH = {
    'Blocked':   { color: '#d9433f', rank: 0 },
    'Overdue':   { color: '#ee8a3c', rank: 1 },
    'No update': { color: '#e9b021', rank: 2 },
    'On track':  { color: '#22a55a', rank: 3 },
    'On hold':   { color: '#8b80c9', rank: 4 },
    'Closed':    { color: '#a3a9b1', rank: 5 },
};
const OPEN = ['Blocked', 'Overdue', 'No update', 'On track'];

const STEP = {
    'Done':        { icon: '✓', cls: 'done' },
    'In progress': { icon: '●', cls: 'progress' },
    'Blocked':     { icon: '!', cls: 'blocked' },
    'Next':        { icon: '○', cls: 'next' },
};

const ZOOM = { Day: 32, Week: 12, Month: 4.5, Quarter: 1.8 };   // px per day
const FORMER_STAFF = ['kamal'];
const EST_DAYS = 14;

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/** Parse a sheet date to local midnight. SheetJS dates can be a few seconds off, so round. */
function toDay(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) {
        if (isNaN(v)) return null;
        const r = new Date(v.getTime() + 12 * 3600000);
        return new Date(r.getFullYear(), r.getMonth(), r.getDate());
    }
    if (typeof v === 'number' && v > 30000) return toDay(new Date(Math.round((v - 25569) * DAY)));
    const s = String(v).trim();
    const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s);
    return isNaN(d) ? null : toDay(d);
}

const todayDate = () => { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); };
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const daysBetween = (a, b) => Math.round((b - a) / DAY);

const fmtShort = (d) => d ? `${d.getDate()} ${MONTHS[d.getMonth()]}` : '';
const fmtLong = (d) => d ? `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}` : '';

function fmtDur(days) {
    if (days < 14) return `${Math.max(1, days)}d`;
    if (days < 60) return `${Math.round(days / 7)}w`;
    return `${Math.round(days / 30)}mo`;
}

function normHealth(v) {
    const s = String(v || '').trim().toLowerCase();
    return Object.keys(HEALTH).find(h => h.toLowerCase() === s) || (s ? 'On track' : 'No update');
}

function normStep(v) {
    const s = String(v || '').trim().toLowerCase();
    if (s.startsWith('done') || s === 'complete' || s === 'completed') return 'Done';
    if (s.includes('progress')) return 'In progress';
    if (s.startsWith('block')) return 'Blocked';
    return 'Next';
}

/** WhaTap staff names only: drop brackets and former staff. */
function cleanOwners(v) {
    return String(v || '')
        .replace(/[[\]()]/g, '')
        .split(/[,/&]+/)
        .map(s => s.trim())
        .filter(s => s && !FORMER_STAFF.includes(s.toLowerCase()))
        .join(', ');
}

function parsePct(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return v > 1 ? v / 100 : v;
    const n = parseFloat(String(v).replace('%', ''));
    if (isNaN(n)) return null;
    return String(v).includes('%') || n > 1 ? n / 100 : n;
}

const isTrue = (v) => v === true || String(v).trim().toUpperCase() === 'TRUE';

/** Map a row's headers by a normalised name (letters and % only). */
function columnMap(rows) {
    const keys = rows.length ? Object.keys(rows[0]) : [];
    const map = {};
    keys.forEach(k => { map[k.toLowerCase().replace(/[^a-z%]/g, '')] = k; });
    return (name) => map[name];
}

/* ═══════════════════════════════════════════════════════════════
   Data model
   ═══════════════════════════════════════════════════════════════ */

function buildSteps(stepRows) {
    const byProject = new Map();
    if (!stepRows || !stepRows.length) return byProject;
    const col = columnMap(stepRows);
    const k = {
        project: col('project'), step: col('step'), status: col('status'),
        start: col('start'), end: col('endortarget') || col('end') || col('target'),
        owner: col('owner'), withWho: col('workingwith'), note: col('note'),
    };
    stepRows.forEach(r => {
        const project = String(r[k.project] || '').trim();
        const name = String(r[k.step] || '').trim();
        if (!project || !name) return;
        const key = project.toLowerCase();
        if (!byProject.has(key)) byProject.set(key, []);
        byProject.get(key).push({
            name,
            status: normStep(r[k.status]),
            start: toDay(r[k.start]),
            end: toDay(r[k.end]),
            owner: cleanOwners(r[k.owner]),
            withWho: String(r[k.withWho] || '').trim(),
            note: String(r[k.note] || '').trim(),
        });
    });
    byProject.forEach(estimateSteps);
    return byProject;
}

/**
 * Fill missing step dates next to their neighbours (~14 days each) and flag them
 * as estimated. Real dates are never changed.
 */
function estimateSteps(steps) {
    steps.forEach(s => { s.estStart = !s.start; s.estEnd = !s.end; s.eStart = s.start; s.eEnd = s.end; });
    for (let i = 0; i < steps.length; i++) {
        const s = steps[i], prev = steps[i - 1], next = steps.slice(i + 1).find(n => n.start || n.end);
        if (!s.eStart && prev?.eEnd) s.eStart = prev.eEnd;
        if (!s.eEnd && s.eStart) {
            const nextStart = next?.start && next.start > s.eStart ? next.start : null;
            s.eEnd = nextStart || addDays(s.eStart, EST_DAYS);
        }
        if (!s.eStart && s.eEnd) s.eStart = addDays(s.eEnd, -EST_DAYS);
        if (!s.eStart && !s.eEnd && next) {
            s.eEnd = next.start || next.end;
            s.eStart = addDays(s.eEnd, -EST_DAYS);
        }
    }
}

/**
 * @param {Object[]} projectRows - PROJECTS sheet rows
 * @param {Object[]} stepRows    - STEPS sheet rows (may be empty)
 */
export function buildProjectModel(projectRows, stepRows) {
    const rows = (projectRows || []).filter(r => Object.values(r).some(v => v !== '' && v !== null));
    const col = columnMap(rows);
    const k = {
        date: col('activitydate'), project: col('project'), country: col('country'), type: col('type'),
        owner: col('owner'), stage: col('stagetask'), progress: col('progress'), blocker: col('openblocker'),
        side: col('blockerside'), eta: col('eta'), next: col('nextstep'), health: col('health'),
        latest: col('latest'), pct: col('%done'), shortNote: col('shortnote'),
    };
    const stepsBy = buildSteps(stepRows);

    const groups = new Map();
    let asOf = null;
    rows.forEach(r => {
        const name = String(r[k.project] || '').trim();
        const type = String(r[k.type] || '').trim();
        if (!name || type.toUpperCase() === 'POC') return;
        const date = toDay(r[k.date]);
        if (date && (!asOf || date > asOf)) asOf = date;
        if (!groups.has(name)) groups.set(name, []);
        groups.get(name).push({ r, date });
    });

    const today = todayDate();
    const projects = [];
    groups.forEach((list, name) => {
        list.sort((a, b) => (a.date || 0) - (b.date || 0));
        const latest = [...list].reverse().find(x => isTrue(x.r[k.latest])) || list[list.length - 1];
        const r = latest.r;

        // ETA: latest row's, else carry forward the most recent non-blank ETA.
        let eta = toDay(r[k.eta]);
        if (!eta) {
            const older = [...list].reverse().find(x => toDay(x.r[k.eta]));
            eta = older ? toDay(older.r[k.eta]) : null;
        }

        const health = normHealth(r[k.health]);
        const type = String(r[k.type] || '').trim();
        const stage = String(r[k.stage] || '').trim();
        const steps = stepsBy.get(name.toLowerCase()) || [];
        let pct = parsePct(r[k.pct]);
        if (pct === null && steps.length) pct = steps.filter(s => s.status === 'Done').length / steps.length;
        const live = pct === null && !steps.length && /maintenance/i.test(type);

        const blocker = String(r[k.blocker] || '').trim();
        const progressText = String(r[k.progress] || '').trim();
        const shortNote = String(r[k.shortNote] || '').trim() || blocker || stage;

        projects.push({
            name, type, health, stage, steps, pct, live, eta, blocker,
            country: String(r[k.country] || '').trim(),
            owner: cleanOwners(r[k.owner]),
            blockerSide: String(r[k.side] || '').trim(),
            nextStep: String(r[k.next] || '').trim(),
            progressText, shortNote,
            start: list.find(x => x.date)?.date || null,
            lastUpdate: latest.date,
            etaLate: !!eta && eta < today && health !== 'Closed',
            result: /won/i.test(stage) ? 'Won' : (/lost|drop/i.test(stage) ? 'Lost' : 'Closed'),
        });
    });

    return { projects, asOf, today };
}

/** Urgency order: health rank, furthest past ETA, soonest upcoming date, stalest update. */
function urgencySort(a, b, today) {
    const r = HEALTH[a.health].rank - HEALTH[b.health].rank;
    if (r) return r;
    const late = (p) => p.eta && p.eta < today ? daysBetween(p.eta, today) : -1;
    if (late(a) !== late(b)) return late(b) - late(a);
    const upcoming = (p) => {
        const dates = [p.eta, ...p.steps.map(s => s.end)].filter(d => d && d >= today);
        return dates.length ? Math.min(...dates.map(d => d.getTime())) : Infinity;
    };
    const ua = upcoming(a), ub = upcoming(b);
    if (ua !== ub) return ua - ub;
    return (a.lastUpdate || 0) - (b.lastUpdate || 0);
}

/* ═══════════════════════════════════════════════════════════════
   View state & rendering
   ═══════════════════════════════════════════════════════════════ */

const LEFT_COLS = [
    { key: 'project', label: 'Project', w: 240 },
    { key: 'status', label: 'Status', w: 100 },
    { key: 'progress', label: 'Progress', w: 96 },
    { key: 'eta', label: 'ETA', w: 72 },
    { key: 'owner', label: 'Owner', w: 104 },
    { key: 'notes', label: 'Notes', w: 230 },
];
const HEAD_H = 50;

const state = {
    filter: 'open',      // 'open' | health name | 'avg'
    mode: 'active',      // 'active' | 'closed'
    zoom: 'Month',
    expanded: new Set(),
    showEst: true,
    hideNotes: false,
    centered: false,
};

try {
    const saved = JSON.parse(localStorage.getItem('ps-prefs') || '{}');
    if (ZOOM[saved.zoom]) state.zoom = saved.zoom;
    if (typeof saved.hideNotes === 'boolean') state.hideNotes = saved.hideNotes;
    if (typeof saved.showEst === 'boolean') state.showEst = saved.showEst;
} catch { /* defaults */ }

function savePrefs() {
    try { localStorage.setItem('ps-prefs', JSON.stringify({ zoom: state.zoom, hideNotes: state.hideNotes, showEst: state.showEst })); }
    catch { /* ignore */ }
}

/**
 * Render the Project Status view into a container.
 * @param {HTMLElement} container
 * @param {Object} workbookData
 */
export function renderProjectStatus(container, workbookData) {
    const model = buildProjectModel(workbookData['PROJECTS'], workbookData['STEPS']);
    state.centered = false;

    const draw = () => {
        const prevScroll = container.querySelector('.ps-scroll');
        const keep = prevScroll ? { left: prevScroll.scrollLeft, top: prevScroll.scrollTop } : null;
        const view = buildView(model);
        container.innerHTML = view.html;
        bind(container, model, view, draw);
        const sc = container.querySelector('.ps-scroll');
        if (!sc) return;
        if (!state.centered) {
            sc.scrollLeft = Math.max(0, view.xToday - visibleTl(sc, view) / 2);
            state.centered = true;
        } else if (keep) {
            sc.scrollLeft = keep.left;
            sc.scrollTop = keep.top;
        }
    };
    draw();
}

function visibleProjects(model) {
    const { projects, today } = model;
    let list;
    if (state.mode === 'closed') {
        list = projects.filter(p => p.health === 'Closed' || p.health === 'On hold')
            .sort((a, b) => (b.lastUpdate || 0) - (a.lastUpdate || 0));
    } else if (state.filter === 'avg') {
        list = projects.filter(p => OPEN.includes(p.health) && p.pct !== null).sort((a, b) => a.pct - b.pct);
    } else {
        const set = state.filter === 'open' ? OPEN : [state.filter];
        list = projects.filter(p => set.includes(p.health)).sort((a, b) => urgencySort(a, b, today));
    }
    return list;
}

function buildView(model) {
    const { projects, asOf, today } = model;
    const open = projects.filter(p => OPEN.includes(p.health));
    const count = (h) => open.filter(p => p.health === h).length;
    const withPct = open.filter(p => p.pct !== null);
    const avg = withPct.length ? Math.round(withPct.reduce((s, p) => s + p.pct, 0) / withPct.length * 100) : null;
    const closedCount = projects.filter(p => p.health === 'Closed' || p.health === 'On hold').length;

    /* ── Timeline geometry (over all projects, so filters don't shift the axis) ── */
    const allDates = [today];
    projects.forEach(p => {
        [p.start, p.eta, p.lastUpdate].forEach(d => d && allDates.push(d));
        p.steps.forEach(s => [s.eStart, s.eEnd].forEach(d => d && allDates.push(d)));
    });
    const minD = new Date(Math.min(...allDates));
    const maxD = new Date(Math.max(...allDates, addDays(today, 90)));
    const rangeStart = new Date(minD.getFullYear(), minD.getMonth() - 1, 1);
    const rangeEnd = new Date(maxD.getFullYear(), maxD.getMonth() + 2, 1);
    const ppd = ZOOM[state.zoom];
    const x = (d) => daysBetween(rangeStart, d) * ppd;
    const tlW = Math.ceil(x(rangeEnd));
    const cols = LEFT_COLS.filter(c => !(state.hideNotes && c.key === 'notes'));
    const leftW = cols.reduce((s, c) => s + c.w, 0);
    const gridCols = cols.map(c => c.w + 'px').join(' ');
    const xToday = x(today);

    const tips = [];
    const tip = (html) => { tips.push(html); return tips.length - 1; };

    /* ── Score cards ── */
    const cards = [
        { id: 'open', label: 'Open projects', value: open.length, color: '#243041' },
        { id: 'On track', label: 'On track', value: count('On track'), color: HEALTH['On track'].color },
        { id: 'Blocked', label: 'Blocked', value: count('Blocked'), color: HEALTH.Blocked.color },
        { id: 'Overdue', label: 'Overdue', value: count('Overdue'), color: HEALTH.Overdue.color },
        { id: 'No update', label: 'No update', value: count('No update'), color: HEALTH['No update'].color },
        { id: 'avg', label: 'Average completion', value: avg === null ? '—' : avg + '%', color: '#3b82d6' },
    ];
    const cardsHtml = cards.map(c => {
        const sel = state.mode === 'active' && state.filter === c.id && c.id !== 'open';
        return `<button class="ps-card${sel ? ' selected' : ''}" data-filter="${esc(c.id)}" style="--c:${c.color};">
            <span class="ps-card-value" style="color:${c.color};">${c.value}</span>
            <span class="ps-card-label">${esc(c.label)}</span>
        </button>`;
    }).join('');

    const viewName = state.mode === 'closed' ? 'Closed & on hold'
        : state.filter === 'open' ? 'All open projects'
        : state.filter === 'avg' ? 'Open projects by completion (lowest first)'
        : state.filter;

    /* ── Header ── */
    const months = [];
    for (let m = new Date(rangeStart); m < rangeEnd; m = new Date(m.getFullYear(), m.getMonth() + 1, 1)) months.push(m);
    const monthCells = months.map((m, i) => {
        const next = new Date(m.getFullYear(), m.getMonth() + 1, 1);
        const label = MONTHS[m.getMonth()] + (m.getMonth() === 0 || i === 0 ? ' ' + m.getFullYear() : '');
        return `<div class="ps-hm" style="left:${x(m)}px; width:${x(next) - x(m)}px;">${label}</div>`;
    }).join('');

    let subCells = '';
    const lines = [];
    months.forEach(m => lines.push(`<div class="ps-gl month" style="left:${x(m)}px;"></div>`));
    if (state.zoom === 'Day') {
        for (let d = new Date(rangeStart); d < rangeEnd; d = addDays(d, 1)) {
            const wknd = d.getDay() === 0 || d.getDay() === 6;
            subCells += `<div class="ps-hs${wknd ? ' wknd' : ''}" style="left:${x(d)}px; width:${ppd}px;">${d.getDate()}</div>`;
        }
    }
    if (state.zoom !== 'Quarter') {
        const firstMon = addDays(rangeStart, (8 - rangeStart.getDay()) % 7);
        for (let d = firstMon; d < rangeEnd; d = addDays(d, 7)) {
            lines.push(`<div class="ps-gl week" style="left:${x(d)}px;"></div>`);
            if (state.zoom !== 'Day') {
                const label = state.zoom === 'Week' ? fmtShort(d) : d.getDate();
                subCells += `<div class="ps-hs" style="left:${x(d)}px; width:${7 * ppd}px;">${label}</div>`;
            }
        }
    }

    const headLeft = cols.map(c => `<div class="ps-hc">${esc(c.label)}</div>`).join('');
    const header = `
        <div class="ps-row ps-head" style="height:${HEAD_H}px;">
            <div class="ps-left" style="width:${leftW}px; grid-template-columns:${gridCols};">${headLeft}</div>
            <div class="ps-tl" style="width:${tlW}px;">
                <div class="ps-hrow top">${monthCells}</div>
                <div class="ps-hrow sub">${subCells}</div>
                <div class="ps-today-tag" style="left:${xToday}px;">Today</div>
            </div>
        </div>`;

    /* ── Rows ── */
    const list = visibleProjects(model);
    const rowsHtml = list.map(p => projectRows(p, { x, today, gridCols, leftW, tlW, tip })).join('')
        || `<div class="ps-row"><div class="ps-left ps-empty" style="width:${leftW}px;">No projects in this view.</div><div class="ps-tl" style="width:${tlW}px;"></div></div>`;

    const legend = `
        <div class="ps-legend">
            <span><i class="lg-bar"></i>Project bar</span>
            <span><i class="lg-step done"></i>Step done</span>
            <span><i class="lg-step progress"></i>In progress</span>
            <span><i class="lg-step blocked"></i>Blocked</span>
            <span><i class="lg-step next"></i>Planned</span>
            <span><i class="lg-step est"></i>Estimated</span>
            <span><i class="lg-diamond"></i>ETA</span>
            <span><i class="lg-today"></i>Today</span>
        </div>`;

    const html = `
        <div class="ps-view">
            <div class="ps-title-row">
                <h2 class="ps-title">Project Status</h2>
                <span class="ps-asof">As of ${asOf ? fmtLong(asOf) : '—'}</span>
            </div>
            <div class="ps-cards">${cardsHtml}</div>
            <div class="ps-panel">
                <div class="ps-toolbar">
                    <span class="ps-viewname">${esc(viewName)}</span>
                    <button class="ps-btn${state.mode === 'closed' ? ' active' : ''}" data-act="closed">${state.mode === 'closed' ? '← Open projects' : `Closed &amp; on hold (${closedCount})`}</button>
                    <span class="ps-sep"></span>
                    <span class="ps-tlabel">Zoom</span>
                    <div class="ps-seg">${Object.keys(ZOOM).map(z => `<button class="${state.zoom === z ? 'on' : ''}" data-zoom="${z}">${z}</button>`).join('')}</div>
                    <div class="ps-seg">
                        <button data-act="prev" title="Earlier">‹</button>
                        <button data-act="today">Today</button>
                        <button data-act="next" title="Later">›</button>
                    </div>
                    <span class="ps-sep"></span>
                    <button class="ps-btn" data-act="expand">Expand all</button>
                    <button class="ps-btn" data-act="collapse">Collapse all</button>
                    <label class="ps-check"><input type="checkbox" data-act="est" ${state.showEst ? 'checked' : ''}> Show estimated dates</label>
                    <label class="ps-check"><input type="checkbox" data-act="notes" ${state.hideNotes ? 'checked' : ''}> Hide notes column</label>
                </div>
                ${legend}
                <div class="ps-scroll">
                    <div class="ps-grid" style="width:${leftW + tlW}px;">
                        <div class="ps-lines" style="left:${leftW}px; top:${HEAD_H}px;">${lines.join('')}</div>
                        <div class="ps-today" style="left:${leftW + xToday}px; top:${HEAD_H}px;"></div>
                        ${header}
                        ${rowsHtml}
                    </div>
                </div>
            </div>
            <div class="ps-tip" hidden></div>
        </div>`;

    return { html, tips, leftW, xToday, list };
}

/** Project row + (when expanded) its step rows and info lines. */
function projectRows(p, g) {
    const { x, today, gridCols, leftW, tlW, tip } = g;
    const h = HEALTH[p.health];
    const expanded = state.expanded.has(p.name);
    const closed = p.health === 'Closed';
    const pctTxt = p.live ? 'Live' : (p.pct === null ? '—' : Math.round(p.pct * 100) + '%');
    const statusLabel = closed ? p.result : p.health;

    const noteTip = tip(`<b>${esc(p.name)}</b><br>${esc(p.progressText || p.shortNote)}`);
    const left = `
        <div class="ps-left ps-prow" data-toggle="${esc(p.name)}" style="width:${leftW}px; grid-template-columns:${gridCols};">
            <div class="ps-c ps-pname" title="${esc(`${p.name} · ${p.type}${p.country ? ' · ' + p.country : ''}`)}">
                <span class="ps-arrow${expanded ? ' open' : ''}">▸</span><span class="ps-ellip">${esc(p.name)}</span>
            </div>
            <div class="ps-c"><span class="ps-pill" style="--c:${h.color};"><i></i>${esc(statusLabel)}</span></div>
            <div class="ps-c ps-prog">${p.pct !== null ? `<span class="ps-pbar"><span style="width:${Math.round(p.pct * 100)}%; background:${h.color};"></span></span>` : ''}<span>${pctTxt}</span></div>
            <div class="ps-c${p.etaLate ? ' late' : ''}">${p.eta ? fmtShort(p.eta) : '<span class="ps-muted">—</span>'}</div>
            <div class="ps-c ps-ellip">${esc(p.owner) || '<span class="ps-muted">—</span>'}</div>
            ${state.hideNotes ? '' : `<div class="ps-c ps-ellip ps-note" data-tip="${noteTip}">${esc(p.shortNote)}</div>`}
        </div>`;

    /* Timeline for the project row */
    let tl = '';
    if (closed) {
        const d = p.lastUpdate;
        if (d) {
            const t = tip(`<b>${esc(p.name)}</b><br>${esc(p.result)} · ${fmtLong(d)}<br>${esc(p.owner)}`);
            tl = `<div class="ps-diamond closed" data-tip="${t}" style="left:${x(d)}px;"></div>
                  <div class="ps-blabel" style="left:${x(d) + 12}px;">${esc(p.result)}</div>`;
        }
    } else if (p.health === 'On hold') {
        const from = p.lastUpdate || today;
        const t = tip(`<b>${esc(p.name)}</b><br>On hold since ${fmtLong(from)} (${fmtDur(daysBetween(from, today))})<br>${esc(p.shortNote)}`);
        tl = `<div class="ps-bar hold" data-tip="${t}" style="left:${x(from)}px; width:${Math.max(4, x(today) - x(from))}px;"></div>
              <div class="ps-blabel" style="left:${x(today) + 8}px;">On hold</div>`;
    } else if (p.start) {
        const end = p.eta && p.eta > today ? p.eta : today;
        const w = Math.max(4, x(end) - x(p.start));
        const fill = p.live ? 100 : Math.round((p.pct || 0) * 100);
        const t = tip(`<b>${esc(p.name)}</b><br>${esc(p.owner)}<br>${fmtLong(p.start)} → ${p.eta ? fmtLong(p.eta) + ' (ETA)' : 'today'} · ${fmtDur(daysBetween(p.start, end))}<br>${esc(p.shortNote)}`);
        const labelX = Math.max(x(end), p.eta ? x(p.eta) : 0) + 12;
        tl = `<div class="ps-bar" data-tip="${t}" style="left:${x(p.start)}px; width:${w}px; --c:${h.color};"><span style="width:${fill}%;"></span></div>`;
        if (p.eta) {
            const dt = tip(`<b>ETA · ${esc(p.name)}</b><br>${fmtLong(p.eta)}${p.etaLate ? ` — ${daysBetween(p.eta, today)} days late` : ''}`);
            tl += `<div class="ps-diamond${p.etaLate ? ' late' : ''}" data-tip="${dt}" style="left:${x(p.eta)}px; --c:${h.color};"></div>`;
        }
        tl += `<div class="ps-blabel" style="left:${labelX}px;">${esc([p.live ? 'Live' : (p.pct !== null ? pctTxt : ''), p.owner].filter(Boolean).join(' · '))}</div>`;
    } else {
        tl = `<div class="ps-nodates">no dates in any report</div>`;
    }

    let html = `<div class="ps-row ps-project${expanded ? ' expanded' : ''}">${left}<div class="ps-tl" style="width:${tlW}px;">${tl}</div></div>`;
    if (!expanded) return html;

    /* Step rows */
    if (!p.steps.length) {
        html += infoRow(`<span class="ps-muted">No steps in the STEPS tab for this project yet.</span>`, leftW, tlW);
    }
    p.steps.forEach(s => {
        const st = STEP[s.status];
        const start = s.start || (state.showEst ? s.eStart : null);
        const end = s.end || (state.showEst ? s.eEnd : null);
        const est = (!s.start && start) || (!s.end && end);
        const dateTxt = (d, real) => d ? `${real ? '' : '~'}${fmtShort(d)}` : '—';
        const initials = s.owner ? s.owner.split(/[\s,]+/).filter(Boolean).slice(0, 2).map(w => w[0].toUpperCase()).join('') : '';

        const sLeft = `
            <div class="ps-left ps-srow" style="width:${leftW}px; grid-template-columns:${gridCols};">
                <div class="ps-c ps-sname">
                    <span class="ps-sicon ${st.cls}">${st.icon}</span>
                    <span class="ps-sname-text"><span class="ps-ellip">${esc(s.name)}</span>${s.withWho ? `<span class="ps-with ps-ellip">${esc(s.withWho)}</span>` : ''}</span>
                </div>
                <div class="ps-c ps-sstatus ${st.cls}">${esc(s.status)}</div>
                <div class="ps-c ps-muted">${s.status === 'Done' ? '100%' : '0%'}</div>
                <div class="ps-c ps-sdates">${dateTxt(start, !!s.start)}<br>→ ${dateTxt(end, !!s.end)}</div>
                <div class="ps-c ps-sowner">${initials ? `<span class="ps-avatar">${esc(initials)}</span><span class="ps-ellip">${esc(s.owner)}</span>` : '<span class="ps-muted">—</span>'}</div>
                ${state.hideNotes ? '' : `<div class="ps-c ps-ellip ps-muted" title="${esc(s.note)}">${esc(s.note)}</div>`}
            </div>`;

        let sTl = '';
        if (start && end && (state.showEst || !est)) {
            const dur = fmtDur(daysBetween(start, end));
            const t = tip(`<b>${esc(s.name)}</b><br>${esc(s.owner)}${s.withWho ? ' · ' + esc(s.withWho) : ''}<br>${dateTxt(start, !!s.start)} → ${dateTxt(end, !!s.end)} · ${dur}${est ? ' (estimated)' : ''}<br>${esc(s.note)}`);
            const x0 = x(start), x1 = Math.max(x0 + 4, x(end));
            if (s.status === 'In progress') {
                const xt = Math.min(Math.max(x(today), x0), x1);
                sTl = `<div class="ps-step progress${est ? ' est' : ''}" data-tip="${t}" style="left:${x0}px; width:${Math.max(3, xt - x0)}px;"></div>
                       ${x1 > xt ? `<div class="ps-step next${est ? ' est' : ''}" data-tip="${t}" style="left:${xt}px; width:${x1 - xt}px;"></div>` : ''}`;
            } else if (s.status === 'Blocked') {
                const xt = Math.max(x(today), x0 + 4);
                sTl = `<div class="ps-step blocked${est ? ' est' : ''}" data-tip="${t}" style="left:${x0}px; width:${xt - x0}px;"></div>`;
            } else {
                sTl = `<div class="ps-step ${st.cls}${est ? ' est' : ''}" data-tip="${t}" style="left:${x0}px; width:${x1 - x0}px;"></div>`;
            }
            const labelX = (s.status === 'Blocked' ? Math.max(x(today), x1) : x1) + 8;
            sTl += `<div class="ps-blabel step" style="left:${labelX}px;">${esc([s.name, s.owner, dur].filter(Boolean).join(' · '))}${s.status === 'Blocked' ? ' <b class="late">blocked</b>' : ''}</div>`;
        } else {
            sTl = `<div class="ps-nodates">no dates in any report</div>`;
        }
        html += `<div class="ps-row ps-step-row">${sLeft}<div class="ps-tl" style="width:${tlW}px;">${sTl}</div></div>`;
    });

    if (p.blocker) {
        html += infoRow(`<span class="late"><b>Blocker${p.blockerSide ? ` (${esc(p.blockerSide)})` : ''}:</b> ${esc(p.blocker)}</span>`, leftW, tlW);
    }
    if (p.nextStep) html += infoRow(`<b>Next step:</b> ${esc(p.nextStep)}`, leftW, tlW);
    html += infoRow(`<span class="ps-muted">Last update ${p.lastUpdate ? fmtLong(p.lastUpdate) : '—'}</span>`, leftW, tlW);
    return html;
}

function infoRow(inner, leftW, tlW) {
    return `<div class="ps-row ps-info"><div class="ps-left" style="width:${leftW}px;">${inner}</div><div class="ps-tl" style="width:${tlW}px;"></div></div>`;
}

/* ═══════════════════════════════════════════════════════════════
   Interaction
   ═══════════════════════════════════════════════════════════════ */

/** Width of the timeline area not covered by the frozen table (never tiny on narrow screens). */
const visibleTl = (sc, view) => Math.max(160, sc.clientWidth - view.leftW);

function bind(root, model, view, draw) {
    const sc = root.querySelector('.ps-scroll');
    const tipEl = root.querySelector('.ps-tip');

    root.querySelectorAll('[data-filter]').forEach(b => b.addEventListener('click', () => {
        const f = b.dataset.filter;
        state.mode = 'active';
        state.filter = state.filter === f ? 'open' : f;
        draw();
    }));

    root.querySelectorAll('[data-zoom]').forEach(b => b.addEventListener('click', () => {
        if (state.zoom === b.dataset.zoom) return;
        // Keep the date at the centre of the timeline in place across zooms.
        const tlVisible = visibleTl(sc, view);
        const centerDays = (sc.scrollLeft + tlVisible / 2) / ZOOM[state.zoom];
        state.zoom = b.dataset.zoom;
        savePrefs();
        draw();
        const sc2 = root.querySelector('.ps-scroll');
        sc2.scrollLeft = Math.max(0, centerDays * ZOOM[state.zoom] - tlVisible / 2);
    }));

    root.querySelectorAll('[data-act]').forEach(el => {
        const act = el.dataset.act;
        const evt = el.tagName === 'INPUT' ? 'change' : 'click';
        el.addEventListener(evt, () => {
            const tlVisible = visibleTl(sc, view);
            if (act === 'closed') { state.mode = state.mode === 'closed' ? 'active' : 'closed'; state.filter = 'open'; draw(); }
            else if (act === 'prev') sc.scrollBy({ left: -tlVisible * 0.6, behavior: 'smooth' });
            else if (act === 'next') sc.scrollBy({ left: tlVisible * 0.6, behavior: 'smooth' });
            else if (act === 'today') sc.scrollTo({ left: Math.max(0, view.xToday - tlVisible / 2), behavior: 'smooth' });
            else if (act === 'expand') { view.list.forEach(p => state.expanded.add(p.name)); draw(); }
            else if (act === 'collapse') { state.expanded.clear(); draw(); }
            else if (act === 'est') { state.showEst = el.checked; savePrefs(); draw(); }
            else if (act === 'notes') { state.hideNotes = el.checked; savePrefs(); draw(); }
        });
    });

    root.querySelectorAll('[data-toggle]').forEach(row => row.addEventListener('click', () => {
        const name = row.dataset.toggle;
        if (state.expanded.has(name)) state.expanded.delete(name); else state.expanded.add(name);
        draw();
    }));

    // One floating tooltip for every bar, diamond and note.
    root.addEventListener('mouseover', (e) => {
        const t = e.target.closest('[data-tip]');
        if (!t || !root.contains(t)) { tipEl.hidden = true; return; }
        tipEl.innerHTML = view.tips[+t.dataset.tip] || '';
        tipEl.hidden = false;
    });
    root.addEventListener('mousemove', (e) => {
        if (tipEl.hidden) return;
        const pad = 14;
        const w = tipEl.offsetWidth, h = tipEl.offsetHeight;
        let left = e.clientX + pad, top = e.clientY + pad;
        if (left + w > window.innerWidth - 8) left = e.clientX - w - pad;
        if (top + h > window.innerHeight - 8) top = e.clientY - h - pad;
        tipEl.style.left = left + 'px';
        tipEl.style.top = top + 'px';
    });
    root.addEventListener('mouseleave', () => { tipEl.hidden = true; });
}
