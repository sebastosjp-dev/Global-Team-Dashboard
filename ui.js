/**
 * ui.js - HTML template generators for dashboard components
 */
import { formatCurrency, parseCurrency, sortCountriesByAmount, sortCountriesByCount } from './utils.js';
import { CONFIG } from './config.js';
/**
 * @deprecated Styles now live in styles.css. Kept as no-op for backward compat.
 */
export function injectServiceAnalysisStyles() {
    /* All styles moved to styles.css — nothing to inject */
}

export function injectPipelineTooltipStyles() {
    injectServiceAnalysisStyles();
}

const PIPELINE_STAGE_THEME = {
    'Discovery':   { bg: '#eef2ff', fg: '#4338ca', border: '#c7d2fe' },
    'PoC':         { bg: '#fef3c7', fg: '#b45309', border: '#fcd34d' },
    'Quotation':   { bg: '#dbeafe', fg: '#1e40af', border: '#93c5fd' },
    'Negotiation': { bg: '#fce7f3', fg: '#be185d', border: '#f9a8d4' },
    'Lost':        { bg: '#fee2e2', fg: '#b91c1c', border: '#fca5a5' },
    'Won':         { bg: '#d1fae5', fg: '#047857', border: '#6ee7b7' },
    'Unknown':     { bg: '#f3f4f6', fg: '#6b7280', border: '#d1d5db' }
};
function getStageTheme(stage) {
    if (!stage) return PIPELINE_STAGE_THEME['Unknown'];
    const k = String(stage).trim();
    return PIPELINE_STAGE_THEME[k] || { bg: '#f3f4f6', fg: '#6b7280', border: '#d1d5db' };
}
function renderStageBadge(stage, opts = {}) {
    const t = getStageTheme(stage);
    const fontSize = opts.fontSize || '0.62rem';
    const padding = opts.padding || '2px 8px';
    return `<span style="display:inline-block; background:${t.bg}; color:${t.fg}; border:1px solid ${t.border}; font-size:${fontSize}; font-weight:800; padding:${padding}; border-radius:8px; text-transform:uppercase; letter-spacing:0.04em; line-height:1.2; white-space:nowrap;">${stage || 'Unknown'}</span>`;
}

// Globally expose tooltip functions
window.showQuarterTooltip = function (event, element) {
    const tooltip = document.getElementById('pipeline-quarter-tooltip');
    if (!tooltip) return;

    const quarter = element.getAttribute('data-q');
    const deals = JSON.parse(element.getAttribute('data-deals'));

    let rowsHtml = deals.slice(0, 5).map(d => `
        <tr>
            <td style="font-weight: 600;">${d.n}</td>
            <td style="text-align: right; color: #10b981; font-weight: 700;">$${d.a}</td>
        </tr>
    `).join('');

    if (deals.length > 5) {
        rowsHtml += `<tr><td colspan="2" style="text-align:center; padding: 8px; color: #9CA3AF; font-size: 0.7rem;">Click to see all ${deals.length} deals...</td></tr>`;
    }

    if (deals.length === 0) {
        rowsHtml = '<tr><td colspan="2" style="text-align:center; padding: 15px; color: #9CA3AF;">No deals found</td></tr>';
    }

    tooltip.innerHTML = `
        <div class="pipeline-tooltip-header" style="padding: 8px 12px; font-size: 0.8rem;">
            <span>${quarter} Preview</span>
        </div>
        <div class="pipeline-tooltip-content">
            <table class="pipeline-tooltip-table">
                <tbody>
                    ${rowsHtml}
                </tbody>
            </table>
        </div>
    `;

    tooltip.style.display = 'block';

    const rect = element.getBoundingClientRect();
    const tooltipHeight = tooltip.offsetHeight;
    const tooltipWidth = tooltip.offsetWidth;

    let top = rect.top - tooltipHeight - 10;
    if (top < 10) top = rect.bottom + 10;

    let left = rect.left + (rect.width / 2) - (tooltipWidth / 2);
    if (left < 10) left = 10;
    if (left + tooltipWidth > window.innerWidth - 10) left = window.innerWidth - tooltipWidth - 10;

    tooltip.style.top = top + 'px';
    tooltip.style.left = left + 'px';
};

window.hideQuarterTooltip = function () {
    const tooltip = document.getElementById('pipeline-quarter-tooltip');
    if (tooltip) tooltip.style.display = 'none';
};

window.showPocTooltip = function (event, element, color = '#007AFF') {
    let tooltip = document.getElementById('poc-hover-tooltip');
    if (!tooltip) {
        tooltip = document.createElement('div');
        tooltip.id = 'poc-hover-tooltip';
        tooltip.style.position = 'fixed';
        tooltip.style.display = 'none';
        tooltip.style.zIndex = '10000';
        tooltip.style.background = '#FFFFFF';
        tooltip.style.borderRadius = '12px';
        tooltip.style.boxShadow = '0 10px 25px rgba(0,0,0,0.15)';
        tooltip.style.width = '350px';
        tooltip.style.maxHeight = 'none';
        tooltip.style.overflow = 'visible';
        tooltip.style.pointerEvents = 'none';
        tooltip.style.transition = 'opacity 0.2s';
        document.body.appendChild(tooltip);
    }

    let names = [];
    try {
        names = JSON.parse(decodeURIComponent(element.getAttribute('data-names')));
    } catch (e) { }
    const title = element.getAttribute('data-title') || 'POCs';

    tooltip.style.border = '1px solid ' + color;

    let rowsHtml = names.map((n, i) => `
        <tr style="transition: background 0.2s;">
            <td style="padding: 10px 16px; border-bottom: 1px solid #F3F4F6; color: #374151; font-weight: 600; font-size: 0.75rem;">${i + 1}. ${n}</td>
        </tr>
    `).join('');

    if (names.length === 0) {
        rowsHtml = '<tr><td style="padding: 15px; text-align: center; color: #9CA3AF;">No data</td></tr>';
    }

    tooltip.innerHTML = `
        <div style="background: ${color}; color: white; padding: 12px 16px; font-weight: 700; font-size: 0.9rem; display: flex; justify-content: space-between;">
            <span>${title}</span>
            <span style="background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 12px; font-size: 0.75rem;">${names.length}</span>
        </div>
        <div style="padding: 0;">
            <table style="width: 100%; border-collapse: collapse;">
                <tbody style="background: #FFFFFF;">${rowsHtml}</tbody>
            </table>
        </div>
    `;

    tooltip.style.display = 'block';
    tooltip.style.opacity = '1';

    const rect = element.getBoundingClientRect();
    const tooltipHeight = tooltip.offsetHeight;
    const tooltipWidth = tooltip.offsetWidth;

    let top = rect.top - tooltipHeight - 10;
    if (top < 10) top = rect.bottom + 10;

    let left = rect.left + (rect.width / 2) - (tooltipWidth / 2);
    if (left < 10) left = 10;
    if (left + tooltipWidth > window.innerWidth - 10) left = window.innerWidth - tooltipWidth - 10;

    tooltip.style.top = top + 'px';
    tooltip.style.left = left + 'px';
};

window.hidePocTooltip = function () {
    const tooltip = document.getElementById('poc-hover-tooltip');
    if (tooltip) {
        tooltip.style.display = 'none';
        tooltip.style.opacity = '0';
    }
};

window.copyDecisionList = function () {
    const table = document.getElementById('decision-required-table');
    if (!table) return;
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    const lines = rows.map(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        const name = cells[1]?.innerText?.trim() || '';
        const partner = cells[2]?.innerText?.trim() || '';
        const country = cells[3]?.innerText?.trim() || '';
        const status = cells[4]?.innerText?.trim() || '';
        const startDate = cells[5]?.innerText?.trim() || '';
        const elapsed = cells[6]?.innerText?.trim() || '';
        return `${name} | ${partner} | ${country} | ${status} | Start: ${startDate} | Elapsed: ${elapsed}`;
    });
    const text = `Decision Required POCs (2+ months)\n${'='.repeat(50)}\n${lines.join('\n')}`;
    navigator.clipboard.writeText(text).then(() => {
        const btn = document.querySelector('button[onclick="copyDecisionList()"]');
        if (btn) {
            const original = btn.innerHTML;
            btn.innerHTML = '<i class="fa-solid fa-check"></i> Copied!';
            btn.style.background = '#34C759';
            setTimeout(() => { btn.innerHTML = original; btn.style.background = '#A855F7'; }, 2000);
        }
    });
};

const QUARTER_MONTHS = {
    Q1: [{ num: 1, label: 'Jan' }, { num: 2, label: 'Feb' }, { num: 3, label: 'Mar' }],
    Q2: [{ num: 4, label: 'Apr' }, { num: 5, label: 'May' }, { num: 6, label: 'Jun' }],
    Q3: [{ num: 7, label: 'Jul' }, { num: 8, label: 'Aug' }, { num: 9, label: 'Sep' }],
    Q4: [{ num: 10, label: 'Oct' }, { num: 11, label: 'Nov' }, { num: 12, label: 'Dec' }],
};

function renderQuarterPanelHtml() {
    const state = window.__quarterPanelState;
    if (!state) return '';
    const { quarter, deals, showCountry, selectedMonth } = state;

    const COUNTRY_FLAGS = {
        'Indonesia': '🇮🇩', 'Thailand': '🇹🇭', 'Malaysia': '🇲🇾',
        'USA': '🇺🇸', 'Philippines': '🇵🇭', 'Vietnam': '🇻🇳',
        'Singapore': '🇸🇬', 'Turkey': '🇹🇷', 'Japan': '🇯🇵',
        'India': '🇮🇳', 'Australia': '🇦🇺', 'Taiwan': '🇹🇼',
        'Hong Kong': '🇭🇰',
    };

    const months = QUARTER_MONTHS[quarter] || [];
    const monthStats = months.map(m => {
        const md = deals.filter(d => d.m === m.num);
        return {
            ...m,
            count: md.length,
            tcv: md.reduce((s, d) => s + (Number(d.t) || 0), 0),
            weighted: md.reduce((s, d) => s + (Number(d.w) || 0), 0),
        };
    });
    const noDateDeals = deals.filter(d => d.m == null);

    const filteredDeals = selectedMonth == null
        ? deals
        : selectedMonth === 'none'
            ? noDateDeals
            : deals.filter(d => d.m === selectedMonth);

    const activeStyle = 'background: #10B981; color: #FFFFFF; border-color: #10B981; box-shadow: 0 4px 10px rgba(16,185,129,0.25);';
    const inactiveStyle = 'background: #FFFFFF; color: #1E293B; border-color: #E5E7EB;';

    const allBtn = `
        <button type="button" onclick="filterQuarterByMonth(null)"
            style="cursor:pointer; flex:0 0 auto; min-width:96px; padding:10px 14px; border-radius:10px; border:1px solid; text-align:left; font-family:inherit; transition:all 0.15s; ${selectedMonth == null ? activeStyle : inactiveStyle}">
            <div style="font-size:0.7rem; font-weight:800; letter-spacing:0.06em; text-transform:uppercase; opacity:0.85;">All</div>
            <div style="font-size:1.05rem; font-weight:900; margin-top:2px;">${deals.length} <span style="font-size:0.65rem; font-weight:700; opacity:0.8;">deals</span></div>
        </button>
    `;

    const monthBtns = monthStats.map(m => {
        const active = selectedMonth === m.num;
        return `
            <button type="button" onclick="filterQuarterByMonth(${m.num})"
                style="cursor:pointer; flex:1 1 0; min-width:140px; padding:10px 14px; border-radius:10px; border:1px solid; text-align:left; font-family:inherit; transition:all 0.15s; ${active ? activeStyle : inactiveStyle}">
                <div style="display:flex; align-items:center; justify-content:space-between; gap:8px;">
                    <span style="font-size:0.72rem; font-weight:800; letter-spacing:0.06em; text-transform:uppercase; opacity:0.9;">${m.label}</span>
                    <span style="font-size:0.6rem; font-weight:800; padding:1px 7px; border-radius:8px; background:${active ? 'rgba(255,255,255,0.25)' : 'rgba(16,185,129,0.12)'}; color:${active ? '#FFFFFF' : '#059669'};">${m.count}</span>
                </div>
                <div style="display:flex; justify-content:space-between; margin-top:6px; gap:6px; font-size:0.65rem;">
                    <span style="opacity:0.7;">TCV</span>
                    <span style="font-weight:800; color:${active ? '#FFFFFF' : '#EF4444'};">$${formatCurrency(m.tcv)}</span>
                </div>
                <div style="display:flex; justify-content:space-between; margin-top:2px; gap:6px; font-size:0.65rem;">
                    <span style="opacity:0.7;">Weighted</span>
                    <span style="font-weight:800; color:${active ? '#FFFFFF' : '#10B981'};">$${formatCurrency(m.weighted)}</span>
                </div>
            </button>
        `;
    }).join('');

    const noDateBtn = noDateDeals.length > 0 ? `
        <button type="button" onclick="filterQuarterByMonth('none')"
            style="cursor:pointer; flex:0 0 auto; min-width:110px; padding:10px 14px; border-radius:10px; border:1px dashed; text-align:left; font-family:inherit; transition:all 0.15s; ${selectedMonth === 'none' ? activeStyle : inactiveStyle}">
            <div style="font-size:0.7rem; font-weight:800; letter-spacing:0.06em; text-transform:uppercase; opacity:0.85;">No Date</div>
            <div style="font-size:1.05rem; font-weight:900; margin-top:2px;">${noDateDeals.length} <span style="font-size:0.65rem; font-weight:700; opacity:0.8;">deals</span></div>
        </button>
    ` : '';

    const monthBarHtml = `
        <div style="background:#F8FAFC; padding:14px 24px; border-bottom:1px solid #E5E7EB; display:flex; gap:10px; flex-wrap:wrap; align-items:stretch;">
            ${allBtn}${monthBtns}${noDateBtn}
        </div>
    `;

    let rowsHtml = filteredDeals.map((d, index) => {
        const flag = d.c ? (COUNTRY_FLAGS[d.c] || '') : '';
        const countryBadge = (showCountry && d.c)
            ? `<span title="${d.c}" style="display: inline-block; font-size: 1.1rem; margin-right: 10px; vertical-align: middle; line-height: 1;">${flag || d.c}</span>`
            : '';
        const stageBadge = renderStageBadge(d.s || 'Unknown', { fontSize: '0.7rem', padding: '4px 10px' });
        const monthLabel = d.m == null
            ? '<span style="color:#94A3B8; font-weight:700; font-size:0.7rem;">—</span>'
            : `<span style="display:inline-block; padding:3px 9px; border-radius:8px; background:#EEF2FF; color:#4338CA; font-weight:800; font-size:0.7rem; letter-spacing:0.04em;">${(QUARTER_MONTHS[quarter] || []).find(mm => mm.num === d.m)?.label || d.m}</span>`;
        return `
        <tr style="border-bottom: 1px solid #F3F4F6; transition: background 0.2s;">
            <td style="padding: 16px 24px; color: #94A3B8; font-weight: 700; width: 60px; font-family: monospace;">${String(index + 1).padStart(2, '0')}</td>
            <td style="padding: 16px 24px; color: #1E293B; font-weight: 700; font-size: 0.95rem;">${countryBadge}${d.n}</td>
            <td style="padding: 16px 24px; text-align: center;">${monthLabel}</td>
            <td style="padding: 16px 24px; text-align: center;">${stageBadge}</td>
            <td style="padding: 16px 24px; text-align: right; color: #EF4444; font-weight: 800; font-size: 1.1rem; letter-spacing: -0.02em;">$${d.tf != null ? d.tf : formatCurrency(d.t || 0)}</td>
            <td style="padding: 16px 24px; text-align: right; color: #10B981; font-weight: 800; font-size: 1.1rem; letter-spacing: -0.02em;">$${d.a}</td>
        </tr>
    `;
    }).join('');

    const totalWeighted = filteredDeals.reduce((sum, d) => sum + (Number(d.w) || 0), 0);
    const totalTcv = filteredDeals.reduce((sum, d) => sum + (Number(d.t) || 0), 0);
    const totalLabel = selectedMonth == null
        ? 'Total Pipeline Value'
        : selectedMonth === 'none'
            ? 'Total (No Date)'
            : `Total — ${(QUARTER_MONTHS[quarter] || []).find(mm => mm.num === selectedMonth)?.label || ''}`;
    const totalRowHtml = filteredDeals.length > 0 ? `
        <tr style="background: #FEF2F2; border-top: 2px solid #FCA5A5;">
            <td style="padding: 18px 24px;"></td>
            <td colspan="3" style="padding: 18px 24px; color: #DC2626; font-weight: 800; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">${totalLabel}</td>
            <td style="padding: 18px 24px; text-align: right; color: #DC2626; font-weight: 900; font-size: 1.2rem; letter-spacing: -0.02em;">$${formatCurrency(totalTcv)}</td>
            <td style="padding: 18px 24px; text-align: right; color: #DC2626; font-weight: 900; font-size: 1.2rem; letter-spacing: -0.02em;">$${formatCurrency(totalWeighted)}</td>
        </tr>
    ` : '';

    if (filteredDeals.length === 0) {
        rowsHtml = '<tr><td colspan="6" style="padding: 60px 20px; text-align: center; color: #94A3B8; font-style: italic;">No active deals found for this period.</td></tr>';
    }

    const headerSubtitle = selectedMonth == null
        ? 'Complete breakdown of weighted pipeline value for the period'
        : selectedMonth === 'none'
            ? 'Deals in this quarter without a recorded date'
            : `Showing deals dated in ${(QUARTER_MONTHS[quarter] || []).find(mm => mm.num === selectedMonth)?.label || ''}`;

    return `
        <div style="background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.08);">
            <div style="background: linear-gradient(135deg, #10B981 0%, #059669 100%); padding: 20px 28px; display: flex; justify-content: space-between; align-items: center; color: white; border-bottom: 1px solid rgba(0,0,0,0.05);">
                <div style="display: flex; align-items: center; gap: 14px;">
                    <div style="width: 40px; height: 40px; background: rgba(255,255,255,0.2); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem;">
                        <i class="fa-solid fa-list-check"></i>
                    </div>
                    <div>
                        <h3 style="margin: 0; font-size: 1.2rem; font-weight: 800; letter-spacing: -0.01em;">${quarter} Detailed Pipeline</h3>
                        <p style="margin: 2px 0 0 0; opacity: 0.8; font-size: 0.8rem; font-weight: 500;">${headerSubtitle}</p>
                    </div>
                </div>
                <div style="background: rgba(255,255,255,0.2); backdrop-filter: blur(4px); padding: 8px 16px; border-radius: 12px; font-size: 0.9rem; font-weight: 800; border: 1px solid rgba(255,255,255,0.1);">
                    ${filteredDeals.length} DEALS
                </div>
            </div>
            ${monthBarHtml}
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; text-align: left;">
                    <thead>
                        <tr style="background: #F8FAFC; border-bottom: 2px solid #F1F5F9;">
                            <th style="padding: 16px 24px; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em; width: 80px;">NO.</th>
                            <th style="padding: 16px 24px; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em;">CRM DEAL NAME</th>
                            <th style="padding: 16px 24px; text-align: center; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em; width: 90px;">MONTH</th>
                            <th style="padding: 16px 24px; text-align: center; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em; width: 140px;">DEAL STAGE</th>
                            <th style="padding: 16px 24px; text-align: right; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em;">TCV (USD)</th>
                            <th style="padding: 16px 24px; text-align: right; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em;">WEIGHTED PIPELINE VALUE (USD)</th>
                        </tr>
                    </thead>
                    <tbody style="background: #FFFFFF;">
                        ${rowsHtml}
                    </tbody>
                    ${totalRowHtml ? `<tfoot>${totalRowHtml}</tfoot>` : ''}
                </table>
            </div>
        </div>
    `;
}

window.selectQuarter = function (element) {
    const quarter = element.getAttribute('data-q');
    const deals = JSON.parse(element.getAttribute('data-deals'))
        .slice()
        .sort((a, b) => (Number(b.w) || 0) - (Number(a.w) || 0));
    const uniqueCountries = new Set(deals.map(d => d.c).filter(Boolean));
    const showCountry = uniqueCountries.size > 1
        || element.getAttribute('data-show-country') === 'true';
    const container = document.getElementById('pipeline-selected-quarter-container');
    if (!container) return;

    // Reset all cards styling
    document.querySelectorAll('.quarter-card').forEach(c => {
        c.style.borderTop = '3px solid #10b981';
        c.style.background = '#F9FAFB';
        c.classList.remove('active-quarter');
        c.style.transform = 'none';
        c.style.boxShadow = 'none';
    });

    // Highlight selected card
    element.style.borderTop = '6px solid #10b981';
    element.style.background = '#FFFFFF';
    element.classList.add('active-quarter');
    element.style.transform = 'translateY(-4px)';
    element.style.boxShadow = '0 10px 25px rgba(16, 185, 129, 0.15)';

    window.__quarterPanelState = { quarter, deals, showCountry, selectedMonth: null };
    container.innerHTML = renderQuarterPanelHtml();
    container.style.display = 'block';

    setTimeout(() => {
        const yOffset = -20;
        const y = container.getBoundingClientRect().top + window.pageYOffset + yOffset;
        window.scrollTo({ top: y, behavior: 'smooth' });
    }, 50);
};

window.filterQuarterByMonth = function (month) {
    if (!window.__quarterPanelState) return;
    window.__quarterPanelState.selectedMonth = month;
    const container = document.getElementById('pipeline-selected-quarter-container');
    if (!container) return;
    container.innerHTML = renderQuarterPanelHtml();
};


export function getServiceAnalysisHTML(stats, filterCountry = 'All') {
    injectServiceAnalysisStyles();

    /* ── Country selector ── */
    let html = `
        <div class="stat-card" style="display:flex; align-items:center; gap:12px; padding:10px 16px; background:#FFFFFF; border:1px solid rgba(99,102,241,0.2); border-left:4px solid #6366f1; margin-bottom:20px;">
            <label style="font-size:0.8rem; color:#6366f1; font-weight:700; text-transform:uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right:8px;"></i>Select Country</label>
            <select id="csm-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:6px 12px; border-radius:8px; width:200px; font-size:0.85rem;">
                ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
            </select>
            <span style="font-size:0.72rem; color:#64748b; margin-left:auto;">Metrics for ${filterCountry || 'All Regions'}</span>
        </div>
    `;

    if (!stats) {
        return html + '<p style="padding:40px; text-align:center; color:#6B7280;">No active service data found for the selected country.</p>';
    }

    /* ══════════════════════════════════════
       SECTION 1 — Health Overview (5 KPI cards)
       ══════════════════════════════════════ */
    const healthTotal = stats.healthGreen + stats.healthYellow + stats.healthRed;
    const healthGreenPct = healthTotal > 0 ? Math.round((stats.healthGreen / healthTotal) * 100) : 0;
    const retentionColor = stats.arrRetentionRate >= 90 ? '#059669' : stats.arrRetentionRate >= 75 ? '#d97706' : '#dc2626';
    const churnColor = stats.churnRate <= 5 ? '#059669' : stats.churnRate <= 15 ? '#d97706' : '#dc2626';

    html += `
        <div style="margin-bottom:8px; display:flex; align-items:center; gap:8px;">
            <div style="width:4px; height:18px; background:#6366f1; border-radius:2px;"></div>
            <span style="font-size:0.78rem; font-weight:700; color:#374151; text-transform:uppercase; letter-spacing:0.06em;">① Health Overview</span>
        </div>
        <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:14px; margin-bottom:20px;">
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid #6366f1; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:6px;">Customer Portfolio</div>
                <div style="font-size:1.8rem; font-weight:800; color:#111827; line-height:1;">${stats.totalEndUsers}</div>
                <div style="display:flex; gap:10px; margin-top:8px; font-size:0.75rem;">
                    <span style="color:#059669;">Active ${stats.activeCount}</span>
                    <span style="color:#9ca3af;">Inactive ${stats.inactiveCount}</span>
                </div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid ${stats.healthRed > 0 ? '#ef4444' : stats.healthYellow > 0 ? '#f59e0b' : '#10b981'}; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:6px;">Health Score</div>
                <div style="font-size:1.8rem; font-weight:800; color:#111827; line-height:1;">${healthGreenPct}<span style="font-size:1rem; font-weight:600; color:#6b7280;">%</span></div>
                <div style="display:flex; gap:8px; margin-top:8px; font-size:0.72rem;">
                    <span style="color:#059669; background:#f0fdf4; padding:2px 7px; border-radius:10px;">✓ ${stats.healthGreen}</span>
                    <span style="color:#d97706; background:#fffbeb; padding:2px 7px; border-radius:10px;">⚠ ${stats.healthYellow}</span>
                    <span style="color:#dc2626; background:#fef2f2; padding:2px 7px; border-radius:10px;">✕ ${stats.healthRed}</span>
                </div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid ${retentionColor}; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:6px;">ARR Retention</div>
                <div style="font-size:1.8rem; font-weight:800; color:${retentionColor}; line-height:1;">${stats.arrRetentionRate}<span style="font-size:1rem; font-weight:600;">%</span></div>
                <div style="margin-top:8px; font-size:0.72rem; color:#6b7280;">Active ARR: $${formatCurrency(stats.activeArr)}</div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid ${churnColor}; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:6px;">Churn Rate</div>
                <div style="font-size:1.8rem; font-weight:800; color:${churnColor}; line-height:1;">${stats.churnRate}<span style="font-size:1rem; font-weight:600;">%</span></div>
                <div style="margin-top:8px; font-size:0.72rem; color:#6b7280;">ARR at Risk: $${formatCurrency(stats.atRiskArr)}</div>
            </div>
        </div>
    `;

    /* ══════════════════════════════════════
       SECTION 2 — Health Score Detail
       ══════════════════════════════════════ */
    const atRiskRows = stats.atRiskCustomers.slice(0, 12).map(c => {
        const scoreColor = c.healthColor === 'red' ? '#dc2626' : '#d97706';
        const scoreBg = c.healthColor === 'red' ? '#fef2f2' : '#fffbeb';
        const badgeText = c.healthColor === 'red' ? 'Critical' : 'At Risk';
        const reasonStr = c.reasons.join(' · ') || 'Single service';
        return `<tr style="border-bottom:1px solid #f3f4f6;">
            <td style="padding:9px 10px; font-weight:600; font-size:0.8rem; color:#111827;">${c.name}</td>
            <td style="padding:9px 10px; font-size:0.78rem; color:#6b7280;">${c.country}</td>
            <td style="padding:9px 10px; text-align:center;">
                <span style="display:inline-block; width:36px; height:36px; border-radius:50%; background:${scoreBg}; color:${scoreColor}; font-weight:800; font-size:0.85rem; line-height:36px; text-align:center;">${c.score}</span>
            </td>
            <td style="padding:9px 10px;"><span style="background:${scoreBg}; color:${scoreColor}; font-size:0.68rem; font-weight:700; padding:2px 8px; border-radius:10px; text-transform:uppercase;">${badgeText}</span></td>
            <td style="padding:9px 10px; font-size:0.75rem; color:#6b7280;">${reasonStr}</td>
            <td style="padding:9px 10px; text-align:right; font-weight:600; font-size:0.78rem; color:#374151;">$${formatCurrency(c.arr)}</td>
        </tr>`;
    }).join('');

    html += `
        <div style="margin-bottom:8px; display:flex; align-items:center; gap:8px;">
            <div style="width:4px; height:18px; background:#ef4444; border-radius:2px;"></div>
            <span style="font-size:0.78rem; font-weight:700; color:#374151; text-transform:uppercase; letter-spacing:0.06em;">② Customer Health Detail</span>
        </div>
        <div style="display:grid; grid-template-columns: 280px 1fr; gap:16px; margin-bottom:20px;">
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow:0 2px 8px rgba(0,0,0,0.06); display:flex; flex-direction:column; align-items:center; justify-content:center;">
                <div style="font-size:0.8rem; font-weight:700; color:#374151; margin-bottom:12px;">Health Distribution</div>
                <div style="height:200px; width:100%; position:relative;"><canvas id="health-score-donut"></canvas></div>
                <div style="display:flex; flex-direction:column; gap:6px; margin-top:12px; width:100%;">
                    <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.78rem;">
                        <span style="display:flex; align-items:center; gap:6px;"><span style="width:10px; height:10px; background:#10b981; border-radius:2px; display:inline-block;"></span>Healthy</span>
                        <span style="font-weight:700; color:#059669;">${stats.healthGreen} customers</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.78rem;">
                        <span style="display:flex; align-items:center; gap:6px;"><span style="width:10px; height:10px; background:#f59e0b; border-radius:2px; display:inline-block;"></span>At Risk</span>
                        <span style="font-weight:700; color:#d97706;">${stats.healthYellow} customers</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.78rem;">
                        <span style="display:flex; align-items:center; gap:6px;"><span style="width:10px; height:10px; background:#ef4444; border-radius:2px; display:inline-block;"></span>Critical</span>
                        <span style="font-weight:700; color:#dc2626;">${stats.healthRed} customers</span>
                    </div>
                </div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.85rem; font-weight:700; color:#111827; margin-bottom:12px; display:flex; align-items:center; gap:8px;">
                    <i class="fa-solid fa-triangle-exclamation" style="color:#f59e0b;"></i> At-Risk Customers
                    <span style="font-size:0.7rem; background:#fef3c7; color:#92400e; padding:2px 8px; border-radius:10px; font-weight:700;">${stats.atRiskCustomers.length} need attention</span>
                </div>
                ${atRiskRows.length > 0 ? `
                <div style="overflow-x:auto;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                        <thead><tr style="background:#f9fafb; text-align:left;">
                            <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Customer</th>
                            <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Country</th>
                            <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase; text-align:center;">Score</th>
                            <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Risk Level</th>
                            <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Reason</th>
                            <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase; text-align:right;">ARR</th>
                        </tr></thead>
                        <tbody>${atRiskRows}</tbody>
                    </table>
                </div>` : '<p style="color:#9ca3af; text-align:center; padding:24px; font-size:0.85rem;">All customers are healthy!</p>'}
            </div>
        </div>
    `;

    /* ══════════════════════════════════════
       SECTION 3 — Churn & Renewal Risk
       ══════════════════════════════════════ */
    const renewalBarTotal = stats.expiringCount || 1;
    const renewalRows = stats.expiringCustomers.map(c => {
        const urgencyColors = {
            critical: { bg: 'rgba(244,63,94,0.06)', text: '#e11d48', badge: '#fecdd3' },
            warning: { bg: 'rgba(245,158,11,0.06)', text: '#d97706', badge: '#fef3c7' },
            normal: { bg: 'transparent', text: '#374151', badge: '#d1fae5' }
        };
        const uc = urgencyColors[c.urgency];
        return `<tr style="border-bottom:1px solid #f3f4f6; background:${uc.bg};">
            <td style="padding:9px 12px; font-weight:600; font-size:0.8rem; color:#111827;">${c.name}</td>
            <td style="padding:9px 12px; font-size:0.78rem; color:#6b7280;">${c.country}</td>
            <td style="padding:9px 12px;"><span style="padding:2px 8px; border-radius:10px; font-size:0.68rem; font-weight:700; background:${c.status === 'Active' ? 'rgba(16,185,129,0.1)' : 'rgba(107,114,128,0.1)'}; color:${c.status === 'Active' ? '#059669' : '#6b7280'}; text-transform:uppercase;">${c.status}</span></td>
            <td style="padding:9px 12px; font-family:monospace; font-size:0.78rem; color:#1e293b; font-weight:600;">${c.endDateStr}</td>
            <td style="padding:9px 12px;"><span style="font-weight:800; color:${uc.text}; font-size:0.82rem; padding:3px 10px; border-radius:8px; background:${uc.badge};">${c.dDay}</span></td>
            <td style="padding:9px 12px; text-align:right; font-weight:600; font-size:0.78rem; color:#374151;">$${formatCurrency(c.tcv)}</td>
            <td style="padding:9px 12px; text-align:right; font-weight:600; font-size:0.78rem; color:#6366f1;">$${formatCurrency(c.arr)}</td>
            <td style="padding:9px 12px; font-size:0.75rem; color:#6b7280; max-width:140px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${c.services}</td>
        </tr>`;
    }).join('');

    html += `
        <div style="margin-bottom:8px; display:flex; align-items:center; gap:8px;">
            <div style="width:4px; height:18px; background:#f59e0b; border-radius:2px;"></div>
            <span style="font-size:0.78rem; font-weight:700; color:#374151; text-transform:uppercase; letter-spacing:0.06em;">③ Churn & Renewal Risk</span>
        </div>
        <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:14px; margin-bottom:16px;">
            <div class="stat-card" style="background:#FFF; padding:16px; border-left:4px solid #ef4444; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:4px;">Critical ≤30d</div>
                <div style="font-size:2rem; font-weight:800; color:#dc2626;">${stats.criticalCount}</div>
                <div style="font-size:0.72rem; color:#9ca3af; margin-top:4px;">Immediate action needed</div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-left:4px solid #f59e0b; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:4px;">Warning ≤90d</div>
                <div style="font-size:2rem; font-weight:800; color:#d97706;">${stats.warningCount}</div>
                <div style="font-size:0.72rem; color:#9ca3af; margin-top:4px;">Plan renewal outreach</div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-left:4px solid #10b981; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:4px;">Monitor ≤180d</div>
                <div style="font-size:2rem; font-weight:800; color:#059669;">${stats.normalExpCount}</div>
                <div style="font-size:0.72rem; color:#9ca3af; margin-top:4px;">Schedule QBR</div>
            </div>
        </div>
        ${renewalRows.length > 0 ? `
        <div class="stat-card" style="background:#FFF; padding:16px; margin-bottom:20px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
            <div style="font-size:0.85rem; font-weight:700; color:#111827; margin-bottom:12px; display:flex; align-items:center; gap:8px;">
                <i class="fa-solid fa-calendar-check" style="color:#f59e0b;"></i> Renewal Pipeline — Next 6 Months
                <span style="font-size:0.7rem; background:#fef3c7; color:#92400e; padding:2px 8px; border-radius:10px; font-weight:700;">${stats.expiringCount} contracts</span>
            </div>
            <div style="overflow-x:auto;">
                <table style="width:100%; border-collapse:collapse; min-width:800px; font-size:0.8rem;">
                    <thead><tr style="background:#f9fafb; text-align:left; border-bottom:2px solid #f1f5f9;">
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Customer</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Country</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Status</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">End Date</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">D-Day</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase; text-align:right;">TCV</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase; text-align:right;">ARR</th>
                        <th style="padding:9px 12px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Services</th>
                    </tr></thead>
                    <tbody>${renewalRows}</tbody>
                </table>
            </div>
        </div>` : '<div style="margin-bottom:20px;"></div>'}
    `;

    /* ══════════════════════════════════════
       SECTION 4 — Expansion Revenue
       ══════════════════════════════════════ */
    const expansionRows = stats.upsellTargets.slice(0, 12).map((t, i) => {
        const rank = i + 1;
        const rankColor = rank === 1 ? '#f59e0b' : rank === 2 ? '#9ca3af' : rank === 3 ? '#b45309' : '#e5e7eb';
        return `<tr style="border-bottom:1px solid #f3f4f6;">
            <td style="padding:9px 10px; text-align:center;">
                <span style="display:inline-block; width:22px; height:22px; background:${rankColor}; color:${rank <= 3 ? '#fff' : '#6b7280'}; border-radius:50%; font-size:0.68rem; font-weight:700; line-height:22px; text-align:center;">${rank}</span>
            </td>
            <td style="padding:9px 10px; font-weight:600; font-size:0.8rem; color:#111827;">${t.name}</td>
            <td style="padding:9px 10px; font-size:0.78rem; color:#6b7280;">${t.country || 'N/A'}</td>
            <td style="padding:9px 10px;"><span style="background:rgba(99,102,241,0.1); color:#6366f1; font-size:0.72rem; font-weight:700; padding:2px 8px; border-radius:10px;">${t.service}</span></td>
            <td style="padding:9px 10px; text-align:right; font-weight:700; font-size:0.8rem; color:#059669;">$${formatCurrency(t.tcv)}</td>
        </tr>`;
    }).join('');

    const multiPct = stats.totalCustomers > 0 ? Math.round((stats.multiServiceCustomers / stats.totalCustomers) * 100) : 0;

    html += `
        <div style="margin-bottom:8px; display:flex; align-items:center; gap:8px;">
            <div style="width:4px; height:18px; background:#10b981; border-radius:2px;"></div>
            <span style="font-size:0.78rem; font-weight:700; color:#374151; text-transform:uppercase; letter-spacing:0.06em;">④ Expansion Revenue</span>
        </div>
        <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:14px; margin-bottom:16px;">
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid #10b981; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:4px;">Upsell Targets</div>
                <div style="font-size:2rem; font-weight:800; color:#111827;">${stats.singleServiceCustomers}</div>
                <div style="font-size:0.72rem; color:#9ca3af; margin-top:4px;">Single-service customers</div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid #10b981; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:4px;">Expansion Opportunity</div>
                <div style="font-size:2rem; font-weight:800; color:#059669;">$${formatCurrency(stats.expansionOpportunity)}</div>
                <div style="font-size:0.72rem; color:#9ca3af; margin-top:4px;">TCV from upsell targets</div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; border-top:3px solid #6366f1; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; margin-bottom:4px;">Multi-Service Rate</div>
                <div style="font-size:2rem; font-weight:800; color:#6366f1;">${multiPct}<span style="font-size:1rem; font-weight:600;">%</span></div>
                <div style="font-size:0.72rem; color:#9ca3af; margin-top:4px;">${stats.multiServiceCustomers} of ${stats.totalCustomers} customers</div>
            </div>
        </div>
        <div class="stat-card" style="background:#FFF; padding:16px; margin-bottom:20px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
            <div style="font-size:0.85rem; font-weight:700; color:#111827; margin-bottom:12px; display:flex; align-items:center; gap:8px;">
                <i class="fa-solid fa-arrow-trend-up" style="color:#10b981;"></i> Top Upsell Targets
            </div>
            ${expansionRows.length > 0 ? `
            <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                <thead><tr style="background:#f9fafb; text-align:left; border-bottom:2px solid #f1f5f9;">
                    <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">#</th>
                    <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Customer</th>
                    <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Country</th>
                    <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase;">Current Service</th>
                    <th style="padding:8px 10px; font-size:0.68rem; color:#9ca3af; font-weight:700; text-transform:uppercase; text-align:right;">TCV</th>
                </tr></thead>
                <tbody>${expansionRows}</tbody>
            </table>` : '<p style="color:#9ca3af; text-align:center; padding:16px; font-size:0.85rem;">No single-service customers found.</p>'}
        </div>
    `;

    /* ══════════════════════════════════════
       SECTION 5 — Service Adoption
       ══════════════════════════════════════ */
    html += `
        <div style="margin-bottom:8px; display:flex; align-items:center; gap:8px;">
            <div style="width:4px; height:18px; background:#8b5cf6; border-radius:2px;"></div>
            <span style="font-size:0.78rem; font-weight:700; color:#374151; text-transform:uppercase; letter-spacing:0.06em;">⑤ Service Adoption</span>
        </div>
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:16px; margin-bottom:20px;">
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.85rem; font-weight:700; color:#111827; margin-bottom:12px;">Service Combination Ranking</div>
                <div style="height:280px;"><canvas id="service-donut-chart"></canvas></div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
                <div style="font-size:0.85rem; font-weight:700; color:#111827; margin-bottom:12px;">Adoption Breakdown</div>
                <div style="display:flex; flex-direction:column; gap:10px; margin-bottom:16px;">
                    <div>
                        <div style="display:flex; justify-content:space-between; font-size:0.78rem; margin-bottom:4px;">
                            <span style="color:#374151; font-weight:600;">Multi-Service</span>
                            <span style="color:#6366f1; font-weight:700;">${stats.multiServiceCustomers} (${multiPct}%)</span>
                        </div>
                        <div style="height:8px; background:#f1f5f9; border-radius:4px; overflow:hidden;">
                            <div style="height:100%; width:${multiPct}%; background:linear-gradient(90deg,#6366f1,#8b5cf6); border-radius:4px; transition:width 0.6s;"></div>
                        </div>
                    </div>
                    <div>
                        <div style="display:flex; justify-content:space-between; font-size:0.78rem; margin-bottom:4px;">
                            <span style="color:#374151; font-weight:600;">Single-Service (Upsell Opportunity)</span>
                            <span style="color:#f59e0b; font-weight:700;">${stats.singleServiceCustomers} (${100 - multiPct}%)</span>
                        </div>
                        <div style="height:8px; background:#f1f5f9; border-radius:4px; overflow:hidden;">
                            <div style="height:100%; width:${100 - multiPct}%; background:linear-gradient(90deg,#f59e0b,#f97316); border-radius:4px; transition:width 0.6s;"></div>
                        </div>
                    </div>
                </div>
                <div style="border-top:1px solid #f3f4f6; padding-top:12px;">
                    <div style="font-size:0.75rem; color:#6b7280; font-weight:700; text-transform:uppercase; margin-bottom:8px;">Top Combinations</div>
                    ${stats.sortedCombos.slice(0, 6).map((c, i) => `
                    <div style="display:flex; justify-content:space-between; align-items:center; padding:5px 0; border-bottom:1px solid #f9fafb;">
                        <span style="font-size:0.78rem; color:#374151; flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; padding-right:8px;">${c[0]}</span>
                        <span style="font-size:0.78rem; font-weight:700; color:#6366f1; flex-shrink:0;">${c[1]} customers</span>
                    </div>`).join('')}
                </div>
            </div>
        </div>
    `;

    return html;
}


export function getRenewalHTML(filtered) {
    if (filtered.length === 0) {
        return '<div style="padding:40px; text-align:center; color:#6B7280; grid-column:1/-1;">No renewals found in the next 6 months.</div>';
    }

    const headers = ['End User', 'Country', 'Status', 'End License Date', 'D-Day', 'TCV Amount', 'ARR Amount', 'Probability'];
    let tableHtml = `<div class="stat-card" style="grid-column:1/-1; padding:16px; background:#FFF; border: 1px solid #F3F4F6; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch;">
        <h2 style="font-size:1.1rem; font-weight:800; color:#111827; margin-bottom:12px; display:flex; align-items:center; gap:10px;">
            <i class="fa-solid fa-calendar-check" style="color:#ef4444;"></i> License Renewal Schedule (Next 6 Months)
        </h2>
        <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; min-width:1000px;">
                <thead><tr style="background:#F9FAFB; text-align:left; border-bottom:2px solid #F3F4F6;">`;

    headers.forEach(h => { tableHtml += `<th style="padding:10px; font-size:0.75rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">${h}</th>`; });
    tableHtml += `</tr></thead><tbody>`;

    filtered.forEach((row, i) => {
        const dDayColor = row.diffDays <= 30 ? '#ef4444' : (row.diffDays <= 90 ? '#f59e0b' : '#374151');
        const statusBg = row['Status'] === 'Closed' ? 'rgba(52, 199, 89, 0.1)' : 'rgba(0,122,255,0.1)';
        const statusColor = row['Status'] === 'Closed' ? '#34c759' : '#007AFF';

        tableHtml += `<tr style="border-bottom:1px solid #F3F4F6; background:${i % 2 === 0 ? 'transparent' : '#F9FBFF'}; transition: background 0.2s;">
            <td style="padding:10px; font-weight:700; color:#111827;">${row['End User'] || ''}</td>
            <td style="padding:10px; color:#4b5563;">${row['Country'] || ''}</td>
            <td style="padding:10px;"><span style="padding:4px 10px; border-radius:12px; font-size:0.7rem; font-weight:700; background:${statusBg}; color:${statusColor}; text-transform:uppercase;">${row['Status'] || ''}</span></td>
            <td style="padding:10px; color:#4b5563; font-family: monospace;">${row['endDateFormatted']}</td>
            <td style="padding:10px; font-weight:800; color:${dDayColor}">${row['D-Day']}</td>
            <td style="padding:10px; font-weight:600;">$${formatCurrency(row['TCV Amount'])}</td>
            <td style="padding:10px; font-weight:600;">$${formatCurrency(row['ARR Amount'])}</td>
            <td style="padding:10px; font-weight:700; color:#6366f1;">${row['Probability']}%</td>
        </tr>`;
    });

    tableHtml += `</tbody></table></div></div>`;
    return tableHtml;
}

/* ═══════════════════════════════════════════════════════════════
   Quarterly Forecast Panel (per-country, Q1–Q4)
   New (Booked) + New (Forecast, weighted) + Renewal targets
   ═══════════════════════════════════════════════════════════════ */
export function getQuarterlyForecastHTML(stats) {
    if (!stats) return '';
    const { country, currentYear, currentQuarter, quarters, kpiTotals } = stats;
    const qOrder = ['Q1', 'Q2', 'Q3', 'Q4'];
    const qIdx = q => qOrder.indexOf(q);
    const qIdxNow = qIdx(currentQuarter);

    const kpiCard = (label, value, color, sub) => `
        <div style="background:#fff; border-radius:10px; padding:14px 16px; border-left:4px solid ${color}; box-shadow:0 1px 4px rgba(0,0,0,0.04); display:flex; flex-direction:column; gap:4px; min-width:0;">
            <div style="font-size:0.66rem; font-weight:800; color:${color}; text-transform:uppercase; letter-spacing:0.06em;">${label}</div>
            <div style="font-size:1.35rem; font-weight:800; color:#111827; line-height:1.1;">$${formatCurrency(value)}</div>
            <div style="font-size:0.62rem; color:#9ca3af; font-weight:600;">${sub}</div>
        </div>
    `;

    const annualKpiStrip = `
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:8px;">
            <span style="font-size:0.62rem; font-weight:800; color:#6366f1; text-transform:uppercase; letter-spacing:0.1em; background:rgba(99,102,241,0.1); padding:3px 10px; border-radius:6px;">Annual Total · ${currentYear}</span>
            <span style="height:1px; flex:1; background:#E5E7EB;"></span>
        </div>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr)); gap:12px; margin-bottom:18px;">
            ${kpiCard('Booked TCV (Annual)', kpiTotals.bookedTcv, '#0ea5e9', 'Closed-won, Q1–Q4')}
            ${kpiCard('Booked ARR (Annual)', kpiTotals.bookedArr, '#10b981', 'Closed-won, Q1–Q4')}
            ${kpiCard('Weighted Pipeline ARR', kpiTotals.weightedArr, '#f59e0b', 'Stage-prob × ARR/yr')}
            ${kpiCard('Renewal Target ARR', kpiTotals.renewalArr, '#a855f7', 'Contracts ending in ' + currentYear)}
        </div>
    `;

    const miniKpiRow = (label, value, color) => `
        <div style="display:flex; justify-content:space-between; align-items:center; padding:5px 9px; border-bottom:1px solid #F1F5F9; gap:8px;">
            <span style="font-size:0.56rem; font-weight:800; color:${color}; text-transform:uppercase; letter-spacing:0.04em; white-space:nowrap;">${label}</span>
            <span style="font-size:0.78rem; font-weight:800; color:#111827; line-height:1;">$${formatCurrency(value)}</span>
        </div>
    `;

    const qKpiCol = (q) => {
        const data = quarters[q];
        const qi = qIdx(q);
        const isPast = qi < qIdxNow;
        const isCurrent = q === currentQuarter;
        const bg = isCurrent ? 'linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%)' : '#FAFAFA';
        const borderColor = isCurrent ? '#3B82F6' : '#E5E7EB';
        const labelColor = isCurrent ? '#1D4ED8' : '#374151';
        const tagText = isPast ? '✓ Closed' : (isCurrent ? '• Active' : 'Upcoming');
        const tagBg = isPast ? '#E5E7EB' : (isCurrent ? '#DBEAFE' : '#FEF3C7');
        const tagColor = isPast ? '#374151' : (isCurrent ? '#1E40AF' : '#92400E');
        return `
            <div onclick="window.openQuarterlyForecastModal('${q}')"
                 style="background:${bg}; border:1px solid ${borderColor}; border-radius:8px; padding:10px 6px 8px; opacity:${isPast ? 0.85 : 1}; min-width:0; cursor:pointer; transition:transform 0.15s, box-shadow 0.15s, border-color 0.15s;"
                 onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 10px 24px rgba(59,130,246,0.18)'; this.style.borderColor='#3B82F6';"
                 onmouseout="this.style.transform='none'; this.style.boxShadow='none'; this.style.borderColor='${borderColor}';"
                 title="Click to view ${q} deal details">
                <div style="display:flex; align-items:center; justify-content:space-between; padding:0 4px 8px; margin-bottom:4px; border-bottom:1px solid ${borderColor};">
                    <span style="font-size:0.75rem; font-weight:800; color:${labelColor}; letter-spacing:0.1em;">${q}</span>
                    <span style="display:flex; align-items:center; gap:6px;">
                        <span style="font-size:0.55rem; font-weight:800; padding:2px 7px; border-radius:8px; background:${tagBg}; color:${tagColor}; white-space:nowrap;">${tagText}</span>
                        <i class="fa-solid fa-up-right-and-down-left-from-center" style="font-size:0.6rem; color:#6366f1;" title="Click to expand"></i>
                    </span>
                </div>
                ${miniKpiRow('Booked TCV', data.booked.tcv, '#0ea5e9')}
                ${miniKpiRow('Booked ARR', data.booked.arr, '#10b981')}
                ${miniKpiRow('wPipeline TCV', data.forecast.wTcv, '#f59e0b')}
                ${miniKpiRow('wPipeline ARR', data.forecast.wArr, '#f59e0b')}
                ${miniKpiRow('Renewal ARR', data.renewal.arr, '#a855f7')}
                <div style="text-align:center; padding-top:6px; margin-top:4px; border-top:1px dashed ${borderColor}; font-size:0.55rem; font-weight:700; color:#6366f1; letter-spacing:0.05em;">
                    <i class="fa-solid fa-hand-pointer" style="font-size:0.55rem;"></i> Click for details
                </div>
            </div>
        `;
    };

    const quarterKpiRow = `
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:8px;">
            <span style="font-size:0.62rem; font-weight:800; color:#0ea5e9; text-transform:uppercase; letter-spacing:0.1em; background:rgba(14,165,233,0.1); padding:3px 10px; border-radius:6px;">Quarterly KPIs</span>
            <span style="height:1px; flex:1; background:#E5E7EB;"></span>
        </div>
        <div style="display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:12px; margin-bottom:14px;">
            ${qOrder.map(qKpiCol).join('')}
        </div>
    `;

    return `
        <div style="display:block; width:100%; padding:20px 22px; background:#fff; border-radius:14px; border:1px solid #E5E7EB; box-shadow:0 4px 16px rgba(0,0,0,0.04); box-sizing:border-box;">
            <div style="display:flex; align-items:center; gap:14px; margin-bottom:16px; padding-bottom:14px; border-bottom:1px solid #F3F4F6; flex-wrap:wrap; row-gap:8px;">
                <div style="background:rgba(99,102,241,0.1); color:#6366f1; width:42px; height:42px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;"><i class="fa-solid fa-chart-line" style="font-size:1.05rem;"></i></div>
                <div style="flex-shrink:0;">
                    <h3 style="margin:0; font-size:0.7rem; color:#6366f1; font-weight:800; text-transform:uppercase; letter-spacing:0.08em; line-height:1.2;">QUARTERLY FORECAST · ${currentYear}</h3>
                    <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827; line-height:1.3;">${country} — Q1 to Q4 (New + Renewal)</h2>
                </div>
                <div style="margin-left:auto; font-size:0.65rem; color:#6b7280; font-weight:600;">w = stage-weighted (확도 반영) · prev = existing ARR · <span style="color:#6366f1; font-weight:700;">Click any quarter to view deal details</span></div>
            </div>
            ${annualKpiStrip}
            ${quarterKpiRow}
        </div>
        <div id="qf-modal"
             style="display:none; position:fixed; inset:0; z-index:9999; align-items:center; justify-content:center; background:rgba(0,0,0,0.55); padding:24px; backdrop-filter:blur(2px);"
             onclick="if(event.target===this) window.closeQuarterlyForecastModal()">
            <div id="qf-modal-body" style="background:#fff; border-radius:16px; max-width:1100px; width:100%; max-height:88vh; overflow:auto; padding:0; position:relative; box-shadow:0 25px 60px rgba(0,0,0,0.25);"></div>
        </div>
    `;
}

/* ── Modal: full-detail Quarter view ─────────────────────────── */
window.openQuarterlyForecastModal = function (qId) {
    const stats = window.__qForecastStats;
    if (!stats) return;
    const q = stats.quarters[qId];
    if (!q) return;

    const qOrder = ['Q1', 'Q2', 'Q3', 'Q4'];
    const qIdxNow = qOrder.indexOf(stats.currentQuarter);
    const qi = qOrder.indexOf(qId);
    const isPast = qi < qIdxNow;
    const isCurrent = qId === stats.currentQuarter;
    const tagText = isPast ? '✓ Closed' : (isCurrent ? '• Active' : 'Upcoming');
    const tagBg = isPast ? '#E5E7EB' : (isCurrent ? '#DBEAFE' : '#FEF3C7');
    const tagColor = isPast ? '#374151' : (isCurrent ? '#1E40AF' : '#92400E');
    const accent = isCurrent ? '#3B82F6' : (isPast ? '#9CA3AF' : '#F59E0B');

    const fmt = window.__qfFmt || (n => Math.round(n || 0).toLocaleString('en-US'));

    const fullBookedRows = q.booked.deals.length === 0
        ? `<tr><td colspan="3" style="padding:18px; text-align:center; color:#9ca3af; font-style:italic;">No booked deals</td></tr>`
        : q.booked.deals.map((d, i) => `
            <tr style="border-bottom:1px solid #F3F4F6;">
                <td style="padding:9px 12px; color:#94A3B8; font-family:monospace; font-weight:700; width:30px;">${String(i + 1).padStart(2, '0')}</td>
                <td style="padding:9px 12px; color:#111827; font-weight:600;">${d.name}</td>
                <td style="padding:9px 12px; text-align:right; white-space:nowrap;">
                    <span style="color:#0EA5E9; font-weight:800;">TCV $${fmt(d.tcv)}</span>
                    <span style="color:#10B981; font-weight:700; margin-left:10px;">ARR $${fmt(d.arr)}</span>
                </td>
            </tr>
        `).join('');

    const fullForecastRows = q.forecast.deals.length === 0
        ? `<tr><td colspan="4" style="padding:18px; text-align:center; color:#9ca3af; font-style:italic;">No forecast deals</td></tr>`
        : q.forecast.deals.map((d, i) => `
            <tr style="border-bottom:1px solid #F3F4F6;">
                <td style="padding:9px 12px; color:#94A3B8; font-family:monospace; font-weight:700; width:30px;">${String(i + 1).padStart(2, '0')}</td>
                <td style="padding:9px 12px; color:#111827; font-weight:600;">${d.name}</td>
                <td style="padding:9px 12px;">${renderStageBadge(d.stage || 'Unknown', { fontSize: '0.62rem', padding: '2px 8px' })}</td>
                <td style="padding:9px 12px; text-align:right; white-space:nowrap;">
                    <div>
                        <span style="color:#0EA5E9; font-weight:700;">TCV $${fmt(d.tcv)}</span>
                        <span style="color:#10B981; font-weight:700; margin-left:10px;">ARR $${fmt(d.arr)}</span>
                    </div>
                    <div style="margin-top:3px;">
                        <span style="color:#fb923c; font-weight:800;">wTCV $${fmt(d.weighted)}</span>
                        <span style="color:#F59E0B; font-weight:700; margin-left:10px;">wARR $${fmt(d.wArr)}</span>
                    </div>
                </td>
            </tr>
        `).join('');

    const fullRenewalRows = q.renewal.deals.length === 0
        ? `<tr><td colspan="5" style="padding:18px; text-align:center; color:#9ca3af; font-style:italic;">No renewals</td></tr>`
        : q.renewal.deals.map((d, i) => `
            <tr style="border-bottom:1px solid #F3F4F6;">
                <td style="padding:9px 12px; color:#94A3B8; font-family:monospace; font-weight:700; width:30px;">${String(i + 1).padStart(2, '0')}</td>
                <td style="padding:9px 12px; color:#111827; font-weight:600;">${d.name}</td>
                <td style="padding:9px 12px; color:#6B7280; font-size:0.78rem;">${d.endDate}</td>
                <td style="padding:9px 12px;"><span style="background:rgba(168,85,247,0.12); color:#7e22ce; font-size:0.65rem; font-weight:800; padding:2px 8px; border-radius:8px;">${d.dDay}</span></td>
                <td style="padding:9px 12px; text-align:right; white-space:nowrap;">
                    <span style="color:#a855f7; font-weight:800;">Target $${fmt(d.targetArr)}</span>
                    <span style="color:#9ca3af; font-weight:600; margin-left:10px;">prev $${fmt(d.currentArr)}</span>
                </td>
            </tr>
        `).join('');

    const sectionWrap = (title, color, total, headerCols, rows) => `
        <div style="margin-bottom:22px;">
            <div style="display:flex; align-items:center; justify-content:space-between; padding:10px 14px; background:${color}11; border-left:4px solid ${color}; border-radius:8px; margin-bottom:8px;">
                <span style="font-size:0.78rem; font-weight:800; color:${color}; text-transform:uppercase; letter-spacing:0.05em;">${title}</span>
                <span style="font-size:0.78rem; font-weight:800; color:#111827;">${total}</span>
            </div>
            <div style="overflow:auto; border:1px solid #F3F4F6; border-radius:8px;">
                <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                    <thead style="background:#F9FAFB;">
                        <tr>${headerCols.map(h => `<th style="padding:9px 12px; text-align:${h.right ? 'right' : 'left'}; font-size:0.66rem; font-weight:800; color:#6b7280; text-transform:uppercase; letter-spacing:0.06em;">${h.label}</th>`).join('')}</tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
        </div>
    `;

    const body = document.getElementById('qf-modal-body');
    if (!body) return;

    body.innerHTML = `
        <div style="position:sticky; top:0; background:#fff; padding:20px 26px 16px; border-bottom:1px solid #F3F4F6; z-index:1; display:flex; align-items:center; gap:14px;">
            <div style="background:${accent}1A; color:${accent}; width:46px; height:46px; border-radius:12px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                <span style="font-size:1.15rem; font-weight:800;">${qId}</span>
            </div>
            <div style="flex:1; min-width:0;">
                <div style="font-size:0.7rem; color:#6b7280; font-weight:700; text-transform:uppercase; letter-spacing:0.06em;">${stats.country} · ${stats.currentYear}</div>
                <div style="font-size:1.15rem; font-weight:800; color:#111827; line-height:1.2;">${qId} — All Deals</div>
            </div>
            <span style="font-size:0.7rem; font-weight:800; padding:5px 12px; border-radius:14px; background:${tagBg}; color:${tagColor};">${tagText}</span>
            <button onclick="window.closeQuarterlyForecastModal()" style="background:#F3F4F6; border:none; width:36px; height:36px; border-radius:10px; cursor:pointer; display:flex; align-items:center; justify-content:center; color:#374151; font-size:1rem;" onmouseover="this.style.background='#E5E7EB'" onmouseout="this.style.background='#F3F4F6'">
                <i class="fa-solid fa-xmark"></i>
            </button>
        </div>
        <div style="padding:22px 26px 26px;">
            ${sectionWrap('NEW · Booked',  '#10b981',
                `TCV $${fmt(q.booked.tcv)} · ARR $${fmt(q.booked.arr)} · ${q.booked.deals.length} deals`,
                [{label:'#'},{label:'Deal Name'},{label:'Value', right:true}],
                fullBookedRows)}
            ${sectionWrap('NEW · Forecast (Stage-Weighted)', '#f59e0b',
                `TCV $${fmt(q.forecast.tcv)} · ARR $${fmt(q.forecast.arr)} · wTCV $${fmt(q.forecast.wTcv)} · wARR $${fmt(q.forecast.wArr)} · ${q.forecast.deals.length} deals`,
                [{label:'#'},{label:'Deal Name'},{label:'Stage'},{label:'Weighted Value', right:true}],
                fullForecastRows)}
            ${sectionWrap('Renewal Target', '#a855f7',
                `$${fmt(q.renewal.arr)} · ${q.renewal.deals.length} contracts`,
                [{label:'#'},{label:'Contract'},{label:'End Date'},{label:'D-Day'},{label:'ARR', right:true}],
                fullRenewalRows)}
        </div>
    `;
    document.getElementById('qf-modal').style.display = 'flex';
};

window.closeQuarterlyForecastModal = function () {
    const m = document.getElementById('qf-modal');
    if (m) m.style.display = 'none';
};

if (typeof window !== 'undefined' && !window.__qfEscBound) {
    window.__qfEscBound = true;
    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') window.closeQuarterlyForecastModal();
    });
    window.__qfFmt = (n) => Math.round(n || 0).toLocaleString('en-US');
}

export function getOrderSheetHTML(stats, filterCountry = null) {
    const currentYear = new Date().getFullYear();
    return `
        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 20px;">
            <div class="stat-card" style="border-left: 5px solid #0ea5e9; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; min-height: 140px;">
                <div class="metric-title-row"><h3 style="color:#0ea5e9; font-size:0.75rem; font-weight:700; margin:0;">ACCUMULATED TCV</h3><span class="metric-globe" data-tooltip="Local currency basis — this card sums contracts in each country's original currency. All other cards on this page are on Korea (USD) basis."><i class="fa-solid fa-globe"></i></span><span class="metric-info" data-tooltip="Total Contract Value of all closed-won deals, summed across all countries in local currency.">i</span></div>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">${formatCurrency(stats.sumLocalTcv)}</h2>
                <div style="font-size: 0.75rem; color: #6B7280; margin-bottom: 8px;">${stats.dealCount} Deals Total</div>
                <div style="flex: 1; position: relative; min-height: 70px;">
                    <canvas id="tcv-yearly-bar"></canvas>
                </div>
            </div>
            <div class="stat-card" onclick="window.openAccumulatedTcvModal()" style="border-left: 5px solid #6366f1; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; min-height: 140px; cursor: pointer; transition: transform 0.15s ease, box-shadow 0.15s ease;" onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 8px 16px -4px rgba(99,102,241,0.25)';" onmouseout="this.style.transform=''; this.style.boxShadow='0 4px 6px -1px rgba(0,0,0,0.1)';">
                <div class="metric-title-row"><h3 style="color:#6366f1; font-size:0.75rem; font-weight:700; margin:0;">ACCUMULATED KTCV</h3><span class="metric-info" data-tooltip="Total Contract Value converted to USD (Korea headquarters currency), aggregating global revenue in a single comparable unit.">i</span><span style="margin-left:auto; font-size:0.6rem; font-weight:700; color:#6366f1; background:rgba(99,102,241,0.1); padding:2px 8px; border-radius:10px; display:inline-flex; align-items:center; gap:4px;"><i class="fa-solid fa-up-right-and-down-left-from-center" style="font-size:0.55rem;"></i>Compare vs Local TCV</span></div>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.sumKorTcv)}</h2>
                <div style="font-size: 0.75rem; color: #6B7280; margin-bottom: 8px;">&nbsp;</div>
                <div style="flex: 1; position: relative; min-height: 70px;">
                    <canvas id="ktcv-yearly-bar"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #8b5cf6; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; position: relative; min-height: 120px;">
                <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 8px;">
                    <div class="metric-title-row"><h3 style="color:#8b5cf6; font-size:0.75rem; font-weight:700; margin:0;">ACCU ARR</h3><span class="metric-info" data-tooltip="Cumulative Annual Recurring Revenue — the annualized value of all active subscription contracts currently in force, snapshot at each year-end.">i</span><span class="metric-yoy-badge" data-tooltip="Year-over-Year (YoY) growth rate — labels above each yearly point show the % change in cumulative ARR vs. the previous year. Green = increase, red = decrease, gray N/A = prior-year value is zero or missing so a comparison can't be computed.">YoY %</span></div>
                    <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.sumArr)}</h2>
                </div>
                <div style="flex: 1; height: 80px; margin-top: auto; position: relative;">
                    <canvas id="arr-sparkline"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #8b5cf6; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; position: relative; min-height: 120px;">
                <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 8px;">
                    <div class="metric-title-row"><h3 style="color:#8b5cf6; font-size:0.75rem; font-weight:700; margin:0;">ARR GROWTH (YoY)</h3><span class="metric-info" data-tooltip="New ARR added each year — the year-over-year delta in cumulative ARR. Reveals which years contributed the most subscription growth, independent of cumulative scale.">i</span></div>
                    <h2 id="arr-growth-headline" style="font-size:1.6rem; font-weight:800; margin: 4px 0;">—</h2>
                </div>
                <div style="flex: 1; height: 80px; margin-top: auto; position: relative;">
                    <canvas id="arr-growth-bar"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #a855f7; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; position: relative; min-height: 120px;">
                <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 8px;">
                    <div class="metric-title-row"><h3 style="color:#a855f7; font-size:0.75rem; font-weight:700; margin:0;">ACCU MRR</h3><span class="metric-info" data-tooltip="Cumulative Monthly Recurring Revenue — the monthly value of active subscription contracts (ARR ÷ 12), snapshot at each year-end.">i</span><span class="metric-yoy-badge" data-tooltip="Year-over-Year (YoY) growth rate — labels above each yearly point show the % change in cumulative MRR vs. the previous year. Green = increase, red = decrease, gray N/A = prior-year value is zero or missing so a comparison can't be computed.">YoY %</span></div>
                    <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.sumMrr)}</h2>
                </div>
                <div style="flex: 1; height: 80px; margin-top: auto; position: relative;">
                    <canvas id="mrr-sparkline"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #a855f7; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; position: relative; min-height: 120px;">
                <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 8px;">
                    <div class="metric-title-row"><h3 style="color:#a855f7; font-size:0.75rem; font-weight:700; margin:0;">MRR GROWTH (YoY)</h3><span class="metric-info" data-tooltip="New MRR added each year — the year-over-year delta in cumulative MRR. Reveals which years contributed the most subscription growth, independent of cumulative scale.">i</span></div>
                    <h2 id="mrr-growth-headline" style="font-size:1.6rem; font-weight:800; margin: 4px 0;">—</h2>
                </div>
                <div style="flex: 1; height: 80px; margin-top: auto; position: relative;">
                    <canvas id="mrr-growth-bar"></canvas>
                </div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #f59e0b; display: flex; flex-direction: column; align-items: stretch;">
                <div class="metric-title-row" style="margin-bottom:8px;"><h3 style="color:#f59e0b; font-size:0.75rem; font-weight:700; margin:0;">QUARTERLY TCV (${currentYear})</h3><span class="metric-info" data-tooltip="Total Contract Value broken down by quarter for the current year, showing seasonal revenue distribution.">i</span></div>
                <div style="height:160px; position:relative;"><canvas id="quarterly-tcv-bar"></canvas></div>
            </div>
            <div class="stat-card" style="grid-column: span 2; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #0ea5e9; display: flex; flex-direction: column; align-items: stretch;">
                <div class="metric-title-row" style="margin-bottom:8px;"><h3 style="color:#0ea5e9; font-size:0.75rem; font-weight:700; margin:0;">TCV CAGR</h3><span class="metric-info" data-tooltip="Compound Annual Growth Rate — the smoothed annual growth rate from the baseline year to each subsequent year. Top number shows the full-period CAGR.">i</span></div>
                <div style="display:flex; gap:20px; align-items:stretch;">
                    <div style="flex:1.2; display:flex; flex-direction:column; min-width:0;">
                        <div id="tcv-cagr-headline" style="font-size:1.6rem; font-weight:800; margin: 0 0 4px;">—</div>
                        <div id="tcv-cagr-sub" style="font-size:0.7rem; color:#64748b; margin-bottom:6px;"></div>
                        <div style="flex:1; height:160px; position:relative;"><canvas id="tcv-cagr-chart"></canvas></div>
                    </div>
                    <div id="tcv-cagr-insights" style="flex:1; border-left:1px solid #e2e8f0; padding-left:16px; font-size:0.72rem; color:#475569; line-height:1.55;"></div>
                </div>
            </div>
            ${(!filterCountry || filterCountry === 'All') ? `
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #6366f1; display:flex; flex-direction:column;">
                <div class="metric-title-row" style="margin-bottom:12px;"><h3 style="color:#6366f1; font-size:0.75rem; font-weight:700; margin:0;">ACCUMULATED KTCV / COUNTRY</h3><span class="metric-info" data-tooltip="Geographic breakdown of total USD TCV, showing each country's share of global revenue.">i</span></div>
                <div style="display: flex; gap: 20px; align-items: center; flex: 1; min-height: 220px;">
                    <div style="position: relative; width: 180px; height: 180px; flex-shrink: 0;">
                        <canvas id="country-tcv-donut"></canvas>
                    </div>
                    <div id="country-tcv-legend" style="flex: 1; min-width: 0;"></div>
                </div>
            </div>` : ''}
            ${(!filterCountry || filterCountry === 'All') ? `
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #f97316; display:flex; flex-direction:column;">
                <div class="metric-title-row" style="margin-bottom:12px;"><h3 style="color:#f97316; font-size:0.75rem; font-weight:700; margin:0;">YoY KTCV GROWTH BY COUNTRY</h3><span class="metric-info" data-tooltip="Year-over-year KTCV growth rate by country, identifying which markets are expanding or contracting.">i</span></div>
                <div style="flex: 1; min-height: 220px; position: relative;">
                    <canvas id="country-yoy-bar"></canvas>
                </div>
            </div>` : ''}
        </div>
        <div id="accum-tcv-modal"
             style="display:none; position:fixed; inset:0; z-index:9999; align-items:center; justify-content:center; background:rgba(0,0,0,0.55); padding:24px; backdrop-filter:blur(2px);"
             onclick="if(event.target===this) window.closeAccumulatedTcvModal()">
            <div id="accum-tcv-modal-body" style="background:#fff; border-radius:16px; max-width:1100px; width:100%; max-height:88vh; overflow:auto; padding:0; position:relative; box-shadow:0 25px 60px rgba(0,0,0,0.25);"></div>
        </div>
    `;
}

/* ── Modal: ACCUMULATED TCV vs ACCUMULATED KTCV comparison ───────── */
window.openAccumulatedTcvModal = function () {
    const stats = window.__orderSheetStats;
    if (!stats) return;
    const body = document.getElementById('accum-tcv-modal-body');
    if (!body) return;

    const years = Object.keys(stats.yearlyTcv || {}).sort();
    const localYearly = years.map(y => stats.yearlyTcv[y].local || 0);
    const koreaYearly = years.map(y => stats.yearlyTcv[y].korea || 0);

    body.innerHTML = `
        <div style="position:sticky; top:0; background:#fff; padding:20px 26px 16px; border-bottom:1px solid #F3F4F6; z-index:1; display:flex; align-items:center; gap:14px;">
            <div style="background:rgba(99,102,241,0.1); color:#6366f1; width:46px; height:46px; border-radius:12px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                <i class="fa-solid fa-scale-balanced" style="font-size:1.15rem;"></i>
            </div>
            <div style="flex:1; min-width:0;">
                <div style="font-size:0.7rem; color:#6b7280; font-weight:700; text-transform:uppercase; letter-spacing:0.06em;">Currency Comparison</div>
                <div style="font-size:1.15rem; font-weight:800; color:#111827; line-height:1.2;">ACCUMULATED TCV vs ACCUMULATED KTCV</div>
            </div>
            <button onclick="window.closeAccumulatedTcvModal()" style="background:#F3F4F6; border:none; width:36px; height:36px; border-radius:10px; cursor:pointer; display:flex; align-items:center; justify-content:center; color:#374151; font-size:1rem;" onmouseover="this.style.background='#E5E7EB'" onmouseout="this.style.background='#F3F4F6'">
                <i class="fa-solid fa-xmark"></i>
            </button>
        </div>
        <div style="padding:22px 26px 26px;">
            <div style="font-size:0.78rem; color:#475569; background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px 14px; margin-bottom:18px; line-height:1.55;">
                <b style="color:#0ea5e9;">ACCUMULATED TCV</b> is in each country's <b>local currency</b> (not directly comparable across countries).
                <b style="color:#6366f1;"> ACCUMULATED KTCV</b> converts every contract to <b>USD (Korea HQ basis)</b> for a single comparable revenue view.
            </div>
            <div style="display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:18px;">
                <div class="stat-card" style="border-left:5px solid #0ea5e9; background:#FFF; padding:18px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.08); display:flex; flex-direction:column; align-items:stretch; min-height:340px;">
                    <div class="metric-title-row"><h3 style="color:#0ea5e9; font-size:0.75rem; font-weight:700; margin:0;">ACCUMULATED TCV</h3><span class="metric-globe" data-tooltip="Local currency basis"><i class="fa-solid fa-globe"></i></span></div>
                    <div style="font-size:0.65rem; color:#94a3b8; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; margin-top:2px;">Local currency · ${stats.dealCount} deals</div>
                    <h2 style="font-size:1.8rem; font-weight:800; margin:6px 0 4px;">${formatCurrency(stats.sumLocalTcv)}</h2>
                    <div style="flex:1; position:relative; min-height:240px; margin-top:8px;">
                        <canvas id="tcv-yearly-bar-modal"></canvas>
                    </div>
                </div>
                <div class="stat-card" style="border-left:5px solid #6366f1; background:#FFF; padding:18px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.08); display:flex; flex-direction:column; align-items:stretch; min-height:340px;">
                    <div class="metric-title-row"><h3 style="color:#6366f1; font-size:0.75rem; font-weight:700; margin:0;">ACCUMULATED KTCV</h3></div>
                    <div style="font-size:0.65rem; color:#94a3b8; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; margin-top:2px;">Korea (USD) basis</div>
                    <h2 style="font-size:1.8rem; font-weight:800; margin:6px 0 4px;">US$ ${formatCurrency(stats.sumKorTcv)}</h2>
                    <div style="flex:1; position:relative; min-height:240px; margin-top:8px;">
                        <canvas id="ktcv-yearly-bar-modal"></canvas>
                    </div>
                </div>
            </div>
        </div>
    `;

    document.getElementById('accum-tcv-modal').style.display = 'flex';

    /* Render stacked yearly bars inside the modal — mirrors the dashboard cards. */
    setTimeout(() => _renderAccumTcvCompareCharts(years, localYearly, koreaYearly), 50);
};

window.closeAccumulatedTcvModal = function () {
    const m = document.getElementById('accum-tcv-modal');
    if (m) m.style.display = 'none';
    if (window.__accumTcvCharts) {
        Object.values(window.__accumTcvCharts).forEach(c => { try { c.destroy(); } catch {} });
        window.__accumTcvCharts = null;
    }
};

if (typeof window !== 'undefined' && !window.__accumTcvEscBound) {
    window.__accumTcvEscBound = true;
    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') window.closeAccumulatedTcvModal();
    });
}

/**
 * Draw the two stacked cumulative bars (Local TCV + KTCV) inside the modal.
 * Mirrors the chart style used on the main ORDER SHEET dashboard so the
 * comparison is visually consistent with the small cards the user clicked from.
 */
function _renderAccumTcvCompareCharts(years, localYearly, koreaYearly) {
    if (typeof Chart === 'undefined') return;
    if (window.__accumTcvCharts) {
        Object.values(window.__accumTcvCharts).forEach(c => { try { c.destroy(); } catch {} });
    }
    window.__accumTcvCharts = {};

    const yearShade = (rgb, idx, total) => {
        const minOpacity = 0.28, maxOpacity = 0.92;
        const t = total <= 1 ? 1 : idx / (total - 1);
        const opacity = (minOpacity + t * (maxOpacity - minOpacity)).toFixed(2);
        return `rgba(${rgb}, ${opacity})`;
    };

    const buildStackedDatasets = (rgb, yearlyValues) => years.map((year, yearIdx) => ({
        label: year,
        data: years.map((_, barIdx) => barIdx >= yearIdx ? (yearlyValues[yearIdx] || 0) : 0),
        backgroundColor: yearShade(rgb, yearIdx, years.length),
        borderColor: `rgb(${rgb})`,
        borderWidth: 0,
        borderRadius: (ctx) => {
            const isTopOfBar = ctx.datasetIndex === ctx.dataIndex;
            const isBottomOfBar = ctx.datasetIndex === 0;
            return {
                topLeft: isTopOfBar ? 4 : 0,
                topRight: isTopOfBar ? 4 : 0,
                bottomLeft: isBottomOfBar ? 4 : 0,
                bottomRight: isBottomOfBar ? 4 : 0
            };
        },
        borderSkipped: false
    }));

    const stackedOptions = (yearlyValues, currencyPrefix) => ({
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: { display: false },
            tooltip: {
                mode: 'index',
                intersect: false,
                backgroundColor: '#1e293b',
                padding: 10,
                cornerRadius: 6,
                filter: (ctx) => (ctx.parsed.y || 0) > 0,
                itemSort: (a, b) => b.datasetIndex - a.datasetIndex,
                callbacks: {
                    label: (ctx) => {
                        const yearly = ctx.parsed.y || 0;
                        const prev = yearlyValues[ctx.datasetIndex - 1] || 0;
                        const growth = (ctx.datasetIndex > 0 && prev > 0)
                            ? `  (${(((yearly / prev) - 1) * 100).toFixed(1)}% YoY)`
                            : '';
                        return ` ${ctx.dataset.label}: ${currencyPrefix}${formatCurrency(yearly)}${growth}`;
                    },
                    footer: (items) => {
                        const total = items.reduce((s, it) => s + (it.parsed.y || 0), 0);
                        return `Cumulative: ${currencyPrefix}${formatCurrency(total)}`;
                    }
                }
            }
        },
        scales: {
            x: { stacked: true, grid: { display: false }, border: { display: false }, ticks: { color: '#94a3b8', font: { size: 10, weight: '700' } } },
            y: { stacked: true, display: false }
        },
        layout: { padding: { top: 22, bottom: 0, left: 4, right: 4 } }
    });

    const yoyLabelsPlugin = (yearlyValues) => ({
        id: 'accumTcvYoyLabels',
        afterDatasetsDraw(chart) {
            const { ctx } = chart;
            ctx.save();
            ctx.font = '700 11px Inter, system-ui, sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            for (let i = 1; i < yearlyValues.length; i++) {
                const meta = chart.getDatasetMeta(i);
                const bar = meta && meta.data && meta.data[i];
                if (!bar) continue;
                const prev = yearlyValues[i - 1] || 0;
                const curr = yearlyValues[i] || 0;
                let text, color;
                if (prev <= 0) {
                    if (curr <= 0) continue;
                    text = 'N/A'; color = '#94a3b8';
                } else {
                    const growth = ((curr / prev) - 1) * 100;
                    const sign = growth >= 0 ? '+' : '';
                    text = `${sign}${growth.toFixed(1)}%`;
                    color = growth >= 0 ? '#10b981' : '#ef4444';
                }
                ctx.fillStyle = color;
                ctx.fillText(text, bar.x, bar.y - 4);
            }
            ctx.restore();
        }
    });

    const tcvCanvas = document.getElementById('tcv-yearly-bar-modal');
    if (tcvCanvas && years.length > 0) {
        window.__accumTcvCharts.tcv = new Chart(tcvCanvas, {
            type: 'bar',
            data: { labels: years, datasets: buildStackedDatasets('14, 165, 233', localYearly) },
            options: stackedOptions(localYearly, ''),
            plugins: [yoyLabelsPlugin(localYearly)]
        });
    }

    const ktcvCanvas = document.getElementById('ktcv-yearly-bar-modal');
    if (ktcvCanvas && years.length > 0) {
        window.__accumTcvCharts.ktcv = new Chart(ktcvCanvas, {
            type: 'bar',
            data: { labels: years, datasets: buildStackedDatasets('99, 102, 241', koreaYearly) },
            options: stackedOptions(koreaYearly, 'US$ '),
            plugins: [yoyLabelsPlugin(koreaYearly)]
        });
    }
}


/**
 * Helper to get country flag image or font-awesome icon
 * @param {string} country 
 * @param {string} defaultIcon 
 * @returns {string} HTML string
 */
function getCountryFlagHTML(country, defaultIcon = 'fa-globe') {
    if (!country || country === 'All') {
        return `<i class="fa-solid ${defaultIcon}"></i>`;
    }

    const flags = {
        'Indonesia': 'id', 'Thailand': 'th', 'Malaysia': 'my', 'USA': 'us',
        'Philippines': 'ph', 'Singapore': 'sg', 'Vietnam': 'vn', 'Turkey': 'tr'
    };

    const code = flags[country];
    if (code) {
        return `<img src="https://flagcdn.com/w160/${code}.png" style="width: 100%; height: 100%; object-fit: cover;" alt="${country}">`;
    }
    return `<i class="fa-solid ${defaultIcon}"></i>`;
}


export function getPipelineHTML(stats, filterCountry, tabName) {
    const currentYear = new Date().getFullYear();
    const pipelineItemsHtml = stats.sortedPipeline.map(([country, values]) => `
        <div style="display: flex; flex-direction: column; padding: 10px; background: #F9FAFB; border-radius: 8px; border-left: 3px solid #10b981;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                <span style="font-weight: 700; color: #374151; font-size: 0.8rem;">${filterCountry ? 'Total Summary' : country}</span>
                <span style="background: rgba(16,185,129,0.12); color: #059669; font-size: 0.68rem; font-weight: 800; padding: 2px 8px; border-radius: 10px;">${values.count || 0} deals</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.7rem; margin-bottom: 2px;">
                <span style="color: #ef4444; opacity: 0.8;">TCV (${currentYear})</span>
                <span style="color: #ef4444; font-weight: 600;">$${formatCurrency(values.tcv || 0)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.7rem; margin-bottom: 2px;">
                <span style="color: var(--text-muted);">PIPELINE</span>
                <span style="color: #34C759; font-weight: 600;">$${formatCurrency(values.amount)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.7rem; margin-bottom: 2px;">
                <span style="color: var(--text-muted);">WEIGHTED PIPELINE VALUE</span>
                <span style="color: #007AFF; font-weight: 600;">$${formatCurrency(values.weighted)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.7rem;">
                <span style="color: #a855f7; opacity: 0.85;">ARR (${currentYear})</span>
                <span style="color: #a855f7; font-weight: 600;">$${formatCurrency(values.arr || 0)}</span>
            </div>
        </div>
    `).join('');

    const quarterlyItemsHtml = stats.sortedQuarterly.map(([q, qData]) => {
        const countryEntries = Object.entries(qData.countries);
        const qTotalAmount = countryEntries.reduce((acc, curr) => acc + curr[1].amount, 0);
        const qTotalWeighted = countryEntries.reduce((acc, curr) => acc + curr[1].weighted, 0);
        const qTotalTcv = countryEntries.reduce((acc, curr) => acc + (curr[1].tcv || 0), 0);
        const qTotalArr = countryEntries.reduce((acc, curr) => acc + (curr[1].arr || 0), 0);
        const qTotalCount = countryEntries.reduce((acc, curr) => acc + (curr[1].count || 0), 0);
        const currentYear = new Date().getFullYear();

        const countryBreakdown = countryEntries
            .sort(sortCountriesByAmount)
            .map(([country, values]) => `
                <div style="margin-top: 6px; padding: 6px; background: rgba(255, 255, 255, 0.03); border-radius: 4px; border: 1px solid #F3F4F6;">
                    <div style="font-weight: 600; color: #111827; font-size: 0.72rem; margin-bottom: 4px; display: flex; align-items: center; justify-content: space-between;">
                        <span style="display: flex; align-items: center; gap: 4px;"><i class="fa-solid fa-location-dot" style="font-size: 0.6rem; color: #34C759;"></i> ${filterCountry ? 'Total' : country}</span>
                        <span style="background: rgba(99,102,241,0.1); color: #6366f1; font-size: 0.6rem; font-weight: 700; padding: 1px 6px; border-radius: 8px;">${values.count || 0}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem; margin-bottom: 2px;">
                        <span style="color: #ef4444; opacity: 0.8;">TCV (${currentYear})</span>
                        <span style="color: #ef4444; font-weight: 600;">$${formatCurrency(values.tcv || 0)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem; margin-bottom: 2px;">
                        <span style="color: var(--text-muted);">PIPELINE</span>
                        <span style="color: #34C759;">$${formatCurrency(values.amount)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem; margin-bottom: 2px;">
                        <span style="color: var(--text-muted);">WEIGHTED PIPELINE VALUE</span>
                        <span style="color: #007AFF;">$${formatCurrency(values.weighted)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem;">
                        <span style="color: #a855f7; opacity: 0.85;">ARR (${currentYear})</span>
                        <span style="color: #a855f7; font-weight: 600;">$${formatCurrency(values.arr || 0)}</span>
                    </div>
                </div>
            `).join('');

        const dealListJson = JSON.stringify(qData.deals.slice(0, 50).map(d => ({ n: d.name, a: formatCurrency(d.weighted), w: d.weighted, t: d.amount, tf: formatCurrency(d.amount), c: d.country || '', s: d.stage || 'Unknown', m: d.month == null ? null : d.month })));
        const dealListAttr = dealListJson
            .replace(/&/g, '&amp;')
            .replace(/'/g, '&apos;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');

        return `
            <div class="quarter-card"
                 data-q="${q}"
                 data-show-country="${filterCountry ? 'false' : 'true'}"
                 data-deals='${dealListAttr}'
                 style="display: flex; flex-direction: column; padding: 10px; background: #F9FAFB; border-radius: 8px; border-top: 3px solid #10b981; cursor: pointer; transition: all 0.2s; height: 100%; min-height: 0;"
                 onmouseover="showQuarterTooltip(event, this)" 
                 onmouseout="hideQuarterTooltip()"
                 onclick="selectQuarter(this)">
                <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px; border-bottom: 1px solid #E5E7EB; padding-bottom: 6px;">
                    <div style="display: flex; flex-direction: column; gap: 2px;">
                        <span style="font-weight: 800; color: #111827; font-size: 0.85rem;">${q}</span>
                        <span style="background: rgba(16,185,129,0.12); color: #059669; font-size: 0.65rem; font-weight: 800; padding: 2px 8px; border-radius: 10px; text-align: center;">${qTotalCount} deals</span>
                    </div>
                    <div style="text-align: right; display: flex; flex-direction: column; gap: 2px;">
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px;">
                            <span style="font-size: 0.6rem; color: #ef4444; text-transform: uppercase;">TCV (${currentYear})</span>
                            <span style="font-size: 0.85rem; color: #ef4444; font-weight: 800;">$${formatCurrency(qTotalTcv)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px;">
                            <span style="font-size: 0.6rem; color: var(--text-secondary); text-transform: uppercase;">PIPELINE</span>
                            <span style="font-size: 0.85rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalAmount)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px;">
                            <span style="font-size: 0.6rem; color: var(--text-secondary); text-transform: uppercase;">WEIGHTED PIPELINE VALUE</span>
                            <span style="font-size: 0.85rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalWeighted)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px;">
                            <span style="font-size: 0.6rem; color: #a855f7; text-transform: uppercase;">ARR (${currentYear})</span>
                            <span style="font-size: 0.85rem; color: #a855f7; font-weight: 800;">$${formatCurrency(qTotalArr)}</span>
                        </div>
                    </div>
                </div>
                ${filterCountry ? `
                <div style="background: rgba(0,0,0,0.04); border-radius: 8px; padding: 8px 10px; margin-top: 8px; flex: 1; display: flex; flex-direction: column; min-height: 0;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                        <span style="font-size: 0.6rem; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">All Deals</span>
                        <span style="font-size: 0.6rem; color: #6b7280; font-weight: 700; background: rgba(99,102,241,0.1); padding: 1px 6px; border-radius: 8px;">${qData.deals.length}</span>
                    </div>
                    ${qData.deals.length === 0 ? `
                        <div style="text-align: center; padding: 16px 8px; color: #9CA3AF; font-size: 0.7rem; font-style: italic;">No active deals</div>
                    ` : `
                        <div style="overflow-y: auto; flex: 1; max-height: 360px; padding-right: 2px;">
                            ${qData.deals.map((d, idx) => `
                                <div style="display: flex; justify-content: space-between; align-items: center; gap: 6px; padding: 5px 4px; border-bottom: 1px solid #EEF2F6; font-size: 0.7rem;">
                                    <span style="color: #94A3B8; font-family: monospace; font-weight: 700; flex-shrink: 0; width: 18px;">${String(idx + 1).padStart(2, '0')}</span>
                                    <span style="color: #1F2937; font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${String(d.name).replace(/"/g, '&quot;')}">${d.name}</span>
                                    ${renderStageBadge(d.stage || 'Unknown', { fontSize: '0.55rem', padding: '1px 6px' })}
                                    <span style="display:flex; flex-direction:column; align-items:flex-end; gap:1px; flex-shrink:0; line-height:1;">
                                        <span style="color: #10B981; font-weight: 800; font-size:0.7rem;" title="Weighted">W $${formatCurrency(d.weighted)}</span>
                                        <span style="color: #a855f7; font-weight: 700; font-size:0.6rem;" title="ARR">ARR $${formatCurrency(d.arr || 0)}</span>
                                    </span>
                                </div>
                            `).join('')}
                        </div>
                    `}
                </div>
                ` : `
                <div style="background: rgba(0,0,0,0.05); border-radius: 8px; padding: 10px; margin-top: 8px; flex: 1; display: flex; flex-direction: column; min-height: 0;">
                    <div style="flex: 1; overflow-y: auto; min-height: 0;">
                        ${countryBreakdown}
                    </div>
                </div>
                `}
            </div>
        `;
    }).join('');

    injectPipelineTooltipStyles();

    // ── Pipeline by Deal Stage (aggregated across all current-year quarters) ──
    const STAGE_ORDER = ['Discovery', 'PoC', 'Quotation', 'Negotiation', 'Lost', 'Won', 'Unknown'];
    const stageEntries = (stats.sortedStage || []).slice().sort((a, b) => {
        const ai = STAGE_ORDER.indexOf(a[0]); const bi = STAGE_ORDER.indexOf(b[0]);
        const ax = ai === -1 ? 999 : ai; const bx = bi === -1 ? 999 : bi;
        if (ax !== bx) return ax - bx;
        return b[1].weighted - a[1].weighted;
    });
    const stageGrandPipeline = stageEntries.reduce((s, [, v]) => s + (v.amount || 0), 0);
    const stageGrandWeighted = stageEntries.reduce((s, [, v]) => s + (v.weighted || 0), 0);
    const stageGrandCount = stageEntries.reduce((s, [, v]) => s + (v.count || 0), 0);

    const stageTableRowsHtml = stageEntries.length === 0
        ? `<tr><td colspan="4" style="padding:14px; text-align:center; color:#9ca3af; font-size:0.78rem;">No stage data available.</td></tr>`
        : stageEntries.map(([stageName, v]) => `
            <tr style="border-bottom:1px solid #f1f5f9;">
                <td style="padding:10px 14px;">${renderStageBadge(stageName, { fontSize: '0.7rem', padding: '4px 10px' })}</td>
                <td style="padding:10px 14px; text-align:center; color:#475569; font-weight:700; font-size:0.8rem;">${v.count || 0}</td>
                <td style="padding:10px 14px; text-align:right; color:#34C759; font-weight:800; font-size:0.9rem;">$${formatCurrency(v.amount || 0)}</td>
                <td style="padding:10px 14px; text-align:right; color:#007AFF; font-weight:800; font-size:0.9rem;">$${formatCurrency(v.weighted || 0)}</td>
            </tr>
        `).join('');

    const stageTableFooterHtml = stageEntries.length === 0 ? '' : `
        <tr style="background:#f8fafc; border-top:2px solid #cbd5e1;">
            <td style="padding:12px 14px; font-size:0.75rem; font-weight:900; color:#111827; text-transform:uppercase; letter-spacing:0.04em;">Total</td>
            <td style="padding:12px 14px; text-align:center; color:#111827; font-weight:900; font-size:0.85rem;">${stageGrandCount}</td>
            <td style="padding:12px 14px; text-align:right; color:#047857; font-weight:900; font-size:0.95rem;">$${formatCurrency(stageGrandPipeline)}</td>
            <td style="padding:12px 14px; text-align:right; color:#1e40af; font-weight:900; font-size:0.95rem;">$${formatCurrency(stageGrandWeighted)}</td>
        </tr>
    `;

    const stageSummaryHtml = `
        <div style="border-top: 1px solid #E5E7EB; padding-top: 12px;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom: 12px; flex-wrap: wrap; gap: 8px;">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div class="stat-icon" style="width: 32px; height: 32px; font-size: 0.9rem; background: rgba(139, 92, 246, 0.15); color: #8b5cf6;"><i class="fa-solid fa-layer-group"></i></div>
                    <div>
                        <h2 style="font-size: 0.95rem; font-weight: 700; color: #111827; margin: 0;">Pipeline by Deal Stage (${currentYear})</h2>
                        <p style="font-size: 0.7rem; color: #6b7280; margin: 2px 0 0;">Aggregated across Q1–Q4 — counts, Pipeline (TCV) and Weighted Pipeline Value per stage.</p>
                    </div>
                </div>
            </div>
            <div style="background: #FFFFFF; border-radius: 12px; border: 1px solid #e5e7eb; box-shadow: 0 2px 6px rgba(0,0,0,0.03); overflow-x: auto;">
                <table style="width:100%; border-collapse:collapse; min-width:520px;">
                    <thead>
                        <tr style="background:#f8fafc; border-bottom:2px solid #cbd5e1;">
                            <th style="padding:10px 14px; text-align:left; font-size:0.68rem; font-weight:800; color:#374151; letter-spacing:0.06em; text-transform:uppercase;">Deal Stage</th>
                            <th style="padding:10px 14px; text-align:center; font-size:0.68rem; font-weight:800; color:#374151; letter-spacing:0.06em; text-transform:uppercase;">Deals</th>
                            <th style="padding:10px 14px; text-align:right; font-size:0.68rem; font-weight:800; color:#34C759; letter-spacing:0.06em; text-transform:uppercase;">Pipeline</th>
                            <th style="padding:10px 14px; text-align:right; font-size:0.68rem; font-weight:800; color:#007AFF; letter-spacing:0.06em; text-transform:uppercase;">Weighted Pipeline Value</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${stageTableRowsHtml}
                        ${stageTableFooterHtml}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    // ── New Pipeline Volume Matrix (Country × Stage × Quarter, ARR & TCV) ──
    // Build a country → stage → quarter aggregate from per-deal rows so each
    // country can be drilled down into its deal-stage composition while still
    // rolling up to the same per-country subtotal beneath it.
    const stageMatrix = {};
    stats.sortedQuarterly.forEach(([q, qData]) => {
        qData.deals.forEach(d => {
            const country = d.country || 'Other';
            const stage = d.stage || 'Unknown';
            if (!stageMatrix[country]) stageMatrix[country] = {};
            if (!stageMatrix[country][stage]) {
                stageMatrix[country][stage] = {
                    Q1: { tcv: 0, arr: 0, count: 0 },
                    Q2: { tcv: 0, arr: 0, count: 0 },
                    Q3: { tcv: 0, arr: 0, count: 0 },
                    Q4: { tcv: 0, arr: 0, count: 0 }
                };
            }
            stageMatrix[country][stage][q].tcv += d.amount || 0;
            stageMatrix[country][stage][q].arr += d.arr || 0;
            stageMatrix[country][stage][q].count += 1;
        });
    });

    const matrixCountrySet = new Set();
    stats.sortedQuarterly.forEach(([, qData]) => Object.keys(qData.countries).forEach(c => matrixCountrySet.add(c)));
    Object.keys(stageMatrix).forEach(c => matrixCountrySet.add(c));
    const matrixCountries = Array.from(matrixCountrySet).sort();

    const stageRank = (s) => {
        const i = STAGE_ORDER.indexOf(s);
        return i === -1 ? 999 : i;
    };

    const matrixRowData = matrixCountries.map(country => {
        const byQ = {};
        let totalTcv = 0, totalArr = 0, totalCount = 0;
        stats.sortedQuarterly.forEach(([q, qData]) => {
            const v = qData.countries[country] || { amount: 0, arr: 0, count: 0 };
            byQ[q] = { tcv: v.amount || 0, arr: v.arr || 0, count: v.count || 0 };
            totalTcv += v.amount || 0;
            totalArr += v.arr || 0;
            totalCount += v.count || 0;
        });

        const stagesPresent = Object.keys(stageMatrix[country] || {})
            .sort((a, b) => stageRank(a) - stageRank(b));
        const stageRows = stagesPresent.map(stage => {
            const sd = stageMatrix[country][stage];
            let sTcv = 0, sArr = 0, sCount = 0;
            ['Q1', 'Q2', 'Q3', 'Q4'].forEach(q => {
                sTcv += sd[q].tcv;
                sArr += sd[q].arr;
                sCount += sd[q].count;
            });
            return { stage, byQ: sd, totalTcv: sTcv, totalArr: sArr, totalCount: sCount };
        });

        return { country, byQ, totalTcv, totalArr, totalCount, stageRows };
    }).sort((a, b) => b.totalTcv - a.totalTcv);

    const matrixColumnTotals = { Q1: { tcv: 0, arr: 0, count: 0 }, Q2: { tcv: 0, arr: 0, count: 0 }, Q3: { tcv: 0, arr: 0, count: 0 }, Q4: { tcv: 0, arr: 0, count: 0 } };
    let matrixGrandTcv = 0, matrixGrandArr = 0, matrixGrandCount = 0;
    matrixRowData.forEach(r => {
        ['Q1', 'Q2', 'Q3', 'Q4'].forEach(q => {
            const c = r.byQ[q] || { tcv: 0, arr: 0, count: 0 };
            matrixColumnTotals[q].tcv += c.tcv;
            matrixColumnTotals[q].arr += c.arr;
            matrixColumnTotals[q].count += c.count;
        });
        matrixGrandTcv += r.totalTcv;
        matrixGrandArr += r.totalArr;
        matrixGrandCount += r.totalCount;
    });

    const matrixQuarters = ['Q1', 'Q2', 'Q3', 'Q4'];
    // Per-quarter color palette — distinct hue for each Q so the eye can group columns at a glance
    const QUARTER_THEME = {
        Q1: { header: '#1e40af', headerBg: '#dbeafe', cellBg: '#f5f9ff', divider: '#93c5fd', emoji: '①' },
        Q2: { header: '#047857', headerBg: '#d1fae5', cellBg: '#f3fbf7', divider: '#6ee7b7', emoji: '②' },
        Q3: { header: '#b45309', headerBg: '#fef3c7', cellBg: '#fffbf2', divider: '#fcd34d', emoji: '③' },
        Q4: { header: '#6d28d9', headerBg: '#ede9fe', cellBg: '#faf7ff', divider: '#c4b5fd', emoji: '④' }
    };
    const Q_DIVIDER = '3px solid';
    const TOTAL_DIVIDER = '4px double #10b981';

    const matrixHeaderHtml = matrixQuarters.map((q) => {
        const t = QUARTER_THEME[q];
        return `
            <th colspan="2" style="padding: 10px 8px; text-align: center; font-size: 0.78rem; font-weight: 900; color: ${t.header}; background: ${t.headerBg}; letter-spacing: 0.05em; border-left: ${Q_DIVIDER} ${t.divider}; border-bottom: 2px solid ${t.divider};">
                <span style="display:inline-block; background: ${t.header}; color:#fff; width:18px; height:18px; line-height:18px; border-radius:50%; font-size:0.7rem; margin-right:6px; text-align:center; font-weight:900;">${q.charAt(1)}</span>${q} PIPELINE
            </th>
        `;
    }).join('');

    const matrixSubHeaderHtml = matrixQuarters.map((q) => {
        const t = QUARTER_THEME[q];
        return `
            <th style="padding: 6px 8px; text-align: right; font-size: 0.62rem; font-weight: 800; color: #b91c1c; background: ${t.cellBg}; border-left: ${Q_DIVIDER} ${t.divider}; border-bottom: 1px solid ${t.divider};">TCV</th>
            <th style="padding: 6px 8px; text-align: right; font-size: 0.62rem; font-weight: 800; color: #4338ca; background: ${t.cellBg}; border-bottom: 1px solid ${t.divider};">ARR</th>
        `;
    }).join('');

    const fmtCell = (v) => v > 0 ? `$${formatCurrency(v)}` : '<span style="color:#cbd5e1;">—</span>';

    // Each country renders as a group: one row per deal-stage, then a country
    // subtotal row. Stage rows use a quiet white background, the subtotal row
    // is shaded so it visually closes the country block.
    const totalColumnCount = 3 + matrixQuarters.length * 2 + 2; // Country + Stage + Deals + 8 + 2

    const renderQuarterCells = (byQ, opts = {}) => matrixQuarters.map(q => {
        const c = byQ[q] || { tcv: 0, arr: 0 };
        const t = QUARTER_THEME[q];
        const fontSize = opts.fontSize || '0.78rem';
        const fontWeight = opts.fontWeight || '600';
        const bg = opts.bg || t.cellBg;
        return `
            <td style="padding: 8px 10px; font-size: ${fontSize}; text-align: right; color: #b91c1c; font-weight: ${fontWeight}; background: ${bg}; border-left: ${Q_DIVIDER} ${t.divider}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(c.tcv)}</td>
            <td style="padding: 8px 10px; font-size: ${fontSize}; text-align: right; color: #4338ca; font-weight: ${fontWeight}; background: ${bg}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(c.arr)}</td>
        `;
    }).join('');

    const matrixBodyHtml = matrixRowData.length === 0 ? `
        <tr><td colspan="${totalColumnCount}" style="padding: 16px; text-align: center; color: #9ca3af; font-size: 0.78rem;">No pipeline deals tagged with a quarter for ${currentYear}.</td></tr>
    ` : matrixRowData.map(r => {
        const stageRowCount = r.stageRows.length;
        const groupRowSpan = stageRowCount + 1; // stages + subtotal

        const stageRowsHtml = r.stageRows.map((s, sIdx) => {
            const isFirst = sIdx === 0;
            const countryCell = isFirst ? `
                <td rowspan="${groupRowSpan}" style="padding: 10px 12px; font-size: 0.82rem; font-weight: 800; color: #111827; white-space: nowrap; border-right: 2px solid #e5e7eb; vertical-align: top; background: #f8fafc;">${r.country}</td>
            ` : '';
            return `
                <tr style="border-bottom: 1px solid #f1f5f9; background: #ffffff;">
                    ${countryCell}
                    <td style="padding: 8px 10px; font-size: 0.72rem; white-space: nowrap;">${renderStageBadge(s.stage, { fontSize: '0.62rem', padding: '2px 8px' })}</td>
                    <td style="padding: 8px 10px; font-size: 0.72rem; color: #6b7280; text-align: center; border-right: ${Q_DIVIDER} #cbd5e1;">${s.totalCount}</td>
                    ${renderQuarterCells(s.byQ)}
                    <td style="padding: 8px 12px; font-size: 0.78rem; text-align: right; color: #b91c1c; font-weight: 700; background: #fff5f5; border-left: ${TOTAL_DIVIDER}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(s.totalTcv)}</td>
                    <td style="padding: 8px 12px; font-size: 0.78rem; text-align: right; color: #4338ca; font-weight: 700; background: #f5f7ff; overflow: hidden; text-overflow: ellipsis;">${fmtCell(s.totalArr)}</td>
                </tr>
            `;
        }).join('');

        // When a country has no stage breakdown (e.g. only order-sheet TCV with
        // no pipeline deals), still render one subtotal row that spans the
        // country column.
        const fallbackCountryCell = stageRowCount === 0 ? `
            <td style="padding: 10px 12px; font-size: 0.82rem; font-weight: 800; color: #111827; white-space: nowrap; border-right: 2px solid #e5e7eb; background: #f8fafc;">${r.country}</td>
        ` : '';

        const subtotalRow = `
            <tr style="border-bottom: 2px solid #cbd5e1; background: #f1f5f9;">
                ${fallbackCountryCell}
                <td style="padding: 9px 10px; font-size: 0.7rem; font-weight: 800; color: #475569; text-transform: uppercase; letter-spacing: 0.05em;">Subtotal</td>
                <td style="padding: 9px 10px; font-size: 0.78rem; color: #1e293b; font-weight: 800; text-align: center; border-right: ${Q_DIVIDER} #94a3b8;">${r.totalCount}</td>
                ${renderQuarterCells(r.byQ, { fontSize: '0.8rem', fontWeight: '800', bg: '#eef2f7' })}
                <td style="padding: 9px 12px; font-size: 0.85rem; text-align: right; color: #b91c1c; font-weight: 900; background: #fee2e2; border-left: ${TOTAL_DIVIDER}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(r.totalTcv)}</td>
                <td style="padding: 9px 12px; font-size: 0.85rem; text-align: right; color: #4338ca; font-weight: 900; background: #e0e7ff; overflow: hidden; text-overflow: ellipsis;">${fmtCell(r.totalArr)}</td>
            </tr>
        `;

        return stageRowsHtml + subtotalRow;
    }).join('');

    const matrixFooterHtml = matrixRowData.length === 0 ? '' : `
        <tr style="background: #ecfdf5; border-top: 3px solid #10b981;">
            <td colspan="2" style="padding: 12px 12px; font-size: 0.82rem; font-weight: 900; color: #047857; border-right: 2px solid #cbd5e1; text-transform: uppercase; letter-spacing: 0.05em;">Grand Total</td>
            <td style="padding: 12px 10px; font-size: 0.78rem; font-weight: 900; color: #475569; text-align: center; border-right: ${Q_DIVIDER} #94a3b8;">${matrixGrandCount}</td>
            ${matrixQuarters.map(q => {
                const t = QUARTER_THEME[q];
                return `
                    <td style="padding: 12px 10px; font-size: 0.82rem; text-align: right; color: #b91c1c; font-weight: 900; background: ${t.headerBg}; border-left: ${Q_DIVIDER} ${t.divider}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(matrixColumnTotals[q].tcv)}</td>
                    <td style="padding: 12px 10px; font-size: 0.82rem; text-align: right; color: #4338ca; font-weight: 900; background: ${t.headerBg}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(matrixColumnTotals[q].arr)}</td>
                `;
            }).join('')}
            <td style="padding: 12px 12px; font-size: 0.9rem; text-align: right; color: #b91c1c; font-weight: 900; background: #fecaca; border-left: ${TOTAL_DIVIDER}; overflow: hidden; text-overflow: ellipsis;">${fmtCell(matrixGrandTcv)}</td>
            <td style="padding: 12px 12px; font-size: 0.9rem; text-align: right; color: #4338ca; font-weight: 900; background: #c7d2fe; overflow: hidden; text-overflow: ellipsis;">${fmtCell(matrixGrandArr)}</td>
        </tr>
    `;

    return `
        <div style="padding: 16px; background: #EDFAF1; border-radius: 16px; border: 1px solid rgba(16, 185, 129, 0.15); display: flex; flex-direction: column; gap: 16px;">
            ${tabName === 'PIPELINE' ? `
            <div class="stat-card" style="display:flex; align-items:center; gap:12px; padding: 10px 16px; background: #FFFFFF; border: 1px solid rgba(16, 185, 129, 0.2); border-left: 4px solid #10b981; margin-bottom: 0;">
                <label style="font-size:0.8rem; color:#34C759; font-weight:700; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right: 8px;"></i>Select Country</label>
                <select id="pipeline-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:6px 12px; border-radius:8px; width: 200px; font-size: 0.85rem;">
                    ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
                <span style="font-size: 0.72rem; color: var(--text-secondary); margin-left: auto;">Metrics for ${filterCountry || 'All Regions'}</span>
            </div>
            ` : ''}

            <div style="background: rgba(52,199,89,0.08); border: 1px solid rgba(16, 185, 129, 0.2); border-radius: 12px; padding: 12px; display: flex; align-items: center; justify-content: space-between;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <div class="stat-icon" style="width: 40px; height: 40px; font-size: 1.1rem; background: rgba(16, 185, 129, 0.2); color: #34C759; overflow: hidden; display: flex; align-items: center; justify-content: center;">
                        ${getCountryFlagHTML(filterCountry)}
                    </div>
                    <div>
                        <h2 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin: 0;">${filterCountry ? 'Total Pipeline' : 'Global Total Pipeline'}</h2>
                        <p style="font-size: 0.75rem; color: var(--text-secondary); margin: 0;">Aggregated metrics</p>
                    </div>
                </div>
                <div style="display: flex; gap: 24px; text-align: right; align-items: flex-start;">
                    <div style="background: rgba(16,185,129,0.08); border: 1px solid rgba(16,185,129,0.2); border-radius: 12px; padding: 8px 14px; display: flex; flex-direction: column; align-items: center;">
                        <span style="font-size: 0.65rem; color: #059669; text-transform: uppercase; font-weight: 700;">DEALS</span>
                        <h2 style="font-size: 1.5rem; font-weight: 900; color: #059669; margin: 0; line-height: 1.1;">${stats.globalTotalCount || 0}</h2>
                    </div>
                    <div>
                        <span style="font-size: 0.7rem; color: #34C759; text-transform: uppercase;">PIPELINE</span>
                        <h2 style="font-size: 1.25rem; font-weight: 800; color: #111827; margin: 0;">US$ ${formatCurrency(stats.globalTotalAmount)}</h2>
                    </div>
                    <div>
                        <span style="font-size: 0.7rem; color: #007AFF; text-transform: uppercase;">WEIGHTED PIPELINE VALUE</span>
                        <h2 style="font-size: 1.25rem; font-weight: 800; color: #111827; margin: 0;">US$ ${formatCurrency(stats.globalTotalWeighted)}</h2>
                    </div>
                </div>
            </div>

            ${!filterCountry ? `
            <div>
                <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 12px;">
                    <div class="stat-icon" style="width: 32px; height: 32px; font-size: 0.9rem; background: rgba(16, 185, 129, 0.15); color: #34C759;"><i class="fa-solid fa-earth-americas"></i></div>
                    <h2 style="font-size: 0.95rem; font-weight: 600; color: #111827;">${currentYear} Pipeline by Country</h2>
                </div>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 10px;">
                    ${pipelineItemsHtml}
                </div>
            </div>
            ` : ''}

            <div style="border-top: 1px solid #E5E7EB; margin-top: 4px; padding-top: 12px;">
                <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 12px;">
                    <div class="stat-icon" style="width: 32px; height: 32px; font-size: 0.9rem; background: rgba(20, 184, 166, 0.15); color: #14b8a6;"><i class="fa-solid fa-globe"></i></div>
                    <h2 style="font-size: 0.95rem; font-weight: 600; color: #111827;">${new Date().getFullYear()} Pipeline Quarter</h2>
                </div>

                <!-- Quarter Cards (full width) -->
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px;">
                    ${quarterlyItemsHtml}
                </div>

                <!-- Pipeline % by Quarter — horizontal bar (placed below quarter cards) -->
                <div class="stat-card" style="background: #FFFFFF; padding: 14px 18px; border-radius: 12px; border: 1px solid rgba(16, 185, 129, 0.1); box-shadow: 0 4px 12px rgba(0,0,0,0.03); margin-top: 14px;">
                    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px;">
                        <h3 style="font-size: 0.9rem; font-weight: 700; color: #111827; margin: 0;">Pipeline % by Quarter</h3>
                        <span style="font-size: 0.7rem; color: #6B7280; background: #F3F4F6; padding: 2px 8px; border-radius: 10px; font-weight: 600;">VALUE BASE</span>
                    </div>
                    <div style="position: relative; height: 200px;">
                        <canvas id="pipeline-quarter-pie-chart"></canvas>
                    </div>
                </div>

                <div id="pipeline-selected-quarter-container" style="margin-top: 20px; display: none;"></div>
            </div>
            <div id="pipeline-quarter-tooltip" class="pipeline-tooltip" style="width: 280px; pointer-events: none;"></div>

            ${stageSummaryHtml}

            <div style="border-top: 1px solid #E5E7EB; padding-top: 12px;">
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; flex-wrap: wrap; gap: 8px;">
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <div class="stat-icon" style="width: 32px; height: 32px; font-size: 0.9rem; background: rgba(99, 102, 241, 0.15); color: #6366f1;"><i class="fa-solid fa-table-cells"></i></div>
                        <div>
                            <h2 style="font-size: 0.95rem; font-weight: 700; color: #111827; margin: 0;">New Pipeline Volume by Country (${currentYear})</h2>
                            <p style="font-size: 0.7rem; color: #6b7280; margin: 2px 0 0;">Quarterly pipeline — TCV & ARR side by side. ARR = TCV ÷ Contract Yr (PIPELINE 시트의 Contract Yr 컬럼 값. 비어있으면 1년 가정)</p>
                        </div>
                    </div>
                    <div style="display: flex; gap: 8px; font-size: 0.65rem; font-weight: 700;">
                        <span style="background:#fef2f2; color:#ef4444; padding: 3px 10px; border-radius: 10px; border:1px solid #fecaca;">TCV</span>
                        <span style="background:#eef2ff; color:#6366f1; padding: 3px 10px; border-radius: 10px; border:1px solid #c7d2fe;">ARR</span>
                    </div>
                </div>
                <div style="background: #FFFFFF; border-radius: 12px; border: 1px solid #e5e7eb; box-shadow: 0 2px 6px rgba(0,0,0,0.03); overflow-x: auto;">
                    <table style="width: 100%; border-collapse: collapse; font-family: inherit; table-layout: fixed; min-width: 1320px;">
                        <colgroup>
                            <col style="width: 9%;">
                            <col style="width: 9%;">
                            <col style="width: 5%;">
                            <col style="width: 7.7%;"><col style="width: 7.7%;">
                            <col style="width: 7.7%;"><col style="width: 7.7%;">
                            <col style="width: 7.7%;"><col style="width: 7.7%;">
                            <col style="width: 7.7%;"><col style="width: 7.7%;">
                            <col style="width: 7.7%;"><col style="width: 7.7%;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th rowspan="2" style="padding: 12px 14px; text-align: left; font-size: 0.72rem; font-weight: 800; color: #374151; border-bottom: 2px solid #cbd5e1; background: #f8fafc; letter-spacing: 0.06em; text-transform: uppercase; border-right: 2px solid #e5e7eb;">Country</th>
                                <th rowspan="2" style="padding: 12px 10px; text-align: left; font-size: 0.72rem; font-weight: 800; color: #374151; border-bottom: 2px solid #cbd5e1; background: #f8fafc; letter-spacing: 0.06em; text-transform: uppercase; border-right: 2px solid #e5e7eb;">Deal Stage</th>
                                <th rowspan="2" style="padding: 12px 10px; text-align: center; font-size: 0.72rem; font-weight: 800; color: #374151; border-bottom: 2px solid #cbd5e1; background: #f8fafc; letter-spacing: 0.06em; text-transform: uppercase; border-right: 3px solid #cbd5e1;">Deals</th>
                                ${matrixHeaderHtml}
                                <th colspan="2" style="padding: 10px 8px; text-align: center; font-size: 0.78rem; font-weight: 900; color: #047857; border-bottom: 2px solid #10b981; background: #a7f3d0; letter-spacing: 0.05em; border-left: ${TOTAL_DIVIDER};">
                                    <span style="display:inline-block; background:#047857; color:#fff; padding:2px 10px; border-radius: 12px; font-size:0.72rem; font-weight:900;">∑ TOTAL</span>
                                </th>
                            </tr>
                            <tr>
                                ${matrixSubHeaderHtml}
                                <th style="padding: 6px 10px; text-align: right; font-size: 0.62rem; font-weight: 900; color: #b91c1c; background: #fee2e2; border-bottom: 1px solid #fca5a5; border-left: ${TOTAL_DIVIDER};">TCV</th>
                                <th style="padding: 6px 10px; text-align: right; font-size: 0.62rem; font-weight: 900; color: #4338ca; background: #c7d2fe; border-bottom: 1px solid #a5b4fc;">ARR</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${matrixBodyHtml}
                            ${matrixFooterHtml}
                        </tbody>
                    </table>
                </div>
            </div>

            <div style="border-top: 1px solid #E5E7EB; padding-top: 16px;">
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px;">
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <div class="stat-icon" style="background: rgba(34, 197, 94, 0.15); color: #34C759; width: 32px; height: 32px;"><i class="fa-solid fa-chart-line"></i></div>
                        <div>
                            <h3 style="font-size: 0.95rem; font-weight: 700; color: #111827; margin: 0;">New Influx Analysis (${currentYear} Monthly)</h3>
                        </div>
                    </div>
                </div>
                <div style="position: relative; height: 280px;">
                    <canvas id="pipeline-influx-chart"></canvas>
                </div>
            </div>
        </div>
    `;
}

/**
 * Generate HTML for the COLLECTION category dashboard
 * @param {Object} stats - Basic statistics
 * @param {Object} detailedStats - Detailed analysis (rows and summary)
 * @param {boolean} showOnlyUnpaid - Whether to filter for balance > 0
 */
export function getCollectionHTML(stats, detailedStats, showOnlyUnpaid = false) {
    const currentYear = new Date().getFullYear();
    const totalCollectionRate = stats.globalTotalTcv > 0
        ? Math.round((stats.globalTotalReceived / stats.globalTotalTcv) * 100)
        : 0;

    return `
        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 20px;">
            <div class="stat-card" style="border-left: 4px solid #6366f1; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: flex-start;">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">ACCUMULATED KTCV</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.globalTotalTcv)}</h2>
            </div>
            <div class="stat-card" style="border-left: 4px solid #8b5cf6; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: flex-start;">
                <h3 style="color:#8b5cf6; font-size:0.75rem; font-weight:700;">ACCUMULATED KTCV RECEIVED</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.globalTotalReceived)}</h2>
                <div style="font-size: 0.7rem; color: #94a3b8;">Total Progress: ${totalCollectionRate}% of TCV</div>
            </div>
        </div>

        <div class="stat-card" style="background:#FFF; padding: 16px; margin-bottom: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: stretch;">
            <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 15px; display: flex; align-items: center; gap: 8px;">
                <i class="fa-solid fa-chart-line" style="color: #6366f1;"></i> Yearly Collection Performance Graph
            </h3>
            <div style="height: 280px; position: relative;"><canvas id="collection-performance-chart"></canvas></div>
            <!-- Container for Year Specific Detail List -->
            <div id="collection-year-detail-container" style="margin-top: 20px; display: none;"></div>
        </div>

        <div class="stat-card" style="background:#FFF; padding: 16px; margin-top: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: stretch;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                <h3 style="font-size: 1rem; font-weight: 700; margin: 0; color: #6366f1; display: flex; align-items: center; gap: 8px;">
                    <i class="fa-solid fa-list-check"></i> Collection Fulfillment Status (by Contract Start)
                </h3>
                <button id="collection-unpaid-toggle" style="background: ${showOnlyUnpaid ? '#ef4444' : '#6366f1'}; color: white; border: none; padding: 6px 12px; border-radius: 6px; font-size: 0.75rem; font-weight: 700; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <i class="fa-solid ${showOnlyUnpaid ? 'fa-filter-circle-xmark' : 'fa-filter'}"></i>
                    ${showOnlyUnpaid ? 'Show All Items' : 'Show Only Unpaid (Balance > 0)'}
                </button>
            </div>
            <div style="overflow-x: auto;">
                <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                    <thead>
                        <tr style="background:#F8FAFC; text-align:left; border-bottom: 1px solid #E2E8F0;">
                            <th style="padding:10px; color:#475569; font-weight:700;">Start</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">Distributor</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">End User</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">Yr</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">${currentYear - 2}</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">${currentYear - 1}</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">${currentYear}</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">TCV</th>
                            <th style="padding:10px; color:#475569; font-weight:700;">Balance</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${detailedStats.rows.map(r => `
                            <tr style="border-bottom: 1px solid #F1F5F9;">
                                <td style="padding:6px 10px; font-family: monospace; white-space: nowrap;">${r.contractStartDateStr}</td>
                                <td style="padding:6px 10px; font-weight:600;">${r.distributor}</td>
                                <td style="padding:6px 10px;">${r.endUser}</td>
                                <td style="padding:6px 10px; text-align:center;">${r.contractYr}</td>
                                <td style="padding:6px 10px; font-weight:700; color: ${r['status' + (currentYear - 2)].includes('✅') ? '#10b981' : (r['status' + (currentYear - 2)].includes('❌') ? '#ef4444' : '#94a3b8')}">${r['status' + (currentYear - 2)]}</td>
                                <td style="padding:6px 10px; font-weight:700; color: ${r['status' + (currentYear - 1)].includes('✅') ? '#10b981' : (r['status' + (currentYear - 1)].includes('❌') ? '#ef4444' : '#94a3b8')}">${r['status' + (currentYear - 1)]}</td>
                                <td style="padding:6px 10px; font-weight:700; color: ${r['status' + currentYear].includes('✅') ? '#10b981' : (r['status' + currentYear].includes('❌') ? '#ef4444' : '#94a3b8')}">${r['status' + currentYear]}</td>
                                <td style="padding:6px 10px; font-weight:700; color:#1e293b;">$${formatCurrency(r.totalTcv)}</td>
                                <td style="padding:6px 10px; font-weight:800; color: ${r.balance > 0 ? '#ef4444' : '#10b981'};">$${formatCurrency(r.balance)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>

        <div class="stat-card" style="background:#FFF; padding: 16px; margin-top: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); width: 100%; max-width: 600px; display: flex; flex-direction: column; align-items: stretch;">
            <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 12px; color: #1e293b; display: flex; align-items: center; gap: 8px;">
                <i class="fa-solid fa-calculator"></i> Total Balance by Distributor (Descending)
            </h3>
            <div style="overflow-x: auto;">
                <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                    <thead>
                        <tr style="background:#F8FAFC; text-align:left; border-bottom: 1px solid #E2E8F0;">
                            <th style="padding:10px; color:#475569; font-weight:700;">Distributor</th>
                            <th style="padding:10px; color:#475569; font-weight:700; text-align:right;">Total Balance</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${detailedStats.summary.map(s => `
                            <tr style="border-bottom: 1px solid #F1F5F9;">
                                <td style="padding:8px 10px; font-weight:600;">${s.name}</td>
                                <td style="padding:8px 10px; font-weight:800; text-align:right; color: ${s.balance > 0 ? '#ef4444' : '#10b981'};">$${formatCurrency(s.balance)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

/**
 * 특정 연도 클릭 시 대리점별 상세 목표 리스트 생성
 */
window.renderCollectionYearDetail = function (year, distributorData) {
    const container = document.getElementById('collection-year-detail-container');
    if (!container) return;

    const sortedDistributors = Object.entries(distributorData).sort((a, b) => b[1] - a[1]);
    const totalAmount = sortedDistributors.reduce((acc, curr) => acc + curr[1], 0);

    let rowsHtml = sortedDistributors.map(([name, amount]) => `
        <tr style="border-bottom: 1px solid #F1F5F9;">
            <td style="padding: 10px; font-weight: 600; color: #1e293b;">${name}</td>
            <td style="padding: 10px; text-align: right; color: #6366f1; font-weight: 700;">$${formatCurrency(amount)}</td>
        </tr>
    `).join('');

    container.innerHTML = `
        <div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 16px; animation: fadeIn 0.3s ease-out;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; border-bottom: 2px solid #E2E8F0; padding-bottom: 8px;">
                <h4 style="margin: 0; font-size: 0.95rem; color: #1e293b; font-weight: 800;">
                    <i class="fa-solid fa-calendar-day" style="color: #6366f1; margin-right: 6px;"></i> ${year} Collection Target Breakdown
                </h4>
                <span style="font-size: 0.85rem; font-weight: 800; color: #6366f1;">Total: $${formatCurrency(totalAmount)}</span>
            </div>
            <div style="max-height: 300px; overflow-y: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.8rem;">
                    <thead>
                        <tr style="text-align: left; color: #64748b; font-weight: 700; text-transform: uppercase; font-size: 0.7rem;">
                            <th style="padding: 8px 10px;">Distributor</th>
                            <th style="padding: 8px 10px; text-align: right;">Target Amount (ARR)</th>
                        </tr>
                    </thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            </div>
        </div>
    `;
    container.style.display = 'block';
};

export function getPartnerHTML(stats, filterCountry, tabName) {
    const displayCountries = CONFIG.COUNTRIES.filter(c => (!filterCountry || filterCountry === 'All') || c === filterCountry);
    const totalPartners = CONFIG.COUNTRIES.reduce((sum, c) => sum + (stats.counts[c] || 0), 0);

    const globalCardHtml = (!filterCountry || filterCountry === 'All') ? `
        <div class="stat-card" style="margin:0; padding: 16px; background: #FFFFFF; border: 1px solid #10B981; border-radius: 16px; display: flex; align-items: center; gap: 12px; position: relative; overflow: hidden; box-shadow: 0 4px 12px rgba(16, 185, 129, 0.15);">
            <div class="stat-icon" style="width: 40px; height: 40px; min-width: 40px; border-radius: 50%; background: #10B981; display: flex; align-items: center; justify-content: center; color: white;">
                <i class="fa-solid fa-earth-americas" style="font-size: 1.2rem;"></i>
            </div>
            <div>
                <h4 style="margin: 0; font-size: 0.7rem; color: #10B981; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 800;">GLOBAL</h4>
                <div style="display: flex; align-items: baseline; gap: 4px;">
                    <span style="font-size: 1.5rem; font-weight: 800; color: #111827; line-height: 1;">${totalPartners}</span>
                    <span style="font-size: 0.75rem; color: #9CA3AF; font-weight: 500;">Partners</span>
                </div>
            </div>
        </div>
    ` : '';

    const statsCardsHtml = displayCountries.map(c => {
        const count = stats.counts[c] || 0;
        return `
            <div class="stat-card" style="margin:0; padding: 12px 16px; background: #FFFFFF; border: 1px solid #F3F4F6; border-radius: 12px; display: flex; align-items: center; gap: 12px; position: relative; overflow: hidden;">
                <div class="stat-icon" style="width: 36px; height: 36px; min-width: 36px; border-radius: 50%; overflow: hidden; border: 1px solid #E5E7EB; padding: 0; background: #000; display: flex; align-items: center; justify-content: center;">
                    ${getCountryFlagHTML(c)}
                </div>
                <div>
                    <h4 style="margin: 0; font-size: 0.7rem; color: #6B7280; text-transform: uppercase; font-weight: 700;">${c}</h4>
                    <div style="display: flex; align-items: baseline; gap: 4px;">
                        <span style="font-size: 1.4rem; font-weight: 800; color: #111827; line-height: 1;">${count}</span>
                        <span style="font-size: 0.72rem; color: #9CA3AF; font-weight: 500;">Partners</span>
                    </div>
                </div>
            </div>
        `;
    }).join('');

    const yearFilter = stats.partnerYearFilter || 'all';
    const curY = stats.currentYear;
    const prevY = stats.previousYear;
    const yearLabel = yearFilter === 'current' ? `${curY}` : yearFilter === 'previous' ? `${prevY}` : 'All Years';

    const buildRankingRows = (list, valueCellFn) => list.slice(0, 10).map((p, idx) => `
        <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
            <td style="padding: 8px 12px; font-weight: 800; color: ${idx < 3 ? '#fbbf24' : '#94a3b8'}; width: 40px;">
                ${idx + 1}${idx < 3 ? ' <i class="fa-solid fa-crown" style="font-size: 0.65rem; margin-left: 2px;"></i>' : ''}
            </td>
            <td style="padding: 8px 12px; color: #111827; font-weight: 600;">${p.name}</td>
            ${valueCellFn(p)}
        </tr>
    `).join('');

    const pocRankingRowsHtml = buildRankingRows(stats.sortedP, p => `
        <td style="padding: 8px 12px; text-align: center;">
            <span style="background: rgba(0,122,255,0.1); color: #007AFF; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.75rem;">${p.count} POCs</span>
        </td>
        <td style="padding: 8px 12px; text-align: right; color: #34C759; font-weight: 700;">$${formatCurrency(p.sumValue)}</td>
    `);

    const revenueRankingRowsHtml = buildRankingRows(stats.sortedByRevenue || [], p => `
        <td style="padding: 8px 12px; text-align: right; color: #34C759; font-weight: 700;">$${formatCurrency(p.sumValue)}</td>
        <td style="padding: 8px 12px; text-align: center;">
            <span style="background: rgba(0,122,255,0.1); color: #007AFF; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.75rem;">${p.count} POCs</span>
        </td>
    `);

    const customerRankingRowsHtml = buildRankingRows(stats.sortedByCustomers || [], p => `
        <td style="padding: 8px 12px; text-align: center;">
            <span style="background: rgba(168,85,247,0.12); color: #a855f7; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.75rem;">${p.customerCount} Accounts</span>
        </td>
        <td style="padding: 8px 12px; text-align: center;">
            <span style="background: rgba(0,122,255,0.1); color: #007AFF; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.75rem;">${p.count} POCs</span>
        </td>
    `);

    const rankingCard = (title, iconBg, iconColor, icon, col1, col2, rowsHtml, isEmpty) => `
        <div class="stat-card highlight-card" style="padding: 16px; background: #FFFFFF; border: 1px solid #F3F4F6; display: flex; flex-direction: column; align-items: stretch;">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 12px;">
                <div class="stat-icon" style="background: ${iconBg}; color: ${iconColor}; width: 32px; height: 32px;"><i class="fa-solid ${icon}"></i></div>
                <div>
                    <h3 style="font-size: 0.95rem; font-weight: 700; color: #111827; margin: 0;">${title}</h3>
                    <span style="font-size: 0.68rem; color: #9CA3AF; font-weight: 600;">${yearLabel}</span>
                </div>
            </div>
            <div style="background: #FFFFFF; border-radius: 12px; overflow: hidden; border: 1px solid #F3F4F6;">
                <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 0.78rem;">
                    <thead>
                        <tr style="background: #F9FAFB; border-bottom: 1px solid #E5E7EB;">
                            <th style="padding: 8px 12px; color: #6B7280; font-weight: 600;">Rank</th>
                            <th style="padding: 8px 12px; color: #6B7280; font-weight: 600;">Partner</th>
                            <th style="padding: 8px 12px; color: #6B7280; font-weight: 600; text-align: ${col1.align || 'center'};">${col1.label}</th>
                            <th style="padding: 8px 12px; color: #6B7280; font-weight: 600; text-align: ${col2.align || 'center'};">${col2.label}</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rowsHtml}
                        ${isEmpty ? '<tr><td colspan="4" style="padding: 20px; text-align: center; color: #9CA3AF;">No data</td></tr>' : ''}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    const rankingsEmpty = !stats.sortedP || stats.sortedP.length === 0;

    return `
        <div style="grid-column: 1 / -1; display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 20px;">
            ${globalCardHtml}
            ${statsCardsHtml}
        </div>

        ${tabName === 'PARTNER' ? `
        <div class="stat-card" style="grid-column: 1 / -1; display: flex; flex-wrap: wrap; align-items: center; gap: 12px; padding: 10px 16px; background: #FFFFFF; border: 1px solid rgba(0, 122, 255, 0.2); border-left: 4px solid #007AFF; margin-bottom: 16px;">
            <label style="font-size:0.8rem; color:#007AFF; font-weight:700; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right: 8px;"></i>Country</label>
            <select id="partner-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #D1D5DB; padding:6px 12px; border-radius:8px; width: 200px; font-size: 0.85rem;">
                ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
            </select>
            <label style="font-size:0.8rem; color:#007AFF; font-weight:700; text-transform: uppercase; margin-left: 8px;"><i class="fa-solid fa-calendar" style="margin-right: 8px;"></i>Year</label>
            <select id="partner-filter-year" style="background:#F9FAFB; color:#111827; border:1px solid #D1D5DB; padding:6px 12px; border-radius:8px; width: 180px; font-size: 0.85rem;">
                <option value="all" ${yearFilter === 'all' ? 'selected' : ''}>All Years</option>
                <option value="current" ${yearFilter === 'current' ? 'selected' : ''}>Current Year (${curY})</option>
                <option value="previous" ${yearFilter === 'previous' ? 'selected' : ''}>Previous Year (${prevY})</option>
            </select>
            <div style="margin-left: auto; text-align: right;">
                <span style="font-size: 0.72rem; color: #111827; font-weight: 600;">${filterCountry || 'All Regions'} · ${yearLabel}</span>
            </div>
        </div>
        ` : ''}

        <div style="grid-column: 1 / -1; display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 12px; margin-bottom: 20px;">
            ${rankingCard(
                'Top 10 Partners — by POC',
                'rgba(0,122,255,0.1)', '#007AFF', 'fa-flask',
                { label: 'Running', align: 'center' },
                { label: 'Value (USD)', align: 'right' },
                pocRankingRowsHtml, rankingsEmpty
            )}
            ${rankingCard(
                'Top 10 Partners — by Revenue',
                'rgba(52,199,89,0.12)', '#34C759', 'fa-dollar-sign',
                { label: 'Value (USD)', align: 'right' },
                { label: 'Running', align: 'center' },
                revenueRankingRowsHtml, rankingsEmpty
            )}
            ${rankingCard(
                'Top 10 Partners — by Account (Customers)',
                'rgba(168,85,247,0.12)', '#a855f7', 'fa-users',
                { label: 'Accounts', align: 'center' },
                { label: 'Running', align: 'center' },
                customerRankingRowsHtml, rankingsEmpty
            )}
        </div>

    `;
}

export function getPartnerNetworkDetailsHTML(stats, filterCountry) {
    const groupedListsHtml = stats.sortedCountries.map(country => {
        const partners = stats.partnerGroups[country];
        const partnerItemsHtml = partners.map(p => {
            const name = p[stats.pNameKey] || 'N/A';
            return `
                <div style="padding: 8px 12px; background: rgba(255, 255, 255, 0.03); border-radius: 8px; border: 1px solid #F3F4F6;">
                    <div style="color: #111827; font-weight: 600; font-size: 0.85rem;">${name}</div>
                    <div style="color: #6B7280; font-size: 0.7rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${Object.values(p)[1] || ''}</div>
                </div>
            `;
        }).join('');

        return `
            <div style="background: #F9FAFB; border: 1px solid #F3F4F6; border-radius: 12px; padding: 12px; display: flex; flex-direction: column; gap: 8px;">
                <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E5E7EB; padding-bottom: 6px;">
                    <h3 style="color: #111827; font-size: 0.9rem; font-weight: 700; margin: 0;">${country}</h3>
                    <span style="background: rgba(0,122,255,0.1); color: #007AFF; font-size: 0.65rem; font-weight: 700; padding: 2px 6px; border-radius: 10px;">${partners.length}</span>
                </div>
                <div style="display: grid; grid-template-columns: 1fr; gap: 6px;">
                    ${partnerItemsHtml}
                </div>
            </div>
        `;
    }).join('');

    return `
        <div style="grid-column: 1 / -1; margin-bottom: 24px;">
            <div style="padding: 20px; background: #F9FAFB; border: 1px solid #E5E7EB; border-radius: 20px; display: flex; flex-direction: column; gap: 20px;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <div class="stat-icon" style="background: rgba(0,122,255,0.1); color: #007AFF; width: 40px; height: 40px; font-size: 1.1rem;"><i class="fa-solid fa-handshake"></i></div>
                    <div>
                        <h2 style="font-size: 1.25rem; font-weight: 700; color: #111827; margin: 0;">Partner Network Details</h2>
                    </div>
                </div>

                ${!filterCountry ? `
                <div class="stat-card" style="margin:0; padding: 16px; background: #FFFFFF; border: 1px solid #F3F4F6; border-radius: 16px; display: flex; flex-direction: column; align-items: stretch;">
                    <h3 style="color: #111827; font-size: 1rem; font-weight: 600; margin-bottom: 12px;">Distribution Ranking</h3>
                    <div style="position: relative; height: 300px;">
                        <canvas id="partner-country-chart"></canvas>
                    </div>
                </div>
                ` : ''}

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px;">
                    ${groupedListsHtml}
                </div>
            </div>
        </div>
    `;
}

export function getGenericCountryHTML(stats, filterCountry) {
    if (!stats) return '';
    const totalHtml = stats.sortedTotal.map(([c, count]) => `
        <div style="display:flex; justify-content:space-between; align-items:center; padding:6px 10px; background:rgba(0,0,0,0.05); border-radius:6px;">
            <span style="font-size:0.75rem; color:#4B5563;"><i class="fa-solid fa-earth-americas" style="margin-right:6px;"></i>${filterCountry ? 'Total Deals' : c}</span>
            <span style="font-weight:700; color:#111827;">${count}</span>
        </div>
    `).join('');

    let yearlyHtml = '';
    stats.sortedYears.forEach(y => {
        const items = Object.entries(stats.yearlyCounts[y]).sort((a, b) => b[1] - a[1]).map(([c, count]) => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:4px 0; border-bottom:1px solid #F9FAFB;">
                <span style="font-size:0.75rem; color:#6B7280;">${filterCountry ? 'Deals' : c}</span>
                <span style="font-size:0.75rem; font-weight:600; color:#374151;">${count}</span>
            </div>
        `).join('');
        yearlyHtml += `<div style="margin-top:12px; border-top:1px solid #F3F4F6; padding-top:8px;"><h4 style="font-size:0.75rem; font-weight:800; color:#6366f1; margin-bottom:4px; text-transform:uppercase;">${y} PERFORMANCE</h4>${items}</div>`;
    });

    return `
        <div class="stat-card" style="padding:20px; background:#FFF; border:1px solid #F3F4F6; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px; padding-bottom:12px; border-bottom:1px solid #F3F4F6;">
                <div class="stat-icon" style="background:rgba(99,102,241,0.1); color:#6366f1; width:36px; height:36px; font-size:1rem;"><i class="fa-solid fa-handshake"></i></div>
                <div class="stat-details"><h3 style="margin:0; font-size:0.8rem; color:#6B7280;">CLOSED DEALS</h3><h2 style="margin:0; font-size:0.95rem; font-weight:700; color:#111827;">Summary by Country/Year</h2></div>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:24px; align-items:start;">
                <div>
                    <h4 style="font-size:0.7rem; font-weight:800; color:#6B7280; text-transform:uppercase; letter-spacing:0.05em; margin:0 0 10px;">By Country</h4>
                    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:8px;">${totalHtml}</div>
                </div>
                <div style="border-left:1px solid #F3F4F6; padding-left:24px;">
                    <h4 style="font-size:0.7rem; font-weight:800; color:#6B7280; text-transform:uppercase; letter-spacing:0.05em; margin:0 0 4px;">Year over Year</h4>
                    <div style="max-height:260px; overflow-y:auto; padding-right:4px;">${yearlyHtml}</div>
                </div>
            </div>
        </div>
    `;
}

export function getExpiringContractsHTML(stats) {
    if (!stats) return '';
    const items = stats.slice(0, 5).map(d => `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; background: rgba(239, 68, 68, 0.08); border-radius: 10px; border-left: 4px solid #ef4444; margin-bottom: 8px; transition: transform 0.2s;" onmouseover="this.style.transform='translateX(4px)'" onmouseout="this.style.transform='translateX(0)'">
            <div style="display: flex; flex-direction: column; gap: 4px;">
                <span style="font-size: 0.9rem; font-weight: 700; color: #111827;">${d.name}</span>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="font-size: 0.75rem; color: #DC2626; font-weight: 500;"><i class="fa-regular fa-calendar-alt"></i> End: ${d.date}</span>
                    ${d.year ? `<span style="font-size: 0.7rem; background: rgba(239, 68, 68, 0.15); color: #B91C1C; padding: 2px 6px; border-radius: 4px; font-weight: 600;">${d.year} Yr</span>` : ''}
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 8px;">
                <span style="font-size: 0.65rem; font-weight: 800; color: #ef4444; text-transform: uppercase; background: rgba(239, 68, 68, 0.1); padding: 4px 8px; border-radius: 6px; letter-spacing: 0.05em;">Expiring Soon</span>
                <i class="fa-solid fa-chevron-right" style="color: #ef4444; font-size: 0.7rem; opacity: 0.5;"></i>
            </div>
        </div>
    `).join('');

    return `
        <div class="stat-card" style="display: flex; flex-direction: column; align-items: stretch; padding: 24px; border: 1px solid rgba(239, 68, 68, 0.2); background: #FFF; box-shadow: 0 4px 15px rgba(239, 68, 68, 0.05); border-radius: 12px;">
            <div style="display: flex; align-items: center; gap: 14px; margin-bottom: 20px; border-bottom: 2px solid #FEF2F2; padding-bottom: 12px;">
                <div class="stat-icon" style="background: rgba(239, 68, 68, 0.1); color: #ef4444; width: 42px; height: 42px; border-radius: 10px; display: flex; align-items: center; justify-content: center;"><i class="fa-solid fa-clock-rotate-left"></i></div>
                <div class="stat-details">
                    <h3 style="margin:0; font-size: 0.8rem; color: #ef4444; font-weight:800; text-transform:uppercase; letter-spacing: 0.05em;">EXPIRING SOON</h3>
                    <h2 style="font-size: 1.1rem; font-weight: 800; color: #111827; margin: 0;">Contracts Renewals (Within 3 Months)</h2>
                </div>
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 12px;">
                ${items}
            </div>
            ${stats.length > 5 ? `<div style="text-align: center; font-size: 0.75rem; color: #ef4444; margin-top: 12px; font-weight: 600; cursor: pointer; padding: 8px; border-radius: 8px; background: rgba(239, 68, 68, 0.05);"><i class="fa-solid fa-plus-circle"></i> View ${stats.length - 5} more expiring contracts</div>` : ''}
        </div>
    `;
}

/* ─── Churn Risk Alert ─── */
export function getChurnRiskHTML(stats) {
    if (!stats) return '';

    const { critical, warning, overdue, criticalArr, warningArr, overdueArr, totalArrAtRisk } = stats;

    function fmtArr(v) { return v >= 1000 ? `$${(v / 1000).toFixed(1)}K` : `$${Math.round(v)}`; }
    function daysLabel(d) {
        if (d < 0) return `D+${Math.abs(d)}`;
        if (d === 0) return 'D-Day';
        return `D-${d}`;
    }

    const tierConfig = {
        critical: {
            rowBg: 'rgba(239,68,68,0.06)', border: '#ef4444',
            chipBg: 'rgba(239,68,68,0.12)', dayColor: '#b91c1c',
            headerBg: 'linear-gradient(135deg, #fef2f2 0%, #fff5f5 100%)',
            headerColor: '#ef4444', countBg: '#fee2e2', countColor: '#b91c1c',
            label: 'Critical', sublabel: 'Within 30 Days', icon: 'fa-circle-exclamation'
        },
        warning: {
            rowBg: 'rgba(245,158,11,0.06)', border: '#f59e0b',
            chipBg: 'rgba(245,158,11,0.14)', dayColor: '#92400e',
            headerBg: 'linear-gradient(135deg, #fffbeb 0%, #fefce8 100%)',
            headerColor: '#d97706', countBg: '#fef3c7', countColor: '#92400e',
            label: 'Renew Soon', sublabel: '30–90 Days', icon: 'fa-clock'
        },
        overdue: {
            rowBg: 'rgba(107,114,128,0.05)', border: '#9ca3af',
            chipBg: 'rgba(107,114,128,0.12)', dayColor: '#374151',
            headerBg: 'linear-gradient(135deg, #f9fafb 0%, #f3f4f6 100%)',
            headerColor: '#6b7280', countBg: '#e5e7eb', countColor: '#374151',
            label: 'Overdue', sublabel: 'Revenue Leak', icon: 'fa-triangle-exclamation'
        }
    };

    function buildRow(d, tier) {
        const cfg = tierConfig[tier];
        return `
            <div style="display:grid; grid-template-columns:1fr auto; gap:10px; align-items:center; padding:9px 12px; background:${cfg.rowBg}; border-radius:8px; border-left:3px solid ${cfg.border}; transition:transform 0.15s;" onmouseover="this.style.transform='translateX(3px)'" onmouseout="this.style.transform='translateX(0)'">
                <div style="min-width:0;">
                    <div style="font-size:0.8rem; font-weight:700; color:#111827; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; margin-bottom:3px;" title="${d.name}">${d.name}</div>
                    <div style="display:flex; align-items:center; gap:8px; font-size:0.66rem; color:#6b7280; white-space:nowrap;">
                        ${d.country ? `<span><i class="fa-solid fa-location-dot"></i> ${d.country}</span>` : ''}
                        <span style="font-weight:800; color:${cfg.dayColor}; background:${cfg.chipBg}; padding:1px 7px; border-radius:4px; letter-spacing:0.02em;">${daysLabel(d.daysLeft)}</span>
                        <span style="overflow:hidden; text-overflow:ellipsis;"><i class="fa-regular fa-calendar"></i> ${d.date}</span>
                    </div>
                </div>
                ${d.arr > 0 ? `<div style="text-align:right; white-space:nowrap;"><span style="font-size:0.82rem; font-weight:800; color:#111827;">${fmtArr(d.arr)}</span> <span style="font-size:0.6rem; font-weight:700; color:#9ca3af; letter-spacing:0.05em;">ARR</span></div>` : '<div></div>'}
            </div>
        `;
    }

    function buildSection(tier, list, arrSum) {
        if (!list || list.length === 0) return '';
        const cfg = tierConfig[tier];
        const SHOW = 6;
        const visible = list.slice(0, SHOW);
        const hidden = list.slice(SHOW);
        const sectionId = `churn-${tier}-${Math.random().toString(36).slice(2, 8)}`;

        const expandable = hidden.length > 0 ? `
                <div id="${sectionId}-hidden" style="display:none; flex-direction:column; gap:5px;">
                    ${hidden.map(d => buildRow(d, tier)).join('')}
                </div>
                <button type="button"
                    onclick="(function(btn){var box=document.getElementById('${sectionId}-hidden');var open=box.style.display==='flex';box.style.display=open?'none':'flex';btn.querySelector('span').textContent=open?('+ '+${hidden.length}+' more'):'Show less';btn.querySelector('i').className=open?'fa-solid fa-plus':'fa-solid fa-chevron-up';})(this)"
                    style="cursor:pointer; background:transparent; border:1px dashed ${cfg.border}66; color:${cfg.headerColor}; font-size:0.7rem; padding:7px 12px; font-weight:700; border-radius:8px; margin-top:4px; transition:background 0.15s; display:flex; align-items:center; justify-content:center; gap:6px;"
                    onmouseover="this.style.background='${cfg.rowBg}'"
                    onmouseout="this.style.background='transparent'">
                    <i class="fa-solid fa-plus"></i><span>+ ${hidden.length} more</span>
                </button>
            ` : '';

        return `
            <div style="display:flex; flex-direction:column; gap:10px; min-width:0;">
                <div style="display:flex; align-items:center; justify-content:space-between; gap:10px; padding:11px 14px; background:${cfg.headerBg}; border-radius:10px; border:1px solid ${cfg.border}33;">
                    <div style="display:flex; align-items:center; gap:10px; min-width:0;">
                        <div style="width:30px; height:30px; border-radius:8px; background:#fff; display:flex; align-items:center; justify-content:center; flex-shrink:0; box-shadow:0 1px 3px rgba(0,0,0,0.06);">
                            <i class="fa-solid ${cfg.icon}" style="color:${cfg.headerColor}; font-size:0.85rem;"></i>
                        </div>
                        <div style="min-width:0;">
                            <div style="font-size:0.72rem; font-weight:800; color:${cfg.headerColor}; text-transform:uppercase; letter-spacing:0.06em; line-height:1.1;">${cfg.label}</div>
                            <div style="font-size:0.62rem; color:#6b7280; line-height:1.2;">${cfg.sublabel}</div>
                        </div>
                    </div>
                    <div style="display:flex; align-items:center; gap:8px; flex-shrink:0;">
                        <span style="background:${cfg.countBg}; color:${cfg.countColor}; font-size:0.68rem; font-weight:800; padding:3px 9px; border-radius:12px;">${list.length}</span>
                        ${arrSum > 0 ? `<span style="font-size:0.78rem; font-weight:800; color:#111827;">${fmtArr(arrSum)}</span>` : ''}
                    </div>
                </div>
                <div style="display:flex; flex-direction:column; gap:5px;">
                    ${visible.map(d => buildRow(d, tier)).join('')}
                </div>
                ${expandable}
            </div>
        `;
    }

    const sectionsList = [
        { tier: 'critical', list: critical, arr: criticalArr },
        { tier: 'warning',  list: warning,  arr: warningArr  },
        { tier: 'overdue',  list: overdue,  arr: overdueArr  },
    ].filter(s => s.list && s.list.length > 0);

    const sectionsHTML = sectionsList.map(s => buildSection(s.tier, s.list, s.arr)).join('');
    const gridCols = `repeat(auto-fit, minmax(min(100%, 320px), 1fr))`;

    const pillCritical = critical.length > 0 ? `<span style="background:rgba(239,68,68,0.1); color:#b91c1c; font-size:0.68rem; font-weight:800; padding:4px 10px; border-radius:20px; white-space:nowrap;"><i class="fa-solid fa-circle-exclamation"></i> ${critical.length} Critical</span>` : '';
    const pillWarning  = warning.length  > 0 ? `<span style="background:rgba(245,158,11,0.1); color:#92400e; font-size:0.68rem; font-weight:800; padding:4px 10px; border-radius:20px; white-space:nowrap;"><i class="fa-solid fa-clock"></i> ${warning.length} Renew Soon</span>` : '';
    const pillOverdue  = overdue.length  > 0 ? `<span style="background:rgba(107,114,128,0.1); color:#4b5563; font-size:0.68rem; font-weight:800; padding:4px 10px; border-radius:20px; white-space:nowrap;"><i class="fa-solid fa-triangle-exclamation"></i> ${overdue.length} Overdue</span>` : '';

    return `
        <div class="stat-card" style="display:block; padding:20px 22px; border:1px solid rgba(239,68,68,0.18); background:#fff; border-radius:14px; box-shadow:0 4px 16px rgba(239,68,68,0.06);">
            <div style="display:flex; align-items:center; gap:14px; margin-bottom:16px; padding-bottom:14px; border-bottom:1px solid #fef2f2; flex-wrap:wrap; row-gap:10px;">
                <div style="background:rgba(239,68,68,0.1); color:#ef4444; width:42px; height:42px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;"><i class="fa-solid fa-shield-halved" style="font-size:1.05rem;"></i></div>
                <div style="flex-shrink:0;">
                    <h3 style="margin:0; font-size:0.7rem; color:#ef4444; font-weight:800; text-transform:uppercase; letter-spacing:0.08em; line-height:1.2;">CHURN RISK ALERT</h3>
                    <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827; line-height:1.3;">Contract Renewal Monitor</h2>
                </div>
                <div style="margin-left:auto; display:flex; align-items:center; gap:8px; flex-wrap:wrap; justify-content:flex-end;">
                    ${pillCritical}${pillWarning}${pillOverdue}
                    ${totalArrAtRisk > 0 ? `<span style="background:#fef3c7; color:#92400e; font-size:0.7rem; font-weight:800; padding:4px 12px; border-radius:20px; border:1px solid #fde68a; white-space:nowrap;"><i class="fa-solid fa-dollar-sign"></i> ${fmtArr(totalArrAtRisk)} ARR at Risk</span>` : ''}
                </div>
            </div>
            <div style="display:grid; grid-template-columns:${gridCols}; gap:18px; align-items:flex-start;">
                ${sectionsHTML}
            </div>
        </div>
    `;
}

/* ─── Partner ROI ─── */
export function getPartnerROIHTML(stats) {
    if (!stats || !stats.partners || stats.partners.length === 0) return '';
    const { partners, avgWinRate } = stats;

    function fmtK(v) { return v >= 1000 ? `$${(v / 1000).toFixed(1)}K` : v > 0 ? `$${v}` : '–'; }

    const efficiencyBadge = (e, wr) => {
        if (e === 'efficient') return `<span style="background:#dcfce7; color:#15803d; font-size:0.6rem; font-weight:800; padding:2px 8px; border-radius:12px; letter-spacing:0.05em;">EFFICIENT</span>`;
        if (e === 'low-win')   return `<span style="background:#fee2e2; color:#b91c1c; font-size:0.6rem; font-weight:800; padding:2px 8px; border-radius:12px; letter-spacing:0.05em;">LOW WIN RATE</span>`;
        return '';
    };

    const rows = partners.map((p, i) => {
        const wrColor = p.winRate === null ? '#9ca3af' : p.winRate >= avgWinRate ? '#15803d' : p.winRate >= avgWinRate - 15 ? '#d97706' : '#dc2626';
        const barPct = p.total > 0 ? Math.round(p.won / p.total * 100) : 0;
        return `
        <tr style="border-bottom:1px solid #f3f4f6; background:${i % 2 === 0 ? '#fff' : '#fafafa'}; transition:background 0.15s;" onmouseover="this.style.background='#f0f9ff'" onmouseout="this.style.background='${i % 2 === 0 ? '#fff' : '#fafafa'}'">
            <td style="padding:10px 12px; font-weight:700; color:#111827; max-width:180px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                ${p.name}
                <div style="margin-top:3px;">${efficiencyBadge(p.efficiency)}</div>
            </td>
            <td style="padding:10px 12px; text-align:center; font-weight:700; color:#374151;">${p.total}</td>
            <td style="padding:10px 12px; text-align:center;">
                <div style="display:flex; align-items:center; gap:6px;">
                    <div style="flex:1; height:6px; background:#f3f4f6; border-radius:3px; overflow:hidden;">
                        <div style="height:100%; width:${barPct}%; background:#10b981; border-radius:3px;"></div>
                    </div>
                    <span style="font-weight:700; color:#10b981; font-size:0.78rem; min-width:20px;">${p.won}</span>
                </div>
            </td>
            <td style="padding:10px 12px; text-align:center; font-weight:600; color:#ef4444;">${p.drop}</td>
            <td style="padding:10px 12px; text-align:center; font-weight:600; color:#f59e0b;">${p.running}</td>
            <td style="padding:10px 12px; text-align:center;">
                ${p.winRate !== null
                    ? `<span style="font-weight:800; color:${wrColor};">${p.winRate}%</span>`
                    : '<span style="color:#9ca3af; font-size:0.75rem;">–</span>'}
            </td>
            <td style="padding:10px 12px; text-align:right; font-weight:700; color:#111827;">${fmtK(p.wonValue)}</td>
            <td style="padding:10px 12px; text-align:right; color:#6b7280; font-size:0.78rem;">${fmtK(p.valuePerPoc)}</td>
        </tr>`;
    }).join('');

    const effCount = partners.filter(p => p.efficiency === 'efficient').length;
    const lowCount = partners.filter(p => p.efficiency === 'low-win').length;

    return `
        <div class="stat-card" style="padding:22px; background:#fff; border:1px solid #f3f4f6; border-radius:14px; box-shadow:0 2px 8px rgba(0,0,0,0.04); display:block;">
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px; border-bottom:1px solid #f3f4f6; padding-bottom:14px; flex-wrap:wrap; row-gap:8px;">
                <div style="background:rgba(99,102,241,0.1); color:#6366f1; width:40px; height:40px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;"><i class="fa-solid fa-chart-bar" style="font-size:1rem;"></i></div>
                <div>
                    <h3 style="margin:0; font-size:0.72rem; color:#6366f1; font-weight:800; text-transform:uppercase; letter-spacing:0.08em;">PARTNER ROI</h3>
                    <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827;">POC Efficiency by Partner</h2>
                </div>
                <div style="margin-left:auto; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span style="background:#f0fdf4; color:#15803d; font-size:0.68rem; font-weight:800; padding:3px 10px; border-radius:20px;">${effCount} Efficient</span>
                    ${lowCount > 0 ? `<span style="background:#fef2f2; color:#b91c1c; font-size:0.68rem; font-weight:800; padding:3px 10px; border-radius:20px;">${lowCount} Low Win Rate</span>` : ''}
                    <span style="background:#f3f4f6; color:#374151; font-size:0.68rem; font-weight:700; padding:3px 10px; border-radius:20px;">Avg Win Rate: ${avgWinRate}%</span>
                </div>
            </div>
            <div>
                <table style="width:100%; border-collapse:collapse;">
                    <thead>
                        <tr style="background:#f9fafb; text-align:left;">
                            <th style="padding:8px 12px; font-size:0.68rem; color:#6b7280; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">Partner</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#6b7280; font-weight:700; text-transform:uppercase; text-align:center;">Total POCs</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#10b981; font-weight:700; text-transform:uppercase;">Won</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#ef4444; font-weight:700; text-transform:uppercase; text-align:center;">Drop</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#f59e0b; font-weight:700; text-transform:uppercase; text-align:center;">Running</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#6b7280; font-weight:700; text-transform:uppercase; text-align:center;">Win Rate</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#6b7280; font-weight:700; text-transform:uppercase; text-align:right;">Won Value</th>
                            <th style="padding:8px 12px; font-size:0.68rem; color:#6b7280; font-weight:700; text-transform:uppercase; text-align:right;">Value / POC</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
            <div style="margin-top:12px; padding:8px 12px; background:#f9fafb; border-radius:8px; font-size:0.7rem; color:#6b7280;">
                <i class="fa-solid fa-circle-info" style="color:#6366f1;"></i>
                <strong>Efficient</strong>: Win rate ≥${avgWinRate + 10}% &nbsp;·&nbsp;
                <strong>Low Win Rate</strong>: Win rate ≤${Math.max(0, avgWinRate - 15)}% with ≥3 POCs &nbsp;·&nbsp;
                <strong>Value/POC</strong>: Won value ÷ total POC attempts
            </div>
        </div>
    `;
}

/* ─── Pipeline Coverage Ratio ─── */
export function getPipelineCoverageHTML(stats) {
    if (!stats) return '';
    const { quarters, currentQ, totalWeighted, totalBooked, totalTarget, annualCoverage } = stats;

    function fmtM(v) {
        if (v >= 1_000_000) return `$${(v / 1_000_000).toFixed(2)}M`;
        if (v >= 1000)      return `$${(v / 1000).toFixed(1)}K`;
        return v > 0 ? `$${Math.round(v)}` : '–';
    }

    function coverageColor(pct) {
        if (pct === null) return '#9ca3af';
        if (pct >= 150) return '#059669';
        if (pct >= 100) return '#10b981';
        if (pct >= 70)  return '#f59e0b';
        return '#ef4444';
    }
    function coverageBg(pct) {
        if (pct === null) return '#f3f4f6';
        if (pct >= 150) return '#d1fae5';
        if (pct >= 100) return '#ecfdf5';
        if (pct >= 70)  return '#fef3c7';
        return '#fee2e2';
    }

    const quarterCards = quarters.map(q => {
        const color = coverageColor(q.coverage);
        const bg    = coverageBg(q.coverage);
        const label = q.isCurrent ? `${q.q} ← Now` : q.q;
        const pctDisplay = q.coverage !== null ? `${q.coverage}%` : 'N/A';
        const barW = q.coverage !== null ? Math.min(100, q.coverage) : 0;
        const ringColor = q.isCurrent ? '#6366f1' : '#e5e7eb';

        return `
        <div style="flex:1; min-width:130px; background:${q.isCurrent ? '#f5f3ff' : '#fafafa'}; border:${q.isCurrent ? '2px solid #6366f1' : '1px solid #e5e7eb'}; border-radius:12px; padding:14px 16px; position:relative;">
            <div style="font-size:0.72rem; font-weight:800; color:${q.isCurrent ? '#6366f1' : '#9ca3af'}; text-transform:uppercase; letter-spacing:0.06em; margin-bottom:6px;">${label}</div>
            <div style="font-size:2rem; font-weight:900; color:${color}; line-height:1; margin-bottom:4px;">${pctDisplay}</div>
            <div style="height:5px; background:#e5e7eb; border-radius:3px; margin-bottom:8px; overflow:hidden;">
                <div style="height:100%; width:${barW}%; background:${color}; border-radius:3px; transition:width 0.6s;"></div>
            </div>
            <div style="font-size:0.65rem; color:#6b7280; line-height:1.5;">
                ${q.target > 0 ? `<div>Target (LY): <strong>${fmtM(q.target)}</strong></div>` : '<div style="color:#9ca3af;">No LY baseline</div>'}
                ${q.booked > 0 ? `<div>Booked: <strong style="color:#10b981;">${fmtM(q.booked)}</strong></div>` : ''}
                ${!q.isPast && q.weighted > 0 ? `<div>Pipeline: <strong style="color:#6366f1;">${fmtM(q.weighted)}</strong></div>` : ''}
                ${!q.isPast && q.count > 0 ? `<div style="color:#9ca3af;">${q.count} deals in pipeline</div>` : ''}
            </div>
        </div>`;
    }).join('');

    const annualColor = coverageColor(annualCoverage);

    return `
        <div class="stat-card" style="padding:22px; background:#fff; border:1px solid #ede9fe; border-radius:14px; box-shadow:0 2px 10px rgba(99,102,241,0.07);">
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:18px; border-bottom:1px solid #f5f3ff; padding-bottom:14px; flex-wrap:wrap; row-gap:8px;">
                <div style="background:rgba(99,102,241,0.1); color:#6366f1; width:40px; height:40px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;"><i class="fa-solid fa-bullseye" style="font-size:1rem;"></i></div>
                <div>
                    <h3 style="margin:0; font-size:0.72rem; color:#6366f1; font-weight:800; text-transform:uppercase; letter-spacing:0.08em;">PIPELINE COVERAGE RATIO</h3>
                    <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827;">Quarterly Target vs. Pipeline</h2>
                </div>
                <div style="margin-left:auto; display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
                    ${annualCoverage !== null ? `
                    <div style="text-align:right;">
                        <div style="font-size:0.65rem; color:#6b7280; font-weight:600; text-transform:uppercase; letter-spacing:0.05em;">Annual Coverage</div>
                        <div style="font-size:1.4rem; font-weight:900; color:${annualColor}; line-height:1;">${annualCoverage}%</div>
                    </div>` : ''}
                </div>
            </div>
            <div style="display:flex; gap:12px; flex-wrap:wrap; margin-bottom:16px;">
                ${quarterCards}
            </div>
            <div style="padding:10px 14px; background:#f5f3ff; border-radius:8px; font-size:0.7rem; color:#6b7280; display:flex; flex-wrap:wrap; gap:16px;">
                <span><i class="fa-solid fa-circle" style="color:#10b981; font-size:0.5rem;"></i> <strong>Booked TCV</strong> = closed deals this quarter</span>
                <span><i class="fa-solid fa-circle" style="color:#6366f1; font-size:0.5rem;"></i> <strong>Pipeline</strong> = weighted pipeline in PIPELINE sheet</span>
                <span><i class="fa-solid fa-circle" style="color:#9ca3af; font-size:0.5rem;"></i> <strong>Target</strong> = same quarter last year (YoY baseline)</span>
                <span><strong style="color:#059669;">≥150%</strong> Strong &nbsp;·&nbsp; <strong style="color:#10b981;">≥100%</strong> On track &nbsp;·&nbsp; <strong style="color:#f59e0b;">≥70%</strong> Watch &nbsp;·&nbsp; <strong style="color:#ef4444;">&lt;70%</strong> At Risk</span>
            </div>
        </div>
    `;
}

/**
 * Render the Pipeline Change Log table for a selected country.
 * Shows snapshots over time with delta arrows highlighting suspicious drops.
 * @param {string} filterCountry
 * @param {Array} history - newest-last array of snapshots
 * @param {Array} dealDiffs - same length as oldest-first sorted history; each entry
 *   is null (for the initial) or `{ added, removed, modified }`
 * @returns {string} HTML
 */
export function getPipelineChangeLogHTML(filterCountry, history, dealDiffs = []) {
    const headerLeft = `
        <div style="display:flex; align-items:center; gap:10px;">
            <div class="stat-icon" style="width:36px; height:36px; font-size:1rem; background:rgba(245,158,11,0.15); color:#f59e0b; border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-clock-rotate-left"></i></div>
            <div>
                <h3 style="margin:0; font-size:0.72rem; color:#b45309; font-weight:800; text-transform:uppercase; letter-spacing:0.08em;">PIPELINE CHANGE LOG</h3>
                <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827;">${filterCountry} — historical snapshots</h2>
            </div>
        </div>`;

    if (!history || history.length === 0) {
        return `
        <div style="padding:22px; background:#FFFFFF; border:1px solid #fde68a; border-radius:14px; box-shadow:0 2px 10px rgba(245,158,11,0.05);">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:14px; border-bottom:1px solid #fef3c7; padding-bottom:12px; flex-wrap:wrap; gap:10px;">
                ${headerLeft}
            </div>
            <p style="color:#6B7280; font-size:0.8rem; margin:0;">No snapshots yet. The first snapshot will be recorded automatically the next time pipeline values for ${filterCountry} change.</p>
        </div>`;
    }

    // Sort oldest-first to compute pairwise changes, then reverse for display.
    const oldestFirst = [...history].sort((a, b) => new Date(a.date) - new Date(b.date));
    const initial = oldestFirst[0];
    const latest = oldestFirst[oldestFirst.length - 1];

    const fmtDateLong = (d) => new Date(d).toLocaleString('en-US', { year: 'numeric', month: 'short', day: '2-digit', hour: '2-digit', minute: '2-digit' });

    // Compact baseline strip — sits at the very top of the body so it's the
    // anchor reference everything else is compared against.
    const initialQuarterPills = ['Q1', 'Q2', 'Q3', 'Q4'].map(q => {
        const v = initial.byQuarter?.[q] || { amount: 0, count: 0 };
        return `<span style="background:#fff; border:1px solid #ddd6fe; padding:3px 8px; border-radius:6px; font-size:0.65rem; color:#4338ca; font-weight:600;">${q} <span style="color:#111827; font-weight:700;">$${formatCurrency(v.amount || 0)}</span> <span style="color:#9CA3AF;">·${v.count || 0}</span></span>`;
    }).join('');
    const baselineStrip = `
        <div style="display:flex; align-items:center; gap:14px; padding:10px 14px; background:#f5f3ff; border:1px solid #ddd6fe; border-radius:10px; margin-bottom:14px; flex-wrap:wrap;">
            <span style="background:#4338ca; color:#fff; font-size:0.6rem; font-weight:800; padding:3px 8px; border-radius:6px; letter-spacing:0.05em; white-space:nowrap;"><i class="fa-solid fa-flag" style="margin-right:4px;"></i>INITIAL BASELINE</span>
            <span style="font-size:0.7rem; color:#6B7280; white-space:nowrap;">${fmtDateLong(initial.date)}</span>
            <span style="font-size:0.72rem; color:#374151; white-space:nowrap;"><strong style="color:#111827;">${initial.count || 0}</strong> deals</span>
            <span style="font-size:0.72rem; color:#374151; white-space:nowrap;">Pipeline <strong style="color:#10b981;">$${formatCurrency(initial.amount || 0)}</strong></span>
            <span style="font-size:0.72rem; color:#374151; white-space:nowrap;">Weighted <strong style="color:#007AFF;">$${formatCurrency(initial.weighted || 0)}</strong></span>
            <span style="font-size:0.72rem; color:#374151; white-space:nowrap;">TCV <strong style="color:#ef4444;">$${formatCurrency(initial.tcv || 0)}</strong></span>
            <span style="display:flex; gap:5px; flex-wrap:wrap; margin-left:auto;">${initialQuarterPills}</span>
        </div>`;

    // Build per-change events: what shifted between consecutive snapshots.
    const changeEvents = [];
    for (let i = 1; i < oldestFirst.length; i++) {
        const before = oldestFirst[i - 1];
        const after = oldestFirst[i];
        const quartersChanged = ['Q1', 'Q2', 'Q3', 'Q4'].map(q => {
            const b = before.byQuarter?.[q] || { amount: 0, weighted: 0, count: 0 };
            const a = after.byQuarter?.[q] || { amount: 0, weighted: 0, count: 0 };
            return {
                q,
                beforeAmt: b.amount || 0,
                afterAmt: a.amount || 0,
                delta: (a.amount || 0) - (b.amount || 0),
                beforeCount: b.count || 0,
                afterCount: a.count || 0
            };
        }).filter(x => Math.abs(x.delta) >= 1 || x.beforeCount !== x.afterCount);
        const dealDiff = dealDiffs[i] || { added: [], removed: [], modified: [] };
        changeEvents.push({ before, after, quartersChanged, dealDiff });
    }
    changeEvents.reverse(); // newest first

    // Helper to render a single secondary metric chip (weighted / TCV / deals)
    // when those also moved alongside pipeline.
    const secondaryChip = (label, color, before, after, isCurrency = true) => {
        const d = (after || 0) - (before || 0);
        if (Math.abs(d) < 1) return '';
        const arrow = d > 0 ? '▲' : '▼';
        const sign = d > 0 ? '+' : '−';
        const c = d > 0 ? '#10b981' : '#ef4444';
        const cur = isCurrency ? `$${formatCurrency(after)}` : String(after);
        const dStr = isCurrency ? `$${formatCurrency(Math.abs(d))}` : String(Math.abs(d));
        return `<span style="display:inline-flex; align-items:center; gap:5px; background:#F9FAFB; border:1px solid #E5E7EB; padding:3px 8px; border-radius:6px; font-size:0.66rem;">
            <span style="color:${color}; font-weight:700; text-transform:uppercase; letter-spacing:0.03em;">${label}</span>
            <span style="color:#111827; font-weight:700;">${cur}</span>
            <span style="color:${c}; font-weight:800;">${arrow} ${sign}${dStr}</span>
        </span>`;
    };

    let changeListHTML;
    if (changeEvents.length === 0) {
        changeListHTML = `<div style="padding:18px; text-align:center; font-size:0.78rem; color:#9CA3AF; font-style:italic; background:#FAFAFA; border:1px dashed #E5E7EB; border-radius:10px;">No changes recorded yet — the baseline above is the only snapshot so far.</div>`;
    } else {
        changeListHTML = changeEvents.map(ev => {
            const dAmt = (ev.after.amount || 0) - (ev.before.amount || 0);
            const arrow = dAmt > 0 ? '▲' : (dAmt < 0 ? '▼' : '–');
            const color = dAmt > 0 ? '#10b981' : (dAmt < 0 ? '#ef4444' : '#6B7280');
            const sign = dAmt > 0 ? '+' : (dAmt < 0 ? '−' : '');
            const deltaPct = (ev.before.amount || 0) > 0
                ? ((dAmt / ev.before.amount) * 100).toFixed(1)
                : null;
            const accentBg = dAmt < 0 ? '#FEF2F2' : (dAmt > 0 ? '#F0FDF4' : '#F9FAFB');
            const accentBorder = dAmt < 0 ? '#fecaca' : (dAmt > 0 ? '#bbf7d0' : '#E5E7EB');

            const secondaryChips = [
                secondaryChip('Deals', '#6B7280', ev.before.count, ev.after.count, false),
                secondaryChip('Weighted', '#007AFF', ev.before.weighted, ev.after.weighted, true),
                secondaryChip('TCV', '#ef4444', ev.before.tcv, ev.after.tcv, true)
            ].filter(Boolean).join(' ');

            const quarterChips = ev.quartersChanged.length === 0
                ? `<span style="color:#9CA3AF; font-size:0.72rem; font-style:italic;">Totals shifted but no single quarter moved by $1+ — likely a rounding-level edit.</span>`
                : ev.quartersChanged.map(qc => {
                    const qArrow = qc.delta > 0 ? '▲' : (qc.delta < 0 ? '▼' : '–');
                    const qColor = qc.delta > 0 ? '#10b981' : (qc.delta < 0 ? '#ef4444' : '#6B7280');
                    const qSign = qc.delta > 0 ? '+' : (qc.delta < 0 ? '−' : '');
                    const qBg = qc.delta < 0 ? '#FEF2F2' : (qc.delta > 0 ? '#F0FDF4' : '#FFFBEB');
                    const qBorder = qc.delta < 0 ? '#fecaca' : (qc.delta > 0 ? '#bbf7d0' : '#fde68a');
                    const countNote = qc.beforeCount !== qc.afterCount
                        ? ` <span style="color:#6B7280; font-size:0.62rem;">${qc.beforeCount}→${qc.afterCount} deals</span>`
                        : '';
                    return `<span style="display:inline-flex; align-items:center; gap:6px; background:${qBg}; border:1px solid ${qBorder}; padding:6px 10px; border-radius:8px; font-size:0.72rem; line-height:1.3;">
                        <strong style="color:#111827; font-size:0.78rem; background:#fff; border:1px solid ${qBorder}; padding:1px 6px; border-radius:5px;">${qc.q}</strong>
                        <span style="color:#9CA3AF;">$${formatCurrency(qc.beforeAmt)} →</span>
                        <span style="color:#111827; font-weight:700;">$${formatCurrency(qc.afterAmt)}</span>
                        <span style="color:${qColor}; font-weight:800;">${qArrow} ${qSign}$${formatCurrency(Math.abs(qc.delta))}</span>${countNote}
                    </span>`;
                }).join(' ');

            // Deal-level diff chips. Surfaces deal swaps that totals/quarters can hide.
            const diff = ev.dealDiff || { added: [], removed: [], modified: [] };
            const renderDealChip = (mode, deal, beforeDeal) => {
                const bg = mode === 'added' ? '#F0FDF4' : (mode === 'removed' ? '#FEF2F2' : '#FFFBEB');
                const border = mode === 'added' ? '#bbf7d0' : (mode === 'removed' ? '#fecaca' : '#fde68a');
                const tagBg = mode === 'added' ? '#10b981' : (mode === 'removed' ? '#ef4444' : '#f59e0b');
                const tagText = mode === 'added' ? 'NEW' : (mode === 'removed' ? 'GONE' : 'CHG');
                const label = deal.customer ? `${deal.customer} — ${deal.name}` : deal.name;
                let body;
                if (mode === 'modified') {
                    const parts = [];
                    if ((beforeDeal.amount || 0) !== (deal.amount || 0)) {
                        const d = (deal.amount || 0) - (beforeDeal.amount || 0);
                        const ar = d > 0 ? '▲' : '▼'; const cl = d > 0 ? '#10b981' : '#ef4444';
                        parts.push(`<span style="color:#9CA3AF;">$${formatCurrency(beforeDeal.amount || 0)} →</span> <span style="font-weight:700;">$${formatCurrency(deal.amount || 0)}</span> <span style="color:${cl}; font-weight:800;">${ar} $${formatCurrency(Math.abs(d))}</span>`);
                    }
                    if ((beforeDeal.quarter || '') !== (deal.quarter || '')) {
                        parts.push(`<span style="color:#7c3aed; font-weight:700;">${beforeDeal.quarter || '–'} → ${deal.quarter || '–'}</span>`);
                    }
                    body = parts.join(' · ');
                } else {
                    const qBadge = deal.quarter ? `<span style="background:#fff; border:1px solid ${border}; color:#374151; font-weight:700; font-size:0.62rem; padding:1px 5px; border-radius:4px;">${deal.quarter}</span>` : '';
                    body = `${qBadge} <span style="font-weight:700;">$${formatCurrency(deal.amount || 0)}</span>`;
                }
                return `<span style="display:inline-flex; align-items:center; gap:6px; background:${bg}; border:1px solid ${border}; padding:5px 9px; border-radius:8px; font-size:0.7rem; line-height:1.3; max-width:100%;">
                    <span style="background:${tagBg}; color:#fff; font-size:0.55rem; font-weight:800; padding:1px 5px; border-radius:4px; letter-spacing:0.04em;">${tagText}</span>
                    <span style="color:#111827; font-weight:600; max-width:220px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${label.replace(/"/g, '&quot;')}">${label}</span>
                    ${body}
                </span>`;
            };

            const totalDealChanges = diff.added.length + diff.removed.length + diff.modified.length;
            const dealSection = totalDealChanges === 0 ? '' : `
                <div>
                    <div style="font-size:0.62rem; color:#92400e; font-weight:800; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:6px;"><i class="fa-solid fa-list-check" style="margin-right:4px;"></i>Deal-level changes (${totalDealChanges} — ${diff.removed.length} removed · ${diff.added.length} added · ${diff.modified.length} modified)</div>
                    <div style="display:flex; gap:6px; flex-wrap:wrap;">
                        ${diff.removed.map(d => renderDealChip('removed', d)).join('')}
                        ${diff.added.map(d => renderDealChip('added', d)).join('')}
                        ${diff.modified.map(m => renderDealChip('modified', m.after, m.before)).join('')}
                    </div>
                </div>`;

            return `
            <div style="display:grid; grid-template-columns: minmax(220px, 260px) 1fr; gap:0; border:1px solid ${accentBorder}; border-radius:12px; overflow:hidden; background:#fff;">
                <div style="padding:14px 16px; background:${accentBg}; border-right:1px solid ${accentBorder}; display:flex; flex-direction:column; gap:4px;">
                    <div style="font-size:0.68rem; color:#6B7280; font-weight:600;">${fmtDateLong(ev.after.date)}</div>
                    <div style="font-size:0.62rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em; margin-top:6px;">Pipeline total</div>
                    <div style="font-size:0.72rem; color:#9CA3AF;">$${formatCurrency(ev.before.amount || 0)} →</div>
                    <div style="font-size:1.15rem; color:#111827; font-weight:800; line-height:1.1;">$${formatCurrency(ev.after.amount || 0)}</div>
                    <div style="font-size:0.95rem; color:${color}; font-weight:800; margin-top:2px;">${arrow} ${sign}$${formatCurrency(Math.abs(dAmt))}${deltaPct !== null ? ` <span style="font-size:0.7rem;">(${sign}${Math.abs(deltaPct)}%)</span>` : ''}</div>
                </div>
                <div style="padding:14px 16px; display:flex; flex-direction:column; gap:10px;">
                    <div>
                        <div style="font-size:0.62rem; color:#92400e; font-weight:800; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:6px;"><i class="fa-solid fa-bullseye" style="margin-right:4px;"></i>Affected quarters (${ev.quartersChanged.length})</div>
                        <div style="display:flex; gap:6px; flex-wrap:wrap;">${quarterChips}</div>
                    </div>
                    ${dealSection}
                    ${secondaryChips ? `
                    <div>
                        <div style="font-size:0.62rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:6px;">Other metrics that moved</div>
                        <div style="display:flex; gap:6px; flex-wrap:wrap;">${secondaryChips}</div>
                    </div>` : ''}
                </div>
            </div>`;
        }).join('');
    }

    // Cumulative drift cards (initial vs latest).
    let quarterlyDrift = '';
    if (initial && latest && initial !== latest && initial.byQuarter && latest.byQuarter) {
        const cards = ['Q1', 'Q2', 'Q3', 'Q4'].map(q => {
            const o = initial.byQuarter[q] || { amount: 0 };
            const n = latest.byQuarter[q] || { amount: 0 };
            const dAmt = (n.amount || 0) - (o.amount || 0);
            const color = dAmt > 0 ? '#10b981' : (dAmt < 0 ? '#ef4444' : '#6B7280');
            const arrow = dAmt > 0 ? '▲' : (dAmt < 0 ? '▼' : '–');
            const sign = dAmt > 0 ? '+' : (dAmt < 0 ? '−' : '');
            return `
            <div style="flex:1; min-width:130px; background:#FAFAFA; border:1px solid #E5E7EB; border-radius:10px; padding:10px 12px;">
                <div style="font-size:0.65rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:4px;">${q} cumulative drift</div>
                <div style="font-size:1rem; font-weight:800; color:${color}; line-height:1.2;">${arrow} ${sign}$${formatCurrency(Math.abs(dAmt))}</div>
                <div style="font-size:0.62rem; color:#9CA3AF; margin-top:2px;">$${formatCurrency(o.amount || 0)} → $${formatCurrency(n.amount || 0)}</div>
            </div>`;
        }).join('');
        quarterlyDrift = `
        <div style="margin-top:14px; padding-top:14px; border-top:1px dashed #E5E7EB;">
            <div style="font-size:0.7rem; color:#6B7280; font-weight:700; margin-bottom:8px;">Drift from initial baseline → latest snapshot</div>
            <div style="display:flex; gap:10px; flex-wrap:wrap;">${cards}</div>
        </div>`;
    }

    return `
        <div style="padding:22px; background:#FFFFFF; border:1px solid #fde68a; border-radius:14px; box-shadow:0 2px 10px rgba(245,158,11,0.05);">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:14px; border-bottom:1px solid #fef3c7; padding-bottom:12px; flex-wrap:wrap; gap:10px;">
                ${headerLeft}
                <div style="display:flex; align-items:center; gap:8px;">
                    <span style="font-size:0.7rem; color:#6B7280;">${history.length} snapshot${history.length === 1 ? '' : 's'} · ${changeEvents.length} change${changeEvents.length === 1 ? '' : 's'}</span>
                    <button id="pipeline-changelog-reset" style="background:#fff; border:1px solid #E5E7EB; color:#6B7280; padding:6px 10px; border-radius:8px; font-size:0.7rem; cursor:pointer; font-weight:600;"><i class="fa-solid fa-trash" style="margin-right:4px;"></i>Clear Log</button>
                </div>
            </div>
            ${baselineStrip}
            <div style="display:flex; flex-direction:column; gap:10px;">${changeListHTML}</div>
            ${quarterlyDrift}
            <div style="margin-top:12px; padding:8px 12px; background:#FFFBEB; border:1px solid #fef3c7; border-radius:8px; font-size:0.68rem; color:#92400e;">
                <i class="fa-solid fa-circle-info" style="margin-right:6px;"></i>A new snapshot is recorded automatically whenever any pipeline total changes. Stored locally in this browser.
            </div>
        </div>
    `;
}

/**
 * Render the country's full deal-level pipeline list. Sits at the bottom of
 * the country pipeline page as the live "current state" anchor that the
 * change log diffs against.
 * @param {string} filterCountry
 * @param {Array<{key,name,customer,quarter,amount,weighted}>} deals
 * @returns {string} HTML
 */
export function getCurrentPipelineListHTML(filterCountry, deals) {
    const sorted = [...(deals || [])].sort((a, b) => (b.amount || 0) - (a.amount || 0));
    const totalAmount = sorted.reduce((acc, d) => acc + (d.amount || 0), 0);
    const totalWeighted = sorted.reduce((acc, d) => acc + (d.weighted || 0), 0);

    const qBadge = (q) => {
        if (!q) return '<span style="color:#9CA3AF; font-size:0.62rem;">–</span>';
        const map = { Q1: '#6366f1', Q2: '#0ea5e9', Q3: '#f59e0b', Q4: '#ef4444' };
        const c = map[q] || '#6B7280';
        return `<span style="background:${c}1a; color:${c}; border:1px solid ${c}40; font-size:0.62rem; font-weight:800; padding:2px 7px; border-radius:6px; letter-spacing:0.04em;">${q}</span>`;
    };

    const rows = sorted.length === 0
        ? `<tr><td colspan="5" style="padding:18px; text-align:center; font-size:0.78rem; color:#9CA3AF; font-style:italic;">No active pipeline deals for ${filterCountry}.</td></tr>`
        : sorted.map((d, i) => `
            <tr style="border-bottom:1px solid #F3F4F6;">
                <td style="padding:9px 12px; font-size:0.72rem; color:#9CA3AF; text-align:right; width:36px;">${i + 1}</td>
                <td style="padding:9px 12px; font-size:0.78rem; color:#111827; font-weight:600;">${d.customer || '<span style="color:#9CA3AF;">—</span>'}</td>
                <td style="padding:9px 12px; font-size:0.74rem; color:#374151;">${d.name}</td>
                <td style="padding:9px 12px; text-align:center;">${qBadge(d.quarter)}</td>
                <td style="padding:9px 12px; font-size:0.78rem; text-align:right; color:#10b981; font-weight:700; white-space:nowrap;">$${formatCurrency(d.amount || 0)}</td>
                <td style="padding:9px 12px; font-size:0.78rem; text-align:right; color:#007AFF; font-weight:600; white-space:nowrap;">$${formatCurrency(d.weighted || 0)}</td>
            </tr>`).join('');

    return `
        <div style="padding:22px; background:#FFFFFF; border:1px solid #d1fae5; border-radius:14px; box-shadow:0 2px 10px rgba(16,185,129,0.05);">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:14px; border-bottom:1px solid #ecfdf5; padding-bottom:12px; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                    <div style="width:36px; height:36px; font-size:1rem; background:rgba(16,185,129,0.15); color:#10b981; border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-list"></i></div>
                    <div>
                        <h3 style="margin:0; font-size:0.72rem; color:#059669; font-weight:800; text-transform:uppercase; letter-spacing:0.08em;">CURRENT PIPELINE DEALS</h3>
                        <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827;">${filterCountry} — live snapshot (${sorted.length} deals)</h2>
                    </div>
                </div>
                <div style="display:flex; gap:18px; font-size:0.72rem; color:#374151; flex-wrap:wrap;">
                    <span>Pipeline <strong style="color:#10b981; font-size:0.95rem;">$${formatCurrency(totalAmount)}</strong></span>
                    <span>Weighted <strong style="color:#007AFF; font-size:0.95rem;">$${formatCurrency(totalWeighted)}</strong></span>
                </div>
            </div>
            <div style="overflow-x:auto;">
                <table style="width:100%; border-collapse:collapse; min-width:680px;">
                    <thead style="background:#F9FAFB;">
                        <tr>
                            <th style="padding:9px 12px; text-align:right; font-size:0.62rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">#</th>
                            <th style="padding:9px 12px; text-align:left; font-size:0.62rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">Customer</th>
                            <th style="padding:9px 12px; text-align:left; font-size:0.62rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">Deal</th>
                            <th style="padding:9px 12px; text-align:center; font-size:0.62rem; color:#6B7280; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">Q</th>
                            <th style="padding:9px 12px; text-align:right; font-size:0.62rem; color:#10b981; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">Amount</th>
                            <th style="padding:9px 12px; text-align:right; font-size:0.62rem; color:#007AFF; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">Weighted</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
            <div style="margin-top:12px; padding:8px 12px; background:#ecfdf5; border:1px solid #d1fae5; border-radius:8px; font-size:0.68rem; color:#065f46;">
                <i class="fa-solid fa-circle-info" style="margin-right:6px;"></i>This list is the live anchor. Each visit compares it against the previous snapshot — added, removed, or modified deals show up in the change log above.
            </div>
        </div>
    `;
}

export function getPartnerPerformanceHTML() {
    return `
        <div class="stat-card" style="padding: 24px; background: #FFFFFF; border: 1px solid #F3F4F6;">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 24px; border-bottom: 1px solid #F3F4F6; padding-bottom: 16px;">
                <div class="stat-icon" style="background: rgba(0,122,255,0.1); color: #007AFF; width: 36px; height: 36px;"><i class="fa-solid fa-ranking-star"></i></div>
                <div>
                    <h3 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin: 0;">Partner Performance Summary (Top 10)</h3>
                    <p style="font-size: 0.75rem; color: #6B7280; margin-top: 2px;">Ranked by Accumulated TCV</p>
                </div>
            </div>
            <div style="position: relative; height: 400px;"><canvas id="partner-top-performer-chart"></canvas></div>
        </div>
    `;
}

/**
 * Build per-account status board (timeline cards) like the reference design.
 * Renders one card per running/hold/decision-required POC.
 * @param {Object} stats - getPocStats result
 * @returns {string} HTML
 */
export function getPocStatusBoardHTML(stats) {
    const list = (stats.runningList || []).slice();
    if (list.length === 0) return '';

    const now = new Date();
    const todayMs = now.getTime();
    const dayMs = 86400 * 1000;

    const cards = list.map(r => {
        const startMs = r.startDateMs;
        const startDate = startMs ? new Date(startMs) : null;

        // Scheduled end: License End if available, else Start + max(120, days+30) days as projected end
        let endMs = r.licenseEndMs;
        let endLabel = r.licenseEndDisplay ? `License end: <b>${r.licenseEndDisplay}</b>` : null;
        if (!endMs && startMs) {
            const projectedDays = Math.max(120, (r.days || 0) + 30);
            endMs = startMs + projectedDays * dayMs;
            endLabel = `Should finish: <b>${new Date(endMs).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' })}</b>`;
        }
        if (!endMs) {
            endMs = todayMs + 30 * dayMs;
            endLabel = endLabel || 'Scheduled end: TBD';
        }

        // Status badge logic
        const isOverdue = (r.daysSinceStart != null && r.daysSinceStart > 60) || r.isStalled || (endMs < todayMs);
        const isHold = String(r.status).toLowerCase() === 'hold';
        let badgeColor, badgeBg, badgeIcon, badgeText;
        if (isHold) { badgeColor = '#D97706'; badgeBg = '#FEF3C7'; badgeIcon = 'fa-pause'; badgeText = 'Hold'; }
        else if (isOverdue) { badgeColor = '#DC2626'; badgeBg = '#FEE2E2'; badgeIcon = 'fa-triangle-exclamation'; badgeText = 'Overdue'; }
        else { badgeColor = '#1D4ED8'; badgeBg = '#DBEAFE'; badgeIcon = 'fa-circle-play'; badgeText = 'Running'; }

        // Timeline track positioning (% from left)
        const totalSpan = Math.max(1, (endMs - (startMs || todayMs)));
        const todayPctRaw = startMs ? ((todayMs - startMs) / totalSpan) * 100 : 100;
        const todayPct = Math.max(2, Math.min(98, todayPctRaw));
        const todayPastEnd = todayMs > endMs;

        // Issue dots derived from notes/techComm — simple keyword scan, distributed along timeline
        const noteText = `${r.notes || ''} ${r.techComm || ''}`;
        const issueKeywords = ['issue', 'block', 'hold', 'pause', 'overdue', 'delay', 'license expired', 'pic unreachable', 'pending'];
        const detectedIssues = [];
        const lower = noteText.toLowerCase();
        issueKeywords.forEach(kw => {
            if (lower.includes(kw)) detectedIssues.push(kw);
        });
        const issueDotsHtml = detectedIssues.slice(0, 5).map((kw, idx) => {
            const pct = 15 + (idx + 1) * (Math.max(0, todayPct - 20) / Math.max(1, detectedIssues.length));
            return `<div class="poc-sb-dot poc-sb-dot-issue" style="left:${pct}%" title="${kw}">
                        <span class="poc-sb-dot-label">${kw}</span>
                    </div>`;
        }).join('');

        // Format short dates for axis
        const startShort = startDate ? startDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : '—';
        const endShort = new Date(endMs).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
        const todayShort = now.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

        // Footer bullets — split notes by sentence-like separators
        const footerBits = noteText
            .split(/[•·\.\n\r]|[ ]·[ ]/)
            .map(s => s.trim())
            .filter(s => s.length >= 4)
            .slice(0, 4);
        const footerHtml = footerBits.length
            ? `<div class="poc-sb-footer">${footerBits.map(b => `<span class="poc-sb-footer-bit">${b}</span>`).join('<span class="poc-sb-sep">·</span>')}</div>`
            : '';

        const indSuffix = r.industry && r.industry !== 'Unknown' ? ` — ${r.industry}` : '';

        return `
        <div class="poc-status-card ${isOverdue ? 'is-overdue' : ''}">
            <div class="poc-sb-head">
                <div class="poc-sb-title-wrap">
                    <span class="poc-sb-title-pill">${r.name}${indSuffix}</span>
                </div>
                <div class="poc-sb-badge" style="color:${badgeColor}; background:${badgeBg};">
                    <i class="fa-solid ${badgeIcon}"></i> ${badgeText}
                </div>
            </div>
            <div class="poc-sb-meta">
                <span class="poc-sb-meta-item"><span class="poc-sb-meta-key">Start:</span> <b>${r.startDate || '—'}</b></span>
                <span class="poc-sb-meta-item"><span class="poc-sb-meta-key">Working days:</span> <b>${r.days}</b> days</span>
                <span class="poc-sb-meta-item"><span class="poc-sb-meta-key">Partner:</span> <b>${r.partner}</b></span>
                <span class="poc-sb-meta-item"><span class="poc-sb-meta-key">Est. Value:</span> <b>$${formatCurrency(r.estValue || 0)}</b></span>
            </div>
            <div class="poc-sb-timeline">
                <div class="poc-sb-track ${isOverdue ? 'overdue-track' : ''}">
                    <div class="poc-sb-progress" style="width:${todayPct}%"></div>
                    ${issueDotsHtml}
                    <div class="poc-sb-dot poc-sb-dot-today" style="left:${todayPct}%">
                        <span class="poc-sb-dot-label poc-sb-today-label">Today<br>${todayShort}</span>
                    </div>
                    <div class="poc-sb-dot poc-sb-dot-start" style="left:0%"></div>
                    <div class="poc-sb-dot poc-sb-dot-end ${todayPastEnd ? 'past' : ''}" style="left:100%"></div>
                </div>
                <div class="poc-sb-axis">
                    <span class="poc-sb-axis-l">Start<br>${startShort}</span>
                    <span class="poc-sb-axis-r">${endLabel ? endLabel.replace(/<\/?b>/g, '') : 'End'}<br>${endShort}</span>
                </div>
            </div>
            ${footerHtml}
        </div>`;
    }).join('');

    return `
        <div class="stat-card highlight-card" style="padding: 24px; margin-bottom: 30px; display: block;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 18px;">
                <div>
                    <h3 style="font-size: 1.05rem; font-weight: 700; color: #111827; margin: 0;">
                        <i class="fa-solid fa-chart-gantt" style="margin-right:8px; color:#1D4ED8;"></i>Account Status Board
                    </h3>
                    <p style="font-size: 0.75rem; color: #6B7280; margin: 4px 0 0;">Per-account POC timeline — start, today, scheduled end</p>
                </div>
                <div style="font-size: 0.75rem; color: #6B7280;">${list.length} accounts</div>
            </div>
            <div class="poc-sb-grid">
                ${cards}
            </div>
            <div class="poc-sb-legend">
                <span><span class="poc-sb-legend-line"></span> Active progress</span>
                <span><span class="poc-sb-legend-line overdue"></span> Overdue timeline</span>
                <span><span class="poc-sb-legend-dot issue"></span> Issue/block</span>
                <span><span class="poc-sb-legend-dot today"></span> Today</span>
                <span><span class="poc-sb-legend-dot end"></span> Scheduled end</span>
            </div>
        </div>
    `;
}

export function getPocHTML(stats, filters, uniqueValues) {
    const currentYear = new Date().getFullYear();


    return `
        <div class="stat-card" style="display:flex; flex-wrap: wrap; gap: 20px; padding: 18px; background: #FFFFFF; border: 1px solid #F3F4F6; margin-bottom: 24px;">
            <div style="display:flex; flex-direction:column; gap:8px;">
                <label style="font-size:0.8rem; color: #6B7280; font-weight:600; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right: 6px;"></i>Country</label>
                <select id="poc-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 12px; border-radius:6px; width: 180px;">
                    ${Array.from(uniqueValues.countries).map(c => `<option value="${c}" ${filters.country === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
            </div>
            <div style="display:flex; flex-direction:column; gap:8px;">
                <label style="font-size:0.8rem; color: #6B7280; font-weight:600; text-transform: uppercase;"><i class="fa-solid fa-industry" style="margin-right: 6px;"></i>Industry</label>
                <select id="poc-filter-industry" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 12px; border-radius:6px; width: 240px;">
                    ${Array.from(uniqueValues.industries).map(c => `<option value="${c}" ${filters.industry === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
            </div>
            <div style="display:flex; flex-direction:column; gap:8px;">
                <label style="font-size:0.8rem; color: #6B7280; font-weight:600; text-transform: uppercase;"><i class="fa-solid fa-handshake" style="margin-right: 6px;"></i>Partner</label>
                <select id="poc-filter-partner" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 12px; border-radius:6px; width: 180px;">
                    ${Array.from(uniqueValues.partners).map(c => `<option value="${c}" ${filters.partner === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
            </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 20px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="background: #EBF4FF; border: 1px solid rgba(0,122,255,0.2); padding: 24px; border-left: 5px solid #007AFF; cursor: pointer; transition: all 0.2s;"
                 onmouseover="showPocTooltip(event, this, '#007AFF')"
                 onmouseout="hidePocTooltip()"
                 data-title="Running POCs"
                 data-names="${encodeURIComponent(JSON.stringify(stats.runningNames))}">
                <div class="stat-icon" style="background: rgba(0, 122, 255, 0.15); color: #007AFF; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-play"></i></div>
                <div>
                    <h3 style="color: #007AFF; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Total Running POCs</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.statusStats.running} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Companies</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #FFF9ED; border: 1px solid rgba(245,158,11,0.2); padding: 24px; border-left: 5px solid #F59E0B; cursor: pointer; transition: all 0.2s;"
                 onmouseover="showPocTooltip(event, this, '#F59E0B')"
                 onmouseout="hidePocTooltip()"
                 data-title="Hold POCs"
                 data-names="${encodeURIComponent(JSON.stringify(stats.holdNames))}">
                <div class="stat-icon" style="background: rgba(245, 158, 11, 0.15); color: #F59E0B; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-pause"></i></div>
                <div>
                    <h3 style="color: #D97706; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Hold POCs</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.totalHold} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Companies</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #FFF5F5; border: 1px solid rgba(255,59,48,0.2); padding: 24px; border-left: 5px solid #ef4444; cursor: pointer; transition: all 0.2s;"
                 onmouseover="showPocTooltip(event, this, '#ef4444')"
                 onmouseout="hidePocTooltip()"
                 data-title="Long-term (100+) POCs"
                 data-names="${encodeURIComponent(JSON.stringify(stats.staledRunningList.map(r => r.name)))}">
                <div class="stat-icon" style="background: rgba(239, 68, 68, 0.2); color: #fca5a5; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-hourglass-half"></i></div>
                <div>
                    <h3 style="color: #FF3B30; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Long-term (100+)</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.staledRunningList.length} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Companies</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #FDF2FF; border: 1px solid rgba(168,85,247,0.25); padding: 24px; border-left: 5px solid #A855F7; cursor: pointer; transition: all 0.2s;"
                 onmouseover="showPocTooltip(event, this, '#A855F7')"
                 onmouseout="hidePocTooltip()"
                 data-title="Decision Required (2+ Months)"
                 data-names="${encodeURIComponent(JSON.stringify(stats.overTwoMonthsNames))}">
                <div class="stat-icon" style="background: rgba(168, 85, 247, 0.15); color: #A855F7; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-triangle-exclamation"></i></div>
                <div>
                    <h3 style="color: #A855F7; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Decision Required</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.overTwoMonthsList.length} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Companies</span></h2>
                    <p style="color: #7C3AED; font-size: 0.72rem; margin: 4px 0 0; font-weight: 500;">2+ months since start</p>
                </div>
            </div>
        </div>


        <div class="stat-card highlight-card" style="padding: 24px; margin-bottom: 30px; background: #FFFFFF; border: 1px solid #F3F4F6; display: block;">
            <h3 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin-bottom: 20px;">Monthly POC Status (${currentYear})</h3>
            <div style="position: relative; height: 350px;"><canvas id="poc-influx-chart"></canvas></div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="padding: 20px; display: flex; flex-direction: column;">
                <h4 style="font-size: 0.85rem; color: #111827; margin-bottom: 16px;"><i class="fa-solid fa-pie-chart" style="margin-right: 8px;"></i>Status Distribution</h4>
                <div style="position: relative; flex: 1;"><canvas id="poc-status-chart"></canvas></div>
                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin-top: 15px; border-top: 1px solid rgba(0,0,0,0.05); padding-top: 12px;">
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #34C759;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">Won</span></div>
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #FF3B30;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">Drop</span></div>
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #007AFF;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">Running</span></div>
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #FF9500;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">Hold</span></div>
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #9CA3AF;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">Others</span></div>
                </div>
            </div>
            <div class="stat-card highlight-card" style="padding: 20px; display: flex; flex-direction: column;">
                <h4 style="font-size: 0.85rem; color: #111827; margin-bottom: 16px;"><i class="fa-solid fa-clock" style="margin-right: 8px;"></i>Aging (100+ Working Days)</h4>
                <div style="position: relative; flex: 1;"><canvas id="poc-aging-chart"></canvas></div>
                <div style="display: flex; flex-wrap: wrap; justify-content: center; gap: 20px; margin-top: 15px; border-top: 1px solid rgba(0,0,0,0.05); padding-top: 12px;">
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #FF3B30;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">100+</span></div>
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #FF9500;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">60-100</span></div>
                    <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;"><span style="width: 30px; height: 10px; border-radius: 2px; background: #34C759;"></span><span style="font-size: 0.72rem; color: #6B7280; font-weight: 500;">&lt;60</span></div>
                </div>
            </div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="padding: 24px; display: block;">
                <h3 style="font-size: 1.05rem; font-weight: 600; margin-bottom: 12px;">Bottleneck POC</h3>
                <div style="position: relative; height: 260px;"><canvas id="poc-bottleneck-chart"></canvas></div>
            </div>
            <div class="stat-card highlight-card" style="padding: 24px; display: block;">
                <h3 style="font-size: 1.05rem; font-weight: 600; margin-bottom: 12px;">Industry Opportunity Analysis</h3>
                <div style="position: relative; height: 260px;"><canvas id="poc-industry-chart"></canvas></div>
            </div>
        </div>

        ${(() => {
            const onTrack = stats.runningList.filter(r => r.daysSinceStart != null && r.daysSinceStart <= 60).sort((a, b) => (a.daysSinceStart || 0) - (b.daysSinceStart || 0));
            const overdue = stats.runningList.filter(r => r.daysSinceStart == null || r.daysSinceStart > 60).sort((a, b) => (b.daysSinceStart || 0) - (a.daysSinceStart || 0));
            const allRows = [...onTrack, ...overdue];
            const thStyle = `padding: 10px 14px; color: #6B7280; font-weight: 600; font-size: 0.78rem; white-space: nowrap;`;
            const renderRow = (r, i, isDecision) => `
                <tr style="border-bottom: 1px solid ${isDecision ? 'rgba(168,85,247,0.12)' : '#E5E7EB'}; background: ${isDecision ? (i % 2 === 0 ? 'rgba(168,85,247,0.03)' : 'transparent') : (i % 2 === 0 ? '#FAFAFA' : 'transparent')};">
                    <td style="padding: 11px 14px; color: #9CA3AF; font-weight: 500; font-size: 0.78rem;">${i + 1}</td>
                    <td style="padding: 11px 14px; font-weight: 600; color: #111827; font-size: 0.8rem;">
                        ${r.name}
                        ${isDecision ? '<span style="background: rgba(168,85,247,0.15); color: #A855F7; font-size: 0.62rem; padding: 2px 6px; border-radius: 4px; font-weight: 700; margin-left: 6px; vertical-align: middle;">DECISION</span>' : ''}
                    </td>
                    <td style="padding: 11px 14px; color: #374151; font-size: 0.8rem;">${r.partner}</td>
                    <td style="padding: 11px 14px; color: #374151; font-size: 0.8rem;">${r.country}</td>
                    <td style="padding: 11px 14px; text-align: center;">
                        <span style="background: ${r.statusColor}20; color: ${r.statusColor}; padding: 3px 10px; border-radius: 6px; font-weight: 700; font-size: 0.7rem; text-transform: uppercase;">${r.status}</span>
                    </td>
                    <td style="padding: 11px 14px; text-align: center; color: #374151; font-size: 0.78rem;">${r.startDate || '-'}</td>
                    <td style="padding: 11px 14px; text-align: center;">
                        <span style="background: ${isDecision ? 'rgba(168,85,247,0.12)' : 'rgba(52,199,89,0.12)'}; color: ${isDecision ? '#7C3AED' : '#16a34a'}; padding: 3px 10px; border-radius: 12px; font-weight: 700; font-size: 0.78rem;">${r.daysSinceStart != null ? r.daysSinceStart + 'd' : '-'}</span>
                    </td>
                    <td style="padding: 11px 14px; text-align: center;">
                        <span style="color: ${r.days >= 100 ? '#FF3B30' : (r.days >= 60 ? '#FF9500' : '#34C759')}; font-weight: 700; font-size: 0.8rem;">${r.days}</span>
                    </td>
                    <td style="padding: 11px 14px; max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #6B7280; font-size: 0.78rem;">${r.notes || '-'}</td>
                </tr>`;
            return `
        <div class="stat-card highlight-card" style="padding: 24px; display: block;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;">
                <div>
                    <h3 style="font-size: 1.05rem; font-weight: 700; color: #111827; margin: 0;">All Running POCs</h3>
                    <p style="font-size: 0.75rem; color: #6B7280; margin: 4px 0 0;">
                        <span style="color: #16a34a; font-weight: 600;">${onTrack.length} on track</span>
                        <span style="margin: 0 8px; color: #D1D5DB;">|</span>
                        <span style="color: #A855F7; font-weight: 600;">${overdue.length} decision required (2+ months)</span>
                    </p>
                </div>
                <button onclick="copyDecisionList()" style="background: #A855F7; color: #fff; border: none; padding: 8px 16px; border-radius: 8px; font-size: 0.78rem; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-copy"></i> Copy Decision List
                </button>
            </div>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.8rem; text-align: left;" id="decision-required-table">
                    <thead>
                        <tr style="background: #F3F4F6; border-bottom: 2px solid #E5E7EB;">
                            <th style="${thStyle}">#</th>
                            <th style="${thStyle}">POC Name</th>
                            <th style="${thStyle}">Partner</th>
                            <th style="${thStyle}">Country</th>
                            <th style="${thStyle} text-align: center;">Status</th>
                            <th style="${thStyle} text-align: center;">Start Date</th>
                            <th style="${thStyle} text-align: center;">Days Elapsed</th>
                            <th style="${thStyle} text-align: center;">W.Days</th>
                            <th style="${thStyle}">Notes</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${onTrack.length > 0 ? `
                        <tr><td colspan="9" style="padding: 8px 14px; background: rgba(52,199,89,0.06); border-bottom: 1px solid rgba(52,199,89,0.2);">
                            <span style="font-size: 0.72rem; font-weight: 700; color: #16a34a; text-transform: uppercase; letter-spacing: 0.05em;"><i class="fa-solid fa-circle-check" style="margin-right: 5px;"></i>On Track — Within 60 Days (${onTrack.length})</span>
                        </td></tr>
                        ${onTrack.map((r, i) => renderRow(r, i + 1, false)).join('')}
                        ` : ''}
                        ${overdue.length > 0 ? `
                        <tr><td colspan="9" style="padding: 8px 14px; background: rgba(168,85,247,0.07); border-top: 2px solid rgba(168,85,247,0.2); border-bottom: 1px solid rgba(168,85,247,0.2);">
                            <span style="font-size: 0.72rem; font-weight: 700; color: #7C3AED; text-transform: uppercase; letter-spacing: 0.05em;"><i class="fa-solid fa-triangle-exclamation" style="margin-right: 5px;"></i>Decision Required — Over 60 Days, Oldest First (${overdue.length})</span>
                        </td></tr>
                        ${overdue.map((r, i) => renderRow(r, i + 1, true)).join('')}
                        ` : ''}
                    </tbody>
                </table>
            </div>
        </div>`;
        })()}

        ${getPocStatusBoardHTML(stats)}
    `;
}

export function getProjectHTML(stats, filterCountry, uniqueValues) {
    const currentYear = new Date().getFullYear();
    const country = filterCountry || 'All';

    const escape = (str) => String(str || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
    const fmtDate = (d) => d ? d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' }) : '-';
    const statusColor = (status) => {
        const s = String(status || '').toLowerCase();
        if (s.includes('resolved') || s.includes('complete') || s.includes('done')) return '#10B981';
        if (s.includes('progress')) return '#F59E0B';
        if (s.includes('hold') || s.includes('block')) return '#EF4444';
        return '#6B7280';
    };

    const recent = stats.entries.slice(0, 15);
    const rowsHtml = recent.map((r, i) => `
        <tr style="border-bottom: 1px solid #E5E7EB; background: ${i % 2 === 0 ? '#FAFAFA' : 'transparent'};">
            <td style="padding: 10px 14px; color: #9CA3AF; font-size: 0.78rem;">${i + 1}</td>
            <td style="padding: 10px 14px; color: #374151; font-size: 0.8rem;">${escape(r.country)}</td>
            <td style="padding: 10px 14px; font-weight: 600; color: #111827; font-size: 0.8rem;">${escape(r.poc)}</td>
            <td style="padding: 10px 14px; color: #374151; font-size: 0.78rem;">${escape(r.category)}</td>
            <td style="padding: 10px 14px; text-align: center;">
                <span style="background: ${statusColor(r.status)}20; color: ${statusColor(r.status)}; padding: 3px 10px; border-radius: 6px; font-weight: 700; font-size: 0.7rem; text-transform: uppercase;">${escape(r.status) || '-'}</span>
            </td>
            <td style="padding: 10px 14px; text-align: center; color: #374151; font-size: 0.78rem;">${fmtDate(r.date)}</td>
            <td style="padding: 10px 14px; text-align: center; color: #374151; font-size: 0.78rem;">${fmtDate(r.resolvedDate)}</td>
            <td style="padding: 10px 14px; max-width: 260px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #6B7280; font-size: 0.78rem;">${escape(r.log) || '-'}</td>
        </tr>
    `).join('');
    const thStyle = `padding: 10px 14px; color: #6B7280; font-weight: 600; font-size: 0.78rem; white-space: nowrap; text-align: left;`;

    return `
        <div class="stat-card" style="display:flex; flex-wrap: wrap; gap: 20px; padding: 18px; background: #FFFFFF; border: 1px solid #F3F4F6; margin-bottom: 24px;">
            <div style="display:flex; flex-direction:column; gap:8px;">
                <label style="font-size:0.8rem; color:#6B7280; font-weight:600; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right: 6px;"></i>Country</label>
                <select id="project-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 12px; border-radius:6px; width: 180px;">
                    ${Array.from(uniqueValues.countries).map(c => `<option value="${c}" ${country === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
            </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 20px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="background: #EBF4FF; border: 1px solid rgba(0,122,255,0.2); padding: 24px; border-left: 5px solid #007AFF;">
                <div class="stat-icon" style="background: rgba(0,122,255,0.15); color: #007AFF; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-list-check"></i></div>
                <div>
                    <h3 style="color: #007AFF; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Total Logs</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.totalLogs} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Entries</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #FFF9ED; border: 1px solid rgba(245,158,11,0.2); padding: 24px; border-left: 5px solid #F59E0B;">
                <div class="stat-icon" style="background: rgba(245,158,11,0.15); color: #F59E0B; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-spinner"></i></div>
                <div>
                    <h3 style="color: #D97706; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">In Progress</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.inProgressCount} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Open</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #ECFDF5; border: 1px solid rgba(16,185,129,0.2); padding: 24px; border-left: 5px solid #10B981;">
                <div class="stat-icon" style="background: rgba(16,185,129,0.15); color: #10B981; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-circle-check"></i></div>
                <div>
                    <h3 style="color: #059669; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Resolved</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.resolvedCount} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Closed</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #FDF2FF; border: 1px solid rgba(168,85,247,0.25); padding: 24px; border-left: 5px solid #A855F7;">
                <div class="stat-icon" style="background: rgba(168,85,247,0.15); color: #A855F7; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-stopwatch"></i></div>
                <div>
                    <h3 style="color: #A855F7; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Avg Resolution</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;">${stats.avgResolutionDays !== null ? stats.avgResolutionDays : '-'} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Days</span></h2>
                    <p style="color: #7C3AED; font-size: 0.72rem; margin: 4px 0 0; font-weight: 500;">across resolved logs</p>
                </div>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding: 24px; margin-bottom: 30px; background: #FFFFFF; border: 1px solid #F3F4F6; display: block;">
            <h3 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin-bottom: 20px;">Monthly Project Activity (${currentYear})</h3>
            <div style="position: relative; height: 350px;"><canvas id="project-activity-chart"></canvas></div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="padding: 20px; display: flex; flex-direction: column;">
                <h4 style="font-size: 0.85rem; color: #111827; margin-bottom: 16px;"><i class="fa-solid fa-pie-chart" style="margin-right: 8px;"></i>Category Distribution</h4>
                <div style="position: relative; flex: 1; min-height: 240px;"><canvas id="project-category-chart"></canvas></div>
            </div>
            <div class="stat-card highlight-card" style="padding: 20px; display: flex; flex-direction: column;">
                <h4 style="font-size: 0.85rem; color: #111827; margin-bottom: 16px;"><i class="fa-solid fa-flag" style="margin-right: 8px;"></i>Status Breakdown</h4>
                <div style="position: relative; flex: 1; min-height: 240px;"><canvas id="project-status-chart"></canvas></div>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding: 24px; display: block;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;">
                <div>
                    <h3 style="font-size: 1.05rem; font-weight: 700; color: #111827; margin: 0;">Recent Project Logs</h3>
                    <p style="font-size: 0.75rem; color: #6B7280; margin: 4px 0 0;">Latest ${recent.length} of ${stats.totalLogs} entries</p>
                </div>
            </div>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.8rem;">
                    <thead>
                        <tr style="background: #F3F4F6; border-bottom: 2px solid #E5E7EB;">
                            <th style="${thStyle}">#</th>
                            <th style="${thStyle}">Country</th>
                            <th style="${thStyle}">POC</th>
                            <th style="${thStyle}">Category</th>
                            <th style="${thStyle} text-align: center;">Status</th>
                            <th style="${thStyle} text-align: center;">Date</th>
                            <th style="${thStyle} text-align: center;">Resolved</th>
                            <th style="${thStyle}">Log Details</th>
                        </tr>
                    </thead>
                    <tbody>${rowsHtml || '<tr><td colspan="8" style="padding: 20px; text-align: center; color: #9CA3AF;">No project logs found.</td></tr>'}</tbody>
                </table>
            </div>
        </div>
    `;
}

export function getEventHTML(stats) {
    if (!stats) return '';
    const currentYear = new Date().getFullYear();
    return `
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px;">
            <div style="display: flex; align-items: center; gap: 15px;">
                <div class="stat-icon" style="background: rgba(245, 158, 11, 0.15); color: #f59e0b; width: 48px; height: 48px; font-size: 1.3rem;"><i class="fa-solid fa-calendar-check"></i></div>
                <div><h2 style="font-size: 1.6rem; font-weight: 700; color: #111827; margin: 0;">${currentYear} Event Performance Analytics</h2></div>
            </div>
        </div>
        <div class="stat-card" style="background:#FFF; padding: 24px;"><canvas id="event-roi-chart" style="height: 350px;"></canvas></div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; margin-top: 24px;">
            <div class="stat-card" style="background:#FFF; border-left: 5px solid #007AFF; padding:20px;">
                <h3 style="color:#6B7280; font-size:0.75rem; font-weight:700;">TOTAL SPENDING</h3>
                <h2 style="font-size:1.4rem; font-weight:800; margin: 8px 0;">$ ${formatCurrency(stats.totalSpending)}</h2>
                <div style="font-size: 0.75rem; color: #007AFF;">Across ${stats.eventCount} Events</div>
            </div>
            <div class="stat-card" style="background:#FFF; border-left: 5px solid #10b981; padding:20px;">
                <h3 style="color:#6B7280; font-size:0.75rem; font-weight:700;">TOTAL POC</h3>
                <h2 style="font-size:1.4rem; font-weight:800; margin: 8px 0;">${stats.totalPOC}</h2>
                <div style="font-size: 0.75rem; color: #10b981;">$${formatCurrency(stats.costPerPOC)} Per POC</div>
            </div>
            <div class="stat-card" style="background:#FFF; border-left: 5px solid #f59e0b; padding:20px;">
                <h3 style="color:#6B7280; font-size:0.75rem; font-weight:700;">CONVERTED DEALS</h3>
                <h2 style="font-size:1.4rem; font-weight:800; margin: 8px 0;">${stats.totalDeals}</h2>
                <div style="font-size: 0.75rem; color: #f59e0b;">$${formatCurrency(stats.costPerDeal)} Per Deal</div>
            </div>
        </div>
    `;
}

export function getCountrySpecificHTML(stats, countryName) {
    const items = stats.sortedYears.map(y => `
        <div class="stat-card" style="padding:20px; background:#F9FAFB; border-top:4px solid #6366f1; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
            <h3 style="font-size:0.9rem; font-weight:800; color:#4B5563; border-bottom:1px solid #E5E7EB; padding-bottom:10px; margin-bottom:15px;">${y} Performance</h3>
            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:10px;">
                <span style="font-size:0.8rem; color:#6B7280;">KOR TCV</span><span style="font-weight:800; color:#111827;">$${formatCurrency(stats.summary[y].kTcv)}</span>
            </div>
            <div style="display:flex; justify-content:space-between; align-items:center; margin-top:8px;">
                <span style="font-size:0.8rem; color:#6B7280;">Local ARR</span><span style="font-weight:800; color:#111827;">$${formatCurrency(stats.summary[y].lArr)}</span>
            </div>
        </div>
    `).join('');

    return `
        <div style="margin-bottom:30px;">
            <h2 style="font-size:1.75rem; font-weight:800; color:#111827; letter-spacing:-0.025em; margin-bottom:8px;">${countryName} Market Analysis</h2>
            <p style="color:#6B7280; font-size:0.9rem;">Historical performance and yearly metrics summary.</p>
        </div>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:20px;">${items}</div>
    `;
}

export function getKPIHTML(kpiData, currentKPIYear = new Date().getFullYear(), isAdmin = true, currentUser = 'admin', availableUsers = []) {
    if (!kpiData || !kpiData.categories) return '<p>No KPI data found.</p>';

    const getSubItems = (obj) => {
        const si = obj.subItems;
        if (si && si.length >= 3 && typeof si[0] === 'object') return si;
        return [{ name: "", achievements: [0,0,0,0] }, { name: "", achievements: [0,0,0,0] }, { name: "", achievements: [0,0,0,0] }];
    };

    const computeTotalAch = (subItems) =>
        [0,1,2,3].map(q => subItems.reduce((s, si) => s + ((si.achievements?.[q]) || 0), 0));

    const calculateRateFromSubs = (targets, subItems) => {
        const totalAch = computeTotalAch(subItems);
        const sumT = targets.reduce((a, b) => a + b, 0);
        const sumA = totalAch.reduce((a, b) => a + b, 0);
        if (sumT === 0) return sumA > 0 ? 100 : 0;
        return Math.min(200, Math.round((sumA / sumT) * 100));
    };

    const renderRow = (catId, objId, obj, catCellHtml = '') => {
        const subItems = getSubItems(obj);
        const totalAch = computeTotalAch(subItems);
        const rate = calculateRateFromSubs(obj.targets, subItems);
        const rateColor = rate >= 100 ? '#10B981' : (rate >= 70 ? '#F59E0B' : '#EF4444');

        const roAttr = isAdmin ? '' : 'readonly';
        const roStyle = isAdmin ? '' : 'cursor:default; opacity:0.75;';

        const mainRow = `
            <tr class="kpi-row kpi-main-row" data-cat="${catId}" data-obj="${objId}">
                ${catCellHtml}
                <td class="kpi-objective" style="padding: 0 10px;">
                    <div style="height: 44px; overflow: hidden; display: flex; align-items: center;">
                        <div ${isAdmin ? `contenteditable="true" onblur="this.style.background='transparent'; this.style.borderColor='transparent'; this.style.boxShadow='none'; window.updateKPIObjectiveName(this, ${catId}, ${objId})" onfocus="this.style.background='#FFF'; this.style.borderColor='#6366f1'; this.style.boxShadow='0 0 0 3px rgba(99,102,241,0.1)';" onmouseenter="if(document.activeElement!==this){this.style.borderColor='#CBD5E1';}" onmouseleave="if(document.activeElement!==this){this.style.borderColor='transparent';}"` : ''}
                             style="outline: none; width: 100%; min-width: 0; padding: 4px 6px; border: 1px dashed transparent; border-radius: 4px; transition: background 0.2s, border-color 0.2s; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 700; ${isAdmin ? 'cursor:text;' : ''}"
                             title="${isAdmin ? 'Click to edit' : ''}">${obj.name}</div>
                    </div>
                </td>
                <td class="kpi-indicator" style="padding: 0 15px;">
                    <div style="height: 44px; overflow: hidden; display: flex; align-items: center;">
                        <div ${isAdmin ? `contenteditable="true" onfocus="this.style.background='rgba(0,0,0,0.02)'; this.style.whiteSpace='normal'; this.style.overflow='visible';" onblur="this.style.whiteSpace='nowrap'; this.style.overflow='hidden'; this.style.background='transparent'; window.updateKPIText(this, 'kpis', ${catId}, ${objId});"` : ''}
                             style="outline: none; width: 100%; min-width: 0; transition: all 0.2s; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; ${isAdmin ? 'cursor:text;' : ''}"
                        >${obj.kpis || ''}</div>
                    </div>
                </td>
                ${[0,1,2,3].map(() => `<td style="background: rgba(16, 185, 129, 0.05);"></td>`).join('')}
                <td class="kpi-weight" style="padding: 4px;">
                    <div style="display: flex; align-items: center; justify-content: center; gap: 4px;">
                        <input type="number" min="0" max="100" step="1"
                               style="width: 50px; text-align: right; border: 1px solid ${isAdmin ? '#CBD5E1' : 'transparent'}; background: ${isAdmin ? '#F8FAFC' : 'rgba(0,0,0,0.02)'}; padding: 6px; border-radius: 6px; font-weight: 700; color: inherit; transition: all 0.2s; outline: none; ${roStyle}"
                               ${isAdmin ? `onfocus="this.style.background='#FFF'; this.style.borderColor='#6366f1'; this.style.boxShadow='0 0 0 3px rgba(99,102,241,0.1)';" onblur="this.style.background='#F8FAFC'; this.style.borderColor='#CBD5E1'; this.style.boxShadow='none'; window.updateKPINumber(this, 'weight', ${catId}, ${objId});"` : 'readonly'}
                               value="${obj.weight || 0}">%
                    </div>
                </td>
                <td class="kpi-rate" id="kpi-rate-${catId}-${objId}" style="color: ${rateColor}">${rate}%</td>
            </tr>
        `;

        const subItemRows = subItems.slice(0, 3).map((sub, subIdx) => `
            <tr class="kpi-subitem-row" data-cat="${catId}" data-obj="${objId}">
                <td class="kpi-subitem-cell">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <span class="kpi-subitem-num">${subIdx + 1}.</span>
                        <div class="kpi-subitem-input" ${isAdmin ? `contenteditable="true" data-placeholder="Enter detail..." onfocus="this.style.borderColor='#6366f1'; this.style.background='#FFF';" onblur="this.style.borderColor='transparent'; this.style.background='transparent'; window.updateKPISubItem(this, ${catId}, ${objId}, ${subIdx});"` : ''}
                             style="${!isAdmin ? 'cursor:default; color:#374151;' : ''}"
                        >${sub.name || (isAdmin ? '' : '—')}</div>
                    </div>
                </td>
                <td class="kpi-subitem-empty"></td>
                ${[0,1,2,3].map(qi => `
                    <td style="background: rgba(99,102,241,0.04); padding: 4px;">
                        <input type="text" class="kpi-achieve-input" value="${formatCurrency(sub.achievements?.[qi] || 0)}"
                               ${!isAdmin ? `onchange="window.updateKPISubItemAchievement(this, ${catId}, ${objId}, ${subIdx}, ${qi})"` : 'readonly'}
                               style="${isAdmin ? 'cursor:default; opacity:0.5;' : ''}">
                    </td>
                `).join('')}
                <td colspan="2" class="kpi-subitem-empty"></td>
            </tr>
        `).join('');

        const achievementRow = `
            <tr class="kpi-achieve-row kpi-row" data-cat="${catId}" data-obj="${objId}">
                <td colspan="2" style="text-align: right; font-weight: 700; color: #6366f1; background: rgba(99,102,241,0.06); padding: 6px 15px; font-size: 0.78rem; letter-spacing: 0.04em; text-transform: uppercase; border-top: 2px solid rgba(99,102,241,0.2);">
                    <i class="fa-solid fa-sigma" style="margin-right: 5px;"></i>Achievement (Total)
                </td>
                ${totalAch.map((a, qi) => `
                    <td style="background: rgba(99,102,241,0.10); border-top: 2px solid rgba(99,102,241,0.2); padding: 6px 10px;">
                        <input type="text" class="kpi-achieve-input" id="kpi-ach-total-${catId}-${objId}-${qi}" value="${formatCurrency(a)}" readonly
                               style="background: transparent; font-weight: 700; color: #4338CA; cursor: default;">
                    </td>
                `).join('')}
                <td colspan="2" style="background: rgba(99,102,241,0.06); border-top: 2px solid rgba(99,102,241,0.2);"></td>
            </tr>
        `;

        return mainRow + subItemRows + achievementRow;
    };

    let tableBody = '';
    let totalWeight = 0;
    let totalWeightedRate = 0;

    kpiData.categories.forEach((cat, catIdx) => {
        const rowsPerObj = 5; // 1 main + 3 sub-items + 1 achievement
        cat.objectives.forEach((obj, objIdx) => {
            const catCellHtml = objIdx === 0
                ? `<td rowspan="${cat.objectives.length * rowsPerObj}" class="kpi-cat-cell" style="background: ${cat.color}"><div contenteditable="true" onblur="this.style.background='transparent'; window.updateKPICategoryName(this, ${catIdx})" style="outline: none; min-height: 1.5em; width: 100%; text-align: center; transition: all 0.2s; cursor: text;" onfocus="this.style.background='rgba(255,255,255,0.2)';" title="Click to edit">${cat.name}</div></td>`
                : '';
            tableBody += renderRow(catIdx, objIdx, obj, catCellHtml);
            const subItems = getSubItems(obj);
            const rate = calculateRateFromSubs(obj.targets, subItems);
            totalWeight += (obj.weight || 0);
            totalWeightedRate += rate * (obj.weight || 0) / 100;
        });
    });

    const totalRateColor = totalWeightedRate >= 100 ? '#10B981' : (totalWeightedRate >= 70 ? '#F59E0B' : '#EF4444');
    const weightGap = Math.abs(totalWeight - 100);
    const weightWarning = weightGap > 0.1
        ? `<span style="color:#FCA5A5; font-size:0.72rem; font-weight:600; margin-left:8px;">(Total: ${Math.round(totalWeight)}% — Must be 100%)</span>`
        : '';

    const modeLabel = isAdmin
        ? `<span style="background:#fef3c7; color:#92400e; font-size:0.75rem; font-weight:700; padding:3px 10px; border-radius:20px; border:1px solid #fde68a;">🔐 Admin</span>`
        : `<span style="background:#ede9fe; color:#5b21b6; font-size:0.75rem; font-weight:700; padding:3px 10px; border-radius:20px; border:1px solid #ddd6fe;">👤 ${currentUser}</span>`;

    const userListForDropdown = [...availableUsers];
    if (!isAdmin && currentUser && currentUser !== 'admin' && !userListForDropdown.includes(currentUser)) {
        userListForDropdown.push(currentUser);
    }
    const userOptions = [
        `<option value="admin" ${isAdmin ? 'selected' : ''}>🔐 Admin (Structure & Target Setting)</option>`,
        ...userListForDropdown.map(u => `<option value="${u}" ${!isAdmin && currentUser === u ? 'selected' : ''}>👤 ${u}</option>`)
    ].join('');

    return `
        <div class="kpi-container">
            <div class="kpi-actions" style="flex-wrap: wrap; gap: 10px;">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <label style="font-size: 0.85rem; font-weight: 700; color: #64748B;">Year:</label>
                    <select id="kpi-year-select" style="padding: 8px 12px; border-radius: 8px; border: 1px solid #CBD5E1; font-weight: 700; font-family: inherit; font-size: 0.9rem; background: #FFF; outline: none; cursor: pointer; color: #1E293B;" onchange="window.changeKPIYear(this.value)">
                        ${[2026, 2027, 2028, 2029, 2030].map(y => `<option value="${y}" ${currentKPIYear === y ? 'selected' : ''}>${y}</option>`).join('')}
                    </select>
                </div>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <label style="font-size: 0.85rem; font-weight: 700; color: #64748B;">Mode:</label>
                    <select onchange="window.switchKPIMode(this.value)" style="padding: 8px 12px; border-radius: 8px; border: 1px solid #CBD5E1; font-weight: 600; font-family: inherit; font-size: 0.88rem; background: #FFF; outline: none; cursor: pointer; color: #1E293B; min-width: 180px;">
                        ${userOptions}
                    </select>
                    <button onclick="window.addKPIUser()" title="Add new team member" style="padding: 8px 12px; border-radius: 8px; border: 1px solid #CBD5E1; background: #F8FAFC; font-size: 0.82rem; font-weight: 700; cursor: pointer; color: #475569;">+ Member</button>
                    ${modeLabel}
                </div>
                <div style="flex-grow: 1;"></div>
                ${isAdmin ? `<button class="btn-kpi btn-reset" onclick="window.resetKPIData()"><i class="fa-solid fa-undo"></i> Reset</button>` : ''}
                <button class="btn-kpi btn-export" onclick="window.exportKPIData()"><i class="fa-solid fa-download"></i> Export</button>
                <button class="btn-kpi btn-save" onclick="window.saveKPIData()"><i class="fa-solid fa-save"></i> Save Changes</button>
            </div>
            <table class="kpi-table" style="table-layout: fixed; width: 100%;">
                <thead class="kpi-header">
                    <tr>
                        <th rowspan="2" style="width: 50px;">CAT.</th>
                        <th rowspan="2" style="width: 220px;">STRATEGIC OBJECTIVES</th>
                        <th rowspan="2">KEY PERFORMANCE INDICATORS</th>
                        <th colspan="4">TARGET / ACHIEVEMENT (${currentKPIYear})</th>
                        <th rowspan="2" style="width: 70px;">WEIGHT</th>
                        <th rowspan="2" style="width: 90px;">RATE</th>
                    </tr>
                    <tr>
                        <th style="width: 100px;">Q1</th>
                        <th style="width: 100px;">Q2</th>
                        <th style="width: 100px;">Q3</th>
                        <th style="width: 100px;">Q4</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableBody}
                </tbody>
                <tfoot>
                    <tr style="background: #1E293B; color: white;">
                        <td colspan="7" id="kpi-footer-label" style="text-align: right; padding: 12px 16px; font-weight: 700; font-size: 0.82rem; letter-spacing: 0.06em; text-transform: uppercase;">
                            TOTAL WEIGHT${weightWarning}
                        </td>
                        <td id="kpi-footer-weight" style="text-align: center; font-size: 1.05rem; font-weight: 800; padding: 12px; color: ${weightGap > 0.1 ? '#FCA5A5' : '#86EFAC'};">${Math.round(totalWeight)}%</td>
                        <td id="kpi-footer-rate" style="text-align: center; font-size: 1.05rem; font-weight: 800; color: ${totalRateColor}; padding: 12px;">${Math.round(totalWeightedRate)}%</td>
                    </tr>
                </tfoot>
            </table>
            <div style="margin-top: 20px; padding: 15px; background: rgba(99,102,241,0.05); border-radius: 12px; border-left: 4px solid #6366f1;">
                <p style="margin: 0; font-size: 0.8rem; color: #4F46E5; font-weight: 600;">
                    <i class="fa-solid fa-circle-info"></i> Enter details for each item (1~3) and quarterly Target/Achievement, then click 'Save Changes'. The final RATE is automatically calculated based on each item's Weight. (Sum of all Weights = 100%)
                </p>
            </div>
        </div>
    `;
}

/* ═══════════════════════════════════════════════════════════════
   TCV vs ARR Dashboard
   ═══════════════════════════════════════════════════════════════ */

/**
 * Generate HTML for the TCV vs ARR Revenue Mix dashboard.
 * @param {Object} stats - Output from getTcvArrStats
 * @param {{ country: string, contractYr: string }} filters
 * @returns {string}
 */
export function getTcvArrHTML(stats, filters = {}) {
    if (!stats) {
        return '<p style="padding:40px; text-align:center; color:#6B7280;">No TCV/ARR data found. Ensure the ORDER SHEET contains KOR TCV and End User columns.</p>';
    }

    const gapPct = stats.totalTcv > 0 ? ((stats.totalGap / stats.totalTcv) * 100).toFixed(1) : '0.0';
    const arrPct = stats.totalTcv > 0 ? ((stats.totalRecurringTcv / stats.totalTcv) * 100).toFixed(1) : '0.0';

    /* ── Filter Bar ── */
    const filterHtml = `
        <div class="stat-card" style="display:flex; align-items:center; gap:16px; padding: 12px 18px; background: #FFFFFF; border: 1px solid rgba(30, 64, 175, 0.15); border-left: 4px solid #1e40af; margin-bottom: 20px; flex-wrap: wrap;">
            <div style="display:flex; align-items:center; gap:8px;">
                <label style="font-size:0.78rem; color:#1e40af; font-weight:700; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right:6px;"></i>Country</label>
                <select id="tcvarr-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #CBD5E1; padding:6px 12px; border-radius:8px; width:170px; font-size:0.82rem; font-weight:500;">
                    ${stats.uniqueCountries.map(c => `<option value="${c}" ${(filters.country || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <label style="font-size:0.78rem; color:#1e40af; font-weight:700; text-transform: uppercase;"><i class="fa-solid fa-calendar" style="margin-right:6px;"></i>Contract Yr</label>
                <select id="tcvarr-filter-year" style="background:#F9FAFB; color:#111827; border:1px solid #CBD5E1; padding:6px 12px; border-radius:8px; width:120px; font-size:0.82rem; font-weight:500;">
                    ${stats.uniqueYears.map(y => `<option value="${y}" ${(filters.contractYr || 'All') === y ? 'selected' : ''}>${y}</option>`).join('')}
                </select>
            </div>
            <span style="font-size: 0.72rem; color: #64748b; margin-left: auto;">Showing ${stats.accountCount} accounts · ${filters.country || 'All Regions'} · Yr: ${filters.contractYr || 'All'}</span>
        </div>
    `;

    /* ── KPI Summary Cards ── */
    const kpiHtml = `
        <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px;">
            <div class="stat-card" style="background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(15,23,42,0.30);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.06); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.18); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-file-invoice-dollar" style="font-size:1rem;"></i></div>
                    <h3 style="font-size:0.78rem; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:#ffffff; margin:0;">Total KOR TCV</h3>
                </div>
                <h2 style="font-size:1.8rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em; color:#ffffff;">$${formatCurrency(stats.totalTcv)}</h2>
                <div style="font-size:0.78rem; margin-top:8px; color:#dbeafe; font-weight:600;">${stats.accountCount} accounts</div>
            </div>
            <div class="stat-card" style="background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(30,58,138,0.25);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.06); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.18); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-arrows-rotate" style="font-size:1rem;"></i></div>
                    <h3 style="font-size:0.78rem; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:#ffffff; margin:0;">Total KOR ARR</h3>
                </div>
                <h2 style="font-size:1.8rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em; color:#ffffff;">$${formatCurrency(stats.totalArr)}</h2>
                <div style="font-size:0.78rem; margin-top:8px; color:#dbeafe; font-weight:600;">${arrPct}% of TCV is recurring</div>
            </div>
            <div class="stat-card" style="background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(37,99,235,0.25);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.06); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.18); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-chart-column" style="font-size:1rem;"></i></div>
                    <h3 style="font-size:0.78rem; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:#ffffff; margin:0;">Revenue Gap</h3>
                </div>
                <h2 style="font-size:1.8rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em; color:#ffffff;">$${formatCurrency(stats.totalGap)}</h2>
                <div style="font-size:0.78rem; margin-top:8px; color:#dbeafe; font-weight:600;">${gapPct}% one-time revenue</div>
            </div>
            <div class="stat-card" style="background: linear-gradient(135deg, #0284c7 0%, #0ea5e9 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(14,165,233,0.25);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.10); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.22); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-building" style="font-size:1rem; color:white;"></i></div>
                    <h3 style="font-size:0.78rem; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; color:#ffffff; margin:0;">Account Mix</h3>
                </div>
                <div style="display:flex; gap:18px; margin-top:6px; align-items:baseline;">
                    <div><h2 style="font-size:1.6rem; font-weight:800; margin:0; color:#ffffff;">${stats.recurringCount}</h2><span style="font-size:0.72rem; font-weight:600; color:#e0f2fe;">Recurring</span></div>
                    <div><h2 style="font-size:1.6rem; font-weight:800; margin:0; color:#fde68a;">${stats.perpetualCount}</h2><span style="font-size:0.72rem; font-weight:600; color:#fef3c7;">Perpetual</span></div>
                </div>
            </div>
        </div>
    `;

    /* ── Chart Container ── */
    const chartHtml = `
        <div class="stat-card" style="background:#FFF; padding:20px; margin-bottom:20px; box-shadow: 0 6px 18px rgba(0,0,0,0.06); border:1px solid #F1F5F9; border-radius:14px;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:16px; flex-wrap:wrap; gap:10px;">
                <h3 style="font-size:1.05rem; font-weight:800; color:#111827; margin:0; display:flex; align-items:center; gap:10px;">
                    <i class="fa-solid fa-circle-half-stroke" style="color:#1e40af;"></i> ARR Recurring Rate by Account
                    <span style="background:#eff6ff; color:#1e40af; font-size:0.68rem; font-weight:700; padding:3px 10px; border-radius:12px;">${stats.items.length > 15 ? 'Top 15 by TCV' : stats.items.length + ' accounts'}</span>
                </h3>
                <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                    <span style="display:flex; align-items:center; gap:4px; font-size:0.7rem; color:#64748b;"><span style="width:10px; height:10px; background:#059669; border-radius:2px; display:inline-block;"></span>≥80% Healthy</span>
                    <span style="display:flex; align-items:center; gap:4px; font-size:0.7rem; color:#64748b;"><span style="width:10px; height:10px; background:#2563eb; border-radius:2px; display:inline-block;"></span>60–79%</span>
                    <span style="display:flex; align-items:center; gap:4px; font-size:0.7rem; color:#64748b;"><span style="width:10px; height:10px; background:#d97706; border-radius:2px; display:inline-block;"></span>40–59%</span>
                    <span style="display:flex; align-items:center; gap:4px; font-size:0.7rem; color:#64748b;"><span style="width:10px; height:10px; background:#dc2626; border-radius:2px; display:inline-block;"></span>&lt;40% Risk</span>
                </div>
            </div>
            <div id="tcvarr-chart-container" style="position:relative; width:100%; min-height:400px;">
                <canvas id="tcvarr-bar-chart"></canvas>
            </div>
        </div>
    `;

    /* ── Detail Table ── */
    const tableRows = stats.items.map((item, i) => {
        const arrFill = item.recurringPct || 0;
        const gapColor = item.gapPct > 80 ? '#ef4444' : item.gapPct > 50 ? '#f59e0b' : '#22c55e';
        const typeBadge = item.isPerpetual
            ? '<span style="background:#fef3c7; color:#92400e; font-size:0.62rem; font-weight:700; padding:2px 8px; border-radius:10px;">PERPETUAL</span>'
            : '<span style="background:#d1fae5; color:#065f46; font-size:0.62rem; font-weight:700; padding:2px 8px; border-radius:10px;">RECURRING</span>';

        return `<tr style="border-bottom:1px solid #F3F4F6; transition: background 0.15s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='transparent'">
            <td style="padding:10px 12px; color:#94a3b8; font-weight:700; font-family:monospace; font-size:0.78rem; width:40px;">${String(i + 1).padStart(2, '0')}</td>
            <td style="padding:10px 12px; font-weight:700; color:#1e293b; font-size:0.82rem; max-width:200px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${item.name}</td>
            <td style="padding:10px 12px; font-size:0.75rem; color:#64748b;">${item.country}</td>
            <td style="padding:10px 12px; text-align:right; font-weight:800; color:#1e40af; font-size:0.85rem;">$${formatCurrency(item.tcv)}</td>
            <td style="padding:10px 12px; text-align:right; font-weight:800; color:#2563eb; font-size:0.85rem;">$${formatCurrency(item.arr)}</td>
            <td style="padding:10px 12px; text-align:right; font-weight:700; color:${gapColor}; font-size:0.82rem;">$${formatCurrency(item.gap)}</td>
            <td style="padding:10px 12px; text-align:center;">
                <div style="width:60px; background:#f1f5f9; border-radius:10px; height:8px; overflow:hidden; display:inline-block; vertical-align:middle;" title="Recurring %: ${arrFill.toFixed(1)}%">
                    <div style="width:${Math.min(arrFill, 100)}%; height:100%; background: linear-gradient(90deg, #3b82f6, #1e40af); border-radius:10px; transition: width 0.4s;"></div>
                </div>
                <span style="font-size:0.68rem; color:${gapColor}; font-weight:700; margin-left:4px;">${item.gapPct.toFixed(0)}%</span>
            </td>
            <td style="padding:10px 12px; text-align:center;">${typeBadge}</td>
        </tr>`;
    }).join('');

    const tableHtml = `
        <div class="stat-card" style="background:#FFF; padding:18px; margin-bottom:20px; box-shadow: 0 6px 18px rgba(0,0,0,0.06); border:1px solid #F1F5F9; border-radius:14px;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:14px;">
                <h3 style="font-size:1rem; font-weight:800; color:#111827; margin:0; display:flex; align-items:center; gap:8px;">
                    <i class="fa-solid fa-table-list" style="color:#1e40af;"></i> Revenue Mix Detail
                    <span style="background:#f1f5f9; color:#475569; font-size:0.68rem; font-weight:700; padding:2px 10px; border-radius:12px;">Sorted by TCV ↓</span>
                </h3>
                <button id="tcvarr-table-toggle" onclick="(function(){const el=document.getElementById('tcvarr-table-body');const btn=document.getElementById('tcvarr-table-toggle');const hidden=el.style.display==='none';el.style.display=hidden?'block':'none';btn.textContent=hidden?'Hide Details':'Show Details';})()" style="background:#eff6ff; color:#1e40af; border:1px solid #bfdbfe; padding:5px 14px; border-radius:8px; font-size:0.75rem; font-weight:600; cursor:pointer;">Hide Details</button>
            </div>
            <div id="tcvarr-table-body">
                <div style="overflow-x:auto;">
                    <table style="width:100%; border-collapse:collapse; min-width:900px;">
                        <thead>
                            <tr style="background:#F8FAFC; text-align:left; border-bottom:2px solid #E2E8F0;">
                                <th style="padding:10px 12px; font-size:0.68rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; width:40px;">#</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.06em;">End User</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.06em;">Country</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#1e40af; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; text-align:right;">KOR TCV (USD)</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#1e40af; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; text-align:right;">KOR ARR (USD)</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; text-align:right;">GAP</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; text-align:center;">ARR RATIO</th>
                                <th style="padding:10px 12px; font-size:0.68rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.06em; text-align:center;">TYPE</th>
                            </tr>
                        </thead>
                        <tbody>${tableRows}</tbody>
                    </table>
                </div>
            </div>
        </div>
    `;

    /* ── Strategic Insights (AI-style analysis) ── */
    const strategicInsightsHtml = _buildStrategicInsightsHTML(stats, arrPct, gapPct);

    const segmentationHtml = _buildAccountSegmentationHTML(stats);

    return filterHtml + kpiHtml + chartHtml + tableHtml + strategicInsightsHtml + segmentationHtml;
}

/**
 * Build the Strategic Insights HTML panel for TCV vs ARR view.
 * Pure function — derives all insights from pre-computed stats.
 * @param {Object} stats - Output from getTcvArrStats
 * @param {string} arrPct - Recurring rate percentage string
 * @param {string} gapPct - Gap percentage string
 * @returns {string} HTML string
 */
function _buildStrategicInsightsHTML(stats, arrPct, gapPct) {
    if (!stats || stats.items.length === 0) return '';

    /* ── 1. Revenue Health Score ── */
    const healthRatio = parseFloat(arrPct);
    let healthGrade, healthColor, healthBg, healthBorder, healthIcon, healthDesc;
    if (healthRatio >= 80) {
        healthGrade = 'A'; healthColor = '#059669'; healthBg = '#ecfdf5'; healthBorder = '#a7f3d0';
        healthIcon = 'fa-shield-check';
        healthDesc = 'Excellent recurring base. The business has strong long-term valuation with a predictable revenue floor.';
    } else if (healthRatio >= 60) {
        healthGrade = 'B'; healthColor = '#2563eb'; healthBg = '#eff6ff'; healthBorder = '#93c5fd';
        healthIcon = 'fa-chart-line';
        healthDesc = 'Healthy mix. Recurring revenue forms a solid base but there\'s room to convert more one-time deals into subscriptions.';
    } else if (healthRatio >= 40) {
        healthGrade = 'C'; healthColor = '#d97706'; healthBg = '#fffbeb'; healthBorder = '#fde68a';
        healthIcon = 'fa-triangle-exclamation';
        healthDesc = 'Moderate risk. More than half of the revenue is one-time, meaning it must be re-earned annually. Prioritize converting perpetual licenses.';
    } else {
        healthGrade = 'D'; healthColor = '#dc2626'; healthBg = '#fef2f2'; healthBorder = '#fecaca';
        healthIcon = 'fa-circle-exclamation';
        healthDesc = 'Critical dependency on one-time revenue. Without new contract wins, next year\'s revenue will decline sharply.';
    }

    const healthScoreHtml = `
        <div style="display:flex; gap:16px; align-items:stretch;">
            <div style="min-width:90px; display:flex; flex-direction:column; align-items:center; justify-content:center; background:${healthBg}; border:2px solid ${healthBorder}; border-radius:14px; padding:12px;">
                <div style="font-size:2.5rem; font-weight:900; color:${healthColor}; line-height:1;">${healthGrade}</div>
                <div style="font-size:0.62rem; font-weight:700; color:${healthColor}; text-transform:uppercase; margin-top:4px;">Grade</div>
            </div>
            <div style="flex:1;">
                <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
                    <i class="fa-solid ${healthIcon}" style="color:${healthColor}; font-size:1rem;"></i>
                    <h4 style="margin:0; font-size:0.88rem; font-weight:800; color:#1e293b;">Revenue Health Score</h4>
                    <span style="background:${healthBg}; color:${healthColor}; font-size:0.72rem; font-weight:700; padding:2px 10px; border-radius:10px; border:1px solid ${healthBorder};">${arrPct}% Recurring</span>
                </div>
                <p style="margin:0 0 8px 0; font-size:0.78rem; color:#475569; line-height:1.5;">${healthDesc}</p>
                <div style="display:flex; gap:20px; font-size:0.72rem; color:#64748b;">
                    <span><strong style="color:#1e40af;">TCV:</strong> $${formatCurrency(stats.totalTcv)}</span>
                    <span><strong style="color:#16a34a;">ARR:</strong> $${formatCurrency(stats.totalArr)}</span>
                    <span><strong style="color:#7c3aed;">Gap:</strong> $${formatCurrency(stats.totalGap)} (${gapPct}%)</span>
                </div>
            </div>
        </div>
    `;

    /* ── 2. Top Risk Accounts (Wide Gap) ── */
    const wideGapAccounts = stats.items
        .filter(a => a.gapPct > 50 && a.tcv > 0 && !a.isPerpetual)
        .sort((a, b) => b.gap - a.gap)
        .slice(0, 3);

    const riskStrategies = [
        'Propose a 3-Year subscription migration with a 15-20% discount incentive to convert the non-recurring upfront fee into ARR.',
        'Bundle value-added services (training, premium support) to create a recurring revenue stream alongside the one-time deal.',
        'Offer a "Hybrid Contract" — reduce the upfront fee by 30% and introduce annual maintenance fees to build recurring revenue.'
    ];

    const riskAccountsHtml = wideGapAccounts.length > 0 ? `
        <div style="margin-top:18px;">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                <i class="fa-solid fa-triangle-exclamation" style="color:#dc2626; font-size:1rem;"></i>
                <h4 style="margin:0; font-size:0.88rem; font-weight:800; color:#1e293b;">Top Risk Accounts <span style="font-weight:500; color:#94a3b8; font-size:0.72rem;">(Wide Gap — High TCV, Low ARR)</span></h4>
            </div>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap:10px;">
                ${wideGapAccounts.map((acc, i) => {
                    return `
                    <div style="background:#fef2f2; border:1px solid #fecaca; border-radius:10px; padding:12px 14px; border-left:4px solid #ef4444;">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                            <span style="font-weight:800; color:#991b1b; font-size:0.82rem;">${acc.name}</span>
                            <span style="background:#fee2e2; color:#b91c1c; font-size:0.6rem; font-weight:700; padding:2px 8px; border-radius:8px;">${acc.isPerpetual ? 'PERPETUAL' : 'LOW RECURRING'}</span>
                        </div>
                        <div style="display:flex; gap:12px; font-size:0.7rem; color:#64748b; margin-bottom:8px;">
                            <span>TCV: <strong style="color:#1e40af;">$${formatCurrency(acc.tcv)}</strong></span>
                            <span>ARR: <strong style="color:#16a34a;">$${formatCurrency(acc.arr)}</strong></span>
                            <span>Gap: <strong style="color:#dc2626;">$${formatCurrency(acc.gap)}</strong></span>
                        </div>
                        <div style="background:white; border-radius:6px; padding:8px 10px; border:1px solid #fecaca;">
                            <div style="display:flex; align-items:flex-start; gap:6px;">
                                <i class="fa-solid fa-lightbulb" style="color:#f59e0b; font-size:0.7rem; margin-top:2px; flex-shrink:0;"></i>
                                <span style="font-size:0.7rem; color:#57534e; line-height:1.45;">${riskStrategies[i] || riskStrategies[0]}</span>
                            </div>
                        </div>
                    </div>`;
                }).join('')}
            </div>
        </div>
    ` : '';

    /* ── 3. Expansion Opportunities (Narrow Gap) ── */
    const narrowGapAccounts = stats.items
        .filter(a => !a.isPerpetual && a.gapPct < 30 && a.arr > 0)
        .sort((a, b) => a.gapPct - b.gapPct)
        .slice(0, 3);

    const expansionStrategies = [
        'Already highly recurring — propose premium tier upgrades (e.g., Enterprise license, advanced analytics modules) for incremental ARR growth.',
        'Leverage renewal touchpoints to cross-sell complementary products. This account\'s strong recurring base makes expansion low-risk.',
        'Introduce multi-year lock-in with an annual price escalator (3-5%) to compound recurring growth on this already stable account.'
    ];

    const expansionHtml = narrowGapAccounts.length > 0 ? `
        <div style="margin-top:18px;">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                <i class="fa-solid fa-rocket" style="color:#059669; font-size:1rem;"></i>
                <h4 style="margin:0; font-size:0.88rem; font-weight:800; color:#1e293b;">Expansion Opportunities <span style="font-weight:500; color:#94a3b8; font-size:0.72rem;">(Narrow Gap — High Recurring Stability)</span></h4>
            </div>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap:10px;">
                ${narrowGapAccounts.map((acc, i) => {
                    const arrRatio = acc.recurringPct.toFixed(1);
                    return `
                    <div style="background:#ecfdf5; border:1px solid #a7f3d0; border-radius:10px; padding:12px 14px; border-left:4px solid #10b981;">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                            <span style="font-weight:800; color:#065f46; font-size:0.82rem;">${acc.name}</span>
                            <span style="background:#d1fae5; color:#047857; font-size:0.6rem; font-weight:700; padding:2px 8px; border-radius:8px;">${arrRatio}% RECURRING</span>
                        </div>
                        <div style="display:flex; gap:12px; font-size:0.7rem; color:#64748b; margin-bottom:8px;">
                            <span>TCV: <strong style="color:#1e40af;">$${formatCurrency(acc.tcv)}</strong></span>
                            <span>ARR: <strong style="color:#16a34a;">$${formatCurrency(acc.arr)}</strong></span>
                            <span>Gap: <strong style="color:#059669;">$${formatCurrency(acc.gap)}</strong> only</span>
                        </div>
                        <div style="background:white; border-radius:6px; padding:8px 10px; border:1px solid #a7f3d0;">
                            <div style="display:flex; align-items:flex-start; gap:6px;">
                                <i class="fa-solid fa-arrow-trend-up" style="color:#059669; font-size:0.7rem; margin-top:2px; flex-shrink:0;"></i>
                                <span style="font-size:0.7rem; color:#57534e; line-height:1.45;">${expansionStrategies[i] || expansionStrategies[0]}</span>
                            </div>
                        </div>
                    </div>`;
                }).join('')}
            </div>
        </div>
    ` : '';

    /* ── 4. Operational Forecast (Revenue Cliff) ── */
    const recurringRate = parseFloat(arrPct) / 100;
    const oneTimeRevenue = stats.totalTcv - stats.totalArr;
    const projectedNextYearFloor = stats.totalArr;
    const revenueCliff = oneTimeRevenue;
    const cliffPct = stats.totalTcv > 0 ? ((revenueCliff / stats.totalTcv) * 100).toFixed(1) : '0.0';
    const neededNewDeals = Math.ceil(revenueCliff / (stats.totalTcv / Math.max(stats.accountCount, 1)));

    const forecastHtml = `
        <div style="margin-top:18px;">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                <i class="fa-solid fa-chart-area" style="color:#7c3aed; font-size:1rem;"></i>
                <h4 style="margin:0; font-size:0.88rem; font-weight:800; color:#1e293b;">Operational Forecast <span style="font-weight:500; color:#94a3b8; font-size:0.72rem;">(Revenue Cliff Analysis)</span></h4>
            </div>
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px;">
                <div style="background:linear-gradient(135deg, #f5f3ff 0%, #ede9fe 100%); border:1px solid #c4b5fd; border-radius:12px; padding:16px;">
                    <h5 style="margin:0 0 10px 0; font-size:0.75rem; font-weight:700; color:#6d28d9; text-transform:uppercase; letter-spacing:0.06em;">If No New Contracts Are Signed</h5>
                    <div style="display:flex; align-items:baseline; gap:6px; margin-bottom:8px;">
                        <span style="font-size:1.6rem; font-weight:900; color:#7c3aed;">$${formatCurrency(projectedNextYearFloor)}</span>
                        <span style="font-size:0.72rem; color:#8b5cf6; font-weight:600;">projected floor (next year)</span>
                    </div>
                    <div style="margin-bottom:10px;">
                        <div style="display:flex; justify-content:space-between; font-size:0.68rem; color:#7c3aed; font-weight:600; margin-bottom:3px;">
                            <span>Recurring Coverage</span>
                            <span>${arrPct}%</span>
                        </div>
                        <div style="width:100%; background:rgba(124,58,237,0.1); border-radius:8px; height:10px; overflow:hidden;">
                            <div style="width:${Math.min(parseFloat(arrPct), 100)}%; height:100%; background:linear-gradient(90deg, #8b5cf6, #7c3aed); border-radius:8px; transition: width 0.6s;"></div>
                        </div>
                    </div>
                    <ul style="margin:0; padding:0 0 0 16px; font-size:0.72rem; color:#475569; line-height:1.7;">
                        <li><strong style="color:#dc2626;">$${formatCurrency(revenueCliff)}</strong> one-time revenue at risk (${cliffPct}% of current TCV)</li>
                        <li><strong style="color:#7c3aed;">${stats.perpetualCount}</strong> perpetual-only accounts contribute $0 to next year's recurring floor</li>
                        <li>Need ~<strong style="color:#1e40af;">${neededNewDeals}</strong> new deals of average size to maintain current revenue level</li>
                    </ul>
                </div>
                <div style="background:linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); border:1px solid #86efac; border-radius:12px; padding:16px;">
                    <h5 style="margin:0 0 10px 0; font-size:0.75rem; font-weight:700; color:#166534; text-transform:uppercase; letter-spacing:0.06em;">BizDev Priority Actions</h5>
                    <div style="display:flex; flex-direction:column; gap:8px;">
                        <div style="display:flex; align-items:flex-start; gap:8px; background:white; padding:8px 10px; border-radius:8px; border:1px solid #bbf7d0;">
                            <span style="background:#059669; color:white; font-size:0.58rem; font-weight:800; padding:2px 6px; border-radius:4px; flex-shrink:0; margin-top:1px;">P1</span>
                            <span style="font-size:0.72rem; color:#374151; line-height:1.45;">Convert top ${Math.min(wideGapAccounts.length, 3)} high-risk accounts to subscription. Potential ARR uplift: <strong style="color:#059669;">$${formatCurrency(wideGapAccounts.reduce((s, a) => s + a.gap, 0))}</strong></span>
                        </div>
                        <div style="display:flex; align-items:flex-start; gap:8px; background:white; padding:8px 10px; border-radius:8px; border:1px solid #bbf7d0;">
                            <span style="background:#2563eb; color:white; font-size:0.58rem; font-weight:800; padding:2px 6px; border-radius:4px; flex-shrink:0; margin-top:1px;">P2</span>
                            <span style="font-size:0.72rem; color:#374151; line-height:1.45;">Upsell ${narrowGapAccounts.length} high-stability accounts to premium tiers. These accounts have proven commitment to recurring spend.</span>
                        </div>
                        <div style="display:flex; align-items:flex-start; gap:8px; background:white; padding:8px 10px; border-radius:8px; border:1px solid #bbf7d0;">
                            <span style="background:#7c3aed; color:white; font-size:0.58rem; font-weight:800; padding:2px 6px; border-radius:4px; flex-shrink:0; margin-top:1px;">P3</span>
                            <span style="font-size:0.72rem; color:#374151; line-height:1.45;">Target recurring rate of <strong style="color:#059669;">60%+</strong> by Q4 to improve SaaS valuation multiples and reduce revenue volatility.</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;

    /* ── Assemble Strategic Insights Panel ── */
    return `
        <div class="stat-card" style="background:#FFF; padding:22px; margin-bottom:20px; box-shadow: 0 6px 18px rgba(0,0,0,0.06); border:1px solid #F1F5F9; border-radius:14px; border-top: 4px solid #6366f1;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:18px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
                <h3 style="font-size:1.1rem; font-weight:800; color:#111827; margin:0; display:flex; align-items:center; gap:10px;">
                    <div style="width:36px; height:36px; background:linear-gradient(135deg, #6366f1, #8b5cf6); border-radius:10px; display:flex; align-items:center; justify-content:center;">
                        <i class="fa-solid fa-brain" style="color:white; font-size:0.9rem;"></i>
                    </div>
                    Strategic Insights
                    <span style="background:#ede9fe; color:#6d28d9; font-size:0.62rem; font-weight:700; padding:3px 10px; border-radius:10px;">SaaS Financial Analysis</span>
                </h3>
                <span style="font-size:0.68rem; color:#94a3b8; font-style:italic;">Auto-generated from ${stats.accountCount} accounts</span>
            </div>

            ${healthScoreHtml}
            ${riskAccountsHtml}
            ${expansionHtml}
            ${forecastHtml}
        </div>
    `;
}

/**
 * Build the Account Segmentation panel — groups accounts into 4 strategic buckets.
 * @param {Object} stats - Output from getTcvArrStats
 * @returns {string} HTML string
 */
function _buildAccountSegmentationHTML(stats) {
    if (!stats || stats.items.length === 0) return '';

    /* ── Classify accounts into 4 segments ── */
    const multiYear = [];
    const annualRenewal = [];
    const perpetual = [];

    stats.items.forEach(a => {
        if (a.isPerpetual || a.arr === 0) {
            perpetual.push(a);
        } else {
            const ratio = a.arr > 0 ? a.tcv / a.arr : 1;
            if (ratio >= 1.8) multiYear.push(a);
            else annualRenewal.push(a);
        }
    });

    /* ── Concentration Risk: accounts that together reach ≥70% of total ARR ── */
    const sortedByArr = [...stats.items].filter(a => a.arr > 0).sort((a, b) => b.arr - a.arr);
    const concentrationAccounts = [];
    let cumArr = 0;
    const threshold = stats.totalArr * 0.7;
    for (const a of sortedByArr) {
        concentrationAccounts.push(a);
        cumArr += a.arr;
        if (cumArr >= threshold) break;
    }

    /* ── Helpers ── */
    const pct = (val, total) => total > 0 ? ((val / total) * 100).toFixed(0) + '%' : '0%';

    const segmentBadge = (count, color) =>
        `<span style="background:${color}22; color:${color}; font-size:0.62rem; font-weight:800; padding:2px 8px; border-radius:10px; border:1px solid ${color}44;">${count} accounts</span>`;

    const accountRow = (a, accentColor) => {
        const ratio = a.arr > 0 ? (a.tcv / a.arr).toFixed(1) : '—';
        const arrShare = pct(a.arr, stats.totalArr);
        return `
            <div style="display:flex; align-items:center; gap:8px; padding:7px 10px; border-radius:8px; background:#FAFAFA; border:1px solid #F1F5F9; margin-bottom:5px;">
                <div style="width:6px; height:6px; border-radius:50%; background:${accentColor}; flex-shrink:0;"></div>
                <span style="flex:1; font-size:0.75rem; font-weight:700; color:#1e293b; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${a.name}">${a.name}</span>
                <span style="font-size:0.68rem; color:#1e40af; font-weight:700; flex-shrink:0;">$${formatCurrency(a.tcv)}</span>
                <span style="font-size:0.68rem; color:#16a34a; font-weight:700; flex-shrink:0; min-width:52px; text-align:right;">$${formatCurrency(a.arr)}</span>
                <span style="font-size:0.63rem; color:#64748b; min-width:34px; text-align:right;">${a.arr > 0 ? ratio + 'x' : '—'}</span>
            </div>`;
    };

    const segmentCard = (icon, title, subtitle, color, bg, border, badgeCount, actionText, accounts, extra = '') => `
        <div style="background:${bg}; border:1px solid ${border}; border-radius:14px; overflow:hidden;">
            <div style="background:${color}; padding:12px 16px; display:flex; align-items:center; gap:10px;">
                <div style="width:30px; height:30px; background:rgba(255,255,255,0.18); border-radius:8px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                    <i class="fa-solid ${icon}" style="color:white; font-size:0.82rem;"></i>
                </div>
                <div style="flex:1;">
                    <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                        <span style="font-size:0.85rem; font-weight:800; color:white;">${title}</span>
                        <span style="background:rgba(255,255,255,0.25); color:white; font-size:0.6rem; font-weight:700; padding:1px 7px; border-radius:8px;">${badgeCount} accounts</span>
                    </div>
                    <span style="font-size:0.65rem; color:rgba(255,255,255,0.8);">${subtitle}</span>
                </div>
            </div>
            <div style="padding:12px 14px;">
                <div style="max-height:210px; overflow-y:auto; margin-bottom:10px;">
                    <div style="display:flex; justify-content:flex-end; gap:20px; padding:0 10px 4px 0; font-size:0.6rem; color:#94a3b8; font-weight:600; text-transform:uppercase; letter-spacing:0.05em;">
                        <span style="min-width:52px; text-align:right;">TCV</span>
                        <span style="min-width:52px; text-align:right;">ARR</span>
                        <span style="min-width:34px; text-align:right;">Ratio</span>
                    </div>
                    ${accounts.length > 0 ? accounts.map(a => accountRow(a, color)).join('') : `<p style="font-size:0.75rem; color:#94a3b8; text-align:center; padding:12px 0; margin:0;">No accounts in this segment</p>`}
                </div>
                ${extra}
                <div style="background:white; border-radius:8px; padding:8px 12px; border:1px solid ${border}; margin-top:4px;">
                    <div style="display:flex; align-items:flex-start; gap:7px;">
                        <i class="fa-solid fa-arrow-right" style="color:${color}; font-size:0.68rem; margin-top:3px; flex-shrink:0;"></i>
                        <span style="font-size:0.7rem; color:#374151; line-height:1.5;">${actionText}</span>
                    </div>
                </div>
            </div>
        </div>`;

    /* ── Multi-Year Anchors ── */
    const multiYearTotalArr = multiYear.reduce((s, a) => s + a.arr, 0);
    const multiYearCard = segmentCard(
        'fa-anchor',
        'Multi-Year Anchors',
        'TCV/ARR ≥ 1.8x — long-term contracts locked in',
        '#1e40af', '#eff6ff', '#bfdbfe',
        multiYear.length,
        `Low near-term churn risk. However, renewals arrive in bulk at contract end — map expiry dates now and initiate re-engagement at least 6 months in advance.`,
        multiYear,
        multiYear.length > 0 ? `<div style="background:#dbeafe; border-radius:8px; padding:6px 10px; margin-bottom:8px; font-size:0.68rem; color:#1e40af; font-weight:600;">Total ARR from segment: <strong>$${formatCurrency(multiYearTotalArr)}</strong> (${pct(multiYearTotalArr, stats.totalArr)} of portfolio ARR)</div>` : ''
    );

    /* ── Annual Renewers ── */
    const annualTotalArr = annualRenewal.reduce((s, a) => s + a.arr, 0);
    const annualCard = segmentCard(
        'fa-rotate',
        'Annual Renewers',
        'TCV/ARR < 1.8x — 1-year contracts, renews each year',
        '#d97706', '#fffbeb', '#fde68a',
        annualRenewal.length,
        `Highest churn exposure — every account needs a renewal decision this cycle. Prioritize QBRs and multi-year migration offers for any account with ARR > $20K.`,
        annualRenewal,
        annualRenewal.length > 0 ? `<div style="background:#fef3c7; border-radius:8px; padding:6px 10px; margin-bottom:8px; font-size:0.68rem; color:#92400e; font-weight:600;">Total ARR at renewal risk: <strong>$${formatCurrency(annualTotalArr)}</strong> (${pct(annualTotalArr, stats.totalArr)} of portfolio ARR)</div>` : ''
    );

    /* ── Perpetual / One-Time ── */
    const perpetualTcv = perpetual.reduce((s, a) => s + a.tcv, 0);
    const perpetualCard = segmentCard(
        'fa-box-archive',
        'Perpetual / One-Time',
        'ARR = $0 — no recurring revenue contribution',
        '#7c3aed', '#f5f3ff', '#c4b5fd',
        perpetual.length,
        `These accounts contribute $0 to next year's ARR floor. Introduce a SaaS upgrade path or annual maintenance plan to convert even 30% of TCV into recurring revenue.`,
        perpetual,
        perpetual.length > 0 ? `<div style="background:#ede9fe; border-radius:8px; padding:6px 10px; margin-bottom:8px; font-size:0.68rem; color:#6d28d9; font-weight:600;">One-time TCV at risk: <strong>$${formatCurrency(perpetualTcv)}</strong> — must be re-won next period</div>` : ''
    );

    /* ── Concentration Risk ── */
    const concArrSum = concentrationAccounts.reduce((s, a) => s + a.arr, 0);
    const concPct = pct(concArrSum, stats.totalArr);
    const concentrationCard = segmentCard(
        'fa-warning',
        'Concentration Risk',
        `Top ${concentrationAccounts.length} accounts = ≥70% of total ARR`,
        '#dc2626', '#fef2f2', '#fecaca',
        concentrationAccounts.length,
        `Losing any single account from this group would materially impact ARR. Ensure dedicated CSM coverage and executive sponsorship for each of these accounts.`,
        concentrationAccounts,
        `<div style="background:#fee2e2; border-radius:8px; padding:6px 10px; margin-bottom:8px; font-size:0.68rem; color:#991b1b; font-weight:600;">Combined ARR: <strong>$${formatCurrency(concArrSum)}</strong> — ${concPct} of total portfolio ARR concentrated in ${concentrationAccounts.length} accounts</div>`
    );

    return `
        <div class="stat-card" style="background:#FFF; padding:22px; margin-bottom:20px; box-shadow: 0 6px 18px rgba(0,0,0,0.06); border:1px solid #F1F5F9; border-radius:14px; border-top: 4px solid #0f172a;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:18px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
                <h3 style="font-size:1.1rem; font-weight:800; color:#111827; margin:0; display:flex; align-items:center; gap:10px;">
                    <div style="width:36px; height:36px; background:linear-gradient(135deg, #0f172a, #1e40af); border-radius:10px; display:flex; align-items:center; justify-content:center;">
                        <i class="fa-solid fa-layer-group" style="color:white; font-size:0.9rem;"></i>
                    </div>
                    Account Segmentation
                    <span style="background:#f1f5f9; color:#475569; font-size:0.62rem; font-weight:700; padding:3px 10px; border-radius:10px;">Strategic Grouping by Contract Behavior</span>
                </h3>
                <span style="font-size:0.68rem; color:#94a3b8; font-style:italic;">${stats.accountCount} accounts across 4 segments</span>
            </div>
            <div style="display:grid; grid-template-columns: repeat(2, 1fr); gap:16px;">
                ${multiYearCard}
                ${annualCard}
                ${perpetualCard}
                ${concentrationCard}
            </div>
            <div style="margin-top:14px; padding:10px 14px; background:#f8fafc; border-radius:10px; border:1px solid #e2e8f0;">
                <span style="font-size:0.68rem; color:#64748b; line-height:1.6;">
                    <strong style="color:#374151;">How to read:</strong> &nbsp;
                    <span style="color:#1e40af; font-weight:600;">Multi-Year Anchors</span> = safe now, track renewal dates &nbsp;·&nbsp;
                    <span style="color:#d97706; font-weight:600;">Annual Renewers</span> = recurring but fragile, needs active CSM &nbsp;·&nbsp;
                    <span style="color:#7c3aed; font-weight:600;">Perpetual</span> = no ARR, convert to SaaS &nbsp;·&nbsp;
                    <span style="color:#dc2626; font-weight:600;">Concentration</span> = cannot afford to lose these accounts
                </span>
            </div>
        </div>
    `;
}

/* ═══════════════════════════════════════════════════════════════
   DEAL LOST
   ═══════════════════════════════════════════════════════════════ */

/**
 * Render the Deal Lost analytics dashboard.
 * @param {Object} stats - from getDealLostStats
 * @param {string|null} filterCountry
 * @param {Object} uniqueValues
 * @returns {string} HTML
 */
export function getDealLostHTML(stats, filterCountry, uniqueValues) {
    const country = filterCountry || 'All';
    const escape = (str) => String(str || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
    const fmtDate = (d) => d ? d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' }) : '-';

    const reasonColor = (reason) => {
        const r = String(reason || '').toLowerCase();
        if (r.includes('price') || r.includes('budget')) return '#EF4444';
        if (r.includes('competitor')) return '#F97316';
        if (r.includes('no decision')) return '#9CA3AF';
        if (r.includes('technical')) return '#0EA5E9';
        if (r.includes('timing')) return '#A855F7';
        if (r.includes('internal')) return '#F59E0B';
        if (r.includes('ghosted')) return '#6B7280';
        return '#6366F1';
    };

    const filterBar = `
        <div class="stat-card" style="display:flex; flex-wrap:wrap; gap:20px; padding:18px; background:#FFFFFF; border:1px solid #F3F4F6; margin-bottom:24px;">
            <div style="display:flex; flex-direction:column; gap:8px;">
                <label style="font-size:0.8rem; color:#6B7280; font-weight:600; text-transform:uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right:6px;"></i>Country</label>
                <select id="deallost-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 12px; border-radius:6px; width:180px;">
                    ${Array.from(uniqueValues.countries).map(c => `<option value="${c}" ${country === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
            </div>
        </div>`;

    if (stats.totalDeals === 0) {
        return `
            ${filterBar}
            <div class="stat-card" style="padding:48px; background:#FFFFFF; border:1px solid #F3F4F6; text-align:center;">
                <div style="display:inline-flex; align-items:center; justify-content:center; width:72px; height:72px; border-radius:50%; background:rgba(239,68,68,0.1); color:#EF4444; font-size:2rem; margin-bottom:18px;">
                    <i class="fa-solid fa-circle-xmark"></i>
                </div>
                <h2 style="font-size:1.4rem; font-weight:700; color:#111827; margin:0 0 8px;">No Lost Deals Recorded</h2>
                <p style="color:#6B7280; font-size:0.9rem; max-width:520px; margin:0 auto;">
                    Once a row in the <strong>DEAL LOST</strong> sheet has an Account Name, Lost Date, or Deal Amount,
                    it will be analyzed here. The 8 lost-reason categories will then be broken down by country,
                    industry, and competitor.
                </p>
            </div>`;
    }

    const topReasonName = stats.topReason ? stats.topReason.name : '-';
    const topReasonShare = stats.topReason ? Math.round((stats.topReason.count / stats.totalDeals) * 100) : 0;

    const recent = stats.entries.slice(0, 20);
    const thStyle = `padding:10px 14px; color:#6B7280; font-weight:600; font-size:0.78rem; white-space:nowrap; text-align:left;`;

    const rowsHtml = recent.map((r, i) => `
        <tr style="border-bottom:1px solid #E5E7EB; background:${i % 2 === 0 ? '#FAFAFA' : 'transparent'};">
            <td style="padding:10px 14px; color:#9CA3AF; font-size:0.78rem;">${i + 1}</td>
            <td style="padding:10px 14px; color:#374151; font-size:0.8rem;">${escape(r.country) || '-'}</td>
            <td style="padding:10px 14px; font-weight:600; color:#111827; font-size:0.8rem;">${escape(r.account)}</td>
            <td style="padding:10px 14px; color:#374151; font-size:0.78rem;">${escape(r.industry)}</td>
            <td style="padding:10px 14px; color:#374151; font-size:0.78rem;">${escape(r.partner)}</td>
            <td style="padding:10px 14px; color:#374151; font-size:0.78rem;">${escape(r.competitor) || '-'}</td>
            <td style="padding:10px 14px; text-align:right; font-weight:700; color:#111827; font-size:0.8rem;">${r.amount > 0 ? '$' + formatCurrency(r.amount) : '-'}</td>
            <td style="padding:10px 14px; text-align:center;">
                <span style="background:${reasonColor(r.reason)}1A; color:${reasonColor(r.reason)}; padding:3px 10px; border-radius:6px; font-weight:700; font-size:0.7rem; text-transform:uppercase; white-space:nowrap;">${escape(r.reason)}</span>
            </td>
            <td style="padding:10px 14px; text-align:center; color:#374151; font-size:0.78rem;">${fmtDate(r.lostDate)}</td>
            <td style="padding:10px 14px; max-width:260px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:#6B7280; font-size:0.78rem;" title="${escape(r.desc)}">${escape(r.desc) || '-'}</td>
        </tr>
    `).join('');

    return `
        ${filterBar}

        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(230px, 1fr)); gap:20px; margin-bottom:30px;">
            <div class="stat-card highlight-card" style="background:#FEF2F2; border:1px solid rgba(239,68,68,0.2); padding:24px; border-left:5px solid #EF4444;">
                <div class="stat-icon" style="background:rgba(239,68,68,0.15); color:#EF4444; width:56px; height:56px; font-size:1.5rem;"><i class="fa-solid fa-circle-xmark"></i></div>
                <div>
                    <h3 style="color:#DC2626; font-size:0.8rem; text-transform:uppercase; font-weight:700;">Lost Deals</h3>
                    <h2 style="color:#111827; font-size:2.2rem; font-weight:800; margin:0;">${stats.totalDeals} <span style="font-size:1rem; font-weight:400; opacity:0.7;">Deals</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background:#FFF7ED; border:1px solid rgba(249,115,22,0.25); padding:24px; border-left:5px solid #F97316;">
                <div class="stat-icon" style="background:rgba(249,115,22,0.15); color:#F97316; width:56px; height:56px; font-size:1.5rem;"><i class="fa-solid fa-money-bill-trend-up"></i></div>
                <div>
                    <h3 style="color:#C2410C; font-size:0.8rem; text-transform:uppercase; font-weight:700;">Lost Value (USD)</h3>
                    <h2 style="color:#111827; font-size:2.2rem; font-weight:800; margin:0;">$ ${formatCurrency(stats.totalAmount)}</h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background:#EFF6FF; border:1px solid rgba(0,122,255,0.2); padding:24px; border-left:5px solid #007AFF;">
                <div class="stat-icon" style="background:rgba(0,122,255,0.15); color:#007AFF; width:56px; height:56px; font-size:1.5rem;"><i class="fa-solid fa-coins"></i></div>
                <div>
                    <h3 style="color:#007AFF; font-size:0.8rem; text-transform:uppercase; font-weight:700;">Avg Deal Size</h3>
                    <h2 style="color:#111827; font-size:2.2rem; font-weight:800; margin:0;">$ ${formatCurrency(stats.avgDealSize)}</h2>
                    ${stats.avgDaysInPipeline !== null ? `<p style="color:#2563EB; font-size:0.72rem; margin:4px 0 0; font-weight:500;">avg ${stats.avgDaysInPipeline} days in pipeline</p>` : ''}
                </div>
            </div>
            <div class="stat-card highlight-card" style="background:#FDF2FF; border:1px solid rgba(168,85,247,0.25); padding:24px; border-left:5px solid #A855F7;">
                <div class="stat-icon" style="background:rgba(168,85,247,0.15); color:#A855F7; width:56px; height:56px; font-size:1.5rem;"><i class="fa-solid fa-flag"></i></div>
                <div>
                    <h3 style="color:#A855F7; font-size:0.8rem; text-transform:uppercase; font-weight:700;">Top Lost Reason</h3>
                    <h2 style="color:#111827; font-size:1.4rem; font-weight:800; margin:0; line-height:1.25;">${escape(topReasonName)}</h2>
                    <p style="color:#7C3AED; font-size:0.72rem; margin:4px 0 0; font-weight:500;">${topReasonShare}% of all losses</p>
                </div>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding:24px; margin-bottom:30px; background:#FFFFFF; border:1px solid #F3F4F6; display:block;">
            <h3 style="font-size:1.1rem; font-weight:700; color:#111827; margin-bottom:6px;">Monthly Lost Deals (${new Date().getFullYear()})</h3>
            <p style="font-size:0.78rem; color:#6B7280; margin:0 0 16px;">Bars: deal count · Line: lost value (USD)</p>
            <div style="position:relative; height:340px;"><canvas id="deallost-monthly-chart"></canvas></div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:24px; margin-bottom:30px;">
            <div class="stat-card highlight-card" style="padding:20px; display:flex; flex-direction:column;">
                <h4 style="font-size:0.85rem; color:#111827; margin-bottom:16px;"><i class="fa-solid fa-chart-pie" style="margin-right:8px; color:#EF4444;"></i>Lost Reason Breakdown</h4>
                <div style="position:relative; flex:1; min-height:280px;"><canvas id="deallost-reason-chart"></canvas></div>
            </div>
            <div class="stat-card highlight-card" style="padding:20px; display:flex; flex-direction:column;">
                <h4 style="font-size:0.85rem; color:#111827; margin-bottom:16px;"><i class="fa-solid fa-earth-americas" style="margin-right:8px; color:#0EA5E9;"></i>Lost Value by Country</h4>
                <div style="position:relative; flex:1; min-height:280px;"><canvas id="deallost-country-chart"></canvas></div>
            </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:24px; margin-bottom:30px;">
            <div class="stat-card highlight-card" style="padding:20px; display:flex; flex-direction:column;">
                <h4 style="font-size:0.85rem; color:#111827; margin-bottom:16px;"><i class="fa-solid fa-industry" style="margin-right:8px; color:#A855F7;"></i>Lost Deals by Industry</h4>
                <div style="position:relative; flex:1; min-height:280px;"><canvas id="deallost-industry-chart"></canvas></div>
            </div>
            <div class="stat-card highlight-card" style="padding:20px; display:flex; flex-direction:column;">
                <h4 style="font-size:0.85rem; color:#111827; margin-bottom:16px;"><i class="fa-solid fa-trophy" style="margin-right:8px; color:#F97316;"></i>Top Competitors (by lost value)</h4>
                <div style="position:relative; flex:1; min-height:280px;">${stats.sortedCompetitors.length > 0 ? '<canvas id="deallost-competitor-chart"></canvas>' : '<div style="display:flex; align-items:center; justify-content:center; height:100%; color:#9CA3AF; font-size:0.85rem;">No competitor data recorded yet.</div>'}</div>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding:24px; display:block;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:20px;">
                <div>
                    <h3 style="font-size:1.05rem; font-weight:700; color:#111827; margin:0;">Recent Lost Deals</h3>
                    <p style="font-size:0.75rem; color:#6B7280; margin:4px 0 0;">Showing ${recent.length} of ${stats.totalDeals} deals</p>
                </div>
            </div>
            <div style="overflow-x:auto;">
                <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                    <thead>
                        <tr style="background:#F3F4F6; border-bottom:2px solid #E5E7EB;">
                            <th style="${thStyle}">#</th>
                            <th style="${thStyle}">Country</th>
                            <th style="${thStyle}">Account</th>
                            <th style="${thStyle}">Industry</th>
                            <th style="${thStyle}">Partner</th>
                            <th style="${thStyle}">Competitor</th>
                            <th style="${thStyle} text-align:right;">Amount (USD)</th>
                            <th style="${thStyle} text-align:center;">Reason</th>
                            <th style="${thStyle} text-align:center;">Lost Date</th>
                            <th style="${thStyle}">Notes</th>
                        </tr>
                    </thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            </div>
        </div>
    `;
}
