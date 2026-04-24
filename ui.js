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

window.selectQuarter = function (element) {
    const quarter = element.getAttribute('data-q');
    const deals = JSON.parse(element.getAttribute('data-deals'));
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

    let rowsHtml = deals.map((d, index) => `
        <tr style="border-bottom: 1px solid #F3F4F6; transition: background 0.2s;">
            <td style="padding: 16px 24px; color: #94A3B8; font-weight: 700; width: 60px; font-family: monospace;">${String(index + 1).padStart(2, '0')}</td>
            <td style="padding: 16px 24px; color: #1E293B; font-weight: 700; font-size: 0.95rem;">${d.n}</td>
            <td style="padding: 16px 24px; text-align: right; color: #10B981; font-weight: 800; font-size: 1.1rem; letter-spacing: -0.02em;">$${d.a}</td>
        </tr>
    `).join('');

    if (deals.length === 0) {
        rowsHtml = '<tr><td colspan="3" style="padding: 60px 20px; text-align: center; color: #94A3B8; font-style: italic;">No active deals found for this quarter.</td></tr>';
    }

    container.innerHTML = `
        <div style="background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.08);">
            <div style="background: linear-gradient(135deg, #10B981 0%, #059669 100%); padding: 20px 28px; display: flex; justify-content: space-between; align-items: center; color: white; border-bottom: 1px solid rgba(0,0,0,0.05);">
                <div style="display: flex; align-items: center; gap: 14px;">
                    <div style="width: 40px; height: 40px; background: rgba(255,255,255,0.2); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem;">
                        <i class="fa-solid fa-list-check"></i>
                    </div>
                    <div>
                        <h3 style="margin: 0; font-size: 1.2rem; font-weight: 800; letter-spacing: -0.01em;">${quarter} Detailed Pipeline</h3>
                        <p style="margin: 2px 0 0 0; opacity: 0.8; font-size: 0.8rem; font-weight: 500;">Complete breakdown of weighted pipeline value for the period</p>
                    </div>
                </div>
                <div style="background: rgba(255,255,255,0.2); backdrop-filter: blur(4px); padding: 8px 16px; border-radius: 12px; font-size: 0.9rem; font-weight: 800; border: 1px solid rgba(255,255,255,0.1);">
                    ${deals.length} DEALS
                </div>
            </div>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; text-align: left;">
                    <thead>
                        <tr style="background: #F8FAFC; border-bottom: 2px solid #F1F5F9;">
                            <th style="padding: 16px 24px; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em; width: 80px;">NO.</th>
                            <th style="padding: 16px 24px; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em;">CRM DEAL NAME</th>
                            <th style="padding: 16px 24px; text-align: right; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em;">WEIGHTED PIPELINE VALUE (USD)</th>
                        </tr>
                    </thead>
                    <tbody style="background: #FFFFFF;">
                        ${rowsHtml}
                    </tbody>
                </table>
            </div>
        </div>
    `;
    container.style.display = 'block';

    setTimeout(() => {
        const yOffset = -20;
        const y = container.getBoundingClientRect().top + window.pageYOffset + yOffset;
        window.scrollTo({ top: y, behavior: 'smooth' });
    }, 50);
};


export function getServiceAnalysisHTML(stats, filterCountry = 'All') {
    injectServiceAnalysisStyles();

    let html = `
        <div class="stat-card" style="display:flex; align-items:center; gap:12px; padding: 10px 16px; background: #FFFFFF; border: 1px solid rgba(99, 102, 241, 0.2); border-left: 4px solid #6366f1; margin-bottom: 20px;">
            <label style="font-size:0.8rem; color:#6366f1; font-weight:700; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right: 8px;"></i>Select Country</label>
            <select id="csm-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:6px 12px; border-radius:8px; width: 200px; font-size: 0.85rem;">
                ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
            </select>
            <span style="font-size: 0.72rem; color: #64748b; margin-left: auto;">Metrics for ${filterCountry || 'All Regions'}</span>
        </div>
    `;

    if (!stats) {
        return html + '<p style="padding:40px; text-align:center; color:#6B7280;">No active service data found for the selected country.</p>';
    }

    /* ── Customer Overview Section ── */
    html += `
        <div style="display:grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 20px;">
            <div class="stat-card" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding:20px; border-radius:16px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 24px rgba(102, 126, 234, 0.25);">
                <div style="position:absolute; top:-20px; right:-20px; width:100px; height:100px; background:rgba(255,255,255,0.08); border-radius:50%;"></div>
                <div style="position:absolute; bottom:-30px; right:30px; width:60px; height:60px; background:rgba(255,255,255,0.05); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                    <div style="width:44px; height:44px; background:rgba(255,255,255,0.15); border-radius:12px; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(4px);">
                        <i class="fa-solid fa-users" style="font-size:1.2rem;"></i>
                    </div>
                    <div>
                        <h3 style="font-size:0.72rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; opacity:0.85; margin:0;">Total End Users</h3>
                    </div>
                </div>
                <h2 style="font-size:2.4rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em;">${stats.totalEndUsers}</h2>
                <div style="display:flex; gap:16px; margin-top:12px; font-size:0.78rem;">
                    <span style="display:flex; align-items:center; gap:4px;"><span style="width:8px; height:8px; background:#34d399; border-radius:50%; display:inline-block;"></span> Active: ${stats.activeCount}</span>
                    <span style="display:flex; align-items:center; gap:4px;"><span style="width:8px; height:8px; background:rgba(255,255,255,0.4); border-radius:50%; display:inline-block;"></span> Inactive: ${stats.inactiveCount}</span>
                </div>
            </div>

            <div class="stat-card" style="background: linear-gradient(135deg, ${stats.criticalCount > 0 ? '#f43f5e 0%, #e11d48 100%' : stats.expiringCount > 0 ? '#f59e0b 0%, #d97706 100%' : '#10b981 0%, #059669 100%'}); padding:20px; border-radius:16px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 24px ${stats.criticalCount > 0 ? 'rgba(244, 63, 94, 0.25)' : stats.expiringCount > 0 ? 'rgba(245, 158, 11, 0.25)' : 'rgba(16, 185, 129, 0.25)'};">
                <div style="position:absolute; top:-20px; right:-20px; width:100px; height:100px; background:rgba(255,255,255,0.08); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                    <div style="width:44px; height:44px; background:rgba(255,255,255,0.15); border-radius:12px; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(4px);">
                        <i class="fa-solid fa-clock-rotate-left" style="font-size:1.2rem;"></i>
                    </div>
                    <div>
                        <h3 style="font-size:0.72rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; opacity:0.85; margin:0;">Expiring Within 6 Months</h3>
                    </div>
                </div>
                <h2 style="font-size:2.4rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em;">${stats.expiringCount}</h2>
                <div style="display:flex; gap:12px; margin-top:12px; font-size:0.75rem; flex-wrap:wrap;">
                    <span style="display:flex; align-items:center; gap:4px; background:rgba(255,255,255,0.15); padding:3px 10px; border-radius:20px;"><span style="width:7px; height:7px; background:#fef2f2; border-radius:50%; display:inline-block; box-shadow:0 0 4px rgba(255,255,255,0.5);"></span> ≤30d: ${stats.criticalCount}</span>
                    <span style="display:flex; align-items:center; gap:4px; background:rgba(255,255,255,0.15); padding:3px 10px; border-radius:20px;"><span style="width:7px; height:7px; background:#fef9c3; border-radius:50%; display:inline-block;"></span> ≤90d: ${stats.warningCount}</span>
                    <span style="display:flex; align-items:center; gap:4px; background:rgba(255,255,255,0.15); padding:3px 10px; border-radius:20px;"><span style="width:7px; height:7px; background:#d1fae5; border-radius:50%; display:inline-block;"></span> ≤180d: ${stats.normalExpCount}</span>
                </div>
            </div>
        </div>
    `;

    /* ── Expiring Customers Table ── */
    if (stats.expiringCustomers.length > 0) {
        const expiringRows = stats.expiringCustomers.map((c, i) => {
            const urgencyColors = { critical: { bg: 'rgba(244,63,94,0.08)', text: '#e11d48', badge: '#fecdd3' }, warning: { bg: 'rgba(245,158,11,0.06)', text: '#d97706', badge: '#fef3c7' }, normal: { bg: 'transparent', text: '#374151', badge: '#d1fae5' } };
            const uc = urgencyColors[c.urgency];
            return `<tr style="border-bottom:1px solid #F3F4F6; background:${uc.bg}; transition:background 0.2s;">
                <td style="padding:10px 12px; font-weight:700; color:#1e293b; font-size:0.82rem;">${c.name}</td>
                <td style="padding:10px 12px; color:#4b5563; font-size:0.8rem;">${c.country}</td>
                <td style="padding:10px 12px;"><span style="padding:3px 10px; border-radius:12px; font-size:0.68rem; font-weight:700; background:${c.status === 'Active' ? 'rgba(16,185,129,0.1)' : 'rgba(107,114,128,0.1)'}; color:${c.status === 'Active' ? '#059669' : '#6b7280'}; text-transform:uppercase;">${c.status}</span></td>
                <td style="padding:10px 12px; font-family:monospace; font-size:0.78rem; color:#64748b;">${c.startDateStr}</td>
                <td style="padding:10px 12px; font-family:monospace; font-size:0.78rem; font-weight:600; color:#1e293b;">${c.endDateStr}</td>
                <td style="padding:10px 12px;"><span style="font-weight:800; color:${uc.text}; font-size:0.82rem; padding:3px 10px; border-radius:8px; background:${uc.badge};">${c.dDay}</span></td>
                <td style="padding:10px 12px; font-weight:600; color:#1e293b; text-align:right; font-size:0.8rem;">$${formatCurrency(c.tcv)}</td>
                <td style="padding:10px 12px; font-weight:600; color:#6366f1; text-align:right; font-size:0.8rem;">$${formatCurrency(c.arr)}</td>
                <td style="padding:10px 12px; font-size:0.75rem; color:#64748b; max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${c.services}</td>
            </tr>`;
        }).join('');

        html += `
        <div class="stat-card" style="background:#FFF; padding:16px; margin-bottom:20px; box-shadow:0 4px 12px rgba(0,0,0,0.06); border:1px solid #F3F4F6; border-radius:12px;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:14px;">
                <h3 style="font-size:1rem; font-weight:800; color:#111827; margin:0; display:flex; align-items:center; gap:8px;">
                    <i class="fa-solid fa-triangle-exclamation" style="color:#f59e0b;"></i> Licenses Expiring Within 6 Months
                    <span style="background:#fef3c7; color:#92400e; font-size:0.7rem; font-weight:700; padding:2px 10px; border-radius:12px;">${stats.expiringCount} customers</span>
                </h3>
            </div>
            <div style="overflow-x:auto;">
                <table style="width:100%; border-collapse:collapse; min-width:900px;">
                    <thead>
                        <tr style="background:#F9FAFB; text-align:left; border-bottom:2px solid #F1F5F9;">
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">End User</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">Country</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">Status</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">Active License</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">End License</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">D-Day</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; text-align:right;">TCV</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; text-align:right;">ARR</th>
                            <th style="padding:10px 12px; font-size:0.7rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">Services</th>
                        </tr>
                    </thead>
                    <tbody>${expiringRows}</tbody>
                </table>
            </div>
        </div>
        `;
    }

    /* ── Service Analysis (existing) ── */
    html += `
        <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 20px;">
            <div class="stat-card" style="border-left: 4px solid #6366f1; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: flex-start;">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">ACTIVE CUSTOMERS</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">${stats.totalCustomers}</h2>
            </div>
            <div class="stat-card" style="border-left: 4px solid #10b981; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: flex-start;">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">MULTI-SERVICE</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">${stats.multiServiceCustomers}</h2>
            </div>
            <div class="stat-card" style="border-left: 4px solid #f59e0b; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: flex-start;">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">UPSELL TARGETS</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">${stats.singleServiceCustomers}</h2>
            </div>
        </div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 20px;">
            <div class="stat-card" style="background:#FFF; padding: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column; align-items: stretch;">
                <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 12px;">Service Combination Ranking</h3>
                <div style="height: 280px;"><canvas id="service-donut-chart"></canvas></div>
            </div>
            <div class="stat-card" style="background:#FFF; padding: 16px; display:flex; flex-direction:column; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); align-items: stretch;">
                <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 12px;">Top Upsell Targets</h3>
                <div style="overflow-y:auto; flex:1; max-height:280px;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                        <thead style="background:#F9FAFB; position:sticky; top:0;"><tr><th style="padding:6px 8px; text-align:left;">End User</th><th style="padding:6px 8px; text-align:left;">Service</th><th style="padding:6px 8px; text-align:right;">TCV</th></tr></thead>
                        <tbody>
                            ${stats.upsellTargets.slice(0, 15).map(t => `<tr style="border-top: 1px solid #F3F4F6;"><td style="padding:6px 8px;">${t.name}</td><td style="padding:6px 8px;"><span style="background:rgba(99,102,241,0.1); color:#6366f1; padding:2px 8px; border-radius:10px;">${t.service}</span></td><td style="padding:6px 8px; text-align:right; color:#10b981; font-weight:600;">$${formatCurrency(t.tcv)}</td></tr>`).join('')}
                        </tbody>
                    </table>
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

export function getOrderSheetHTML(stats, filterCountry = null) {
    const currentYear = new Date().getFullYear();
    return `
        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 20px;">
            <div class="stat-card" style="border-left: 5px solid #0ea5e9; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; min-height: 140px;">
                <h3 style="color:#0ea5e9; font-size:0.75rem; font-weight:700;">ACCUMULATED TCV</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">${formatCurrency(stats.sumLocalTcv)}</h2>
                <div style="font-size: 0.75rem; color: #6B7280; margin-bottom: 8px;">${stats.dealCount} Deals Total</div>
                <div style="flex: 1; position: relative; min-height: 70px;">
                    <canvas id="tcv-yearly-bar"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #6366f1; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; min-height: 140px;">
                <h3 style="color:#6366f1; font-size:0.75rem; font-weight:700;">ACCUMULATED KTCV</h3>
                <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.sumKorTcv)}</h2>
                <div style="font-size: 0.75rem; color: #6B7280; margin-bottom: 8px;">&nbsp;</div>
                <div style="flex: 1; position: relative; min-height: 70px;">
                    <canvas id="ktcv-yearly-bar"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #8b5cf6; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; position: relative; min-height: 120px;">
                <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 8px;">
                    <h3 style="color:#8b5cf6; font-size:0.75rem; font-weight:700;">ARR</h3>
                    <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.sumArr)}</h2>
                </div>
                <div style="flex: 1; height: 80px; margin-top: auto; position: relative;">
                    <canvas id="arr-sparkline"></canvas>
                </div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #a855f7; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); display: flex; flex-direction: column; align-items: stretch; position: relative; min-height: 120px;">
                <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 8px;">
                    <h3 style="color:#a855f7; font-size:0.75rem; font-weight:700;">MRR</h3>
                    <h2 style="font-size:1.6rem; font-weight:800; margin: 4px 0;">US$ ${formatCurrency(stats.sumMrr)}</h2>
                </div>
                <div style="flex: 1; height: 80px; margin-top: auto; position: relative;">
                    <canvas id="mrr-sparkline"></canvas>
                </div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #f59e0b; display: flex; flex-direction: column; align-items: stretch;">
                <h3 style="color:#f59e0b; font-size:0.75rem; font-weight:700; margin-bottom: 8px;">QUARTERLY TCV (${currentYear})</h3>
                <div style="height:160px; position:relative;"><canvas id="quarterly-tcv-bar"></canvas></div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #10b981; display: flex; flex-direction: column; align-items: stretch;">
                <h3 style="color:#10b981; font-size:0.75rem; font-weight:700; margin-bottom: 8px;">YEARLY TCV GROWTH</h3>
                <div style="height:160px; position:relative;"><canvas id="tcv-growth-chart"></canvas></div>
            </div>
            ${(!filterCountry || filterCountry === 'All') ? `
            <div class="stat-card" style="grid-column: 1 / -1; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #6366f1;">
                <h3 style="color:#6366f1; font-size:0.75rem; font-weight:700; margin-bottom: 12px;">ACCUMULATED KTCV / COUNTRY</h3>
                <div style="display: flex; gap: 32px; align-items: center;">
                    <div style="position: relative; width: 180px; height: 180px; flex-shrink: 0;">
                        <canvas id="country-tcv-donut"></canvas>
                    </div>
                    <div id="country-tcv-legend" style="flex: 1;"></div>
                </div>
            </div>` : ''}
            ${(!filterCountry || filterCountry === 'All') ? `
            <div class="stat-card" style="grid-column: 1 / -1; background:#FFF; padding:16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #f97316;">
                <h3 style="color:#f97316; font-size:0.75rem; font-weight:700; margin-bottom: 12px;">YoY KTCV GROWTH BY COUNTRY</h3>
                <div style="height: 220px; position: relative;">
                    <canvas id="country-yoy-bar"></canvas>
                </div>
            </div>` : ''}
        </div>
    `;
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
                <span style="color: var(--text-muted);">PIPELINE</span>
                <span style="color: #34C759; font-weight: 600;">$${formatCurrency(values.amount)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.7rem; margin-bottom: 2px;">
                <span style="color: var(--text-muted);">WEIGHTED PIPELINE VALUE</span>
                <span style="color: #007AFF; font-weight: 600;">$${formatCurrency(values.weighted)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.7rem;">
                <span style="color: #ef4444; opacity: 0.8;">TCV (${currentYear})</span>
                <span style="color: #ef4444; font-weight: 600;">$${formatCurrency(values.tcv || 0)}</span>
            </div>
        </div>
    `).join('');

    const quarterlyItemsHtml = stats.sortedQuarterly.map(([q, qData]) => {
        const countryEntries = Object.entries(qData.countries);
        const qTotalAmount = countryEntries.reduce((acc, curr) => acc + curr[1].amount, 0);
        const qTotalWeighted = countryEntries.reduce((acc, curr) => acc + curr[1].weighted, 0);
        const qTotalTcv = countryEntries.reduce((acc, curr) => acc + (curr[1].tcv || 0), 0);
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
                        <span style="color: var(--text-muted);">PIPELINE</span>
                        <span style="color: #34C759;">$${formatCurrency(values.amount)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem; margin-bottom: 2px;">
                        <span style="color: var(--text-muted);">WEIGHTED PIPELINE VALUE</span>
                        <span style="color: #007AFF;">$${formatCurrency(values.weighted)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem;">
                        <span style="color: #ef4444; opacity: 0.8;">TCV (${currentYear})</span>
                        <span style="color: #ef4444; font-weight: 600;">$${formatCurrency(values.tcv || 0)}</span>
                    </div>
                </div>
            `).join('');

        const dealListJson = JSON.stringify(qData.deals.slice(0, 50).map(d => ({ n: d.name, a: formatCurrency(d.weighted) })));

        return `
            <div class="quarter-card" 
                 data-q="${q}" 
                 data-deals='${dealListJson.replace(/'/g, "&apos;")}'
                 style="display: flex; flex-direction: column; padding: 10px; background: #F9FAFB; border-radius: 8px; border-top: 3px solid #10b981; cursor: pointer; transition: all 0.2s;"
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
                            <span style="font-size: 0.6rem; color: var(--text-secondary); text-transform: uppercase;">PIPELINE</span>
                            <span style="font-size: 0.85rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalAmount)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px;">
                            <span style="font-size: 0.6rem; color: var(--text-secondary); text-transform: uppercase;">WEIGHTED PIPELINE VALUE</span>
                            <span style="font-size: 0.85rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalWeighted)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 6px;">
                            <span style="font-size: 0.6rem; color: #ef4444; text-transform: uppercase;">TCV (${currentYear})</span>
                            <span style="font-size: 0.85rem; color: #ef4444; font-weight: 800;">$${formatCurrency(qTotalTcv)}</span>
                        </div>
                    </div>
                </div>
                ${!filterCountry ? `
                <div style="background: rgba(0,0,0,0.05); border-radius: 8px; padding: 10px; margin-top: 8px;">
                    <div style="max-height: 120px; overflow-y: auto;">
                        ${countryBreakdown}
                    </div>
                </div>
                ` : ''}
            </div>
        `;
    }).join('');

    injectPipelineTooltipStyles();

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
                
                <div style="display: grid; grid-template-columns: 1fr; gap: 20px; align-items: start;">
                    <!-- Quarterly Grid and Pie Chart in a flex/grid container -->
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 20px;">
                        <!-- Left: Quarter Cards -->
                        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px;">
                            ${quarterlyItemsHtml}
                        </div>
                        
                        <!-- Right: Pie Chart Card -->
                        <div class="stat-card" style="background: #FFFFFF; padding: 20px; border-radius: 12px; border: 1px solid rgba(16, 185, 129, 0.1); box-shadow: 0 4px 12px rgba(0,0,0,0.03); display: flex; flex-direction: column; height: 100%; min-height: 400px;">
                            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 15px;">
                                <h3 style="font-size: 0.9rem; font-weight: 700; color: #111827; margin: 0;">Pipeline % by Quarter</h3>
                                <span style="font-size: 0.7rem; color: #6B7280; background: #F3F4F6; padding: 2px 8px; border-radius: 10px; font-weight: 600;">VALUE BASE</span>
                            </div>
                            <div style="flex: 1; position: relative; display: flex; align-items: center; justify-content: center;">
                                <canvas id="pipeline-quarter-pie-chart" style="max-height: 320px;"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div id="pipeline-selected-quarter-container" style="margin-top: 20px; display: none;"></div>
            </div>
            <div id="pipeline-quarter-tooltip" class="pipeline-tooltip" style="width: 280px; pointer-events: none;"></div>

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

    const rankingRowsHtml = stats.sortedP.slice(0, 10).map((p, idx) => `
        <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
            <td style="padding: 8px 12px; font-weight: 800; color: ${idx < 3 ? '#fbbf24' : '#94a3b8'}; width: 40px;">
                ${idx + 1}${idx < 3 ? ' <i class="fa-solid fa-crown" style="font-size: 0.65rem; margin-left: 2px;"></i>' : ''}
            </td>
            <td style="padding: 8px 12px; color: #111827; font-weight: 600;">${p.name}</td>
            <td style="padding: 8px 12px; text-align: center;">
                <span style="background: rgba(0,122,255,0.1); color: #007AFF; padding: 2px 8px; border-radius: 10px; font-weight: 700; font-size: 0.75rem;">${p.count} POCs</span>
            </td>
            <td style="padding: 8px 12px; text-align: right; color: #34C759; font-weight: 700;">
                $${formatCurrency(p.sumValue)}
            </td>
        </tr>
    `).join('');

    return `
        <div style="grid-column: 1 / -1; display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 20px;">
            ${globalCardHtml}
            ${statsCardsHtml}
        </div>

        ${tabName === 'PARTNER' ? `
        <div class="stat-card" style="grid-column: 1 / -1; display: flex; align-items: center; gap: 12px; padding: 10px 16px; background: #FFFFFF; border: 1px solid rgba(0, 122, 255, 0.2); border-left: 4px solid #007AFF; margin-bottom: 16px;">
            <label style="font-size:0.8rem; color:#007AFF; font-weight:700; text-transform: uppercase;"><i class="fa-solid fa-earth-americas" style="margin-right: 8px;"></i>Select Country</label>
            <select id="partner-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #D1D5DB; padding:6px 12px; border-radius:8px; width: 200px; font-size: 0.85rem;">
                ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
            </select>
            <div style="margin-left: auto; text-align: right;">
                <span style="font-size: 0.72rem; color: #111827; font-weight: 600;">${filterCountry || 'All Regions'}</span>
            </div>
        </div>
        ` : ''}

        <div style="grid-column: 1 / -1; margin-bottom: 20px;">
            <div class="stat-card highlight-card" style="padding: 16px; background: #FFFFFF; border: 1px solid #F3F4F6; display: flex; flex-direction: column; align-items: stretch;">
                <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 16px;">
                    <div class="stat-icon" style="background: rgba(0,122,255,0.1); color: #007AFF; width: 32px; height: 32px;"><i class="fa-solid fa-ranking-star"></i></div>
                    <div>
                        <h3 style="font-size: 1rem; font-weight: 700; color: #111827; margin: 0;">Partner Real-time Status</h3>
                    </div>
                </div>
                <div style="background: #FFFFFF; border-radius: 12px; overflow: hidden; border: 1px solid #F3F4F6;">
                    <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 0.8rem;">
                        <thead>
                            <tr style="background: #F9FAFB; border-bottom: 1px solid #E5E7EB;">
                                <th style="padding: 8px 12px; color: #6B7280; font-weight: 600;">Rank</th>
                                <th style="padding: 8px 12px; color: #6B7280; font-weight: 600;">Partner Name</th>
                                <th style="padding: 8px 12px; color: #6B7280; font-weight: 600; text-align: center;">Running</th>
                                <th style="padding: 8px 12px; color: #6B7280; font-weight: 600; text-align: right;">Value (USD)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rankingRowsHtml}
                            ${stats.sortedP.length === 0 ? '<tr><td colspan="4" style="padding: 20px; text-align: center; color: #9CA3AF;">No data</td></tr>' : ''}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

    `;
}

export function getPartnerNetworkDetailsHTML(stats, filterCountry) {
    const groupedListsHtml = stats.sortedCountries.map(country => {
        const partners = stats.partnerGroups[country];
        const partnerItemsHtml = partners.slice(0, 10).map(p => {
            const name = p[stats.pNameKey] || 'N/A';
            return `
                <div style="padding: 8px 12px; background: rgba(255, 255, 255, 0.03); border-radius: 8px; border: 1px solid #F3F4F6;">
                    <div style="color: #111827; font-weight: 600; font-size: 0.85rem;">${name}</div>
                    <div style="color: #6B7280; font-size: 0.7rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${Object.values(p)[1] || ''}</div>
                </div>
            `;
        }).join('') + (partners.length > 10 ? `<div style="text-align: center; color: #9CA3AF; font-size: 0.7rem; padding-top: 4px;">+ ${partners.length - 10} more</div>` : '');

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

    function buildRows(list, tier) {
        const cfg = {
            critical: { bg: 'rgba(239,68,68,0.07)', border: '#ef4444', badgeBg: 'rgba(239,68,68,0.12)', badgeColor: '#b91c1c', badgeText: 'CRITICAL', dayColor: '#dc2626' },
            warning:  { bg: 'rgba(245,158,11,0.07)', border: '#f59e0b', badgeBg: 'rgba(245,158,11,0.12)', badgeColor: '#92400e', badgeText: 'RENEW SOON', dayColor: '#d97706' },
            overdue:  { bg: 'rgba(107,114,128,0.07)', border: '#9ca3af', badgeBg: 'rgba(107,114,128,0.1)', badgeColor: '#4b5563', badgeText: 'OVERDUE', dayColor: '#6b7280' },
        }[tier];

        return list.slice(0, 5).map(d => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 14px; background:${cfg.bg}; border-radius:8px; border-left:3px solid ${cfg.border}; margin-bottom:6px; transition:transform 0.15s;" onmouseover="this.style.transform='translateX(3px)'" onmouseout="this.style.transform='translateX(0)'">
                <div style="display:flex; flex-direction:column; gap:3px; min-width:0; flex:1;">
                    <span style="font-size:0.82rem; font-weight:700; color:#111827; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${d.name}</span>
                    <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                        ${d.country ? `<span style="font-size:0.68rem; color:#6b7280;"><i class="fa-solid fa-location-dot"></i> ${d.country}</span>` : ''}
                        <span style="font-size:0.7rem; font-weight:700; color:${cfg.dayColor};">${daysLabel(d.daysLeft)}</span>
                        <span style="font-size:0.68rem; color:#6b7280;"><i class="fa-regular fa-calendar-alt"></i> ${d.date}</span>
                    </div>
                </div>
                <div style="display:flex; flex-direction:column; align-items:flex-end; gap:4px; margin-left:10px; flex-shrink:0;">
                    ${d.arr > 0 ? `<span style="font-size:0.78rem; font-weight:800; color:#111827;">${fmtArr(d.arr)} <span style="font-size:0.62rem; font-weight:600; color:#6b7280;">ARR</span></span>` : ''}
                    <span style="font-size:0.6rem; font-weight:800; color:${cfg.badgeColor}; background:${cfg.badgeBg}; padding:2px 7px; border-radius:5px; letter-spacing:0.05em;">${cfg.badgeText}</span>
                </div>
            </div>
        `).join('') + (list.length > 5 ? `<div style="text-align:center; font-size:0.7rem; color:#9ca3af; padding:6px; font-weight:600;">+ ${list.length - 5} more</div>` : '');
    }

    const sections = [];
    if (critical.length > 0) sections.push({ tier: 'critical', list: critical, label: 'Critical — Within 30 Days', icon: 'fa-circle-exclamation', color: '#ef4444', arr: criticalArr });
    if (warning.length > 0)  sections.push({ tier: 'warning',  list: warning,  label: 'Renew Soon — 30–90 Days',    icon: 'fa-clock',             color: '#f59e0b', arr: warningArr });
    if (overdue.length > 0)  sections.push({ tier: 'overdue',  list: overdue,  label: 'Overdue — Revenue Leak',      icon: 'fa-triangle-exclamation', color: '#9ca3af', arr: overdueArr });

    const sectionsHTML = sections.map(s => `
        <div style="flex:1; min-width:260px;">
            <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                <i class="fa-solid ${s.icon}" style="color:${s.color}; font-size:0.85rem;"></i>
                <span style="font-size:0.72rem; font-weight:800; color:${s.color}; text-transform:uppercase; letter-spacing:0.06em;">${s.label}</span>
                ${s.arr > 0 ? `<span style="margin-left:auto; font-size:0.7rem; font-weight:700; color:#374151;">${fmtArr(s.arr)} ARR</span>` : ''}
            </div>
            ${buildRows(s.list, s.tier)}
        </div>
    `).join('');

    const pillCritical = critical.length > 0 ? `<span style="background:rgba(239,68,68,0.1); color:#b91c1c; font-size:0.68rem; font-weight:800; padding:3px 10px; border-radius:20px;"><i class="fa-solid fa-circle-exclamation"></i> ${critical.length} Critical</span>` : '';
    const pillWarning  = warning.length  > 0 ? `<span style="background:rgba(245,158,11,0.1); color:#92400e; font-size:0.68rem; font-weight:800; padding:3px 10px; border-radius:20px;"><i class="fa-solid fa-clock"></i> ${warning.length} Renew Soon</span>` : '';
    const pillOverdue  = overdue.length  > 0 ? `<span style="background:rgba(107,114,128,0.1); color:#4b5563; font-size:0.68rem; font-weight:800; padding:3px 10px; border-radius:20px;"><i class="fa-solid fa-triangle-exclamation"></i> ${overdue.length} Overdue</span>` : '';

    return `
        <div class="stat-card" style="padding:22px; border:1px solid rgba(239,68,68,0.18); background:#fff; border-radius:14px; box-shadow:0 4px 16px rgba(239,68,68,0.06);">
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px; border-bottom:1px solid #fef2f2; padding-bottom:14px; flex-wrap:wrap; row-gap:8px;">
                <div style="background:rgba(239,68,68,0.1); color:#ef4444; width:40px; height:40px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;"><i class="fa-solid fa-shield-halved" style="font-size:1rem;"></i></div>
                <div>
                    <h3 style="margin:0; font-size:0.72rem; color:#ef4444; font-weight:800; text-transform:uppercase; letter-spacing:0.08em;">CHURN RISK ALERT</h3>
                    <h2 style="margin:0; font-size:1rem; font-weight:800; color:#111827;">Contract Renewal Monitor</h2>
                </div>
                <div style="margin-left:auto; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    ${pillCritical}${pillWarning}${pillOverdue}
                    ${totalArrAtRisk > 0 ? `<span style="background:#fef3c7; color:#92400e; font-size:0.68rem; font-weight:800; padding:3px 10px; border-radius:20px; border:1px solid #fde68a;"><i class="fa-solid fa-dollar-sign"></i> ${fmtArr(totalArrAtRisk)} ARR at Risk</span>` : ''}
                </div>
            </div>
            <div style="display:flex; gap:20px; flex-wrap:wrap; align-items:flex-start;">
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
                ${obj.targets.map((t, i) => `
                    <td style="background: rgba(16, 185, 129, 0.05);">
                        <input type="text" class="kpi-target-input" data-idx="${i}" value="${formatCurrency(t)}"
                               ${isAdmin ? `onchange="window.updateKPICell(this, 'targets', ${catId}, ${objId}, ${i})"` : 'readonly'}
                               style="${roStyle}">
                    </td>
                `).join('')}
                <td class="kpi-weight" style="padding: 4px;">
                    <div style="display: flex; align-items: center; justify-content: center; gap: 4px;">
                        <input type="number" style="width: 50px; text-align: right; border: 1px solid transparent; background: rgba(0,0,0,0.02); padding: 6px; border-radius: 6px; font-weight: 700; color: inherit; transition: all 0.2s; outline: none; ${roStyle}"
                               ${isAdmin ? `onfocus="this.style.background='#FFF'; this.style.borderColor='#6366f1';" onblur="this.style.background='rgba(0,0,0,0.02)'; this.style.borderColor='transparent';" onchange="window.updateKPINumber(this, 'weight', ${catId}, ${objId})"` : 'readonly'}
                               value="${obj.weight || 0}">%
                    </div>
                </td>
                <td class="kpi-rate" style="color: ${rateColor}">${rate}%</td>
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
                ${totalAch.map(a => `
                    <td style="background: rgba(99,102,241,0.10); border-top: 2px solid rgba(99,102,241,0.2); padding: 6px 10px;">
                        <input type="text" class="kpi-achieve-input" value="${formatCurrency(a)}" readonly
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
        ? `<span style="color:#FCA5A5; font-size:0.72rem; font-weight:600; margin-left:8px;">(합계: ${Math.round(totalWeight)}% — 100%가 되어야 합니다)</span>`
        : '';

    const modeLabel = isAdmin
        ? `<span style="background:#fef3c7; color:#92400e; font-size:0.75rem; font-weight:700; padding:3px 10px; border-radius:20px; border:1px solid #fde68a;">🔐 Admin</span>`
        : `<span style="background:#ede9fe; color:#5b21b6; font-size:0.75rem; font-weight:700; padding:3px 10px; border-radius:20px; border:1px solid #ddd6fe;">👤 ${currentUser}</span>`;

    const userOptions = [
        `<option value="admin" ${isAdmin ? 'selected' : ''}>🔐 Admin (구조·목표 설정)</option>`,
        ...availableUsers.map(u => `<option value="${u}" ${!isAdmin && currentUser === u ? 'selected' : ''}>👤 ${u}</option>`)
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
                        <td colspan="7" style="text-align: right; padding: 12px 16px; font-weight: 700; font-size: 0.82rem; letter-spacing: 0.06em; text-transform: uppercase;">
                            TOTAL WEIGHT${weightWarning}
                        </td>
                        <td style="text-align: center; font-size: 1.05rem; font-weight: 800; padding: 12px; color: ${weightGap > 0.1 ? '#FCA5A5' : '#86EFAC'};">${Math.round(totalWeight)}%</td>
                        <td style="text-align: center; font-size: 1.05rem; font-weight: 800; color: ${totalRateColor}; padding: 12px;">${Math.round(totalWeightedRate)}%</td>
                    </tr>
                </tfoot>
            </table>
            <div style="margin-top: 20px; padding: 15px; background: rgba(99,102,241,0.05); border-radius: 12px; border-left: 4px solid #6366f1;">
                <p style="margin: 0; font-size: 0.8rem; color: #4F46E5; font-weight: 600;">
                    <i class="fa-solid fa-circle-info"></i> 각 항목의 세부 내용(1~3번)과 분기별 Target/Achievement를 입력 후 'Save Changes'를 클릭하세요. 최종 RATE는 각 항목 가중치(Weight) 기준으로 자동 계산됩니다. (모든 Weight 합계 = 100%)
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
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.15); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-file-invoice-dollar" style="font-size:1rem;"></i></div>
                    <h3 style="font-size:0.68rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; opacity:0.85; margin:0;">Total KOR TCV</h3>
                </div>
                <h2 style="font-size:1.8rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em;">$${formatCurrency(stats.totalTcv)}</h2>
                <div style="font-size:0.7rem; margin-top:8px; opacity:0.75;">${stats.accountCount} accounts</div>
            </div>
            <div class="stat-card" style="background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(30,58,138,0.25);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.06); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.15); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-arrows-rotate" style="font-size:1rem;"></i></div>
                    <h3 style="font-size:0.68rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; opacity:0.85; margin:0;">Total KOR ARR</h3>
                </div>
                <h2 style="font-size:1.8rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em;">$${formatCurrency(stats.totalArr)}</h2>
                <div style="font-size:0.7rem; margin-top:8px; opacity:0.75;">${arrPct}% of TCV is recurring</div>
            </div>
            <div class="stat-card" style="background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(37,99,235,0.25);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.06); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.15); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-chart-column" style="font-size:1rem;"></i></div>
                    <h3 style="font-size:0.68rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; opacity:0.85; margin:0;">Revenue Gap</h3>
                </div>
                <h2 style="font-size:1.8rem; font-weight:800; margin:0; line-height:1; letter-spacing:-0.02em;">$${formatCurrency(stats.totalGap)}</h2>
                <div style="font-size:0.7rem; margin-top:8px; opacity:0.75;">${gapPct}% one-time revenue</div>
            </div>
            <div class="stat-card" style="background: linear-gradient(135deg, #0284c7 0%, #0ea5e9 100%); padding:18px; border-radius:14px; color:white; position:relative; overflow:hidden; box-shadow: 0 8px 20px rgba(14,165,233,0.25);">
                <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:rgba(255,255,255,0.08); border-radius:50%;"></div>
                <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">
                    <div style="width:38px; height:38px; background:rgba(255,255,255,0.18); border-radius:10px; display:flex; align-items:center; justify-content:center;"><i class="fa-solid fa-building" style="font-size:1rem; color:white;"></i></div>
                    <h3 style="font-size:0.68rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; opacity:0.9; margin:0; color:white;">Account Mix</h3>
                </div>
                <div style="display:flex; gap:14px; margin-top:6px; align-items:baseline;">
                    <div><h2 style="font-size:1.5rem; font-weight:800; margin:0; color:#bfdbfe;">${stats.recurringCount}</h2><span style="font-size:0.65rem; opacity:0.85; color:white;">Recurring</span></div>
                    <div><h2 style="font-size:1.5rem; font-weight:800; margin:0; color:#fde68a;">${stats.perpetualCount}</h2><span style="font-size:0.65rem; opacity:0.85; color:white;">Perpetual</span></div>
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
