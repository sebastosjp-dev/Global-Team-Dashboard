/**
 * ui.js - HTML template generators for dashboard components
 */
import { formatCurrency, parseCurrency, sortCountriesByAmount } from './utils.js';
import { CONFIG } from './config.js';
export function injectServiceAnalysisStyles() {
    if (document.getElementById('insight-tooltip-style')) return;
    const style = document.createElement('style');
    style.id = 'insight-tooltip-style';
    style.textContent = `
        .insight-tooltip-wrap { position: relative; display: inline-flex; align-items: center; }
        .insight-info-icon {
            display: inline-flex; align-items: center; justify-content: center;
            width: 18px; height: 18px; border-radius: 50%;
            background: rgba(148,163,184,0.15); color: #9CA3AF;
            font-size: 11px; cursor: pointer; margin-left: 8px;
            border: 1px solid rgba(148,163,184,0.2);
            transition: background 0.2s, color 0.2s;
        }
        .insight-info-icon:hover { background: rgba(99,102,241,0.25); color: #34C759; border-color: #6366f1; }
        .insight-popup {
            display: none; position: absolute; z-index: 9999;
            top: calc(100% + 10px); left: 50%; transform: translateX(-50%);
            width: 300px; background: #FFFFFF;
            border: 1px solid #E5E7EB; border-radius: 12px;
            padding: 16px; box-shadow: 0 8px 32px rgba(0,0,0,0.12);
            pointer-events: none;
        }
        .insight-info-icon:hover + .insight-popup { display: block; }
        .insight-popup-title { font-size: 0.8rem; font-weight: 700; color: #34C759; margin-bottom: 8px; display: flex; align-items: center; gap: 6px; }
        .insight-popup-body { font-size: 0.75rem; color: #6B7280; line-height: 1.6; }
        
        /* Pipeline Quarter Tooltip */
        .pipeline-tooltip {
            position: fixed;
            display: none;
            z-index: 10000;
            background: #FFFFFF;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15);
            border: 1px solid #10b981;
            padding: 0;
            width: 350px;
            max-height: 400px;
            overflow: hidden;
            pointer-events: none;
            transition: opacity 0.2s ease;
        }
        .pipeline-tooltip-header {
            background: #10b981;
            color: white;
            padding: 12px 16px;
            font-weight: 700;
            font-size: 0.9rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .pipeline-tooltip-content {
            padding: 0;
            max-height: 340px;
            overflow-y: auto;
        }
        .pipeline-tooltip-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.75rem;
        }
        .pipeline-tooltip-table th {
            text-align: left;
            padding: 10px 16px;
            background: #F9fafb;
            color: #6B7280;
            font-weight: 700;
            position: sticky;
            top: 0;
            border-bottom: 1px solid #E5E7EB;
        }
        .pipeline-tooltip-table td {
            padding: 10px 16px;
            border-bottom: 1px solid #F3F4F6;
            color: #374151;
        }
        .pipeline-tooltip-table tr:hover {
            background: #F0FDFA;
        }
        .quarter-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.15);
            border-top-width: 4px !important;
        }

        /* KPI Dashboard Styles */
        .kpi-container {
            padding: 24px;
            background: #F8FAFC;
            border-radius: 20px;
            border: 1px solid #E2E8F0;
        }
        .kpi-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 20px rgba(0,0,0,0.05);
        }
        .kpi-header th {
            background: #1E293B;
            color: white;
            padding: 12px;
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.1);
        }
        .kpi-cat-cell {
            writing-mode: vertical-lr;
            transform: rotate(180deg);
            text-align: center;
            font-weight: 800;
            color: white;
            padding: 15px;
            font-size: 0.9rem;
            letter-spacing: 0.1em;
        }
        .kpi-row td {
            padding: 10px 15px;
            border: 1px solid #F1F5F9;
            font-size: 0.85rem;
            vertical-align: middle;
        }
        .kpi-objective { font-weight: 700; color: #1E293B; }
        .kpi-indicator { color: #64748B; font-size: 0.8rem; line-height: 1.4; }
        .kpi-target-input, .kpi-achieve-input {
            width: 100%;
            border: 1px solid transparent;
            background: rgba(0,0,0,0.03);
            padding: 6px 10px;
            border-radius: 6px;
            text-align: right;
            font-weight: 600;
            transition: all 0.2s;
        }
        .kpi-target-input:focus, .kpi-achieve-input:focus {
            background: white;
            border-color: #6366f1;
            box-shadow: 0 0 0 3px rgba(99,102,241,0.1);
            outline: none;
        }
        .kpi-weight { text-align: center; font-weight: 700; color: #64748B; }
        .kpi-rate { 
            text-align: center; 
            font-weight: 800; 
            color: #ef4444; 
            font-size: 1rem;
        }
        .kpi-actions {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
            justify-content: flex-end;
        }
        .btn-kpi {
            padding: 10px 20px;
            border-radius: 10px;
            font-weight: 700;
            font-size: 0.85rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s;
            border: none;
        }
        .btn-save { background: #10B981; color: white; }
        .btn-save:hover { background: #059669; transform: translateY(-2px); }
        .btn-export { background: #6366f1; color: white; }
        .btn-export:hover { background: #4f46e5; transform: translateY(-2px); }
        .btn-reset { background: #94A3B8; color: white; }
        .btn-reset:hover { background: #64748B; }
    `;
    document.head.appendChild(style);
}

export function injectPipelineTooltipStyles() {
    injectServiceAnalysisStyles();
}

// Globally expose tooltip functions
window.showQuarterTooltip = function(event, element) {
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

window.hideQuarterTooltip = function() {
    const tooltip = document.getElementById('pipeline-quarter-tooltip');
    if (tooltip) tooltip.style.display = 'none';
};

window.showPocTooltip = function(event, element, color = '#007AFF') {
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
        tooltip.style.maxHeight = '400px';
        tooltip.style.overflow = 'hidden';
        tooltip.style.pointerEvents = 'none';
        tooltip.style.transition = 'opacity 0.2s';
        document.body.appendChild(tooltip);
    }
    
    let names = [];
    try {
        names = JSON.parse(decodeURIComponent(element.getAttribute('data-names')));
    } catch(e) {}
    const title = element.getAttribute('data-title') || 'POCs';
    
    tooltip.style.border = '1px solid ' + color;
    
    let rowsHtml = names.map((n, i) => `
        <tr style="transition: background 0.2s;">
            <td style="padding: 10px 16px; border-bottom: 1px solid #F3F4F6; color: #374151; font-weight: 600; font-size: 0.75rem;">${i+1}. ${n}</td>
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
        <div style="padding: 0; max-height: 340px; overflow-y: auto;">
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

window.hidePocTooltip = function() {
    const tooltip = document.getElementById('poc-hover-tooltip');
    if (tooltip) {
        tooltip.style.display = 'none';
        tooltip.style.opacity = '0';
    }
};

window.selectQuarter = function(element) {
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
                        <p style="margin: 2px 0 0 0; opacity: 0.8; font-size: 0.8rem; font-weight: 500;">Complete breakdown of expected deals for the period</p>
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
                            <th style="padding: 16px 24px; text-align: right; color: #64748B; font-weight: 800; text-transform: uppercase; font-size: 0.7rem; letter-spacing: 0.1em;">EXPECTED AMOUNT (USD)</th>
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
        window.scrollTo({top: y, behavior: 'smooth'});
    }, 50);
};


export function getServiceAnalysisHTML(stats) {
    if (!stats) return '<p style="padding:40px; text-align:center; color:#6B7280;">No active service data found.</p>';
    injectServiceAnalysisStyles();
    return `
        <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 24px;">
            <div class="stat-card" style="border-left: 4px solid #6366f1; background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">ACTIVE CUSTOMERS</h3>
                <h2 style="font-size:1.8rem; font-weight:800; margin: 8px 0;">${stats.totalCustomers}</h2>
            </div>
            <div class="stat-card" style="border-left: 4px solid #10b981; background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">MULTI-SERVICE</h3>
                <h2 style="font-size:1.8rem; font-weight:800; margin: 8px 0;">${stats.multiServiceCustomers}</h2>
            </div>
            <div class="stat-card" style="border-left: 4px solid #f59e0b; background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="color:#64748b; font-size:0.75rem; font-weight:700;">UPSELL TARGETS</h3>
                <h2 style="font-size:1.8rem; font-weight:800; margin: 8px 0;">${stats.singleServiceCustomers}</h2>
            </div>
        </div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px;">
            <div class="stat-card" style="background:#FFF; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 15px;">Service Combination Ranking</h3>
                <div style="height: 300px;"><canvas id="service-donut-chart"></canvas></div>
            </div>
            <div class="stat-card" style="background:#FFF; padding: 24px; display:flex; flex-direction:column; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 15px;">Top Upsell Targets</h3>
                <div style="overflow-y:auto; flex:1; max-height:300px;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                        <thead style="background:#F9FAFB; position:sticky; top:0;"><tr><th style="padding:8px; text-align:left;">End User</th><th style="padding:8px; text-align:left;">Service</th><th style="padding:8px; text-align:right;">TCV</th></tr></thead>
                        <tbody>
                            ${stats.upsellTargets.slice(0, 15).map(t => `<tr style="border-top: 1px solid #F3F4F6;"><td style="padding:8px;">${t.name}</td><td style="padding:8px;"><span style="background:rgba(99,102,241,0.1); color:#6366f1; padding:2px 8px; border-radius:10px;">${t.service}</span></td><td style="padding:8px; text-align:right; color:#10b981; font-weight:600;">$${formatCurrency(t.tcv)}</td></tr>`).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}


export function getRenewalHTML(filtered) {
    if (filtered.length === 0) {
        return '<div style="padding:40px; text-align:center; color:#6B7280; grid-column:1/-1;">No renewals found in the next 6 months.</div>';
    }

    const headers = ['End User', 'Country', 'Status', 'End License Date', 'D-Day', 'TCV Amount', 'ARR Amount', 'Probability'];
    let tableHtml = `<div class="stat-card" style="grid-column:1/-1; padding:24px; background:#FFF; border: 1px solid #F3F4F6; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
        <h2 style="font-size:1.25rem; font-weight:800; color:#111827; margin-bottom:20px; display:flex; align-items:center; gap:10px;">
            <i class="fa-solid fa-calendar-check" style="color:#ef4444;"></i> License Renewal Schedule (Next 6 Months)
        </h2>
        <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; min-width:1000px;">
                <thead><tr style="background:#F9FAFB; text-align:left; border-bottom:2px solid #F3F4F6;">`;

    headers.forEach(h => { tableHtml += `<th style="padding:14px; font-size:0.75rem; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.05em;">${h}</th>`; });
    tableHtml += `</tr></thead><tbody>`;

    filtered.forEach((row, i) => {
        const dDayColor = row.diffDays <= 30 ? '#ef4444' : (row.diffDays <= 90 ? '#f59e0b' : '#374151');
        const statusBg = row['Status'] === 'Closed' ? 'rgba(52, 199, 89, 0.1)' : 'rgba(0,122,255,0.1)';
        const statusColor = row['Status'] === 'Closed' ? '#34c759' : '#007AFF';

        tableHtml += `<tr style="border-bottom:1px solid #F3F4F6; background:${i % 2 === 0 ? 'transparent' : '#F9FBFF'}; transition: background 0.2s;">
            <td style="padding:14px; font-weight:700; color:#111827;">${row['End User'] || ''}</td>
            <td style="padding:14px; color:#4b5563;">${row['Country'] || ''}</td>
            <td style="padding:14px;"><span style="padding:4px 10px; border-radius:12px; font-size:0.7rem; font-weight:700; background:${statusBg}; color:${statusColor}; text-transform:uppercase;">${row['Status'] || ''}</span></td>
            <td style="padding:14px; color:#4b5563; font-family: monospace;">${row['endDateFormatted']}</td>
            <td style="padding:14px; font-weight:800; color:${dDayColor}">${row['D-Day']}</td>
            <td style="padding:14px; font-weight:600;">$${formatCurrency(row['TCV Amount'])}</td>
            <td style="padding:14px; font-weight:600;">$${formatCurrency(row['ARR Amount'])}</td>
            <td style="padding:14px; font-weight:700; color:#6366f1;">${row['Probability']}%</td>
        </tr>`;
    });

    tableHtml += `</tbody></table></div></div>`;
    return tableHtml;
}

export function getOrderSheetHTML(stats) {
    return `
        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 24px;">
            <div class="stat-card" style="border-left: 5px solid #0ea5e9; background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
                <h3 style="color:#0ea5e9; font-size:0.75rem; font-weight:700;">ACCUMULATED TCV</h3>
                <h2 style="font-size:1.6rem; font-weight:800;">${formatCurrency(stats.sumLocalTcv)}</h2>
                <div style="font-size: 0.75rem; color: #6B7280;">${stats.dealCount} Deals Total</div>
            </div>
            <div class="stat-card" style="border-left: 5px solid #6366f1; background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
                <h3 style="color:#6366f1; font-size:0.75rem; font-weight:700;">ACCUMULATED KTCV</h3>
                <h2 style="font-size:1.6rem; font-weight:800;">US$ ${formatCurrency(stats.sumKorTcv)}</h2>
            </div>
            <div class="stat-card" style="grid-column: 1 / -1; border-left: 5px solid #8b5cf6; background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
                <h3 style="color:#8b5cf6; font-size:0.75rem; font-weight:700;">ARR & MRR PERFORMANCE</h3>
                <div style="display: flex; gap: 60px;">
                    <div><span style="font-size: 0.7rem; color: #6B7280;">ACCUMULATED ARR</span><h2 style="font-size:1.6rem; font-weight:800;">US$ ${formatCurrency(stats.sumArr)}</h2></div>
                    <div><span style="font-size: 0.7rem; color: #6B7280;">ACCUMULATED MRR</span><h2 style="font-size:1.6rem; font-weight:800;">US$ ${formatCurrency(stats.sumMrr)}</h2></div>
                </div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #f59e0b;">
                <h3 style="color:#f59e0b; font-size:0.75rem; font-weight:700;">QUARTERLY TCV (2026)</h3>
                <div style="height:180px; position:relative;"><canvas id="quarterly-tcv-bar"></canvas></div>
            </div>
            <div class="stat-card" style="background:#FFF; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-left: 5px solid #10b981;">
                <h3 style="color:#10b981; font-size:0.75rem; font-weight:700;">YEARLY TCV GROWTH</h3>
                <div style="height:180px; position:relative;"><canvas id="tcv-growth-chart"></canvas></div>
            </div>
        </div>
    `;
}

export function getPipelineHTML(stats, filterCountry, tabName) {
    const pipelineItemsHtml = stats.sortedPipeline.map(([country, values]) => `
        <div style="display: flex; flex-direction: column; padding: 10px; background: #F9FAFB; border-radius: 8px; border-left: 3px solid #10b981;">
            <span style="font-weight: 700; color: #374151; font-size: 0.8rem; margin-bottom: 6px;">${filterCountry ? 'Total Summary' : country}</span>
            <div style="display: flex; justify-content: space-between; font-size: 0.72rem; margin-bottom: 2px;">
                <span style="color: var(--text-muted);">PIPELINE</span>
                <span style="color: #34C759; font-weight: 600;">$${formatCurrency(values.amount)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.72rem;">
                <span style="color: var(--text-muted);">WEIGHTED</span>
                <span style="color: #007AFF; font-weight: 600;">$${formatCurrency(values.weighted)}</span>
            </div>
        </div>
    `).join('');

    const quarterlyItemsHtml = stats.sortedQuarterly.map(([q, qData]) => {
        const countryEntries = Object.entries(qData.countries);
        const qTotalAmount = countryEntries.reduce((acc, curr) => acc + curr[1].amount, 0);
        const qTotalWeighted = countryEntries.reduce((acc, curr) => acc + curr[1].weighted, 0);

        const countryBreakdown = countryEntries
            .sort(sortCountriesByAmount)
            .map(([country, values]) => `
                <div style="margin-top: 8px; padding: 6px; background: rgba(255, 255, 255, 0.03); border-radius: 4px; border: 1px solid #F3F4F6;">
                    <div style="font-weight: 600; color: #111827; font-size: 0.72rem; margin-bottom: 4px; display: flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-location-dot" style="font-size: 0.6rem; color: #34C759;"></i> ${filterCountry ? 'Total' : country}
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem; margin-bottom: 2px;">
                        <span style="color: var(--text-muted);">PIPELINE</span>
                        <span style="color: #34C759;">$${formatCurrency(values.amount)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem;">
                        <span style="color: var(--text-muted);">WEIGHTED</span>
                        <span style="color: #007AFF;">$${formatCurrency(values.weighted)}</span>
                    </div>
                </div>
            `).join('');

        // Prepare deal list for tooltip
        const dealListJson = JSON.stringify(qData.deals.slice(0, 50).map(d => ({ n: d.name, a: formatCurrency(d.amount) })));

        return `
            <div class="quarter-card" 
                 data-q="${q}" 
                 data-deals='${dealListJson.replace(/'/g, "&apos;")}'
                 style="display: flex; flex-direction: column; padding: 12px; background: #F9FAFB; border-radius: 8px; border-top: 3px solid #10b981; cursor: pointer; transition: all 0.2s;"
                 onmouseover="showQuarterTooltip(event, this)" 
                 onmouseout="hideQuarterTooltip()"
                 onclick="selectQuarter(this)">
                <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; border-bottom: 1px solid #E5E7EB; padding-bottom: 8px;">
                    <span style="font-weight: 800; color: #111827; font-size: 0.9rem; margin-top: 2px;">${q}</span>
                    <div style="text-align: right; display: flex; flex-direction: column; gap: 3px;">
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 8px;">
                            <span style="font-size: 0.65rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em;">PIPELINE</span>
                            <span style="font-size: 0.95rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalAmount)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 8px;">
                            <span style="font-size: 0.65rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em;">WEIGHTED</span>
                            <span style="font-size: 0.95rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalWeighted)}</span>
                        </div>
                    </div>
                </div>
                ${!filterCountry ? `
                <div style="background: rgba(0,0,0,0.15); border-radius: 10px; padding: 12px; margin-top: 10px;">
                    <div style="max-height: 150px; overflow-y: auto; padding-right: 2px;">
                        ${countryBreakdown}
                    </div>
                </div>
                ` : ''}
            </div>
        `;
    }).join('');

    injectPipelineTooltipStyles();

    return `
        <div style="padding: 24px; background: #EDFAF1; border-radius: 16px; border: 1px solid rgba(16, 185, 129, 0.15); display: flex; flex-direction: column; gap: 24px;">
            ${tabName === 'PIPELINE' ? `
            <div class="stat-card" style="display:flex; align-items:center; gap:15px; padding: 14px 20px; background: #FFFFFF; border: 1px solid rgba(16, 185, 129, 0.2); border-left: 4px solid #10b981; margin-bottom: 4px;">
                <label style="font-size:0.85rem; color:#34C759; font-weight:700; text-transform: uppercase; letter-spacing: 0.05em;"><i class="fa-solid fa-earth-americas" style="margin-right: 10px;"></i>Select Country</label>
                <select id="pipeline-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 16px; border-radius:8px; width: 220px; font-family: 'Inter', sans-serif; cursor: pointer; font-weight: 500;">
                    ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
                </select>
                <span style="font-size: 0.75rem; color: var(--text-secondary); margin-left: auto;">Showing pipeline metrics for ${filterCountry || 'All Regions'}</span>
            </div>
            ` : ''}

            <div style="background: rgba(52,199,89,0.08); border: 1px solid rgba(16, 185, 129, 0.2); border-radius: 12px; padding: 16px; display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px;">
                <div style="display: flex; align-items: center; gap: 15px;">
                    <div class="stat-icon" style="width: 45px; height: 45px; font-size: 1.3rem; background: rgba(16, 185, 129, 0.2); color: #34C759;"><i class="fa-solid fa-globe"></i></div>
                    <div>
                        <h2 style="font-size: 1.2rem; font-weight: 700; color: #111827; margin: 0;">${filterCountry ? 'Total Pipeline' : 'Global Total Pipeline'}</h2>
                        <p style="font-size: 0.78rem; color: var(--text-secondary); margin: 2px 0 0 0;">${filterCountry ? 'Aggregated metrics' : 'Aggregated metrics across all regions'}</p>
                    </div>
                </div>
                <div style="display: flex; gap: 30px; text-align: right;">
                    <div>
                        <span style="font-size: 0.75rem; color: #34C759; text-transform: uppercase; letter-spacing: 0.05em;">PIPELINE</span>
                        <h2 style="font-size: 1.4rem; font-weight: 800; color: #111827; margin: 0;">US$ ${formatCurrency(stats.globalTotalAmount)}</h2>
                    </div>
                    <div>
                        <span style="font-size: 0.75rem; color: #007AFF; text-transform: uppercase; letter-spacing: 0.05em;">WEIGHTED</span>
                        <h2 style="font-size: 1.4rem; font-weight: 800; color: #111827; margin: 0;">US$ ${formatCurrency(stats.globalTotalWeighted)}</h2>
                    </div>
                </div>
            </div>

            ${!filterCountry ? `
            <div>
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
                    <div class="stat-icon" style="width: 36px; height: 36px; font-size: 1rem; background: rgba(16, 185, 129, 0.15); color: #34C759;"><i class="fa-solid fa-earth-americas"></i></div>
                    <h2 style="font-size: 1rem; font-weight: 600; color: #111827;">Pipeline by Country</h2>
                </div>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px;">
                    ${pipelineItemsHtml}
                </div>
            </div>
            ` : ''}

            <div style="border-top: 1px solid #E5E7EB; pt: 20px; margin-top: 4px; padding-top: 16px;">
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
                    <div class="stat-icon" style="width: 36px; height: 36px; font-size: 1rem; background: rgba(20, 184, 166, 0.15); color: #14b8a6;"><i class="fa-solid fa-calendar-quarter"></i></div>
                    <h2 style="font-size: 1rem; font-weight: 600; color: #111827;">Pipeline by Quarter (Expected Close)</h2>
                </div>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px;">
                    ${quarterlyItemsHtml}
                </div>
                <!-- Container for the full-width deal list -->
                <div id="pipeline-selected-quarter-container" style="margin-top: 24px; display: none;"></div>
            </div>
            <div id="pipeline-quarter-tooltip" class="pipeline-tooltip" style="width: 280px; pointer-events: none;"></div>

            <div style="border-top: 1px solid #E5E7EB; padding-top: 24px;">
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 20px;">
                    <div class="stat-icon" style="width: 36px; height: 36px; font-size: 1rem; background: rgba(99, 102, 241, 0.15); color: #6366f1;"><i class="fa-solid fa-list-check"></i></div>
                    <h2 style="font-size: 1.1rem; font-weight: 700; color: #111827;">Pipeline Summary by Year & Country</h2>
                </div>
                ${Object.entries(stats.pipelineByYearCountry).sort((a, b) => b[0].localeCompare(a[0])).map(([year, countries]) => {
                    const countryEntries = Object.entries(countries).sort(sortCountriesByAmount);
                    const yearTotalAmount = countryEntries.reduce((acc, curr) => acc + curr[1].amount, 0);
                    const yearTotalWeighted = countryEntries.reduce((acc, curr) => acc + curr[1].weighted, 0);

                    const countryRows = countryEntries.map(([country, values]) => `
                        <tr style="border-bottom: 1px solid #F3F4F6; transition: background 0.2s;">
                            <td style="padding: 12px 20px; color: #111827; font-weight: 600;">${country}</td>
                            <td style="padding: 12px 20px; text-align: right; color: #10b981; font-weight: 700;">$${formatCurrency(values.amount)}</td>
                            <td style="padding: 12px 20px; text-align: right; color: #14b8a6; font-weight: 700;">$${formatCurrency(values.weighted)}</td>
                        </tr>
                    `).join('');

                    return `
                        <div style="margin-bottom: 24px; background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                            <div style="background: #F9FAFB; padding: 14px 20px; border-bottom: 1px solid #E5E7EB; display: flex; justify-content: space-between; align-items: center;">
                                <div style="display: flex; align-items: center; gap: 10px;">
                                    <i class="fa-solid fa-calendar-days" style="color: #6366f1;"></i>
                                    <span style="font-weight: 800; color: #111827; font-size: 1rem;">${year} Summary</span>
                                </div>
                                <div style="display: flex; gap: 24px;">
                                    <div style="text-align: right;">
                                        <span style="font-size: 0.65rem; color: #6B7280; text-transform: uppercase; font-weight: 700; letter-spacing: 0.05em;">Year Total</span>
                                        <span style="font-size: 1rem; color: #10b981; font-weight: 800; display: block;">$${formatCurrency(yearTotalAmount)}</span>
                                    </div>
                                    <div style="text-align: right;">
                                        <span style="font-size: 0.65rem; color: #6B7280; text-transform: uppercase; font-weight: 700; letter-spacing: 0.05em; display: block;">Year Weighted</span>
                                        <span style="font-size: 1rem; color: #14b8a6; font-weight: 800;">$${formatCurrency(yearTotalWeighted)}</span>
                                    </div>
                                </div>
                            </div>
                            <div style="overflow-x: auto;">
                                <table style="width: 100%; border-collapse: collapse; font-size: 0.85rem;">
                                    <thead>
                                        <tr style="background: #FAFAFA; border-bottom: 1px solid #E5E7EB;">
                                            <th style="padding: 12px 20px; text-align: left; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">Country</th>
                                            <th style="padding: 12px 20px; text-align: right; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">PIPELINE (USD)</th>
                                            <th style="padding: 12px 20px; text-align: right; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em;">WEIGHTED (USD)</th>
                                        </tr>
                                    </thead>
                                    <tbody>${countryRows}</tbody>
                                </table>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>

            <div style="border-top: 1px solid #E5E7EB; padding-top: 24px;">
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px;">
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <div class="stat-icon" style="background: rgba(34, 197, 94, 0.15); color: #34C759; width: 36px; height: 36px;"><i class="fa-solid fa-chart-line"></i></div>
                        <div>
                            <h3 style="font-size: 1rem; font-weight: 700; color: #111827; margin: 0;">New Influx & Pipeline Analysis (2026 Monthly)</h3>
                            <p style="font-size: 0.75rem; color: var(--text-secondary); margin-top: 2px;">Monthly Expected Close count vs KOR TCV (USD)</p>
                        </div>
                    </div>
                </div>
                <div style="position: relative; height: 320px;">
                    <canvas id="pipeline-influx-chart"></canvas>
                </div>
            </div>
        </div>
    `;
}

export function getPartnerHTML(stats, filterCountry, tabName) {
    const displayCountries = CONFIG.COUNTRIES.filter(c => (!filterCountry || filterCountry === 'All') || c === filterCountry);

    const totalPartners = CONFIG.COUNTRIES.reduce((sum, c) => sum + (stats.counts[c] || 0), 0);

    const globalCardHtml = (!filterCountry || filterCountry === 'All') ? `
        <div class="stat-card" style="margin:0; padding: 20px; background: #FFFFFF; border: 1px solid #10B981; border-radius: 16px; display: flex; align-items: center; gap: 15px; position: relative; overflow: hidden; box-shadow: 0 4px 12px rgba(16, 185, 129, 0.15);">
            <div class="stat-icon" style="width: 48px; height: 48px; min-width: 48px; border-radius: 50%; overflow: hidden; border: 2px solid rgba(255,255,255,0.15); padding: 0; background: #10B981; display: flex; align-items: center; justify-content: center; color: white;">
                <i class="fa-solid fa-earth-americas" style="font-size: 1.5rem;"></i>
            </div>
            <div>
                <div style="display: flex; align-items: center; gap: 6px;">
                    <h4 style="margin: 0; font-size: 0.75rem; color: #10B981; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 800;">GLOBAL</h4>
                </div>
                <div style="display: flex; align-items: baseline; gap: 4px; margin-top: 2px;">
                    <span style="font-size: 1.8rem; font-weight: 800; color: #111827; line-height: 1;">${totalPartners}</span>
                    <span style="font-size: 0.8rem; color: #9CA3AF; font-weight: 500;">Partners</span>
                </div>
            </div>
        </div>
    ` : '';

    const statsCardsHtml = displayCountries.map(c => {
        const count = stats.counts[c] || 0;
        let flagCode = 'un';
        if (c === 'Indonesia') flagCode = 'id';
        else if (c === 'Thailand') flagCode = 'th';
        else if (c === 'Malaysia') flagCode = 'my';
        else if (c === 'USA') flagCode = 'us';
        else if (c === 'Philippines') flagCode = 'ph';
        else if (c === 'Singapore') flagCode = 'sg';
        else if (c === 'Turkey') flagCode = 'tr';

        const flagUrl = `https://flagcdn.com/w160/${flagCode}.png`;
        return `
            <div class="stat-card" style="margin:0; padding: 20px; background: #FFFFFF; border: 1px solid #F3F4F6; border-radius: 16px; display: flex; align-items: center; gap: 15px; position: relative; overflow: hidden;">
                <div class="stat-icon" style="width: 48px; height: 48px; min-width: 48px; border-radius: 50%; overflow: hidden; border: 2px solid rgba(255,255,255,0.15); padding: 0; background: #000;">
                    <img src="${flagUrl}" style="width: 100%; height: 100%; object-fit: cover; transform: scale(1.1);" alt="${c}">
                </div>
                <div>
                    <div style="display: flex; align-items: center; gap: 6px;">
                        <h4 style="margin: 0; font-size: 0.75rem; color: #6B7280; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 700;">${c}</h4>
                    </div>
                    <div style="display: flex; align-items: baseline; gap: 4px; margin-top: 2px;">
                        <span style="font-size: 1.8rem; font-weight: 800; color: #111827; line-height: 1;">${count}</span>
                        <span style="font-size: 0.8rem; color: #9CA3AF; font-weight: 500;">Partners</span>
                    </div>
                </div>
            </div>
        `;
    }).join('');

    const rankingRowsHtml = stats.sortedP.slice(0, 10).map((p, idx) => `
        <tr style="border-bottom: 1px solid rgba(255,255,255,0.04);">
            <td style="padding: 12px 16px; font-weight: 800; color: ${idx < 3 ? '#fbbf24' : '#94a3b8'}; width: 50px;">
                ${idx + 1}${idx < 3 ? ' <i class="fa-solid fa-crown" style="font-size: 0.7rem; margin-left: 4px;"></i>' : ''}
            </td>
            <td style="padding: 12px 16px; color: #111827; font-weight: 600;">${p.name}</td>
            <td style="padding: 12px 16px; text-align: center;">
                <span style="background: rgba(0,122,255,0.15); color: #007AFF; padding: 2px 10px; border-radius: 12px; font-weight: 700;">${p.count} POCs</span>
            </td>
            <td style="padding: 12px 16px; text-align: right; color: #34C759; font-weight: 700;">
                US$ ${formatCurrency(p.sumValue)}
            </td>
        </tr>
    `).join('');

    const groupedListsHtml = stats.sortedCountries.map(country => {
        const partners = stats.partnerGroups[country];
        const partnerItemsHtml = partners.slice(0, 10).map(p => {
            const name = p[stats.pNameKey] || 'N/A';
            return `
                <div style="padding: 10px 15px; background: rgba(255, 255, 255, 0.03); border-radius: 8px; border: 1px solid #F3F4F6; transition: all 0.2s ease;">
                    <div style="color: #111827; font-weight: 600; font-size: 0.9rem;">${name}</div>
                    <div style="color: #6B7280; font-size: 0.75rem; margin-top: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${Object.values(p)[1] || ''}</div>
                </div>
            `;
        }).join('') + (partners.length > 10 ? `<div style="text-align: center; color: #9CA3AF; font-size: 0.75rem; padding-top: 5px;">...and ${partners.length - 10} more</div>` : '');

        return `
            <div style="background: #F9FAFB; border: 1px solid #F3F4F6; border-radius: 16px; padding: 20px; display: flex; flex-direction: column; gap: 12px;">
                <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(255, 255, 255, 0.1); padding-bottom: 10px; margin-bottom: 4px;">
                    <h3 style="color: #111827; font-size: 1rem; font-weight: 700; margin: 0;">${country}</h3>
                    <span style="background: rgba(0,122,255,0.15); color: #007AFF; font-size: 0.7rem; font-weight: 700; padding: 2px 8px; border-radius: 12px;">${partners.length} TOTAL</span>
                </div>
                <div style="display: grid; grid-template-columns: 1fr; gap: 8px;">
                    ${partnerItemsHtml}
                </div>
            </div>
        `;
    }).join('');

    return `
        <div style="grid-column: 1 / -1; display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin-bottom: 25px;">
            ${globalCardHtml}
            ${statsCardsHtml}
        </div>

        ${tabName === 'PARTNER' ? `
        <div class="stat-card" style="grid-column: 1 / -1; display: flex; align-items: center; gap: 15px; padding: 14px 20px; background: rgba(15, 23, 42, 0.4); border: 1px solid rgba(59, 130, 246, 0.2); border-left: 4px solid #3b82f6; margin-bottom: 20px;">
            <label style="font-size:0.85rem; color:#007AFF; font-weight:700; text-transform: uppercase; letter-spacing: 0.05em;"><i class="fa-solid fa-earth-americas" style="margin-right: 10px;"></i>Select Country</label>
            <select id="partner-filter-country" style="background:#F9FAFB; color:#111827; border:1px solid #334155; padding:8px 16px; border-radius:8px; width: 220px; font-family: 'Inter', sans-serif; cursor: pointer; font-weight: 500;">
                ${['All', ...CONFIG.COUNTRIES].map(c => `<option value="${c}" ${(filterCountry || 'All') === c ? 'selected' : ''}>${c}</option>`).join('')}
            </select>
            <div style="margin-left: auto; text-align: right;">
                <span style="font-size: 0.75rem; color: var(--text-secondary); display: block;">Showing partner metrics for</span>
                <span style="font-size: 0.85rem; color: #111827; font-weight: 600;">${filterCountry || 'All Regions'}</span>
            </div>
        </div>
        ` : ''}

        <div style="grid-column: 1 / -1; margin-bottom: 24px;">
            <div class="stat-card highlight-card" style="padding: 24px; background: #FFFFFF; border: 1px solid #F3F4F6;">
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 24px;">
                    <div class="stat-icon" style="background: rgba(0,122,255,0.1); color: #007AFF; width: 36px; height: 36px;"><i class="fa-solid fa-ranking-star"></i></div>
                    <div>
                        <h3 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin: 0;">Partner Real-time Status (Linked with POC Sheet)</h3>
                        <p style="font-size: 0.75rem; color: var(--text-secondary); margin-top: 2px;">Ranking by 'Running' POC Count and Total Estimated Value (USD)</p>
                    </div>
                </div>
                <div style="background: rgba(0,0,0,0.02); border-radius: 12px; overflow: hidden; border: 1px solid #F3F4F6;">
                    <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 0.85rem;">
                        <thead>
                            <tr style="background: #F3F4F6; border-bottom: 1px solid #E5E7EB;">
                                <th style="padding: 12px 16px; color: #6B7280; font-weight: 600;">Rank</th>
                                <th style="padding: 12px 16px; color: #6B7280; font-weight: 600;">Partner Name</th>
                                <th style="padding: 12px 16px; color: #6B7280; font-weight: 600; text-align: center;">Running Count</th>
                                <th style="padding: 12px 16px; color: #6B7280; font-weight: 600; text-align: right;">Estimated Value (USD)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rankingRowsHtml}
                            ${stats.sortedP.length === 0 ? '<tr><td colspan="4" style="padding: 30px; text-align: center; color: #9CA3AF; font-style: italic;">No running POCs found for this selection</td></tr>' : ''}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div style="grid-column: 1 / -1; margin-bottom: 30px;">
            <div style="padding: 30px; background: #F9FAFB; border: 1px solid #E5E7EB; border-radius: 24px; display: flex; flex-direction: column; gap: 30px;">
                <div style="display: flex; align-items: center; gap: 15px;">
                    <div class="stat-icon" style="background: rgba(0,122,255,0.15); color: #007AFF; width: 48px; height: 48px; font-size: 1.3rem;"><i class="fa-solid fa-handshake"></i></div>
                    <div>
                        <h2 style="font-size: 1.5rem; font-weight: 700; color: #111827; margin: 0;">Partner Network Details</h2>
                        <p style="font-size: 0.85rem; color: #6B7280; margin-top: 4px;">Regional distribution and partner counts</p>
                    </div>
                </div>

                ${!filterCountry ? `
                <div class="stat-card" style="margin:0; padding: 24px; background: rgba(0,0,0,0.15); border: 1px solid #F3F4F6; border-radius: 20px;">
                    <h3 style="color: #111827; font-size: 1.1rem; font-weight: 600; margin-bottom: 20px;">Distribution Ranking</h3>
                    <div style="position: relative; height: 350px;">
                        <canvas id="partner-country-chart"></canvas>
                    </div>
                </div>
                ` : ''}

                <div style="display: flex; align-items: center; gap: 12px; margin-top: 10px;">
                    <div class="stat-icon" style="background: rgba(147, 51, 234, 0.2); color: #34C759; width: 40px; height: 40px; font-size: 1rem;"><i class="fa-solid fa-list-ul"></i></div>
                    <h3 style="font-size: 1.2rem; font-weight: 700; color: #111827; margin: 0;">Regional Partner Lists</h3>
                </div>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 16px;">
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
        const items = Object.entries(stats.yearlyCounts[y]).sort(sortCountriesByCount).map(([c, count]) => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:4px 0; border-bottom:1px solid #F9FAFB;">
                <span style="font-size:0.75rem; color:#6B7280;">${filterCountry ? 'Deals' : c}</span>
                <span style="font-size:0.75rem; font-weight:600; color:#374151;">${count}</span>
            </div>
        `).join('');
        yearlyHtml += `<div style="margin-top:12px; border-top:1px solid #F3F4F6; padding-top:8px;"><h4 style="font-size:0.75rem; font-weight:800; color:#6366f1; margin-bottom:4px; text-transform:uppercase;">${y} PERFORMANCE</h4>${items}</div>`;
    });

    return `
        <div class="stat-card" style="padding:20px; background:#FFF; border:1px solid #F3F4F6; max-width:400px; display:block; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px;">
                <div class="stat-icon" style="background:rgba(99,102,241,0.1); color:#6366f1; width:36px; height:36px; font-size:1rem;"><i class="fa-solid fa-handshake"></i></div>
                <div class="stat-details"><h3 style="margin:0; font-size:0.8rem; color:#6B7280;">CLOSED DEALS</h3><h2 style="margin:0; font-size:0.95rem; font-weight:700; color:#111827;">Summary by Country/Year</h2></div>
            </div>
            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:8px;">${totalHtml}</div>
            <div style="max-height:220px; overflow-y:auto; margin-top:12px; padding-right:4px;">${yearlyHtml}</div>
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

export function getPartnerPerformanceHTML() {
    return `
        <div class="stat-card" style="padding: 24px; background: #FFFFFF; border: 1px solid #F3F4F6; grid-column: 1 / -1; margin-bottom: 24px;">
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
    // Monthly Analysis Table
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const monthlyTableRows = months.map((m, i) => {
        const newCount = stats.influxData[i].count;
        const estVal = stats.influxData[i].estimated;
        const wVal = stats.influxData[i].weighted;
        return `
            <tr style="border-bottom: 1px solid #F3F4F6;">
                <td style="padding: 10px; color: #374151; font-weight: 600;">${m}</td>
                <td style="padding: 10px; text-align: center; color: #10b981; font-weight: 700; background: rgba(16, 185, 129, 0.05);">${newCount}</td>
                <td style="padding: 10px; text-align: right; color: #007AFF; font-weight: 600;">$${formatCurrency(estVal)}</td>
                <td style="padding: 10px; text-align: right; color: #6366f1; font-weight: 600;">$${formatCurrency(wVal)}</td>
            </tr>`;
    }).join('');

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
            <div class="stat-card highlight-card" style="padding: 20px; background: #F3F4F6; border: 1px solid #E5E7EB;">
                <h4 style="color: #111827; font-size: 0.95rem; font-weight: 600; margin-bottom: 12px;"><i class="fa-solid fa-table-list" style="color: #ef4444;"></i> Long-term Summary</h4>
                <div style="max-height: 180px; overflow-y: auto;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 0.78rem;">
                        <thead style="background: #F3F4F6; position: sticky; top: 0;">
                            <tr><th style="padding: 8px; text-align: left;">Name</th><th style="padding: 8px; text-align: right; color: #FF3B30;">Days</th></tr>
                        </thead>
                        <tbody>
                            ${stats.staledRunningList.map(r => `<tr><td style="padding: 8px; border-bottom: 1px solid rgba(0,0,0,0.05);">${r.name}</td><td style="padding: 8px; text-align: right; color: #ef4444; font-weight: 700;">${r.days}</td></tr>`).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding: 24px; margin-bottom: 30px; background: #FFFFFF; border: 1px solid #F3F4F6;">
            <h3 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin-bottom: 20px;">Monthly POC Analysis (2026)</h3>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.85rem;">
                    <thead style="background: #F9FAFB; border-bottom: 2px solid #E5E7EB;">
                        <tr>
                            <th style="padding: 10px; text-align: left; color: #6B7280;">Month</th>
                            <th style="padding: 10px; text-align: center; color: #6B7280;">New POC Starts</th>
                            <th style="padding: 10px; text-align: right; color: #6B7280;">Total Estimated (USD)</th>
                            <th style="padding: 10px; text-align: right; color: #6B7280;">Total Weighted (USD)</th>
                        </tr>
                    </thead>
                    <tbody>${monthlyTableRows}</tbody>
                </table>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding: 24px; margin-bottom: 30px; background: #FFFFFF; border: 1px solid #F3F4F6;">
            <h3 style="font-size: 1.1rem; font-weight: 700; color: #111827; margin-bottom: 20px;">Monthly Influx & Pipeline Analysis (2026)</h3>
            <div style="position: relative; height: 350px;"><canvas id="poc-influx-chart"></canvas></div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="padding: 20px; display: flex; flex-direction: column;">
                <h4 style="font-size: 0.85rem; color: #111827; margin-bottom: 16px;"><i class="fa-solid fa-pie-chart" style="margin-right: 8px;"></i>Status Distribution</h4>
                <div style="position: relative; flex: 1;"><canvas id="poc-status-chart"></canvas></div>
            </div>
            <div class="stat-card highlight-card" style="padding: 20px; display: flex; flex-direction: column;">
                <h4 style="font-size: 0.85rem; color: #111827; margin-bottom: 16px;"><i class="fa-solid fa-clock" style="margin-right: 8px;"></i>Aging (100+ Working Days)</h4>
                <div style="position: relative; flex: 1;"><canvas id="poc-aging-chart"></canvas></div>
            </div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="padding: 24px;">
                <h3 style="font-size: 1.05rem; font-weight: 600; margin-bottom: 12px;">Bottleneck Analysis (Avg Working Days)</h3>
                <div style="position: relative; height: 260px;"><canvas id="poc-bottleneck-chart"></canvas></div>
            </div>
            <div class="stat-card highlight-card" style="padding: 24px;">
                <h3 style="font-size: 1.05rem; font-weight: 600; margin-bottom: 12px;">Industry Opportunity Analysis</h3>
                <div style="position: relative; height: 260px;"><canvas id="poc-industry-chart"></canvas></div>
            </div>
        </div>

        <div class="stat-card highlight-card" style="padding: 24px;">
            <h3 style="font-size: 1.05rem; font-weight: 600; margin-bottom: 20px;">Priority Follow-up List</h3>
            <div style="overflow-x: auto;">
                <table style="width: 100%; border-collapse: collapse; font-size: 0.8rem; text-align: left;">
                    <thead><tr style="background: #F3F4F6;">
                        <th style="padding: 12px;">POC Name</th><th style="padding: 12px;">Partner</th><th style="padding: 12px;">W.Days</th><th style="padding: 12px;">Notes</th>
                    </tr></thead>
                    <tbody>
                        ${stats.runningList.slice(0, 15).map(r => `
                            <tr style="border-bottom: 1px solid #E5E7EB;">
                                <td style="padding: 12px; font-weight: 500;">${r.name}</td>
                                <td style="padding: 12px;">${r.partner}</td>
                                <td style="padding: 12px;"><span style="color: ${r.days >= 100 ? '#ef4444' : '#10b981'}; font-weight: 700;">${r.days}</span></td>
                                <td style="padding: 12px; max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${r.notes || '-'}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

export function getEventHTML(stats) {
    if (!stats) return '';
    return `
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px;">
            <div style="display: flex; align-items: center; gap: 15px;">
                <div class="stat-icon" style="background: rgba(245, 158, 11, 0.15); color: #f59e0b; width: 48px; height: 48px; font-size: 1.3rem;"><i class="fa-solid fa-calendar-check"></i></div>
                <div><h2 style="font-size: 1.6rem; font-weight: 700; color: #111827; margin: 0;">2026 Event Performance Analytics</h2></div>
            </div>
        </div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; margin-bottom: 24px;">
            <div class="stat-card" style="background:#FFF; border-left: 5px solid #007AFF; padding:20px;">
                <h3 style="color:#6B7280; font-size:0.75rem; font-weight:700;">TOTAL SPENDING</h3>
                <h2 style="font-size:1.8rem; font-weight:800; margin: 8px 0;">$ ${formatCurrency(stats.totalSpending)}</h2>
                <div style="font-size: 0.75rem; color: #007AFF;">Across ${stats.eventCount} Events</div>
            </div>
            <div class="stat-card" style="background:#FFF; border-left: 5px solid #10b981; padding:20px;">
                <h3 style="color:#6B7280; font-size:0.75rem; font-weight:700;">TOTAL POC</h3>
                <h2 style="font-size:1.8rem; font-weight:800; margin: 8px 0;">${stats.totalPOC}</h2>
                <div style="font-size: 0.75rem; color: #10b981;">$${formatCurrency(stats.costPerPOC)} Per POC</div>
            </div>
            <div class="stat-card" style="background:#FFF; border-left: 5px solid #f59e0b; padding:20px;">
                <h3 style="color:#6B7280; font-size:0.75rem; font-weight:700;">CONVERTED DEALS</h3>
                <h2 style="font-size:1.8rem; font-weight:800; margin: 8px 0;">${stats.totalDeals}</h2>
                <div style="font-size: 0.75rem; color: #f59e0b;">$${formatCurrency(stats.costPerDeal)} Per Deal</div>
            </div>
        </div>
        <div class="stat-card" style="background:#FFF; padding: 24px;"><canvas id="event-roi-chart" style="height: 350px;"></canvas></div>
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

export function getKPIHTML(kpiData, currentKPIYear = 2026) {
    if (!kpiData || !kpiData.categories) return '<p>No KPI data found.</p>';

    const renderRow = (catId, objId, obj) => {
        const calculateRate = (targets, achievements) => {
            const sumT = targets.reduce((a, b) => a + b, 0);
            const sumA = achievements.reduce((a, b) => a + b, 0);
            if (sumT === 0) return achievements.some(v => v > 0) ? 100 : 0;
            return Math.min(200, Math.round((sumA / sumT) * 100)); // Allow over 100%
        };

        const rate = calculateRate(obj.targets, obj.achievements);
        const rateColor = rate >= 100 ? '#10B981' : (rate >= 70 ? '#F59E0B' : '#EF4444');

        return `
            <tr class="kpi-row" data-cat="${catId}" data-obj="${objId}">
                <td class="kpi-objective">
                    <div contenteditable="true" onblur="this.style.background='transparent'; window.updateKPIObjectiveName(this, ${catId}, ${objId})" style="outline: none; min-height: 1.5em; width: 100%; transition: all 0.2s; cursor: text;" onfocus="this.style.background='rgba(0,0,0,0.02)';" title="Click to edit">${obj.name}</div>
                </td>
                <td class="kpi-indicator" style="padding: 10px 15px;">
                    <div contenteditable="true" onblur="this.style.background='transparent'; window.updateKPIText(this, 'kpis', ${catId}, ${objId})" style="outline: none; min-height: 1.5em; width: 100%; transition: all 0.2s; cursor: text;" onfocus="this.style.background='rgba(0,0,0,0.02)';" title="Click to edit">${obj.kpis || ''}</div>
                </td>
                ${obj.targets.map((t, i) => `
                    <td style="background: rgba(16, 185, 129, 0.05);">
                        <input type="text" class="kpi-target-input" data-idx="${i}" value="${formatCurrency(t)}" onchange="window.updateKPICell(this, 'targets', ${catId}, ${objId}, ${i})">
                    </td>
                `).join('')}
                <td class="kpi-weight" style="padding: 4px;">
                    <div style="display: flex; align-items: center; justify-content: center; gap: 4px;">
                        <input type="number" style="width: 50px; text-align: right; border: 1px solid transparent; background: rgba(0,0,0,0.02); padding: 6px; border-radius: 6px; font-weight: 700; color: inherit; transition: all 0.2s; outline: none;" onfocus="this.style.background='#FFF'; this.style.borderColor='#6366f1';" onblur="this.style.background='rgba(0,0,0,0.02)'; this.style.borderColor='transparent';" onchange="window.updateKPINumber(this, 'weight', ${catId}, ${objId})" value="${obj.weight || 0}">%
                    </div>
                </td>
                <td class="kpi-rate" style="color: ${rateColor}">${rate}%</td>
            </tr>
            <tr class="kpi-row" data-cat="${catId}" data-obj="${objId}">
                <td colspan="2" style="text-align: right; font-weight: 700; color: #64748B; background: #F8FAFC;">Achievement</td>
                ${obj.achievements.map((a, i) => `
                    <td style="background: rgba(99, 102, 241, 0.05);">
                        <input type="text" class="kpi-achieve-input" data-idx="${i}" value="${formatCurrency(a)}" onchange="window.updateKPICell(this, 'achievements', ${catId}, ${objId}, ${i})">
                    </td>
                `).join('')}
                <td colspan="2" style="background: #F8FAFC;"></td>
            </tr>
        `;
    };

    let tableBody = '';
    kpiData.categories.forEach((cat, catIdx) => {
        cat.objectives.forEach((obj, objIdx) => {
            tableBody += `
                <tr class="kpi-row">
                    ${objIdx === 0 ? `<td rowspan="${cat.objectives.length * 2}" class="kpi-cat-cell" style="background: ${cat.color}"><div contenteditable="true" onblur="this.style.background='transparent'; window.updateKPICategoryName(this, ${catIdx})" style="outline: none; min-height: 1.5em; width: 100%; text-align: center; transition: all 0.2s; cursor: text;" onfocus="this.style.background='rgba(255,255,255,0.2)';" title="Click to edit">${cat.name}</div></td>` : ''}
                    ${renderRow(catIdx, objIdx, obj)}
                </tr>
            `;
        });
    });

    return `
        <div class="kpi-container">
            <div class="kpi-actions">
                <div style="display: flex; align-items: center; gap: 10px;">
                    <label style="font-size: 0.85rem; font-weight: 700; color: #64748B;">Target Year:</label>
                    <select id="kpi-year-select" style="padding: 8px 16px; border-radius: 8px; border: 1px solid #CBD5E1; font-weight: 700; font-family: inherit; font-size: 0.95rem; background: #FFF; outline: none; cursor: pointer; color: #1E293B;" onchange="window.changeKPIYear(this.value)">
                        ${[2026, 2027, 2028, 2029, 2030].map(y => `<option value="${y}" ${currentKPIYear === y ? 'selected' : ''}>${y}</option>`).join('')}
                    </select>
                </div>
                <div style="flex-grow: 1;"></div>
                <button class="btn-kpi btn-reset" onclick="window.resetKPIData()"><i class="fa-solid fa-undo"></i> Reset to Default</button>
                <button class="btn-kpi btn-export" onclick="window.exportKPIData()"><i class="fa-solid fa-download"></i> Export JSON</button>
                <button class="btn-kpi btn-save" onclick="window.saveKPIData()"><i class="fa-solid fa-save"></i> Save Changes</button>
            </div>
            <table class="kpi-table" style="table-layout: fixed; width: 100%;">
                <thead class="kpi-header">
                    <tr>
                        <th rowspan="2" style="width: 50px;">Cat.</th>
                        <th rowspan="2" style="width: 150px;">Strategic Objectives</th>
                        <th rowspan="2">Key Performance Indicators</th>
                        <th colspan="4">Targets (${currentKPIYear})</th>
                        <th rowspan="2" style="width: 70px;">Weight</th>
                        <th rowspan="2" style="width: 90px;">Rate</th>
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
            </table>
            <div style="margin-top: 20px; padding: 15px; background: rgba(99,102,241,0.05); border-radius: 12px; border-left: 4px solid #6366f1;">
                <p style="margin: 0; font-size: 0.8rem; color: #4F46E5; font-weight: 600;">
                    <i class="fa-solid fa-circle-info"></i> Tip: Click on any achievement or target value to update it manually. Click 'Save Changes' to persist updates.
                </p>
            </div>
        </div>
    `;
}

