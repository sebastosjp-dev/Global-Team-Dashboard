/*═══════════════════════════════════════════════════════════════
  SECTION: Application State & Configuration
═══════════════════════════════════════════════════════════════*/

/** @type {Object<string, Array<Object>>} Parsed workbook data keyed by sheet name */
let workbookData = {};

/** @type {string|null} Currently active tab name */
let currentTab = null;

/** Global configuration constants */
const CONFIG = {
    COUNTRIES: ['Indonesia', 'Thailand', 'Malaysia', 'USA', 'Philippines', 'Singapore', 'Turkey'],
    COLORS: [
        '#6366f1', '#0ea5e9', '#10b981', '#f59e0b', '#ef4444',
        '#a855f7', '#ec4899', '#14b8a6', '#f97316', '#84cc16'
    ],
    CHART_DEFAULTS: {
        font: "'Inter', sans-serif",
        gridColor: 'rgba(0,0,0,0.05)'
    }
};

/*═══════════════════════════════════════════════════════════════
  SECTION: Shared Utilities
═══════════════════════════════════════════════════════════════*/

/**
 * Parse an Excel cell value into a JavaScript Date.
 * Handles Date objects, Excel serial numbers (>30000), and date strings.
 * @param {*} val - Raw cell value
 * @returns {Date|null} Parsed date or null if unparseable
 */
function parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number' && val > 30000) {
        return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    if (typeof val === 'string' && !isNaN(val) && val.length > 4) {
        return new Date(Math.round((parseFloat(val) - 25569) * 86400 * 1000));
    }
    const parsed = new Date(val);
    return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Find a column key from an array of keys using multiple match strategies.
 * Tries each pattern function in order, returns the first match.
 * @param {string[]} keys - Array of column header keys
 * @param {...function(string): boolean} patterns - Predicate functions to test each key
 * @returns {string|undefined} The first matching key, or undefined
 */
function findKey(keys, ...patterns) {
    for (const pattern of patterns) {
        const found = keys.find(pattern);
        if (found) return found;
    }
    return undefined;
}

/**
 * Check if a status string matches any of the given terms (case-insensitive).
 * @param {string} status - The status string to check
 * @param {...string} terms - Terms to match against
 * @returns {boolean}
 */
function matchStatus(status, ...terms) {
    const lower = String(status || '').trim().toLowerCase();
    return terms.some(t => lower.includes(t));
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Chart Management
═══════════════════════════════════════════════════════════════*/

/** Chart instance registry — ensures proper cleanup on re-render */
const chartRegistry = {
    /** @type {Object<string, Chart>} */
    instances: {},

    /** Register a chart, destroying any existing one with the same id */
    register(id, chart) {
        if (this.instances[id]) this.instances[id].destroy();
        this.instances[id] = chart;
    },

    /** Destroy a single chart by id */
    destroy(id) {
        if (this.instances[id]) {
            this.instances[id].destroy();
            delete this.instances[id];
        }
    },

    /** Destroy all charts whose id starts with the given tag */
    destroyTag(tag) {
        Object.keys(this.instances).forEach(id => {
            if (id.startsWith(tag)) this.destroy(id);
        });
    },

    /** Destroy all registered charts */
    destroyAll() {
        Object.keys(this.instances).forEach(id => this.destroy(id));
    }
};

/*═══════════════════════════════════════════════════════════════
  SECTION: Country Sorting & Formatting Helpers
═══════════════════════════════════════════════════════════════*/

/** Update the "Last Updated" display in the sidebar header */
function updateLastUpdatedDate() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const dateEl = document.getElementById('last-updated-date');
    if (dateEl) {
        dateEl.innerText = `${d}.${m}.${y}`;
    }
}

/** Get sort priority for a country (lower = higher priority) */
const getCountrySortOrder = (c) => {
    if (c === 'Indonesia') return 1;
    if (c === 'Malaysia') return 2;
    if (c === 'Thailand') return 3;
    return 4;
};
const sortCountriesByAmount = (a, b) => {
    const oA = getCountrySortOrder(a[0]), oB = getCountrySortOrder(b[0]);
    if (oA !== oB) return oA - oB;
    return b[1].amount - a[1].amount;
};
const sortCountriesByCount = (a, b) => {
    const oA = getCountrySortOrder(a[0]), oB = getCountrySortOrder(b[0]);
    if (oA !== oB) return oA - oB;
    return b[1] - a[1];
};

/*═══════════════════════════════════════════════════════════════
  SECTION: DOM Element References
═══════════════════════════════════════════════════════════════*/

// DOM Elements
const dropZone = document.getElementById('drop-zone');
const dashboardContainer = document.getElementById('dashboard-container');
const sidebarNav = document.getElementById('sidebar-nav');
const currentTabTitle = document.getElementById('current-tab-title');
const tableHead = document.getElementById('table-head');
const tableBody = document.getElementById('table-body');
const emptyState = document.getElementById('empty-state');
const dataTable = document.getElementById('data-table');
const searchInput = document.getElementById('search-input');

/*═══════════════════════════════════════════════════════════════
  SECTION: File Handling & Data Loading
═══════════════════════════════════════════════════════════════*/

/** Handle manual file uploads (drag & drop) */
function handleFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('Please upload a valid Excel file (.xlsx or .xls)');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        processWorkbook(workbook);
    };
    reader.readAsArrayBuffer(file);
}

/** Load Excel files from the local development server */
async function loadLocalExcel() {
    try {
        // Show loading state if button was clicked
        const btn = document.getElementById('refresh-btn');
        if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';

        // Add a timestamp to bypass browser cache
        const ts = new Date().getTime();

        const [resMain, resMrr] = await Promise.all([
            fetch('2026 Global Rev.01.xlsx?t=' + ts),
            fetch('Global MRR ARR.xlsx?t=' + ts).catch(() => null)
        ]);

        if (!resMain.ok) {
            throw new Error(`Failed to fetch main file. Server returned ${resMain.status}`);
        }

        const arrayBufferMain = await resMain.arrayBuffer();
        const dataMain = new Uint8Array(arrayBufferMain);
        const workbookMain = XLSX.read(dataMain, { type: 'array' });

        if (resMrr && resMrr.ok) {
            const arrayBufferMrr = await resMrr.arrayBuffer();
            const dataMrr = new Uint8Array(arrayBufferMrr);
            const workbookMrr = XLSX.read(dataMrr, { type: 'array' });

            let targetSheetName = workbookMrr.SheetNames.find(name => name.includes('Global(怨꾩빟?쒖젏)'));
            if (targetSheetName) {
                const sheet = workbookMrr.Sheets[targetSheetName];
                const json = XLSX.utils.sheet_to_json(sheet, { defval: "", cellDates: true });
                window.globalMrrData = json;
                window.globalMrrSheet = sheet; // Store raw sheet for direct cell access
                // Inject into main workbook data so it appears in sidebar
                workbookMain.SheetNames.push(targetSheetName);
                workbookMain.Sheets[targetSheetName] = sheet;
                console.log("Injected MRR sheet into main workbook:", targetSheetName);
            }
        }

        processWorkbook(workbookMain);

        if (btn) btn.innerHTML = '<i class="fa-solid fa-arrows-rotate"></i>';
    } catch (error) {
        console.error('Error loading Excel:', error);
        alert('Could not load Excel files:\n' + (error.stack || error.message) + '\n\nPlease ensure the local server is running.');
        const btn = document.getElementById('refresh-btn');
        if (btn) btn.innerHTML = '<i class="fa-solid fa-arrows-rotate"></i>';
    }
}

// Auto-run on load
window.addEventListener('load', () => {
    loadLocalExcel();
});

function processWorkbook(workbook) {
    workbookData = {};
    const sheetNames = workbook.SheetNames;

    sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "", cellDates: true });
        workbookData[sheetName] = json;
    });

    // Setup UI
    dropZone.classList.remove('active');
    dashboardContainer.classList.add('active');

    buildSidebar(sheetNames, selectTab);
    updateLastUpdatedDate();

    // Select first tab by default
    if (sheetNames.length > 0) {
        const firstTab = sheetNames.includes('ORDER SHEET') ? 'ORDER SHEET' : sheetNames[0];
        selectTab(firstTab);
    }
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Sidebar Navigation
═══════════════════════════════════════════════════════════════*/

// --- Dashboard UI ---
/**
 * Build the sidebar navigation from workbook sheet names.
 * Creates expandable parent items for ORDER SHEET and END USER (CSM).
 * @param {string[]} sheetNames - Array of sheet names from the workbook
 */
function buildSidebar(sheetNames) {
    sidebarNav.innerHTML = '';

    sheetNames.forEach(name => {
        if (name.includes('Global(怨꾩빟?쒖젏)') || ['Sheet9', 'Sheet10', 'Sheet12', 'Sheet13', '2026 Q1 Review'].includes(name)) return;

        if (name === 'ORDER SHEET') {
            // Create Parent Item
            const parentItem = document.createElement('div');
            parentItem.className = 'nav-item nav-item-parent';
            parentItem.innerHTML = `
                <div style="display:flex; align-items:center; gap:12px;">
                    <i class="fa-solid fa-folder"></i> <span>${name}</span>
                </div>
                <i class="fa-solid fa-chevron-down toggle-icon" style="font-size: 0.8em; transition: transform 0.3s;"></i>
            `;

            // Hardcoded list of countries requested by user
            const countries = CONFIG.COUNTRIES;

            const subList = document.createElement('div');
            subList.className = 'nav-sublist';

            // "All Countries" option
            const allItem = document.createElement('div');
            allItem.className = 'nav-item sub-item';
            allItem.innerHTML = `<i class="fa-solid fa-earth-americas"></i> <span>All</span>`;
            allItem.onclick = (e) => { e.stopPropagation(); selectTab(name, null); };
            subList.appendChild(allItem);

            // Individual Country options
            countries.forEach(country => {
                const subItem = document.createElement('div');
                subItem.className = 'nav-item sub-item';
                subItem.innerHTML = `<i class="fa-solid fa-location-dot"></i> <span>${country}</span>`;
                subItem.onclick = (e) => { e.stopPropagation(); selectTab(name, country); };
                subList.appendChild(subItem);
            });

            // Toggle logic
            parentItem.onclick = () => {
                subList.classList.toggle('expanded');
                const icon = parentItem.querySelector('.toggle-icon');
                if (subList.classList.contains('expanded')) {
                    icon.style.transform = 'rotate(180deg)';
                } else {
                    icon.style.transform = 'rotate(0deg)';
                }
                // Also select the main tab when clicking the parent
                selectTab(name, null);
            };

            sidebarNav.appendChild(parentItem);
            sidebarNav.appendChild(subList);

        } else if (name === 'END USER (CSM)') {
            // Create Parent Item for CSM
            const parentItem = document.createElement('div');
            parentItem.className = 'nav-item nav-item-parent';
            parentItem.innerHTML = `
                <div style="display:flex; align-items:center; gap:12px;">
                    <i class="fa-solid fa-folder"></i> <span>${name}</span>
                </div>
                <i class="fa-solid fa-chevron-down toggle-icon" style="font-size: 0.8em; transition: transform 0.3s;"></i>
            `;

            const subList = document.createElement('div');
            subList.className = 'nav-sublist';

            // --- Renewal Management Sub-item ---
            const renewalItem = document.createElement('div');
            renewalItem.className = 'nav-item renewal-tab sub-item';
            renewalItem.style.background = 'linear-gradient(90deg, rgba(245, 158, 11, 0.1) 0%, rgba(245, 158, 11, 0) 100%)';
            renewalItem.style.borderLeft = '2px solid #f59e0b';
            renewalItem.innerHTML = `<i class="fa-solid fa-user-clock" style="color: #f59e0b;"></i> <span>Renewal Management</span>`;
            renewalItem.onclick = (e) => { e.stopPropagation(); selectRenewalView(); };
            subList.appendChild(renewalItem);

            // --- Service Analysis (Upsell) Sub-item ---
            const upsellItem = document.createElement('div');
            upsellItem.className = 'nav-item upsell-tab sub-item';
            upsellItem.style.background = 'linear-gradient(90deg, rgba(16, 185, 129, 0.1) 0%, rgba(16, 185, 129, 0) 100%)';
            upsellItem.style.borderLeft = '2px solid #10b981';
            upsellItem.innerHTML = `<i class="fa-solid fa-chart-pie" style="color: #34C759;"></i> <span>Service Analysis (Upsell)</span>`;
            upsellItem.onclick = (e) => { e.stopPropagation(); selectServiceAnalysisView(); };
            subList.appendChild(upsellItem);

            // Toggle logic
            parentItem.onclick = () => {
                subList.classList.toggle('expanded');
                const icon = parentItem.querySelector('.toggle-icon');
                if (subList.classList.contains('expanded')) {
                    icon.style.transform = 'rotate(180deg)';
                } else {
                    icon.style.transform = 'rotate(0deg)';
                }
                selectTab(name, null);
            };

            sidebarNav.appendChild(parentItem);
            sidebarNav.appendChild(subList);

        } else {
            // Standard Item
            const navItem = document.createElement('div');
            navItem.className = 'nav-item';
            const icon = name === 'EVENT' ? 'fa-calendar-check' : 'fa-folder';
            navItem.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${name}</span>`;
            navItem.onclick = () => selectTab(name, null);
            sidebarNav.appendChild(navItem);
        }
    });
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Renewal & CSM Views
═══════════════════════════════════════════════════════════════*/

/** Switch to the Renewal Management view */
function selectRenewalView() {
    currentTab = 'RENEWAL_VIEW';
    const title = 'Churn Prevention & Renewal Management (CSM)';
    currentTabTitle.innerText = title;

    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.renewal-tab')?.classList.add('active');

    renderRenewalTable();
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Service Analysis (Upsell)
═══════════════════════════════════════════════════════════════*/

/* --- SERVICE ANALYSIS HELPERS --- */
function injectServiceAnalysisStyles() {
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
    `;
    document.head.appendChild(style);
}

function getServiceAnalysisStats(data) {
    const activeData = data.filter(r => r['Status'] === 'Active' && r['Services']);
    if (activeData.length === 0) return null;
    const comboCounts = {};
    const upscaleTargetsArr = [];
    activeData.forEach(row => {
        const raw = String(row['Services'] || '').trim();
        if (!raw) return;
        const services = raw.split(',').map(s => s.trim()).filter(Boolean);
        const combo = [...services].sort().join(' + ');
        comboCounts[combo] = (comboCounts[combo] || 0) + 1;
        if (services.length === 1) {
            upscaleTargetsArr.push({ name: row['End User'], country: row['Country'], service: raw, tcv: parseCurrency(row['TCV Amount']) });
        }
    });
    return {
        totalCustomers: activeData.length,
        singleServiceCustomers: upscaleTargetsArr.length,
        multiServiceCustomers: activeData.length - upscaleTargetsArr.length,
        sortedCombos: Object.entries(comboCounts).sort((a, b) => b[1] - a[1]),
        upsellTargets: upscaleTargetsArr.sort((a, b) => b.tcv - a.tcv),
        palette: CONFIG.COLORS
    };
}

function getServiceAnalysisHTML(stats) {
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

function initServiceAnalysisCharts(stats) {
    if (!stats) return;
    chartRegistry.destroyTag('service');
    const ctx = document.getElementById('service-donut-chart');
    if (ctx) {
        chartRegistry.register('service-donut', new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: stats.sortedCombos.slice(0, 8).map(c => c[0]),
                datasets: [{ data: stats.sortedCombos.slice(0, 8).map(c => c[1]), backgroundColor: stats.palette }]
            },
            options: { responsive: true, maintainAspectRatio: false, cutout: '65%', plugins: { legend: { position: 'right', labels: { boxWidth: 10, font: { size: 10 } } } } }
        }));
    }
}

function selectServiceAnalysisView() {
    currentTab = 'SERVICE_ANALYSIS_VIEW';
    currentTabTitle.innerText = 'Service Combination Analysis (Upsell Opportunities)';

    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.upsell-tab')?.classList.add('active');

    renderServiceAnalysisView();
}

function renderServiceAnalysisView() {
    const data = workbookData['END USER (CSM)'] || [];
    const metricsGrid = document.getElementById('tab-metrics-grid');

    // Clear display
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    dataTable.classList.add('hidden');
    emptyState.classList.add('hidden');
    metricsGrid.innerHTML = '';
    metricsGrid.classList.remove('hidden');

    const stats = getServiceAnalysisStats(data);
    metricsGrid.innerHTML = getServiceAnalysisHTML(stats);

    if (stats) {
        setTimeout(() => initServiceAnalysisCharts(stats), 100);
    }
}
/* --- RENEWAL TABLE HELPERS --- */
function getRenewalTableData() {
    const data = workbookData['END USER (CSM)'] || [];
    const today = new Date('2026-03-18');
    const sixMonthsLater = new Date(today);
    sixMonthsLater.setMonth(today.getMonth() + 6);

    const getExpDate = (row) => {
        let endDate = row['End License Date'];
        if (!endDate) return null;
        return (endDate instanceof Date) ? endDate : (typeof endDate === 'number') ? new Date(Math.round((endDate - 25569) * 86400 * 1000)) : new Date(endDate);
    };

    return data.filter(row => {
        const d = getExpDate(row);
        return d && d >= today && d <= sixMonthsLater;
    }).map(row => {
        const d = getExpDate(row);
        const diffDays = Math.ceil((d - today) / (1000 * 60 * 60 * 24));

        return {
            ...row,
            'D-Day': diffDays === 0 ? 'D-Day' : (diffDays > 0 ? `D-${diffDays}` : `D+${Math.abs(diffDays)}`),
            'diffDays': diffDays,
            'endDateFormatted': d.toISOString().split('T')[0]
        };
    }).sort((a, b) => a.diffDays - b.diffDays || (parseCurrency(b['TCV Amount']) - parseCurrency(a['TCV Amount'])));
}

function getRenewalHTML(filtered) {
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

function renderRenewalTable() {
    const filtered = getRenewalTableData();
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    metricsGrid.classList.remove('hidden');
    metricsGrid.innerHTML = getRenewalHTML(filtered);
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Tab Selection & Data Rendering
═══════════════════════════════════════════════════════════════*/

/**
 * Switch to a specific tab, optionally filtering by country.
 * @param {string} tabName - Sheet/tab name
 * @param {string|null} [subTabName=null] - Country filter
 */
function selectTab(tabName, subTabName = null) {
    currentTab = tabName;
    currentTabTitle.innerText = subTabName ? `${tabName} — ${subTabName}` : tabName;
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    renderTableData("", subTabName);
}
function isCountryMatch(row, filterCountry) {
    if (!filterCountry || filterCountry === 'All') return true;
    const k = Object.keys(row).find(k => k.toLowerCase().includes('country') || k.toLowerCase().includes('region') || k.toLowerCase().includes('nation'));
    if (!k) return false;
    return normalizeCountry(row[k]) === filterCountry;
}
function renderTableData(searchTerm = "", filterCountry = null) {
    const data = workbookData[currentTab] || [];
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    if (data.length === 0) return;
    const headers = Object.keys(data[0]).filter(k => !k.startsWith('__EMPTY'));
    const trHead = document.createElement('tr'); headers.forEach(h => { const th = document.createElement('th'); th.innerText = h; trHead.appendChild(th); });
    tableHead.appendChild(trHead);
    const filtered = data.filter(r => isCountryMatch(r, filterCountry) && (!searchTerm || Object.values(r).some(v => String(v).toLowerCase().includes(searchTerm.toLowerCase()))));
    filtered.slice(0, 50).forEach(row => {
        const tr = document.createElement('tr'); headers.forEach(h => { const td = document.createElement('td'); td.innerText = row[h] || ''; tr.appendChild(td); });
        tableBody.appendChild(tr);
    });
    renderTabMetrics(filtered, currentTab, filterCountry);
}

function normalizeCountry(raw) {
    if (!raw) return null;
    let c = String(raw).trim();
    const up = c.toUpperCase();
    if (up === 'IDN' || up.includes('INDONESIA')) return 'Indonesia';
    if (up === 'US' || up === 'USA' || up.includes('UNITED STATES') || up.includes('미국')) return 'USA';
    if (up === 'MA' || up === 'MAL' || up.includes('MALAYSIA')) return 'Malaysia';
    if (up === 'TH' || up === 'THA' || up.includes('THAILAND')) return 'Thailand';
    if (up === 'PH' || up === 'PHI' || up.includes('PHILIPPINES')) return 'Philippines';
    if (up === 'TUR' || up.includes('TURKEY') || up === 'TR') return 'Turkey';
    if (up === 'SIN' || up.includes('SINGAPORE') || up === 'SGP') return 'Singapore';
    return c;
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Metrics Router
═══════════════════════════════════════════════════════════════*/

/**
 * Render tab-specific metrics based on the active tab.
 * Routes to the appropriate rendering function per tab type.
 * @param {Array<Object>} data - Filtered data for the current tab
 * @param {string} tabName - Active tab name
 * @param {string|null} filterCountry - Country filter applied
 */
function renderTabMetrics(data, tabName, filterCountry = null) {
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    let hasMetrics = false;
    const isGlobalTab = tabName && tabName.includes('Global(계약시점)');
    const isCountryTab = tabName && filterCountry === null && !['ORDER SHEET', 'PIPELINE', 'PARTNER', 'POC', 'EVENT', 'END USER (CSM)'].includes(tabName) && !isGlobalTab;

    if (tabName === 'ORDER SHEET' || isGlobalTab) {
        renderOrderSheetMetrics(data, filterCountry, metricsGrid, tabName);
        renderExpiringContracts(data, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'PIPELINE' && workbookData['PIPELINE']) {
        renderPipelineMetrics(workbookData['PIPELINE'], filterCountry, tabName, metricsGrid);
        hasMetrics = true;
    }

    if ((tabName === 'PARTNER' || isCountryTab) && data && data.length > 0) {
        renderPartnerMetricsBlock(data, filterCountry, tabName, metricsGrid);
        renderGenericCountryCounts(data, filterCountry, metricsGrid, tabName);
        renderPartnerTopPerformer(data, metricsGrid);
        hasMetrics = true;
    }

    if ((String(tabName).trim().toUpperCase() === 'POC' || isCountryTab) && data && data.length > 0) {
        renderPocMetricsBlock(data, filterCountry, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'EVENT' && workbookData['EVENT']) {
        renderEventMetrics(workbookData['EVENT'], filterCountry, metricsGrid);
        hasMetrics = true;
    }

    if (hasMetrics) {
        metricsGrid.classList.remove('hidden');
    } else {
        metricsGrid.classList.add('hidden');
    }
}


/*═══════════════════════════════════════════════════════════════
  SECTION: ORDER SHEET (Stats, HTML, Charts, Render)
═══════════════════════════════════════════════════════════════*/

/* --- ORDER SHEET HELPERS --- */
/**
 * Calculate aggregate stats for the ORDER SHEET tab.
 * @param {Array<Object>} data - Current tab data
 * @param {string|null} filterCountry - Country filter
 * @param {string} tabName - Active tab name
 * @returns {{ sumLocalTcv, sumKorTcv, sumArr, sumMrr, dealCount, yearlyTcv, qSums }}
 */
function getOrderSheetStats(data, filterCountry, tabName) {
    const orderData = tabName === 'ORDER SHEET' ? data : (workbookData['ORDER SHEET'] || []);
    const keys = orderData.length > 0 ? Object.keys(orderData[0]) : [];
    const korTcvKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'KORTCV') || keys.find(k => k.toUpperCase().includes('TCV') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
    const arrKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'KORARR') || keys.find(k => k.toUpperCase().includes('ARR') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
    const mrrKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'KORMRR') || keys.find(k => k.toUpperCase().includes('MRR') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
    const startDateKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTSTART')) || keys.find(k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE'));

    let sumLocalTcv = 0, sumKorTcv = 0, sumArr = 0, sumMrr = 0, dealCount = 0;
    const yearlyTcv = {};
    const qSums = { 'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0 };

    orderData.filter(r => isCountryMatch(r, filterCountry)).forEach(row => {
        const lTcv = parseCurrency(row['Local TCV'] || row['Local TCV Amount']);
        const kTcv = korTcvKey ? parseCurrency(row[korTcvKey]) : 0;
        const arrVal = arrKey ? parseCurrency(row[arrKey]) : 0;
        const mrrVal = mrrKey ? parseCurrency(row[mrrKey]) : 0;

        sumLocalTcv += lTcv;
        sumKorTcv += kTcv;
        sumArr += arrVal;
        sumMrr += mrrVal;
        dealCount++;

        const cStart = row[startDateKey];
        if (cStart) {
            let d = (cStart instanceof Date) ? cStart : (typeof cStart === 'number' && cStart > 40000) ? new Date(Math.round((cStart - 25569) * 86400 * 1000)) : new Date(cStart);
            if (d && !isNaN(d.getTime())) {
                const y = d.getFullYear().toString();
                if (!yearlyTcv[y]) yearlyTcv[y] = { local: 0, korea: 0 };
                yearlyTcv[y].local += lTcv;
                yearlyTcv[y].korea += kTcv;
                if (d.getFullYear() === 2026) {
                    qSums[`Q${Math.floor(d.getMonth() / 3) + 1}`] += kTcv;
                }
            }
        }
    });

    return { sumLocalTcv, sumKorTcv, sumArr, sumMrr, dealCount, yearlyTcv, qSums };
}

function getOrderSheetHTML(stats) {
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

function initOrderSheetCharts(stats) {
    chartRegistry.destroyTag('order');
    const barCtx = document.getElementById('quarterly-tcv-bar');
    if (barCtx) {
        chartRegistry.register('order-quarterly', new Chart(barCtx, {
            type: 'bar',
            data: {
                labels: ['Q1', 'Q2', 'Q3', 'Q4'],
                datasets: [{
                    label: 'KOR TCV',
                    data: [stats.qSums.Q1, stats.qSums.Q2, stats.qSums.Q3, stats.qSums.Q4],
                    backgroundColor: 'rgba(245, 158, 11, 0.6)',
                    borderColor: '#f59e0b',
                    borderWidth: 1,
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => '$' + formatCurrency(v) } },
                    x: { grid: { display: false } }
                }
            }
        }));
    }
    const lineCtx = document.getElementById('tcv-growth-chart');
    if (lineCtx) {
        const years = Object.keys(stats.yearlyTcv).sort();
        chartRegistry.register('order-growth', new Chart(lineCtx, {
            type: 'line',
            data: {
                labels: years,
                datasets: [{
                    label: 'KOR TCV',
                    data: years.map(y => stats.yearlyTcv[y].korea),
                    borderColor: '#10b981',
                    backgroundColor: 'rgba(16, 185, 129, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 4
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => '$' + formatCurrency(v) } },
                    x: { grid: { display: false } }
                }
            }
        }));
    }
}

function renderOrderSheetMetrics(data, filterCountry, metricsGrid, tabName) {
    const stats = getOrderSheetStats(data, filterCountry, tabName);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getOrderSheetHTML(stats);
    metricsGrid.appendChild(container);
    setTimeout(() => initOrderSheetCharts(stats), 120);
}

/*═══════════════════════════════════════════════════════════════
  SECTION: PIPELINE (Stats, HTML, Charts, Render)
═══════════════════════════════════════════════════════════════*/

/**
 * Calculate pipeline statistics by country and quarter.
 * @param {Array<Object>} pData - Pipeline data rows
 * @returns {{ pipelineByCountry, pipelineByQuarter, pipelineInfluxData, sortedPipeline, sortedQuarterly, globalTotalAmount, globalTotalWeighted }}
 */
function getPipelineStats(pData) {
    let pipelineByCountry = {};
    let pipelineByQuarter = { 'Q1': {}, 'Q2': {}, 'Q3': {}, 'Q4': {} };
    const pipelineInfluxData = Array(12).fill(0).map(() => ({ count: 0, amount: 0, weighted: 0, value: 0, accounts: [] }));

    pData.forEach(r => {
        const keys = Object.keys(r);
        const c = normalizeCountry(r[Object.keys(r).find(k => k.toLowerCase().includes('country'))]) || 'Other';

        const amtRaw = keys.find(k => (k.toUpperCase().includes('KOR TCV') && k.toUpperCase().includes('USD')) || k === 'Amount') || 'Amount';
        const wAmtRaw = keys.find(k => (k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('KOR TCV')) || k === 'Weighted Amount') || 'Weighted Amount';

        const amt = parseCurrency(r[amtRaw] || r['Amount']);
        const wAmt = parseCurrency(r[wAmtRaw] || r['Weighted Amount']);

        if (!pipelineByCountry[c]) pipelineByCountry[c] = { amount: 0, weighted: 0 };
        pipelineByCountry[c].amount += amt;
        pipelineByCountry[c].weighted += wAmt;

        const qKey = keys.find(k => k.toLowerCase() === 'quarter' || k.toLowerCase().includes('qtr') || k.toLowerCase() === 'q');
        if (qKey && r[qKey]) {
            const qRaw = String(r[qKey]).toUpperCase().trim();
            let qMatch = '';
            if (qRaw.includes('Q1')) qMatch = 'Q1';
            else if (qRaw.includes('Q2')) qMatch = 'Q2';
            else if (qRaw.includes('Q3')) qMatch = 'Q3';
            else if (qRaw.includes('Q4')) qMatch = 'Q4';

            if (qMatch) {
                if (!pipelineByQuarter[qMatch][c]) pipelineByQuarter[qMatch][c] = { amount: 0, weighted: 0 };
                pipelineByQuarter[qMatch][c].amount += amt;
                pipelineByQuarter[qMatch][c].weighted += wAmt;
            }
        }

        const dKey = keys.find(k => k.toLowerCase().includes('date') || k.toLowerCase().includes('start'));
        if (dKey && r[dKey]) {
            let d = null;
            if (r[dKey] instanceof Date) d = r[dKey];
            else if (typeof r[dKey] === 'number' && r[dKey] > 40000) d = new Date(Math.round((r[dKey] - 25569) * 86400 * 1000));
            else {
                const parsed = new Date(r[dKey]);
                if (!isNaN(parsed.getTime())) d = parsed;
            }
            if (d && d.getFullYear() === 2026) {
                const m = d.getMonth();
                pipelineInfluxData[m].count++;
                pipelineInfluxData[m].amount += amt;
                pipelineInfluxData[m].weighted += wAmt;
                pipelineInfluxData[m].value += amt;
                const nKey = keys.find(k => k.toLowerCase().includes('name') || k.toLowerCase().includes('customer') || k.toLowerCase().includes('end user'));
                if (nKey) pipelineInfluxData[m].accounts.push(String(r[nKey]).trim());
            }
        }
    });

    return {
        pipelineByCountry,
        pipelineByQuarter,
        pipelineInfluxData,
        sortedPipeline: Object.entries(pipelineByCountry).sort((a, b) => b[1].amount - a[1].amount),
        sortedQuarterly: Object.entries(pipelineByQuarter).sort((a, b) => a[0].localeCompare(b[0])),
        globalTotalAmount: Object.values(pipelineByCountry).reduce((acc, curr) => acc + curr.amount, 0),
        globalTotalWeighted: Object.values(pipelineByCountry).reduce((acc, curr) => acc + curr.weighted, 0)
    };
}

function getPipelineHTML(stats, filterCountry, tabName) {
    const pipelineItemsHtml = stats.sortedPipeline.map(([country, values]) => `
        <div style="display: flex; flex-direction: column; padding: 10px; background: #F9FAFB; border-radius: 8px; border-left: 3px solid #10b981;">
            <span style="font-weight: 700; color: #374151; font-size: 0.8rem; margin-bottom: 6px;">${filterCountry ? 'Total Summary' : country}</span>
            <div style="display: flex; justify-content: space-between; font-size: 0.72rem; margin-bottom: 2px;">
                <span style="color: var(--text-muted);">Amount</span>
                <span style="color: #34C759; font-weight: 600;">$${formatCurrency(values.amount)}</span>
            </div>
            <div style="display: flex; justify-content: space-between; font-size: 0.72rem;">
                <span style="color: var(--text-muted);">Weighted</span>
                <span style="color: #007AFF; font-weight: 600;">$${formatCurrency(values.weighted)}</span>
            </div>
        </div>
    `).join('');

    const quarterlyItemsHtml = stats.sortedQuarterly.map(([q, countries]) => {
        const countryEntries = Object.entries(countries);
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
                        <span style="color: var(--text-muted);">Amount</span>
                        <span style="color: #34C759;">$${formatCurrency(values.amount)}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 0.68rem;">
                        <span style="color: var(--text-muted);">Weighted</span>
                        <span style="color: #007AFF;">$${formatCurrency(values.weighted)}</span>
                    </div>
                </div>
            `).join('');

        return `
            <div style="display: flex; flex-direction: column; padding: 12px; background: #F9FAFB; border-radius: 8px; border-top: 3px solid #10b981;">
                <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; border-bottom: 1px solid #E5E7EB; padding-bottom: 8px;">
                    <span style="font-weight: 800; color: #111827; font-size: 0.9rem; margin-top: 2px;">${q}</span>
                    <div style="text-align: right; display: flex; flex-direction: column; gap: 3px;">
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 8px;">
                            <span style="font-size: 0.65rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em;">Total Amount</span>
                            <span style="font-size: 0.95rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalAmount)}</span>
                        </div>
                        <div style="display: flex; justify-content: flex-end; align-items: center; gap: 8px;">
                            <span style="font-size: 0.65rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.05em;">Total Weighted</span>
                            <span style="font-size: 0.95rem; color: #34C759; font-weight: 800;">$${formatCurrency(qTotalWeighted)}</span>
                        </div>
                    </div>
                </div>
                ${!filterCountry ? `
                <div style="background: rgba(0,0,0,0.15); border-radius: 10px; padding: 12px; margin-top: 10px;">
                    <div style="max-height: 250px; overflow-y: auto; padding-right: 2px;">
                        ${countryBreakdown}
                    </div>
                </div>
                ` : ''}
            </div>
        `;
    }).join('');

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
                        <span style="font-size: 0.75rem; color: #34C759; text-transform: uppercase; letter-spacing: 0.05em;">Total Amount</span>
                        <h2 style="font-size: 1.4rem; font-weight: 800; color: #111827; margin: 0;">US$ ${formatCurrency(stats.globalTotalAmount)}</h2>
                    </div>
                    <div>
                        <span style="font-size: 0.75rem; color: #007AFF; text-transform: uppercase; letter-spacing: 0.05em;">Total Weighted</span>
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

function initPipelineCharts(stats) {
    const selector = document.getElementById('pipeline-filter-country');
    if (selector) {
        selector.addEventListener('change', (e) => {
            const val = e.target.value;
            const fVal = val === 'All' ? null : val;
            renderTableData(searchInput.value, fVal);
        });
    }

    chartRegistry.destroyTag('pipeline');
    const ctx = document.getElementById('pipeline-influx-chart');
    if (ctx) {
        const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: months,
                datasets: [
                    {
                        label: 'Expected Deals',
                        data: stats.pipelineInfluxData.map(d => d.count),
                        backgroundColor: 'rgba(52,199,89,0.55)',
                        borderColor: '#34C759',
                        borderWidth: 1,
                        borderRadius: 4,
                        yAxisID: 'yCount'
                    },
                    {
                        label: 'Pipeline Value (USD)',
                        data: stats.pipelineInfluxData.map(d => d.value),
                        type: 'line',
                        borderColor: '#f59e0b',
                        backgroundColor: 'rgba(245, 158, 11, 0.1)',
                        borderWidth: 3,
                        tension: 0.4,
                        fill: true,
                        pointRadius: 4,
                        yAxisID: 'yValue'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    yCount: {
                        type: 'linear', position: 'left',
                        grid: { color: 'rgba(0,0,0,0.06)' },
                        ticks: { color: '#6B7280', beginAtZero: true, stepSize: 1 },
                        title: { display: true, text: 'Deal Count', color: '#94a3b8', font: { size: 10 } }
                    },
                    yValue: {
                        type: 'linear', position: 'right',
                        grid: { display: false },
                        ticks: { color: '#f59e0b', callback: (v) => formatCurrency(v) },
                        title: { display: true, text: 'Value (USD)', color: '#f59e0b', font: { size: 10 } }
                    },
                    x: { grid: { display: false }, ticks: { color: '#6B7280', font: { size: 10 } } }
                },
                plugins: {
                    legend: { position: 'top', labels: { color: '#374151', font: { size: 11 }, usePointStyle: true } },
                    tooltip: {
                        backgroundColor: '#FFFFFF', titleColor: '#111827', bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1, padding: 12,
                        callbacks: {
                            label: function (context) {
                                let label = context.dataset.label || '';
                                if (label) label += ': ';
                                if (context.dataset.yAxisID === 'yValue') label += 'US$ ' + formatCurrency(context.parsed.y);
                                else label += context.parsed.y;
                                return label;
                            },
                            afterBody: function (context) {
                                const mIdx = context[0].dataIndex;
                                const accs = stats.pipelineInfluxData[mIdx].accounts;
                                if (accs && accs.length > 0) {
                                    const lines = ['--- Accounts ---'];
                                    accs.slice(0, 10).forEach(a => lines.push(`??${a}`));
                                    if (accs.length > 10) lines.push(`... (+${accs.length - 10} more)`);
                                    return lines;
                                }
                                return '';
                            }
                        }
                    }
                }
            }
        });
        chartRegistry.register('pipeline-influx', chart);
    }
}

function renderPipelineMetrics(pDataRaw, filterCountry, tabName, metricsGrid) {
    const pData = filterCountry ? workbookData['PIPELINE'].filter(r => isCountryMatch(r, filterCountry)) : workbookData['PIPELINE'];
    if (!pData || pData.length === 0) return;

    const stats = getPipelineStats(pData);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.style.marginTop = '12px';
    container.style.marginBottom = '24px';
    container.innerHTML = getPipelineHTML(stats, filterCountry, tabName);
    metricsGrid.appendChild(container);

    setTimeout(() => initPipelineCharts(stats), 100);
}

/*═══════════════════════════════════════════════════════════════
  SECTION: PARTNER (Stats, HTML, Charts, Render)
═══════════════════════════════════════════════════════════════*/

/**
 * Calculate partner statistics with POC cross-reference.
 * @param {Array<Object>} data - Partner data rows
 * @param {string|null} filterCountry - Country filter
 * @returns {{ counts, partnerGroups, sortedCountries, sortedP, pNameKey }}
 */
function getPartnerStats(data, filterCountry) {
    const pKeys = Object.keys(data[0]);
    const pCountryKey = pKeys.find(k => k.toLowerCase().includes('country') || k.toLowerCase().includes('region') || k.toLowerCase() === 'nation');
    const pNameKey = pKeys.find(k => k.toLowerCase().includes('partner') || k.toLowerCase().includes('name') || k.toLowerCase().includes('?뚯궗')) || pKeys[0];

    const counts = {};
    const partnerGroups = {};
    data.forEach(r => {
        let country = normalizeCountry(r[pCountryKey]);
        if (country) {
            if (!partnerGroups[country]) partnerGroups[country] = [];
            partnerGroups[country].push(r);
            counts[country] = (counts[country] || 0) + 1;
        }
    });

    const pocDataForPartner = workbookData['POC'] || [];
    const currentPartnerNames = new Set(data.map(p => String(p[pNameKey] || '').trim().toLowerCase()));

    let pRankingData = {};
    pocDataForPartner.forEach(r => {
        const sKey = Object.keys(r).find(k => k.toLowerCase().includes('status'));
        const statusStr = sKey ? String(r[sKey]).trim().toLowerCase() : '';

        if (statusStr.includes('running') || statusStr.includes('progress') || statusStr.includes('吏꾪뻾') || statusStr.includes('ing')) {
            const pKey = Object.keys(r).find(k => k.toLowerCase() === 'partner');
            let pName = pKey ? String(r[pKey]).trim() : 'Unknown';
            if (!pName || pName.toLowerCase() === 'unknown' || pName === '') pName = 'Direct/Unknown';

            if (filterCountry && pName !== 'Direct/Unknown' && !currentPartnerNames.has(pName.toLowerCase())) return;

            const vKey = Object.keys(r).find(k => k.toUpperCase().includes('ESTIMATED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD'));
            const estValTotal = vKey ? parseCurrency(r[vKey]) : 0;

            if (!pRankingData[pName]) pRankingData[pName] = { count: 0, sumValue: 0 };
            pRankingData[pName].count += 1;
            pRankingData[pName].sumValue += estValTotal;
        }
    });

    return {
        counts,
        partnerGroups,
        sortedCountries: Object.keys(partnerGroups).sort((a, b) => partnerGroups[b].length - partnerGroups[a].length),
        sortedP: Object.entries(pRankingData)
            .map(([name, stats]) => ({ name, ...stats }))
            .sort((a, b) => b.count - a.count || b.sumValue - a.sumValue),
        pNameKey
    };
}

function getPartnerHTML(stats, filterCountry, tabName) {
    const displayCountries = CONFIG.COUNTRIES.filter(c => (!filterCountry || filterCountry === 'All') || c === filterCountry);

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

function initPartnerCharts(stats, filterCountry) {
    const selector = document.getElementById('partner-filter-country');
    if (selector) {
        selector.addEventListener('change', (e) => {
            const val = e.target.value;
            const fVal = val === 'All' ? null : val;
            renderTableData(searchInput.value, fVal);
        });
    }

    if (!filterCountry) {
        const ctx = document.getElementById('partner-country-chart');
        if (ctx) {
            // Destroy existing chart if it exists
            chartRegistry.destroyTag('partner-country');

            const chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: stats.sortedCountries,
                    datasets: [{
                        label: 'Partner Count',
                        data: stats.sortedCountries.map(c => stats.partnerGroups[c].length),
                        backgroundColor: 'rgba(0,122,255,0.5)',
                        borderColor: '#007AFF',
                        borderWidth: 2,
                        borderRadius: 10,
                        hoverBackgroundColor: 'rgba(0,122,255,0.85)',
                        barThickness: 'flex',
                        maxBarThickness: 60
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            backgroundColor: '#FFFFFF', titleColor: '#111827', bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1,
                            padding: 14,
                            titleFont: { size: 14, weight: 'bold' },
                            bodyFont: { size: 13 },
                            cornerRoundness: 8
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            grid: { color: 'rgba(0,0,0,0.05)', drawBorder: false },
                            ticks: { color: '#6B7280', font: { size: 11 }, stepSize: 1 }
                        },
                        x: {
                            grid: { display: false },
                            ticks: { color: '#111827', font: { size: 12, weight: '600' } }
                        }
                    }
                },
                plugins: [{
                    id: 'valueLabel',
                    afterDatasetsDraw(chart) {
                        const { ctx } = chart;
                        ctx.save();
                        chart.data.datasets.forEach((dataset, i) => {
                            const meta = chart.getDatasetMeta(i);
                            meta.data.forEach((element, index) => {
                                const value = dataset.data[index];
                                const { x, y } = element.tooltipPosition();
                                ctx.fillStyle = '#111827';
                                ctx.font = 'bold 12px Inter, sans-serif';
                                ctx.textAlign = 'center';
                                ctx.fillText(value, x, y - 10);
                            });
                        });
                        ctx.restore();
                    }
                }]
            });
            chartRegistry.register('partner-country-dist', chart);
        }
    }
}

function renderPartnerMetricsBlock(data, filterCountry, tabName, metricsGrid) {
    if (!data || data.length === 0) return;

    const stats = getPartnerStats(data, filterCountry);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getPartnerHTML(stats, filterCountry, tabName);
    metricsGrid.appendChild(container);

    setTimeout(() => initPartnerCharts(stats, filterCountry), 100);
}


/* --- GENERIC COUNTRY COUNTS HELPERS --- */
function getGenericCountryStats(data, filterCountry) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const countryKey = keys.find(k => k.toLowerCase().includes('country') || k.toLowerCase().includes('region') || k.toLowerCase() === 'nation');
    if (!countryKey) return null;

    const countryCounts = {};
    const yearlyCounts = {};

    const getYear = (row) => {
        const dKey = keys.find(k => k.toUpperCase().replace(/\s/g, '') === 'CONTRACTSTART') ||
            keys.find(k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE')) ||
            keys.find(k => k.toUpperCase().includes('YEAR'));
        const v = row[dKey];
        if (v instanceof Date) return v.getFullYear().toString();
        if (typeof v === 'number' && v > 40000) return new Date(Math.round((v - 25569) * 86400 * 1000)).getFullYear().toString();
        if (v) { const m = String(v).match(/(20\d{2})/); if (m) return m[1]; }
        return 'Unknown';
    };

    data.forEach(row => {
        const c = normalizeCountry(row[countryKey]);
        if (c) {
            countryCounts[c] = (countryCounts[c] || 0) + 1;
            const y = getYear(row);
            if (!yearlyCounts[y]) yearlyCounts[y] = {};
            yearlyCounts[y][c] = (yearlyCounts[y][c] || 0) + 1;
        }
    });

    const sortedTotal = Object.entries(countryCounts).sort(sortCountriesByCount);
    const sortedYears = Object.keys(yearlyCounts).sort((a, b) => (a === 'Unknown' ? 1 : b === 'Unknown' ? -1 : b - a));

    return { sortedTotal, yearlyCounts, sortedYears };
}

function getGenericCountryHTML(stats, filterCountry) {
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

function renderGenericCountryCounts(data, filterCountry, metricsGrid, tabName) {
    if (data.length === 0 || tabName === 'EVENT') return;
    const stats = getGenericCountryStats(data, filterCountry);
    if (!stats) return;
    const div = document.createElement('div');
    div.innerHTML = getGenericCountryHTML(stats, filterCountry);
    metricsGrid.appendChild(div.firstElementChild);
}


/* --- EXPIRING CONTRACTS HELPERS --- */
function getExpiringContractsStats(data) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const endKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTEND')) || keys.find(k => k.toUpperCase().includes('END'));
    if (!endKey) return null;

    const dealNameKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CRMDEALNAME')) || keys.find(k => k.toUpperCase().includes('DEAL') && k.toUpperCase().includes('NAME')) || keys.find(k => k.toUpperCase().includes('DEAL'));
    const clientKey = keys.find(k => k.toUpperCase().includes('CLIENT') || k.toUpperCase().includes('CUSTOMER') || k.toUpperCase().includes('ACCOUNT'));
    const contractYrKey = keys.find(k => k.toUpperCase().replace(/\s/g, '').includes('CONTRACTYR')) || keys.find(k => k.toUpperCase().includes('YR')) || keys.find(k => k.toUpperCase().includes('YEAR'));

    const now = new Date('2026-03-18');
    const threeMonthsLater = new Date(now);
    threeMonthsLater.setMonth(now.getMonth() + 3);

    const getExpDate = (row) => {
        let d = row[endKey];
        if (d instanceof Date) return d;
        if (typeof d === 'number' && d > 40000) return new Date(Math.round((d - 25569) * 86400 * 1000));
        if (d) { const parsed = new Date(d); return isNaN(parsed.getTime()) ? null : parsed; }
        return null;
    };

    const expiringDeals = data.filter(row => {
        const expDate = getExpDate(row);
        return expDate && expDate >= now && expDate <= threeMonthsLater;
    }).sort((a, b) => (getExpDate(a) || 0) - (getExpDate(b) || 0));

    if (expiringDeals.length === 0) return null;

    return expiringDeals.map(deal => {
        const expDate = getExpDate(deal);
        let name = 'Unknown Deal';
        if (dealNameKey && deal[dealNameKey]) name = deal[dealNameKey];
        else if (clientKey && deal[clientKey]) name = deal[clientKey];
        const yr = contractYrKey ? deal[contractYrKey] : '';
        return { name, date: expDate ? expDate.toISOString().split('T')[0] : 'N/A', year: yr };
    });
}

function getExpiringContractsHTML(stats) {
    if (!stats) return '';
    const items = stats.slice(0, 5).map(d => `
        <div style="display: flex; justify-content: space-between; align-items: flex-start; padding: 8px 10px; background: rgba(239, 68, 68, 0.1); border-radius: 8px; border-left: 3px solid #ef4444; margin-bottom: 6px;">
            <div style="display: flex; flex-direction: column;">
                <span style="font-size: 0.8rem; font-weight: 700; color: #111827;">${d.name}</span>
                <span style="font-size: 0.72rem; color: #fca5a5; margin-top: 4px;">End: ${d.date} ${d.year ? `(${d.year}Yr)` : ''}</span>
            </div>
            <span style="font-size: 0.65rem; font-weight: 800; color: #ef4444; text-transform: uppercase;">Soon</span>
        </div>
    `).join('');

    return `
        <div class="stat-card" style="padding: 16px; border: 1px solid rgba(239, 68, 68, 0.4); background: linear-gradient(135deg, rgba(239, 68, 68, 0.1), rgba(249, 115, 22, 0.1)); max-width: 400px; display: block;">
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;">
                <div class="stat-icon" style="background: rgba(239, 68, 68, 0.2); color: #ef4444; width: 36px; height: 36px;"><i class="fa-solid fa-clock"></i></div>
                <div class="stat-details"><h3 style="margin:0; font-size: 0.8rem; color: #fca5a5;">EXPIRING SOON</h3><h2 style="font-size: 0.95rem; font-weight: 600; color: #111827;">Within 3 Months</h2></div>
            </div>
            <div>${items} ${stats.length > 5 ? `<div style="text-align: center; font-size: 0.7rem; color: #64748b; margin-top: 8px;">+ ${stats.length - 5} more</div>` : ''}</div>
        </div>
    `;
}

function renderExpiringContracts(data, metricsGrid) {
    const stats = getExpiringContractsStats(data);
    if (!stats) return;
    const div = document.createElement('div');
    div.innerHTML = getExpiringContractsHTML(stats);
    metricsGrid.appendChild(div.firstElementChild);
}


/* --- PARTNER PERFORMANCE HELPERS --- */
function getPartnerPerformanceStats(data) {
    if (!data || data.length === 0) return null;
    const keys = Object.keys(data[0]);
    const companyKey = keys.find(k => k.toLowerCase().replace(/\s/g, '') === 'companyname');
    const countryKey = keys.find(k => k.toLowerCase().includes('country'));
    const tcvKey = keys.find(k => k.toLowerCase().replace(/\s/g, '') === 'accumulatedtcv');

    if (!companyKey || !tcvKey) return null;

    const stats = data.map(row => ({
        name: String(row[companyKey] || 'Unknown').trim(),
        country: String(row[countryKey] || 'N/A').trim(),
        tcv: parseCurrency(row[tcvKey])
    })).filter(p => p.tcv > 0).sort((a, b) => b.tcv - a.tcv).slice(0, 10);

    return stats.length > 0 ? stats : null;
}

function getPartnerPerformanceHTML() {
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

function initPartnerPerformanceCharts(stats) {
    const ctx = document.getElementById('partner-top-performer-chart');
    if (ctx) {
        chartRegistry.destroyTag('partner-perf');
        chartRegistry.register('partner-perf-summary', new Chart(ctx, {
            type: 'bar',
            data: {
                labels: stats.map(p => `${p.name} (${p.country})`),
                datasets: [{
                    label: 'TCV (USD)',
                    data: stats.map(p => p.tcv),
                    backgroundColor: 'rgba(0,122,255,0.7)',
                    borderRadius: 4
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#FFFFFF', titleColor: '#111827', bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1, padding: 12,
                        callbacks: { label: (ctx) => ' US$ ' + formatCurrency(ctx.parsed.x) }
                    }
                },
                scales: {
                    x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => '$' + formatCurrency(v) } },
                    y: { grid: { display: false }, ticks: { color: '#111827', font: { weight: '500' } } }
                }
            }
        }));
    }
}

function renderPartnerTopPerformer(data, metricsGrid) {
    const stats = getPartnerPerformanceStats(data);
    if (!stats) return;
    const div = document.createElement('div');
    div.style.gridColumn = '1 / -1';
    div.innerHTML = getPartnerPerformanceHTML();
    metricsGrid.appendChild(div);
    setTimeout(() => initPartnerPerformanceCharts(stats), 100);
}

/*═══════════════════════════════════════════════════════════════
  SECTION: POC (Stats, HTML, Charts, Render)
═══════════════════════════════════════════════════════════════*/

/**
 * Calculate POC analytics: running counts, influx data, status distribution,
 * partner bottlenecks, and industry breakdowns.
 * @param {Array<Object>} data - POC data rows
 * @param {{ country: string, industry: string, partner: string }} filters - Active UI filters
 * @returns {{ stats: Object, uniqueValues: Object }}
 */
function getPocStats(data, filters) {
    const sheet9Data = workbookData['Sheet9'] || [];
    const mergedData = data.map(pRow => {
        const nameKey = Object.keys(pRow).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
        const pName = nameKey ? String(pRow[nameKey]).trim().toLowerCase() : '';
        let matchedS9 = sheet9Data.find(sRow => {
            const sNameKey = Object.keys(sRow).find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
            return sNameKey && String(sRow[sNameKey]).trim().toLowerCase() === pName;
        });
        return { ...pRow, _s9: matchedS9 || {} };
    });

    const uniqueValues = { countries: new Set(['All']), industries: new Set(['All']), partners: new Set(['All']) };
    mergedData.forEach(r => {
        const cKey = Object.keys(r).find(k => k.toLowerCase().includes('country'));
        const iKey = Object.keys(r).find(k => k.toLowerCase().includes('industry'));
        const pKey = Object.keys(r).find(k => k.toLowerCase() === 'partner');
        if (cKey && r[cKey]) uniqueValues.countries.add(normalizeCountry(r[cKey]));
        if (iKey && r[iKey]) uniqueValues.industries.add(String(r[iKey]).trim());
        if (pKey && r[pKey]) uniqueValues.partners.add(String(r[pKey]).trim());
    });

    const displayData = mergedData.filter(r => {
        const cKey = Object.keys(r).find(k => k.toLowerCase().includes('country'));
        const rCountry = cKey ? normalizeCountry(r[cKey]) : '';
        const rIndustry = String(r[Object.keys(r).find(k => k.toLowerCase().includes('industry'))] || '').trim();
        const rPartner = String(r[Object.keys(r).find(k => k.toLowerCase() === 'partner')] || '').trim();
        return (filters.country === 'All' || rCountry === filters.country) &&
            (filters.industry === 'All' || rIndustry === filters.industry) &&
            (filters.partner === 'All' || rPartner === filters.partner);
    });

    let s = {
        longTermCount: 0, midTermCount: 0, normalCount: 0, runningList: [], staledRunningList: [], runningNames: [], partnerRunDays: {},
        statusStats: { won: 0, drop: 0, running: 0, hold: 0, others: 0 }, totalWonDays: 0, wonCountForStats: 0,
        influxData: Array(12).fill(0).map(() => ({ count: 0, estimated: 0, weighted: 0, accounts: [] })), industryVal: {}, partnerRankingData: {},
        monthlyActive: Array(12).fill(0).map(() => ({ count: 0, tcv: 0 }))
    };

    displayData.forEach(r => {
        const allKeys = Object.keys(r);
        const s9Keys = Object.keys(r._s9 || {});

        const wdKey = allKeys.find(k => k.toLowerCase().includes('working days')) || allKeys.find(k => k.toLowerCase().includes('workingdays')) || allKeys.find(k => k.toLowerCase().includes('running days'));
        const s9WdKey = s9Keys.find(k => k.toLowerCase().includes('working days')) || s9Keys.find(k => k.toLowerCase().includes('workingdays')) || s9Keys.find(k => k.toLowerCase().includes('running days'));
        let runningDays = Number(r[wdKey] || (r._s9 && r._s9[s9WdKey]) || 0);

        if (runningDays >= 100) s.longTermCount++; else if (runningDays >= 60) s.midTermCount++; else s.normalCount++;

        const statusKey = allKeys.find(k => k.toLowerCase().includes('current status')) || allKeys.find(k => k.toLowerCase().includes('currentstatus')) || allKeys.find(k => k.toLowerCase().includes('status'));
        const curStatus = String(r[statusKey] || '').trim().toLowerCase();
        const startValForStatus = r[allKeys.find(k => k.toLowerCase().replace(/\s/g, '').includes('pocstart') || k.toLowerCase().replace(/\s/g, '').includes('poc?쒖옉'))];
        let dForStatus = startValForStatus ? (startValForStatus instanceof Date ? startValForStatus : new Date(startValForStatus)) : null;
        if (typeof startValForStatus === 'number' && startValForStatus > 30000) dForStatus = new Date(Math.round((startValForStatus - 25569) * 86400 * 1000));

        if (curStatus.includes('running') || curStatus.includes('progress') || curStatus.includes('ing')) {
            s.statusStats.running++;
            const name = r[allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname')] || 'Unknown';
            s.runningNames.push(name);
        } else if (dForStatus && !isNaN(dForStatus.getTime()) && dForStatus.getFullYear() === 2026) {
            if (curStatus.includes('won') || curStatus.includes('complete') || curStatus.includes('success')) s.statusStats.won++;
            else if (curStatus.includes('drop') || curStatus.includes('cancel') || curStatus.includes('fail')) s.statusStats.drop++;
            else if (curStatus.includes('hold') || curStatus.includes('pause')) s.statusStats.hold++;
            else if (curStatus !== '') s.statusStats.others++;
        }

        if (curStatus.includes('running') || curStatus.includes('progress') || runningDays >= 100) {
            const entry = {
                name: r[allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname')] || 'Unknown',
                partner: r[allKeys.find(k => k.toLowerCase() === 'partner')] || 'Unknown',
                country: r[allKeys.find(k => k.toLowerCase().includes('country'))] || 'Unknown',
                industry: r[allKeys.find(k => k.toLowerCase().includes('industry'))] || 'Unknown',
                days: runningDays,
                notes: r._s9[s9Keys.find(k => k.toLowerCase().includes('notes'))] || '',
                techComm: r[allKeys.find(k => k.toLowerCase().includes('technical comment') || k.toLowerCase().includes('report'))] || '',
                isStalled: runningDays >= 100
            };
            s.runningList.push(entry);
            if (runningDays >= 100) s.staledRunningList.push(entry);
        }

        // Prioritize "POC Start" over "License Start" as per user request
        const startValKey = allKeys.find(k => {
            const n = k.toLowerCase().replace(/[^a-z0-9?쒖옉]/g, '');
            return n === 'pocstart' || (n.includes('poc') && n.includes('start')) || n.includes('poc?쒖옉') || n === 'startdate';
        }) || allKeys.find(k => {
            const n = k.toLowerCase().replace(/[^a-z0-9]/g, '');
            return n.includes('license') && n.includes('start');
        });
        const startVal = r[startValKey];

        const estValKey = allKeys.find(k => k.toUpperCase().includes('ESTIMATED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD'));
        const wValKey = allKeys.find(k => k.toUpperCase().includes('WEIGHTED VALUE') && k.toUpperCase().includes('KOR') && k.toUpperCase().includes('USD')) ||
                        allKeys.find(k => k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('USD'));
        
        const estVal = parseCurrency(r[estValKey] || 0);
        const wVal = parseCurrency(r[wValKey] || 0);

        if (startVal) {
            let d = (startVal instanceof Date) ? startVal : new Date(startVal);
            // Robust check for Excel serial date numbers
            if (typeof startVal === 'number' && startVal > 30000) {
                d = new Date(Math.round((startVal - 25569) * 86400 * 1000));
            } else if (typeof startVal === 'string' && !isNaN(startVal) && startVal.length > 4) {
                d = new Date(Math.round((parseFloat(startVal) - 25569) * 86400 * 1000));
            }

            if (d && !isNaN(d.getTime()) && d.getFullYear() === 2026) {
                const m = d.getMonth();
                s.influxData[m].count++; 
                s.influxData[m].estimated += estVal;
                s.influxData[m].weighted += wVal;
                const pocNameKey = allKeys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
                if (pocNameKey && r[pocNameKey]) s.influxData[m].accounts.push(String(r[pocNameKey]).trim());
            }

            // Monthly Active Analysis (2026)
            if (d && !isNaN(d.getTime())) {
                const year2026 = 2026;
                let pocEndDate = new Date(8640000000000000); // Future date for running/hold
                if (curStatus.includes('won') || curStatus.includes('drop') || curStatus.includes('fail') || curStatus.includes('complete')) {
                    pocEndDate = new Date(d.getTime() + (runningDays * 86400 * 1000));
                }

                for (let m = 0; m < 12; m++) {
                    const monthStart = new Date(year2026, m, 1);
                    const monthEnd = new Date(year2026, m + 1, 0);
                    if (d <= monthEnd && pocEndDate >= monthStart) {
                        s.monthlyActive[m].tcv += estVal;
                        s.monthlyActive[m].count++;
                    }
                }
            }
        }

        const pName = String(r[allKeys.find(k => k.toLowerCase() === 'partner')] || 'Unknown').trim();
        if (pName && pName !== 'Unknown' && pName !== '') {
            if (!s.partnerRunDays[pName]) s.partnerRunDays[pName] = { sum: 0, count: 0 };
            s.partnerRunDays[pName].sum += runningDays; s.partnerRunDays[pName].count++;
        }

        if (curStatus.includes('won') || curStatus.includes('complete') || curStatus.includes('success')) {
            const wd = parseCurrency(r[allKeys.find(k => k.toLowerCase().includes('working days'))] || 0);
            if (wd > 0) { s.totalWonDays += wd; s.wonCountForStats++; }
        }

        if (curStatus.includes('running') || curStatus.includes('progress') || curStatus.includes('ing')) {
            const ind = String(r[allKeys.find(k => k.toLowerCase().includes('industry'))] || 'Other').trim();
            s.industryVal[ind] = (s.industryVal[ind] || 0) + estVal;
            if (!s.partnerRankingData[pName]) s.partnerRankingData[pName] = { count: 0, sumValue: 0 };
            s.partnerRankingData[pName].count++; s.partnerRankingData[pName].sumValue += estVal;
        }
    });

    s.runningList.sort((a, b) => b.days - a.days);
    s.partnerAvg = Object.keys(s.partnerRunDays).map(p => ({ partner: p, avg: Math.round(s.partnerRunDays[p].sum / s.partnerRunDays[p].count) })).sort((a, b) => b.avg - a.avg);
    s.sortedIndustry = Object.entries(s.industryVal).map(([name, val]) => ({ name, val })).sort((a, b) => b.val - a.val);
    s.sortedPartnersRanking = Object.entries(s.partnerRankingData).map(([name, st]) => ({ name, ...st })).sort((a, b) => b.count - a.count || b.sumValue - a.sumValue);

    return { stats: s, uniqueValues };
}

function getPocHTML(stats, filters, uniqueValues) {
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

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; margin-bottom: 30px;">
            <div class="stat-card highlight-card" style="background: #EBF4FF; border: 1px solid rgba(0,122,255,0.2); padding: 24px; border-left: 5px solid #007AFF;">
                <div class="stat-icon" style="background: rgba(0, 122, 255, 0.15); color: #007AFF; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-play"></i></div>
                <div>
                    <h3 style="color: #007AFF; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Total Running POCs</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;" title="${stats.runningNames.join('\n')}">${stats.statusStats.running} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Companies</span></h2>
                </div>
            </div>
            <div class="stat-card highlight-card" style="background: #FFF5F5; border: 1px solid rgba(255,59,48,0.2); padding: 24px; border-left: 5px solid #ef4444;">
                <div class="stat-icon" style="background: rgba(239, 68, 68, 0.2); color: #fca5a5; width: 56px; height: 56px; font-size: 1.5rem;"><i class="fa-solid fa-hourglass-half"></i></div>
                <div>
                    <h3 style="color: #FF3B30; font-size: 0.8rem; text-transform: uppercase; font-weight: 700;">Long-term (100+)</h3>
                    <h2 style="color: #111827; font-size: 2.2rem; font-weight: 800; margin: 0;" title="${stats.staledRunningList.map(r => r.name).join('\n')}">${stats.staledRunningList.length} <span style="font-size: 1rem; font-weight: 400; opacity: 0.7;">Companies</span></h2>
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

function initPocCharts(stats) {
    chartRegistry.destroyTag('poc');

    // Influx Chart
    const ctxInflux = document.getElementById('poc-influx-chart');
    if (ctxInflux) {
        chartRegistry.register('poc-influx', new Chart(ctxInflux, {
            type: 'bar',
            data: {
                labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
                datasets: [
                    { label: 'New POC Starts', data: stats.influxData.map(d => d.count), backgroundColor: 'rgba(52,199,89,0.55)', yAxisID: 'yCount' },
                    { label: 'Estimated Value (USD)', data: stats.influxData.map(d => d.estimated), type: 'line', borderColor: '#007AFF', yAxisID: 'yValue' },
                    { label: 'Weighted Value (USD)', data: stats.influxData.map(d => d.weighted), type: 'line', borderColor: '#6366f1', borderDash: [5, 5], yAxisID: 'yValue' }
                ]
            },
            options: { responsive: true, maintainAspectRatio: false, scales: { yCount: { position: 'left' }, yValue: { position: 'right' } } }
        }));
    }

    // Status Chart
    const ctxStatus = document.getElementById('poc-status-chart');
    if (ctxStatus) {
        chartRegistry.register('poc-status', new Chart(ctxStatus, {
            type: 'doughnut',
            data: {
                labels: ['Won', 'Drop', 'Running', 'Hold', 'Others'],
                datasets: [{ data: [stats.statusStats.won, stats.statusStats.drop, stats.statusStats.running, stats.statusStats.hold, stats.statusStats.others], backgroundColor: ['#34C759', '#FF3B30', '#007AFF', '#FF9500', '#9CA3AF'] }]
            },
            options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { legend: { position: 'bottom' } } },
            plugins: [{
                id: 'doughnutLabelStatus',
                afterDatasetsDraw(chart) {
                    const { ctx, data } = chart;
                    ctx.save();
                    chart.getDatasetMeta(0).data.forEach((element, index) => {
                        const value = data.datasets[0].data[index];
                        if (value > 0) {
                            const position = element.tooltipPosition();
                            ctx.fillStyle = '#FFFFFF';
                            ctx.font = 'bold 13px Inter, sans-serif';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'middle';
                            ctx.fillText(value, position.x, position.y);
                        }
                    });
                    ctx.restore();
                }
            }]
        }));
    }

    // Aging Chart
    const ctxAging = document.getElementById('poc-aging-chart');
    if (ctxAging) {
        chartRegistry.register('poc-aging', new Chart(ctxAging, {
            type: 'doughnut',
            data: {
                labels: ['100+', '60-100', '<60'],
                datasets: [{ data: [stats.longTermCount, stats.midTermCount, stats.normalCount], backgroundColor: ['#FF3B30', '#FF9500', '#34C759'] }]
            },
            options: { responsive: true, maintainAspectRatio: false, cutout: '65%', plugins: { legend: { position: 'bottom' } } }
        }));
    }

    // Bottleneck Chart
    const ctxB = document.getElementById('poc-bottleneck-chart');
    if (ctxB && stats.partnerAvg.length > 0) {
        chartRegistry.register('poc-bottleneck', new Chart(ctxB, {
            type: 'bar',
            data: { labels: stats.partnerAvg.slice(0, 10).map(p => p.partner), datasets: [{ data: stats.partnerAvg.slice(0, 10).map(p => p.avg), backgroundColor: '#007AFF' }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        }));
    }

    // Industry Chart
    const ctxI = document.getElementById('poc-industry-chart');
    if (ctxI && stats.sortedIndustry.length > 0) {
        chartRegistry.register('poc-industry', new Chart(ctxI, {
            type: 'bar',
            data: { labels: stats.sortedIndustry.slice(0, 10).map(i => i.name), datasets: [{ data: stats.sortedIndustry.slice(0, 10).map(i => i.val), backgroundColor: '#a78bfa' }] },
            options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        }));
    }
}

function renderPocMetricsBlock(data, filterCountry, metricsGrid) {
    metricsGrid.innerHTML = '';
    const pocContainer = document.createElement('div');
    pocContainer.id = 'poc-dashboard-container';
    pocContainer.style.gridColumn = '1 / -1';
    metricsGrid.appendChild(pocContainer);

    window.pocFilters = window.pocFilters || { country: 'All', industry: 'All', partner: 'All' };

    window.renderPocUI = function () {
        const { stats, uniqueValues } = getPocStats(data, window.pocFilters);
        const container = document.getElementById('poc-dashboard-container');
        if (container) {
            container.innerHTML = getPocHTML(stats, window.pocFilters, uniqueValues);

            // Attach event listeners
            document.getElementById('poc-filter-country').addEventListener('change', (e) => { window.pocFilters.country = e.target.value; window.renderPocUI(); });
            document.getElementById('poc-filter-industry').addEventListener('change', (e) => { window.pocFilters.industry = e.target.value; window.renderPocUI(); });
            document.getElementById('poc-filter-partner').addEventListener('change', (e) => { window.pocFilters.partner = e.target.value; window.renderPocUI(); });

            setTimeout(() => initPocCharts(stats), 50);
        }
    };

    // Render UI initially
    window.renderPocUI();
}

/*═══════════════════════════════════════════════════════════════
  SECTION: EVENT (Stats, HTML, Charts, Render)
═══════════════════════════════════════════════════════════════*/

/* --- EVENT METRICS HELPERS --- */
/**
 * Calculate event analytics: spending, POC generation, deal conversion.
 * @param {Array<Object>} eventData - Event data rows
 * @param {string|null} filterCountry - Country filter
 * @returns {{ totalSpending, totalPOC, totalDeals, eventCount, comparisonData, costPerPOC, costPerDeal }|null}
 */
function getEventStats(eventData, filterCountry) {
    const filteredEvents = eventData.filter(row => {
        if (filterCountry && !isCountryMatch(row, filterCountry)) return false;
        const dateKey = Object.keys(row).find(k => k.toLowerCase().includes('date'));
        const dateVal = dateKey ? row[dateKey] : null;
        let year = 0;
        if (dateVal instanceof Date) year = dateVal.getFullYear();
        else if (typeof dateVal === 'number' && dateVal > 40000) year = new Date(Math.round((dateVal - 25569) * 86400 * 1000)).getFullYear();
        else { const match = String(dateVal).match(/(20\d{2})/); if (match) year = parseInt(match[1]); }
        return year === 2026;
    });

    if (filteredEvents.length === 0) return null;

    let totalSpending = 0, totalPOC = 0, totalDeals = 0;
    const comparisonData = filteredEvents.map(row => {
        const sKey = Object.keys(row).find(k => k.toLowerCase().includes('spending'));
        const pKey = Object.keys(row).find(k => k.toLowerCase().includes('poc') && k.toLowerCase().includes('generated'));
        const dKey = Object.keys(row).find(k => k.toLowerCase().includes('deal') && (k.toLowerCase().includes('converted') || k.toLowerCase().includes('closed')));
        const nKey = Object.keys(row).find(k => k.toLowerCase().includes('event') || k.toLowerCase().includes('name') || k.toLowerCase().includes('?щ챸')) || Object.keys(row)[0];

        const spend = parseCurrency(row[sKey]);
        const pocs = parseInt(row[pKey]) || 0;
        const deals = parseInt(row[dKey]) || 0;

        totalSpending += spend;
        totalPOC += pocs;
        totalDeals += deals;
        return { name: String(row[nKey] || 'Unnamed').trim(), spend, pocs, deals };
    });

    return { totalSpending, totalPOC, totalDeals, eventCount: filteredEvents.length, comparisonData, costPerPOC: totalPOC > 0 ? (totalSpending / totalPOC) : 0, costPerDeal: totalDeals > 0 ? (totalSpending / totalDeals) : 0 };
}

function getEventHTML(stats) {
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

function initEventCharts(stats) {
    if (!stats) return;
    chartRegistry.destroyTag('event');
    const ctx = document.getElementById('event-roi-chart');
    if (ctx) {
        chartRegistry.register('event-roi', new Chart(ctx, {
            type: 'bar',
            data: {
                labels: stats.comparisonData.map(d => d.name),
                datasets: [
                    { label: 'Spending (USD)', data: stats.comparisonData.map(d => d.spend), backgroundColor: 'rgba(0, 122, 255, 0.65)', yAxisID: 'ySpend' },
                    { label: 'POCs', data: stats.comparisonData.map(d => d.pocs), backgroundColor: 'rgba(52, 199, 89, 0.65)', yAxisID: 'yPoc' }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                scales: {
                    ySpend: { position: 'left', ticks: { callback: v => '$' + formatCurrency(v) } },
                    yPoc: { position: 'right', grid: { display: false } }
                }
            }
        }));
    }
}

function renderEventMetrics(eventData, filterCountry, metricsGrid) {
    const stats = getEventStats(eventData, filterCountry);
    if (!stats) return;

    const eventHeader = document.createElement('div');
    eventHeader.style.gridColumn = '1 / -1';
    eventHeader.style.marginBottom = '25px';
    eventHeader.innerHTML = getEventHTML(stats);
    metricsGrid.prepend(eventHeader);

    setTimeout(() => initEventCharts(stats), 150);
}


/*═══════════════════════════════════════════════════════════════
  SECTION: COUNTRY-SPECIFIC METRICS
═══════════════════════════════════════════════════════════════*/

/* --- COUNTRY SPECIFIC METRICS HELPERS --- */
/**
 * Calculate yearly TCV/ARR breakdown for country-specific views.
 * @param {Array<Object>} data - Order data filtered by country
 * @returns {{ summary: Object, sortedYears: string[] }}
 */
function getCountrySpecificStats(data) {
    const sortedYears = ['2026', '2025', '2024', '2023'];
    const summary = {};
    sortedYears.forEach(y => summary[y] = { kTcv: 0, lArr: 0 });

    data.forEach(row => {
        const dKey = Object.keys(row).find(k => k.toLowerCase().includes('contractstart')) || Object.keys(row).find(k => k.toLowerCase().includes('date'));
        if (!row[dKey]) return;
        let d = (row[dKey] instanceof Date) ? row[dKey] : (typeof row[dKey] === 'number' && row[dKey] > 40000) ? new Date(Math.round((row[dKey] - 25569) * 86400 * 1000)) : new Date(row[dKey]);
        if (d && !isNaN(d.getTime())) {
            const y = d.getFullYear().toString();
            if (summary[y]) {
                summary[y].kTcv += parseCurrency(row['KOR TCV(USD)'] || row['KOR TCV']);
                summary[y].lArr += parseCurrency(row['Local ARR'] || row['ARR']);
            }
        }
    });
    return { summary, sortedYears };
}

function getCountrySpecificHTML(stats, countryName) {
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

function renderCountrySpecificMetrics(data, countryName) {
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    const stats = getCountrySpecificStats(data);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getCountrySpecificHTML(stats, countryName);
    metricsGrid.appendChild(container);
}
/*═══════════════════════════════════════════════════════════════
  SECTION: Currency Utilities (used globally)
═══════════════════════════════════════════════════════════════*/

/**
 * Parse a currency string/number into a numeric value.
 * Strips all non-numeric characters except minus and decimal.
 * @param {*} val - Raw value
 * @returns {number}
 */
function parseCurrency(val) {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    const clean = String(val).replace(/[^0-9.-]+/g, '');
    return parseFloat(clean) || 0;
}

/**
 * Format a number as currency string with locale-appropriate grouping.
 * @param {number} val - Value to format
 * @param {boolean} [isKRW=false] - Use Korean locale formatting
 * @returns {string}
 */
function formatCurrency(val, isKRW = false) {
    if (!val) return '0';
    const locale = isKRW ? 'ko-KR' : 'en-US';
    return new Intl.NumberFormat(locale, {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(Math.round(val));
}

/*═══════════════════════════════════════════════════════════════
  SECTION: Event Listeners
═══════════════════════════════════════════════════════════════*/

/** Global search input handler */
searchInput.addEventListener('input', (e) => {
    if (currentTab) {
        const activeSubSpan = document.querySelector('.nav-sublist .sub-item.active span');
        const country = activeSubSpan ? activeSubSpan.innerText : null;
        renderTableData(e.target.value, country === 'All' ? null : country);
    }
});