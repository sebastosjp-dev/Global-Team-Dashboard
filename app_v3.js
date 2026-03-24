// state
let workbookData = {};
let currentTab = null;

// --- Update Last Updated Date ---
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

// --- Auto Load from Local Server ---
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

            let targetSheetName = workbookMrr.SheetNames.find(name => name.includes('Global(계약시점)'));
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
        alert('Could not load Excel files. Please ensure the local server is running.');
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
        // Convert sheet to JSON array of objects, forcing string evaluation for dates
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "", cellDates: true });
        workbookData[sheetName] = json;
    });

    // Setup UI
    dropZone.classList.remove('active');
    dashboardContainer.classList.add('active');

    buildSidebar(sheetNames);
    updateLastUpdatedDate();

    // Select first tab by default
    if (sheetNames.length > 0) {
        const firstTab = sheetNames.includes('ORDER SHEET') ? 'ORDER SHEET' : sheetNames[0];
        selectTab(firstTab);
    }
}

// --- Dashboard UI ---
function buildSidebar(sheetNames) {
    sidebarNav.innerHTML = '';

    sheetNames.forEach(name => {
        if (name.includes('Global(계약시점)') || ['Sheet9', 'Sheet10', 'Sheet12', 'Sheet13', '2026 Q1 Review'].includes(name)) return;

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
            const countries = ['Indonesia', 'Thailand', 'Malaysia', 'USA', 'Philippines', 'Singapore', 'Turkey'];

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

function selectRenewalView() {
    currentTab = 'RENEWAL_VIEW';
    const title = 'Churn Prevention & Renewal Management (CSM)';
    currentTabTitle.innerText = title;

    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.renewal-tab')?.classList.add('active');

    renderRenewalTable();
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
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    dataTable.classList.add('hidden');
    emptyState.classList.add('hidden');

    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    metricsGrid.classList.remove('hidden');

    // Tally service combinations
    const comboCounts = {};
    const singleServiceCounts = {};
    const allServices = new Set();

    const activeData = data.filter(r => r['Status'] === 'Active' && r['Services']);

    activeData.forEach(row => {
        const raw = String(row['Services'] || '').trim();
        if (!raw) return;
        // Normalize: sort services alphabetically to consolidate duplicates
        const combo = raw.split(',').map(s => s.trim()).filter(Boolean).sort().join(' + ');
        comboCounts[combo] = (comboCounts[combo] || 0) + 1;

        // Individual service usage
        raw.split(',').map(s => s.trim()).filter(Boolean).forEach(s => {
            allServices.add(s);
            singleServiceCounts[s] = (singleServiceCounts[s] || 0) + 1;
        });
    });

    const sortedCombos = Object.entries(comboCounts).sort((a, b) => b[1] - a[1]);
    const totalCustomers = activeData.length;
    const singleServiceCustomers = activeData.filter(r => {
        const services = String(r['Services'] || '').split(',').map(s => s.trim()).filter(Boolean);
        return services.length === 1;
    }).length;
    const multiServiceCustomers = totalCustomers - singleServiceCustomers;

    // Color palette for chart
    const palette = [
        '#6366f1','#0ea5e9','#10b981','#f59e0b','#ef4444','#a855f7',
        '#ec4899','#14b8a6','#f97316','#84cc16'
    ];

    const labels = sortedCombos.map(c => c[0]);
    const counts = sortedCombos.map(c => c[1]);
    const colors = labels.map((_, i) => palette[i % palette.length]);

    // Upsell targets: customers with only 1 service
    const upsellTargets = data.filter(r => {
        if (r['Status'] !== 'Active' || !r['Services']) return false;
        const services = String(r['Services']).split(',').map(s => s.trim()).filter(Boolean);
        return services.length === 1;
    }).map(r => ({
        name: r['End User'],
        country: r['Country'],
        service: String(r['Services']).trim(),
        tcv: parseCurrency(r['TCV Amount'])
    })).sort((a, b) => b.tcv - a.tcv);

    // Inject tooltip CSS once
    if (!document.getElementById('insight-tooltip-style')) {
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
            .insight-popup::before {
                content: ''; position: absolute; bottom: 100%; left: 50%; transform: translateX(-50%);
                border: 6px solid transparent; border-bottom-color: #E5E7EB;
            }
            .insight-info-icon:hover + .insight-popup,
            .insight-popup:hover { display: block; }
            .insight-popup-title { font-size: 0.8rem; font-weight: 700; color: #34C759; margin-bottom: 8px; display: flex; align-items: center; gap: 6px; }
            .insight-popup-body { font-size: 0.75rem; color: #6B7280; line-height: 1.6; }
            .insight-popup-body strong { color: #374151; }
            .insight-popup-tag { display: inline-block; background: rgba(99,102,241,0.15); color: #34C759; padding: 1px 7px; border-radius: 20px; font-size: 0.7rem; font-weight: 600; margin: 2px 2px 0 0; }
        `;
        document.head.appendChild(style);
    }

    // Build layout
    metricsGrid.innerHTML = `
        <!-- Summary Cards -->
        <div class="stat-card" style="border-left: 4px solid #6366f1;">
            <div class="stat-details">
                <h3 style="display:flex; align-items:center;">Active Customers
                    <span class="insight-tooltip-wrap">
                        <span class="insight-info-icon"><i class="fa-solid fa-info"></i></span>
                        <div class="insight-popup">
                            <div class="insight-popup-title"><i class="fa-solid fa-users"></i> Active Customers</div>
                            <div class="insight-popup-body">
                                Total number of customer accounts with <strong>Status = Active</strong> in the END USER (CSM) sheet.<br><br>
                                This is the baseline universe for all upsell and service analysis. Only active accounts are considered for opportunity scoring.
                            </div>
                        </div>
                    </span>
                </h3>
                <h2>${totalCustomers} <span>Accounts</span></h2>
            </div>
            <div class="stat-icon" style="background:rgba(99,102,241,0.1); color:#6366f1;"><i class="fa-solid fa-users"></i></div>
        </div>

        <div class="stat-card" style="border-left: 4px solid #10b981;">
            <div class="stat-details">
                <h3 style="display:flex; align-items:center;">Multi-Service Customers
                    <span class="insight-tooltip-wrap">
                        <span class="insight-info-icon"><i class="fa-solid fa-info"></i></span>
                        <div class="insight-popup">
                            <div class="insight-popup-title"><i class="fa-solid fa-layer-group"></i> Multi-Service Customers</div>
                            <div class="insight-popup-body">
                                Customers already using <strong>2 or more services</strong> (e.g., APM + DBM + SVM).<br><br>
                                These accounts have demonstrated <strong>high platform adoption</strong> and trust. They are strong candidates for <strong>expansion</strong> into newer modules like <span class="insight-popup-tag">K8S</span> <span class="insight-popup-tag">NWM</span> <span class="insight-popup-tag">LOG</span>.
                            </div>
                        </div>
                    </span>
                </h3>
                <h2>${multiServiceCustomers} <span>Accounts</span></h2>
            </div>
            <div class="stat-icon" style="background:rgba(16,185,129,0.1); color:#10b981;"><i class="fa-solid fa-layer-group"></i></div>
        </div>

        <div class="stat-card" style="border-left: 4px solid #f59e0b;">
            <div class="stat-details">
                <h3 style="display:flex; align-items:center;">Upsell Targets (Single Service)
                    <span class="insight-tooltip-wrap">
                        <span class="insight-info-icon"><i class="fa-solid fa-info"></i></span>
                        <div class="insight-popup">
                            <div class="insight-popup-title"><i class="fa-solid fa-arrow-trend-up"></i> Upsell Targets</div>
                            <div class="insight-popup-body">
                                Customers currently using <strong>only 1 service</strong>. These accounts have untapped potential — they are already on the platform but haven't expanded their footprint.<br><br>
                                Priority is ranked by <strong>TCV (Total Contract Value)</strong>: higher TCV means a larger existing relationship and a stronger mandate for a cross-sell conversation.
                            </div>
                        </div>
                    </span>
                </h3>
                <h2>${singleServiceCustomers} <span>Accounts</span></h2>
            </div>
            <div class="stat-icon" style="background:rgba(245,158,11,0.1); color:#f59e0b;"><i class="fa-solid fa-arrow-trend-up"></i></div>
        </div>

        <!-- Charts Row -->
        <div style="grid-column: 1 / -1; display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 8px;">

            <!-- Donut Chart -->
            <div class="stat-card" style="padding: 24px; display: flex; flex-direction: column; margin: 0;">
                <h3 style="color:#111827; font-size:1rem; font-weight:600; margin-bottom: 4px; display:flex; align-items:center;">
                    Service Combination Ranking
                    <span class="insight-tooltip-wrap">
                        <span class="insight-info-icon"><i class="fa-solid fa-info"></i></span>
                        <div class="insight-popup" style="left: auto; right: 0; transform: none;">
                            <div class="insight-popup-title"><i class="fa-solid fa-chart-pie"></i> Service Combination Ranking</div>
                            <div class="insight-popup-body">
                                This donut chart shows <strong>how many active customers use each unique service bundle</strong>.<br><br>
                                Larger slices = more common combinations. Use this to:<br>
                                • Identify your <strong>standard "starter pack"</strong> (most common combo).<br>
                                • Spot which services are almost always paired together.<br>
                                • Find <strong>rare combos</strong> that may indicate specialised use cases worth replicating.
                            </div>
                        </div>
                    </span>
                </h3>
                <p style="color:var(--text-muted); font-size:0.72rem; margin-bottom: 16px;">Distribution of service combos among active customers</p>
                <div style="position:relative; height: 280px;">
                    <canvas id="service-donut-chart"></canvas>
                </div>
            </div>

            <!-- Upsell Target Table -->
            <div class="stat-card" style="padding: 24px; display: flex; flex-direction: column; margin: 0; overflow: hidden;">
                <h3 style="color:#111827; font-size:1rem; font-weight:600; margin-bottom: 4px; display:flex; align-items:center;">
                    Upsell Target List
                    <span class="insight-tooltip-wrap">
                        <span class="insight-info-icon"><i class="fa-solid fa-info"></i></span>
                        <div class="insight-popup" style="left: auto; right: 0; transform: none;">
                            <div class="insight-popup-title"><i class="fa-solid fa-star"></i> Why These Companies?</div>
                            <div class="insight-popup-body">
                                These are <strong>active customers using only a single service</strong>, ranked by <strong>TCV (highest to lowest)</strong>.<br><br>
                                <strong>High TCV = High Opportunity Cost:</strong> A customer already paying a large contract for one service has proven budget and willingness to invest. Adding even one more module could significantly increase ARR.<br><br>
                                <strong>Example:</strong> <span class="insight-popup-tag">K8S-only</span> customers are natural targets for <span class="insight-popup-tag">APM</span> or <span class="insight-popup-tag">SVM</span> — the most common next step among similar customers.
                            </div>
                        </div>
                    </span>
                </h3>
                <p style="color:var(--text-muted); font-size:0.72rem; margin-bottom: 16px;">Single-service customers sorted by TCV (highest opportunity first)</p>
                <div style="overflow-y: auto; flex: 1; max-height: 280px;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.8rem;">
                        <thead style="position:sticky; top:0; background:#F9FAFB;">
                            <tr>
                                <th style="padding:8px 10px; text-align:left; color:#64748b; font-weight:600;">End User</th>
                                <th style="padding:8px 10px; text-align:left; color:#64748b; font-weight:600;">Country</th>
                                <th style="padding:8px 10px; text-align:left; color:#64748b; font-weight:600;">Service</th>
                                <th style="padding:8px 10px; text-align:right; color:#64748b; font-weight:600;">TCV</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${upsellTargets.map(t => `
                                <tr style="border-top: 1px solid #F3F4F6;">
                                    <td style="padding:8px 10px; color:#111827; font-weight:500;">${t.name}</td>
                                    <td style="padding:8px 10px; color:#4B5563;">${t.country}</td>
                                    <td style="padding:8px 10px;"><span style="background:rgba(99,102,241,0.15); color:#818cf8; padding:2px 8px; border-radius:20px; font-size:0.75rem; font-weight:600;">${t.service}</span></td>
                                    <td style="padding:8px 10px; color:#10b981; font-weight:600;">
                                        <div style="display: flex; justify-content: space-between; align-items: baseline; gap: 8px;">
                                            <span style="font-size: 0.7rem; color: #64748b; font-weight: 400;">US$</span>
                                            <span>${formatCurrency(t.tcv)}</span>
                                        </div>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    ${upsellTargets.length === 0 ? '<p style="text-align:center; color:#64748b; padding:40px;">No single-service customers found.</p>' : ''}
                </div>
            </div>
        </div>

        <!-- Combo Breakdown Bar -->
        <div style="grid-column: 1 / -1;">
            <div class="stat-card" style="padding: 24px; margin: 0;">
                <h3 style="color:#111827; font-size:1rem; font-weight:600; margin-bottom: 4px; display:flex; align-items:center;">
                    Combo Breakdown
                    <span class="insight-tooltip-wrap">
                        <span class="insight-info-icon"><i class="fa-solid fa-info"></i></span>
                        <div class="insight-popup">
                            <div class="insight-popup-title"><i class="fa-solid fa-bars-progress"></i> Combo Breakdown</div>
                            <div class="insight-popup-body">
                                A ranked bar view of <strong>all unique service combinations</strong> by customer count.<br><br>
                                Use this to benchmark a prospect against existing customers. If a prospect uses <span class="insight-popup-tag">APM</span> only, the bar shows that most similar customers also adopt <span class="insight-popup-tag">DBM</span> and <span class="insight-popup-tag">SVM</span> — giving you a data-backed upsell pitch.
                            </div>
                        </div>
                    </span>
                </h3>
                <p style="color:var(--text-muted); font-size:0.72rem; margin-bottom: 20px;">All unique service combinations ranked by customer count</p>
                <div style="display: flex; flex-direction: column; gap: 10px;">
                    ${sortedCombos.map(([combo, count], i) => {
                        const pct = totalCustomers > 0 ? Math.round(count / totalCustomers * 100) : 0;
                        const color = palette[i % palette.length];
                        return `
                        <div>
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                                <span style="font-size:0.82rem; color:#e2e8f0; font-weight:500;">${combo}</span>
                                <span style="font-size:0.8rem; color:${color}; font-weight:700;">${count} customers (${pct}%)</span>
                            </div>
                            <div style="height:8px; background:#F3F4F6; border-radius:4px; overflow:hidden;">
                                <div style="height:100%; width:${pct}%; background:${color}; border-radius:4px; transition: width 0.6s ease;"></div>
                            </div>
                        </div>`;
                    }).join('')}
                </div>
            </div>
        </div>
    `;

    // Render donut chart
    setTimeout(() => {
        const ctx = document.getElementById('service-donut-chart');
        if (ctx) {
            if (window.serviceDonutChart) window.serviceDonutChart.destroy();
            window.serviceDonutChart = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels,
                    datasets: [{
                        data: counts,
                        backgroundColor: colors,
                        borderColor: 'rgba(15, 23, 42, 0.8)',
                        borderWidth: 3,
                        hoverOffset: 8
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    cutout: '60%',
                    plugins: {
                        legend: {
                            position: 'right',
                            labels: {
                                color: '#94a3b8',
                                font: { size: 11 },
                                padding: 12,
                                usePointStyle: true,
                                pointStyleWidth: 8
                            }
                        },
                        tooltip: { backgroundColor: '#FFFFFF', titleColor: '#111827', bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1, bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1,
                            bodyColor: '#374151',
                            padding: 12,
                            callbacks: {
                                label: function(context) {
                                    const pct = totalCustomers > 0 ? Math.round(context.parsed / totalCustomers * 100) : 0;
                                    return ` ${context.parsed} customers (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }
    }, 100);
}

function renderRenewalTable() {
    const data = workbookData['END USER (CSM)'] || [];
    tableHead.innerHTML = ''; tableBody.innerHTML = '';
    const today = new Date('2026-03-18');
    const sixMonthsLater = new Date(today);
    sixMonthsLater.setMonth(today.getMonth() + 6);
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = ''; metricsGrid.classList.remove('hidden');
    const headers = ['End User', 'Country', 'Status', 'End License Date', 'D-Day', 'TCV Amount', 'ARR Amount', 'Probability'];
    const filtered = data.filter(row => {
        let endDate = row['End License Date'];
        if (!endDate) return false;
        let d = (endDate instanceof Date) ? endDate : (typeof endDate === 'number') ? new Date(Math.round((endDate - 25569) * 86400 * 1000)) : new Date(endDate);
        return d >= today && d <= sixMonthsLater;
    }).map(row => {
        let endDate = row['End License Date'];
        let d = (endDate instanceof Date) ? endDate : (typeof endDate === 'number') ? new Date(Math.round((endDate - 25569) * 86400 * 1000)) : new Date(endDate);
        const diffDays = Math.ceil((d - today) / (1000 * 60 * 60 * 24));
        return {
            ...row,
            'D-Day': diffDays === 0 ? 'D-Day' : (diffDays > 0 ? `D-${diffDays}` : `D+${Math.abs(diffDays)}`),
            'diffDays': diffDays, 'endDateFormatted': d.toISOString().split('T')[0]
        };
    });
    filtered.sort((a, b) => a.diffDays - b.diffDays);
    const trHead = document.createElement('tr'); headers.forEach(h => { const th = document.createElement('th'); th.innerText = h; trHead.appendChild(th); });
    tableHead.appendChild(trHead);
    filtered.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
            const td = document.createElement('td');
            td.innerText = h === 'End License Date' ? row.endDateFormatted : (row[h] || '');
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });
}
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
    let v = String(row[k] || "").trim().toUpperCase();
    if (v === 'IDN' || v.includes('INDONESIA')) v = 'Indonesia';
    else if (v === 'US' || v === 'USA' || v.includes('UNITED STATES')) v = 'USA';
    else if (v.includes('MALAYSIA')) v = 'Malaysia';
    else if (v.includes('THAILAND')) v = 'Thailand';
    return v === filterCountry;
}
function renderTableData(searchTerm = "", filterCountry = null) {
    const data = workbookData[currentTab] || [];
    tableHead.innerHTML = ''; tableBody.innerHTML = '';
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

function renderTabMetrics(data, tabName, filterCountry = null) {
    const metricsGrid = document.getElementById('tab-metrics-grid'); metricsGrid.innerHTML = '';
    const isGlobalTab = tabName && tabName.includes('Global(계약시점)');
    if (tabName === 'ORDER SHEET' || isGlobalTab) {
        const orderData = tabName === 'ORDER SHEET' ? data : (workbookData['ORDER SHEET'] || []);
        const keys = orderData.length > 0 ? Object.keys(orderData[0]) : [];
        const korTcvKey = keys.find(k => k.toUpperCase().replace(/\s/g,'') === 'KORTCV') || keys.find(k => k.toUpperCase().includes('TCV') && (k.toUpperCase().includes('KOR') || k.toUpperCase().includes('KRW')));
        const startDateKey = keys.find(k => k.toUpperCase().replace(/\s/g,'').includes('CONTRACTSTART')) || keys.find(k => k.toUpperCase().includes('START') && k.toUpperCase().includes('DATE'));
        let sumLocalTcv = 0, sumKorTcv = 0, yearlyTcv = {};
        const qSums = { 'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0 };
        orderData.filter(r => isCountryMatch(r, filterCountry)).forEach(row => {
            const lTcv = parseCurrency(row['Local TCV'] || row['Local TCV Amount']);
            const kTcv = korTcvKey ? parseCurrency(row[korTcvKey]) : 0;
            sumLocalTcv += lTcv; sumKorTcv += kTcv;
            const cStart = row[startDateKey];
            if (cStart) {
                let d = (cStart instanceof Date) ? cStart : (typeof cStart === 'number' && cStart > 40000) ? new Date(Math.round((cStart - 25569) * 86400 * 1000)) : new Date(cStart);
                if (d && !isNaN(d.getTime())) {
                    const y = d.getFullYear().toString();
                    if (!yearlyTcv[y]) yearlyTcv[y] = { local: 0, korea: 0 };
                    yearlyTcv[y].local += lTcv; yearlyTcv[y].korea += kTcv;
                    if (d.getFullYear() === 2026) {
                        qSums[`Q${Math.floor(d.getMonth() / 3) + 1}`] += kTcv;
                    }
                }
            }
        });
        const container = document.createElement('div');
        container.style.gridColumn = '1 / -1';
        container.innerHTML = `
            <div style="display: grid; grid-template-columns: 280px 320px 1fr; gap: 16px; margin-bottom: 24px;">
                <div class="stat-card" style="background:#FFFFFF; border-left: 5px solid #6366f1; padding:20px;">
                    <h3 style="color:#6366f1; font-size:0.75rem; font-weight:700;">ACCUMULATED KOR TCV</h3>
                    <h2 style="font-size:1.6rem; font-weight:800; margin-top:8px;">US$ ${formatCurrency(sumKorTcv)}</h2>
                </div>
                <div class="stat-card" style="background:#FFFFFF; padding:18px;">
                    <h3 style="font-size:0.85rem; font-weight:600; margin-bottom:12px;">Quarterly TCV (2026)</h3>
                    <div style="height:160px; position:relative;"><canvas id="quarterly-tcv-bar"></canvas></div>
                </div>
                <div class="stat-card" style="background:#FFFFFF; padding:18px;">
                    <h3 style="font-size:0.85rem; font-weight:600; margin-bottom:12px;">Yearly KOR TCV Growth</h3>
                    <div style="height:160px; position:relative;"><canvas id="tcv-growth-chart"></canvas></div>
                </div>
            </div>`;
        metricsGrid.appendChild(container);
        setTimeout(() => {
            const barCtx = document.getElementById('quarterly-tcv-bar').getContext('2d');
            new Chart(barCtx, {
                type: 'bar',
                data: { labels: ['Q1', 'Q2', 'Q3', 'Q4'], datasets: [{ label: 'KOR TCV', data: [qSums.Q1, qSums.Q2, qSums.Q3, qSums.Q4], backgroundColor: 'rgba(99, 102, 241, 0.7)', borderRadius: 4 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { callback: (v) => '$' + formatCurrency(v)}}}}
            });
            const lineCtx = document.getElementById('tcv-growth-chart').getContext('2d');
            const gYears = Object.keys(yearlyTcv).sort();
            new Chart(lineCtx, { type: 'line', data: { labels: gYears, datasets: [{ label: 'KOR TCV', data: gYears.map(y => yearlyTcv[y].korea), borderColor: '#6366f1', tension: 0.3 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } } });
        }, 120);
    }
    // Summary Cards for PIPELINE, PARTNER, POC, EVENT (simplified restoration)
    if (tabName === 'PIPELINE' && workbookData['PIPELINE']) {
        const pData = filterCountry ? workbookData['PIPELINE'].filter(r => isCountryMatch(r, filterCountry)) : workbookData['PIPELINE'];
        let total = 0, weighted = 0; pData.forEach(r => { total += parseCurrency(r['KOR TCV(USD)'] || r['Amount']); weighted += parseCurrency(r['Weighted KOR TCV'] || r['Weighted Amount']); });
        const c = document.createElement('div'); c.style.gridColumn='1/-1'; c.innerHTML = `<div class="stat-card" style="background:#F0FFF4; border-left: 5px solid #48BB78; padding:20px;"><h3>PIPELINE TOTAL</h3><h2>US$ ${formatCurrency(total)}</h2><h3>WEIGHTED</h3><h2>US$ ${formatCurrency(weighted)}</h2></div>`;
        metricsGrid.appendChild(c);
    }
}

function renderCountrySpecificMetrics(data, countryName) {
    const metricsGrid = document.getElementById('tab-metrics-grid'); metricsGrid.innerHTML = '';
    const sortedYears = ['2026', '2025', '2024', '2023'];
    const summary = {}; sortedYears.forEach(y => summary[y] = { kTcv: 0, lArr: 0 });
    data.forEach(row => {
        const dKey = Object.keys(row).find(k => k.toLowerCase().includes('contractstart')) || Object.keys(row).find(k => k.toLowerCase().includes('date') || k.toLowerCase().includes('year'));
        if (!row[dKey]) return;
        const d = (row[dKey] instanceof Date) ? row[dKey] : (typeof row[dKey] === 'number' && row[dKey] > 40000) ? new Date(Math.round((row[dKey] - 25569) * 86400 * 1000)) : new Date(row[dKey]);
        if (d && !isNaN(d.getTime())) {
            const y = d.getFullYear().toString();
            if (summary[y]) {
                summary[y].kTcv += parseCurrency(row['KOR TCV(USD)'] || row['KOR TCV']);
                summary[y].lArr += parseCurrency(row['Local ARR'] || row['ARR']);
            }
        }
    });
    const container = document.createElement('div'); container.style.gridColumn = '1 / -1';
    container.innerHTML = `<h2 style="font-size:1.5rem; color:#374151; margin-bottom:12px;">${countryName} Analytics</h2><div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">${sortedYears.map(y => `<div class="stat-card" style="padding:15px; background:#F9FAFB; border-top: 3px solid #6366f1;"><h3>${y} Metrics</h3><div style="display:flex; justify-content:space-between; margin-top:8px;"><span>KOR TCV</span><span style="font-weight:700;">$${formatCurrency(summary[y].kTcv)}</span></div><div style="display:flex; justify-content:space-between; margin-top:4px;"><span>Local ARR</span><span style="font-weight:600;">$${formatCurrency(summary[y].lArr)}</span></div></div>`).join('')}</div>`;
    metricsGrid.appendChild(container);
}
function parseCurrency(val) {
    if (!val) return 0; if (typeof val === 'number') return val;
    const clean = String(val).replace(/[^0-9.-]+/g, "");
    return parseFloat(clean) || 0;
}
function formatCurrency(val, isKRW = false) {
    if (!val) return '0';
    return new Intl.NumberFormat(isKRW ? 'ko-KR' : 'en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(Math.round(val));
}
searchInput.addEventListener('input', (e) => {
    if (currentTab) {
        const activeSubSpan = document.querySelector('.nav-sublist .sub-item.active span');
        const country = activeSubSpan ? activeSubSpan.innerText : null;
        renderTableData(e.target.value, country === 'All' ? null : country);
    }
});
window.addEventListener('DOMContentLoaded', () => {
    loadLocalExcel(); // Start the dashboard on page load
});
