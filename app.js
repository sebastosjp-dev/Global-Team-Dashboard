/**
 * app.js — Main entry point. Orchestrates data loading, sidebar, tab routing, and rendering.
 */
import { CONFIG } from './config.js';
import { parseCurrency, formatCurrency, normalizeCountry, isCountryMatch, sortCountriesByCount } from './utils.js';
import { chartRegistry, initOrderSheetCharts, initPipelineCharts, initPartnerCharts, initPartnerPerformanceCharts, initPocCharts, initEventCharts, initServiceAnalysisCharts } from './charts.js';
import { getOrderSheetStats, getPipelineStats, getPartnerStats, getGenericCountryStats, getExpiringContractsStats, getPartnerPerformanceStats, getPocStats, getEventStats, getCountrySpecificStats, getServiceAnalysisStats } from './services.js';
import { getOrderSheetHTML, getPipelineHTML, getPartnerHTML, getGenericCountryHTML, getExpiringContractsHTML, getPartnerPerformanceHTML, getPocHTML, getEventHTML, getCountrySpecificHTML, getServiceAnalysisHTML, getRenewalHTML, injectServiceAnalysisStyles, getKPIHTML } from './ui.js';

/*═══════════════════════════════════════════════════════════════
  Application State
═══════════════════════════════════════════════════════════════*/
let workbookData = {};
let currentTab = null;
let kpiData = null;
let currentKPIYear = 2026;

/*═══════════════════════════════════════════════════════════════
  DOM Element References
═══════════════════════════════════════════════════════════════*/
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
  Country Sorting & Last Updated
═══════════════════════════════════════════════════════════════*/
function updateLastUpdatedDate() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const dateEl = document.getElementById('last-updated-date');
    if (dateEl) dateEl.innerText = `${d}.${m}.${y}`;
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

/*═══════════════════════════════════════════════════════════════
  File Handling & Data Loading
═══════════════════════════════════════════════════════════════*/
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

async function loadLocalExcel() {
    try {
        const btn = document.getElementById('refresh-btn');
        if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';
        const ts = new Date().getTime();

        const [resMain, resMrr] = await Promise.all([
            fetch('2026 Global Rev.01.xlsx?t=' + ts),
            fetch('Global MRR ARR.xlsx?t=' + ts).catch(() => null)
        ]);

        if (!resMain.ok) throw new Error(`Failed to fetch main file. Server returned ${resMain.status}`);

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
                window.globalMrrSheet = sheet;
                workbookMain.SheetNames.push(targetSheetName);
                workbookMain.Sheets[targetSheetName] = sheet;
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

window.addEventListener('load', () => {
    loadLocalExcel();
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', loadLocalExcel);
    }
});

function processWorkbook(workbook) {
    workbookData = {};
    const sheetNames = workbook.SheetNames;
    sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: "", cellDates: true });
        workbookData[sheetName] = json;
    });

    dropZone.classList.remove('active');
    dashboardContainer.classList.add('active');
    buildSidebar(sheetNames);
    updateLastUpdatedDate();

    if (sheetNames.length > 0) {
        const firstTab = sheetNames.includes('ORDER SHEET') ? 'ORDER SHEET' : sheetNames[0];
        selectTab(firstTab);
    }
}

/*═══════════════════════════════════════════════════════════════
  Sidebar Navigation
═══════════════════════════════════════════════════════════════*/
function buildSidebar(sheetNames) {
    sidebarNav.innerHTML = '';

    sheetNames.forEach(name => {
        if (name.includes('Global(계약시점)') || ['Sheet9', 'Sheet10', 'Sheet12', 'Sheet13', '2026 Q1 Review'].includes(name)) return;

        if (name === 'ORDER SHEET') {
            const parentItem = document.createElement('div');
            parentItem.className = 'nav-item nav-item-parent';
            parentItem.innerHTML = `<div style="display:flex; align-items:center; gap:12px;"><i class="fa-solid fa-folder"></i> <span>${name}</span></div><i class="fa-solid fa-chevron-down toggle-icon" style="font-size: 0.8em; transition: transform 0.3s;"></i>`;

            const subList = document.createElement('div');
            subList.className = 'nav-sublist';

            const allItem = document.createElement('div');
            allItem.className = 'nav-item sub-item';
            allItem.innerHTML = `<i class="fa-solid fa-earth-americas"></i> <span>All</span>`;
            allItem.onclick = (e) => { e.stopPropagation(); selectTab(name, null); };
            subList.appendChild(allItem);

            CONFIG.COUNTRIES.forEach(country => {
                const subItem = document.createElement('div');
                subItem.className = 'nav-item sub-item';
                subItem.innerHTML = `<i class="fa-solid fa-location-dot"></i> <span>${country}</span>`;
                subItem.onclick = (e) => { e.stopPropagation(); selectTab(name, country); };
                subList.appendChild(subItem);
            });

            parentItem.onclick = () => {
                subList.classList.toggle('expanded');
                const icon = parentItem.querySelector('.toggle-icon');
                icon.style.transform = subList.classList.contains('expanded') ? 'rotate(180deg)' : 'rotate(0deg)';
                selectTab(name, null);
            };

            sidebarNav.appendChild(parentItem);
            sidebarNav.appendChild(subList);
        } else if (name === 'END USER (CSM)') {
            const parentItem = document.createElement('div');
            parentItem.className = 'nav-item nav-item-parent';
            parentItem.innerHTML = `<div style="display:flex; align-items:center; gap:12px;"><i class="fa-solid fa-folder"></i> <span>${name}</span></div><i class="fa-solid fa-chevron-down toggle-icon" style="font-size: 0.8em; transition: transform 0.3s;"></i>`;

            const subList = document.createElement('div');
            subList.className = 'nav-sublist';

            const renewalItem = document.createElement('div');
            renewalItem.className = 'nav-item renewal-tab sub-item';
            renewalItem.style.background = 'linear-gradient(90deg, rgba(245, 158, 11, 0.1) 0%, rgba(245, 158, 11, 0) 100%)';
            renewalItem.style.borderLeft = '2px solid #f59e0b';
            renewalItem.innerHTML = `<i class="fa-solid fa-user-clock" style="color: #f59e0b;"></i> <span>Renewal Management</span>`;
            renewalItem.onclick = (e) => { e.stopPropagation(); selectRenewalView(); };
            subList.appendChild(renewalItem);

            const upsellItem = document.createElement('div');
            upsellItem.className = 'nav-item upsell-tab sub-item';
            upsellItem.style.background = 'linear-gradient(90deg, rgba(16, 185, 129, 0.1) 0%, rgba(16, 185, 129, 0) 100%)';
            upsellItem.style.borderLeft = '2px solid #10b981';
            upsellItem.innerHTML = `<i class="fa-solid fa-chart-pie" style="color: #34C759;"></i> <span>Service Analysis (Upsell)</span>`;
            upsellItem.onclick = (e) => { e.stopPropagation(); selectServiceAnalysisView(); };
            subList.appendChild(upsellItem);

            parentItem.onclick = () => {
                subList.classList.toggle('expanded');
                const icon = parentItem.querySelector('.toggle-icon');
                icon.style.transform = subList.classList.contains('expanded') ? 'rotate(180deg)' : 'rotate(0deg)';
                selectTab(name, null);
            };

            sidebarNav.appendChild(parentItem);
            sidebarNav.appendChild(subList);
        } else {
            const navItem = document.createElement('div');
            navItem.className = 'nav-item';
            const icon = name === 'EVENT' ? 'fa-calendar-check' : 'fa-folder';
            navItem.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${name}</span>`;
            navItem.onclick = () => selectTab(name, null);
            sidebarNav.appendChild(navItem);
        }
    });

    // Add KPI Menu at the end
    const kpiItem = document.createElement('div');
    kpiItem.className = 'nav-item kpi-tab';
    kpiItem.style.marginTop = '10px';
    kpiItem.style.borderTop = '1px solid rgba(255,255,255,0.1)';
    kpiItem.style.paddingTop = '15px';
    kpiItem.innerHTML = `<i class="fa-solid fa-bullseye" style="color: #ef4444;"></i> <span style="font-weight: 700;">KPI GOALS</span>`;
    kpiItem.onclick = () => selectKPIView();
    sidebarNav.appendChild(kpiItem);
}

/*═══════════════════════════════════════════════════════════════
  Tab Selection & Rendering
═══════════════════════════════════════════════════════════════*/
function selectTab(tabName, subTabName = null) {
    currentTab = tabName;
    currentTabTitle.innerText = subTabName ? `${tabName} — ${subTabName}` : tabName;
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    renderTableData("", subTabName);
}

function renderTableData(searchTerm = "", filterCountry = null) {
    const data = workbookData[currentTab] || [];
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    if (data.length === 0) return;
    const headers = Object.keys(data[0]).filter(k => !k.startsWith('__EMPTY'));
    const trHead = document.createElement('tr');
    headers.forEach(h => { const th = document.createElement('th'); th.innerText = h; trHead.appendChild(th); });
    tableHead.appendChild(trHead);
    const filtered = data.filter(r => isCountryMatch(r, filterCountry) && (!searchTerm || Object.values(r).some(v => String(v).toLowerCase().includes(searchTerm.toLowerCase()))));
    filtered.slice(0, 50).forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => { const td = document.createElement('td'); td.innerText = row[h] || ''; tr.appendChild(td); });
        tableBody.appendChild(tr);
    });
    renderTabMetrics(filtered, currentTab, filterCountry);
}

/*═══════════════════════════════════════════════════════════════
  Renewal & CSM Views
═══════════════════════════════════════════════════════════════*/
function selectRenewalView() {
    currentTab = 'RENEWAL_VIEW';
    currentTabTitle.innerText = 'Churn Prevention & Renewal Management (CSM)';
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.renewal-tab')?.classList.add('active');
    renderRenewalTable();
}

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
        return { ...row, 'D-Day': diffDays === 0 ? 'D-Day' : (diffDays > 0 ? `D-${diffDays}` : `D+${Math.abs(diffDays)}`), diffDays, endDateFormatted: d.toISOString().split('T')[0] };
    }).sort((a, b) => a.diffDays - b.diffDays || (parseCurrency(b['TCV Amount']) - parseCurrency(a['TCV Amount'])));
}

function renderRenewalTable() {
    const filtered = getRenewalTableData();
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    metricsGrid.classList.remove('hidden');
    metricsGrid.innerHTML = getRenewalHTML(filtered);
}

/*═══════════════════════════════════════════════════════════════
  Service Analysis View
═══════════════════════════════════════════════════════════════*/
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
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    dataTable.classList.add('hidden');
    emptyState.classList.add('hidden');
    metricsGrid.innerHTML = '';
    metricsGrid.classList.remove('hidden');

    const stats = getServiceAnalysisStats(data);
    metricsGrid.innerHTML = getServiceAnalysisHTML(stats);
    if (stats) setTimeout(() => initServiceAnalysisCharts(stats), 100);
}

/*═══════════════════════════════════════════════════════════════
  KPI View & Manual Updates
  Manual updates are stored in localStorage for persistence.
═══════════════════════════════════════════════════════════════*/
const DEFAULT_KPI_DATA = {
    categories: [
        {
            name: "FINANCIAL", color: "#8b5cf6",
            objectives: [
                { name: "Nett New Revenue", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 },
                { name: "Up/cross selling", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
            ]
        },
        {
            name: "CUSTOMER", color: "#f59e0b",
            objectives: [
                { name: "New Strategic Account", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 },
                { name: "Customer Retention", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
            ]
        },
        {
            name: "INTERNAL PROCESSES", color: "#10b981",
            objectives: [
                { name: "Conversion : POC to Deal", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
            ]
        },
        {
            name: "LEARNING & GROWTH", color: "#0ea5e9",
            objectives: [
                { name: "Online Training", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 },
                { name: "Fundamental workshop", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
            ]
        }
    ]
};

function loadKPIData() {
    const key = `global_dashboard_kpi_${currentKPIYear}`;
    let stored = localStorage.getItem(key);
    
    // Fallback for 2026 if new year specific key not found
    if (!stored && currentKPIYear === 2026) {
        stored = localStorage.getItem('global_dashboard_kpi');
    }

    if (stored) {
        try { kpiData = JSON.parse(stored); } catch (e) { kpiData = JSON.parse(JSON.stringify(DEFAULT_KPI_DATA)); }
    } else {
        kpiData = JSON.parse(JSON.stringify(DEFAULT_KPI_DATA));
    }
}

window.saveKPIData = function() {
    const key = `global_dashboard_kpi_${currentKPIYear}`;
    localStorage.setItem(key, JSON.stringify(kpiData));
    
    // Optional: Keep old key if it is 2026 for backward compatibility with other tabs if any
    if (currentKPIYear === 2026) {
        localStorage.setItem('global_dashboard_kpi', JSON.stringify(kpiData));
    }
    
    alert(`KPI Goals for ${currentKPIYear} saved successfully!`);
    renderKPIView();
};

window.resetKPIData = function() {
    if (confirm(`Are you sure you want to reset all KPI data for ${currentKPIYear} to default? This cannot be undone.`)) {
        kpiData = JSON.parse(JSON.stringify(DEFAULT_KPI_DATA));
        const key = `global_dashboard_kpi_${currentKPIYear}`;
        localStorage.setItem(key, JSON.stringify(kpiData));
        
        if (currentKPIYear === 2026) {
            localStorage.setItem('global_dashboard_kpi', JSON.stringify(kpiData));
        }
        
        renderKPIView();
    }
};

window.exportKPIData = function() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(kpiData, null, 2));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", `annual_kpi_goals_${currentKPIYear}.json`);
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
};

window.changeKPIYear = function(year) {
    currentKPIYear = parseInt(year, 10);
    loadKPIData();
    currentTabTitle.innerText = `${currentKPIYear} ANNUAL KPI TARGETS & ACHIEVEMENTS`;
    renderKPIView();
};

window.updateKPICell = function(el, type, catIdx, objIdx, qIdx) {
    const val = parseCurrency(el.value);
    kpiData.categories[catIdx].objectives[objIdx][type][qIdx] = val;
    el.value = formatCurrency(val);
    // Don't auto-save to allow cancel/batch save, but we could if we want
};

window.updateKPIText = function(el, field, catIdx, objIdx) {
    kpiData.categories[catIdx].objectives[objIdx][field] = el.innerText || el.value || '';
};

window.updateKPINumber = function(el, field, catIdx, objIdx) {
    const val = parseFloat(el.value) || 0;
    kpiData.categories[catIdx].objectives[objIdx][field] = val;
    el.value = val;
    renderKPIView(); // Re-render to update the total rate calculation if weight changes
};

window.updateKPICategoryName = function(el, catIdx) {
    kpiData.categories[catIdx].name = el.innerText;
};

window.updateKPIObjectiveName = function(el, catIdx, objIdx) {
    kpiData.categories[catIdx].objectives[objIdx].name = el.innerText;
};

function selectKPIView() {
    currentTab = 'KPI_VIEW';
    currentTabTitle.innerText = `${currentKPIYear} ANNUAL KPI TARGETS & ACHIEVEMENTS`;
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.kpi-tab')?.classList.add('active');
    renderKPIView();
}

function renderKPIView() {
    // always load data here if not loaded, or after year switch.
    if (!kpiData) loadKPIData();
    const metricsGrid = document.getElementById('tab-metrics-grid');
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    dataTable.classList.add('hidden');
    emptyState.classList.add('hidden');
    metricsGrid.innerHTML = getKPIHTML(kpiData, currentKPIYear);
    metricsGrid.classList.remove('hidden');
}

/*═══════════════════════════════════════════════════════════════
  Metrics Router
═══════════════════════════════════════════════════════════════*/
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

    if (hasMetrics) metricsGrid.classList.remove('hidden');
    else metricsGrid.classList.add('hidden');
}

/*═══════════════════════════════════════════════════════════════
  Render Orchestrators (connect Stats → HTML → Charts)
═══════════════════════════════════════════════════════════════*/
function renderOrderSheetMetrics(data, filterCountry, metricsGrid, tabName) {
    const stats = getOrderSheetStats(data, filterCountry, tabName, workbookData);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getOrderSheetHTML(stats);
    metricsGrid.appendChild(container);
    setTimeout(() => initOrderSheetCharts(stats), 120);
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
    setTimeout(() => {
        // Attach filter event listener
        const selector = document.getElementById('pipeline-filter-country');
        if (selector) {
            selector.addEventListener('change', (e) => {
                const val = e.target.value;
                renderTableData(searchInput.value, val === 'All' ? null : val);
            });
        }
        initPipelineCharts(stats);
    }, 100);
}

function renderPartnerMetricsBlock(data, filterCountry, tabName, metricsGrid) {
    if (!data || data.length === 0) return;
    const stats = getPartnerStats(data, filterCountry, workbookData);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getPartnerHTML(stats, filterCountry, tabName);
    metricsGrid.appendChild(container);
    setTimeout(() => {
        // Attach filter event listener
        const selector = document.getElementById('partner-filter-country');
        if (selector) {
            selector.addEventListener('change', (e) => {
                const val = e.target.value;
                renderTableData(searchInput.value, val === 'All' ? null : val);
            });
        }
        initPartnerCharts(stats, filterCountry);
    }, 100);
}

function renderGenericCountryCounts(data, filterCountry, metricsGrid, tabName) {
    if (data.length === 0 || tabName === 'EVENT') return;
    const stats = getGenericCountryStats(data, filterCountry);
    if (!stats) return;
    const div = document.createElement('div');
    div.innerHTML = getGenericCountryHTML(stats, filterCountry);
    metricsGrid.appendChild(div.firstElementChild);
}

function renderExpiringContracts(data, metricsGrid) {
    const stats = getExpiringContractsStats(data);
    if (!stats) return;
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getExpiringContractsHTML(stats);
    metricsGrid.appendChild(container);
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

function renderPocMetricsBlock(data, filterCountry, metricsGrid) {
    metricsGrid.innerHTML = '';
    const pocContainer = document.createElement('div');
    pocContainer.id = 'poc-dashboard-container';
    pocContainer.style.gridColumn = '1 / -1';
    metricsGrid.appendChild(pocContainer);

    window.pocFilters = window.pocFilters || { country: 'All', industry: 'All', partner: 'All' };

    window.renderPocUI = function () {
        const { stats, uniqueValues } = getPocStats(data, window.pocFilters, workbookData);
        const container = document.getElementById('poc-dashboard-container');
        if (container) {
            container.innerHTML = getPocHTML(stats, window.pocFilters, uniqueValues);
            document.getElementById('poc-filter-country').addEventListener('change', (e) => { window.pocFilters.country = e.target.value; window.renderPocUI(); });
            document.getElementById('poc-filter-industry').addEventListener('change', (e) => { window.pocFilters.industry = e.target.value; window.renderPocUI(); });
            document.getElementById('poc-filter-partner').addEventListener('change', (e) => { window.pocFilters.partner = e.target.value; window.renderPocUI(); });
            setTimeout(() => initPocCharts(stats), 50);
        }
    };

    window.renderPocUI();
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
  Event Listeners
═══════════════════════════════════════════════════════════════*/
searchInput.addEventListener('input', (e) => {
    if (currentTab) {
        const activeSubSpan = document.querySelector('.nav-sublist .sub-item.active span');
        const country = activeSubSpan ? activeSubSpan.innerText : null;
        renderTableData(e.target.value, country === 'All' ? null : country);
    }
});