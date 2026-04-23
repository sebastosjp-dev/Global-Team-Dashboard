/**
 * kpi.js — KPI Goals CRUD, persistence (localStorage), and rendering.
 * Extracted from app.js for single-responsibility.
 * @module kpi
 */
import { parseCurrency, formatCurrency } from './utils.js';
import { getKPIHTML } from './ui.js';

/* ═══════════════════════════════════════════════════════════════
   Default Data & State
   ═══════════════════════════════════════════════════════════════ */

/** @type {Object|null} */
let kpiData = null;

/** @type {number} */
let currentKPIYear = new Date().getFullYear();

/**
 * Default KPI structure used when no saved data exists.
 * @constant
 */
const DEFAULT_KPI_DATA = {
    categories: [
        {
            name: "FINANCIAL", color: "#8b5cf6",
            objectives: [
                { name: "Nett New Revenue", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
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
            name: "INTERNAL PROCESS", color: "#3b82f6",
            objectives: [
                { name: "Conversion : POC to Deal", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
            ]
        },
        {
            name: "LEARNING & GROWTH", color: "#22c55e",
            objectives: [
                { name: "Staff Training & Development", kpis: "", targets: [0, 0, 0, 0], achievements: [0, 0, 0, 0], weight: 0 }
            ]
        }
    ]
};

/* ═══════════════════════════════════════════════════════════════
   Persistence Helpers
   ═══════════════════════════════════════════════════════════════ */

/**
 * Load KPI data from localStorage, falling back to defaults.
 */
function loadKPIData() {
    const key = `global_dashboard_kpi_${currentKPIYear}`;
    let stored = localStorage.getItem(key);

    if (!stored && currentKPIYear === 2026) {
        stored = localStorage.getItem('global_dashboard_kpi');
    }

    if (stored) {
        try { kpiData = JSON.parse(stored); } catch (_e) {
            kpiData = JSON.parse(JSON.stringify(DEFAULT_KPI_DATA));
        }
    } else {
        kpiData = JSON.parse(JSON.stringify(DEFAULT_KPI_DATA));
    }
}

/**
 * Persist current KPI data to localStorage.
 */
function persistKPIData() {
    const key = `global_dashboard_kpi_${currentKPIYear}`;
    localStorage.setItem(key, JSON.stringify(kpiData));
    if (currentKPIYear === 2026) {
        localStorage.setItem('global_dashboard_kpi', JSON.stringify(kpiData));
    }
}

/* ═══════════════════════════════════════════════════════════════
   Render
   ═══════════════════════════════════════════════════════════════ */

/**
 * Render the KPI table into the metrics grid.
 */
function renderKPIView() {
    if (!kpiData) loadKPIData();
    const metricsGrid = document.getElementById('tab-metrics-grid');
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');
    const dataTable = document.getElementById('data-table');
    const emptyState = document.getElementById('empty-state');

    if (tableHead) tableHead.innerHTML = '';
    if (tableBody) tableBody.innerHTML = '';
    if (dataTable) dataTable.classList.add('hidden');
    if (emptyState) emptyState.classList.add('hidden');
    if (metricsGrid) {
        metricsGrid.innerHTML = getKPIHTML(kpiData, currentKPIYear);
        metricsGrid.classList.remove('hidden');
    }
}

/* ═══════════════════════════════════════════════════════════════
   Public API — bound to window for inline HTML handlers
   ═══════════════════════════════════════════════════════════════ */

window.saveKPIData = function () {
    persistKPIData();
    alert(`KPI Goals for ${currentKPIYear} saved successfully!`);
    renderKPIView();
};

window.resetKPIData = function () {
    if (confirm(`Are you sure you want to reset all KPI data for ${currentKPIYear} to default? This cannot be undone.`)) {
        kpiData = JSON.parse(JSON.stringify(DEFAULT_KPI_DATA));
        persistKPIData();
        renderKPIView();
    }
};

window.exportKPIData = function () {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(kpiData, null, 2));
    const anchor = document.createElement('a');
    anchor.setAttribute("href", dataStr);
    anchor.setAttribute("download", `annual_kpi_goals_${currentKPIYear}.json`);
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
};

window.changeKPIYear = function (year) {
    currentKPIYear = parseInt(year, 10);
    loadKPIData();
    const titleEl = document.getElementById('current-tab-title');
    if (titleEl) titleEl.innerText = `${currentKPIYear} ANNUAL KPI TARGETS & ACHIEVEMENTS`;
    renderKPIView();
};

window.updateKPICell = function (el, type, catIdx, objIdx, qIdx) {
    const val = parseCurrency(el.value);
    kpiData.categories[catIdx].objectives[objIdx][type][qIdx] = val;
    el.value = formatCurrency(val);
};

window.updateKPIText = function (el, field, catIdx, objIdx) {
    kpiData.categories[catIdx].objectives[objIdx][field] = el.innerText || el.value || '';
};

window.updateKPINumber = function (el, field, catIdx, objIdx) {
    const val = parseFloat(el.value) || 0;
    kpiData.categories[catIdx].objectives[objIdx][field] = val;
    el.value = val;
    renderKPIView();
};

window.updateKPICategoryName = function (el, catIdx) {
    kpiData.categories[catIdx].name = el.innerText;
};

window.updateKPIObjectiveName = function (el, catIdx, objIdx) {
    kpiData.categories[catIdx].objectives[objIdx].name = el.innerText;
};

/* ═══════════════════════════════════════════════════════════════
   Tab Entry Point (called from app.js)
   ═══════════════════════════════════════════════════════════════ */

/**
 * Select and display the KPI view tab.
 * @param {function} setCurrentTab - callback to update app state
 */
export function selectKPIView(setCurrentTab) {
    setCurrentTab('KPI_VIEW');
    const titleEl = document.getElementById('current-tab-title');
    if (titleEl) titleEl.innerText = `${currentKPIYear} ANNUAL KPI TARGETS & ACHIEVEMENTS`;

    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.kpi-tab')?.classList.add('active');

    const dataSection = document.querySelector('.data-section');
    if (dataSection) dataSection.classList.add('hidden');
    document.getElementById('empty-state')?.classList.add('hidden');
    document.getElementById('data-table')?.classList.add('hidden');

    renderKPIView();
}
