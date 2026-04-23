/**
 * app.js — Main entry point. Orchestrates data loading, sidebar, tab routing, and rendering.
 * Delegates to: kpi.js, sidebar.js, views.js, services.js, charts.js, ui.js
 */
import { isCountryMatch } from './utils.js';
import { buildSidebar } from './sidebar.js';
import { selectKPIView } from './kpi.js';
import { selectTrainingView } from './training.js';
import { renderTabMetrics } from './views.js';
import { DATA_SOURCES, AUTH } from './config.js';

function sheetsExportUrl(sheetId) {
    return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
}

/* ═══════════════════════════════════════════════════════════════
   Application State
   ═══════════════════════════════════════════════════════════════ */

/** @type {Object<string, Object[]>} */
let workbookData = {};

/** @type {string|null} */
let currentTab = null;

/* ═══════════════════════════════════════════════════════════════
   DOM Element References
   ═══════════════════════════════════════════════════════════════ */
const dropZone = document.getElementById('drop-zone');
const dashboardContainer = document.getElementById('dashboard-container');
const currentTabTitle = document.getElementById('current-tab-title');
const tableHead = document.getElementById('table-head');
const tableBody = document.getElementById('table-body');
const emptyState = document.getElementById('empty-state');
const dataTable = document.getElementById('data-table');
const searchInput = document.getElementById('search-input');

/* ═══════════════════════════════════════════════════════════════
   Last Updated Date
   ═══════════════════════════════════════════════════════════════ */

/** Update the sidebar "Last Updated" display. */
function updateLastUpdatedDate() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    const dateEl = document.getElementById('last-updated-date');
    if (dateEl) dateEl.innerText = `${d}.${m}.${y}`;
}

/* ═══════════════════════════════════════════════════════════════
   File Handling & Data Loading
   ═══════════════════════════════════════════════════════════════ */

/**
 * Fetch Excel files from local server and process.
 */
async function loadLocalExcel() {
    try {
        const btn = document.getElementById('refresh-btn');
        if (btn) btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i>';

        const [resMain, resMrr] = await Promise.all([
            fetch(sheetsExportUrl(DATA_SOURCES.MAIN_SHEET_ID)),
            fetch(sheetsExportUrl(DATA_SOURCES.MRR_SHEET_ID)).catch(() => null)
        ]);

        if (!resMain.ok) throw new Error(`Failed to fetch main file. Server returned ${resMain.status}`);

        const arrayBufferMain = await resMain.arrayBuffer();
        const dataMain = new Uint8Array(arrayBufferMain);
        const workbookMain = XLSX.read(dataMain, { type: 'array' });

        if (resMrr && resMrr.ok) {
            const arrayBufferMrr = await resMrr.arrayBuffer();
            const dataMrr = new Uint8Array(arrayBufferMrr);
            const workbookMrr = XLSX.read(dataMrr, { type: 'array' });
            const targetSheetName = workbookMrr.SheetNames.find(name => name.includes('Global(Contract Date)'));
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
        alert('Could not load Excel files:\n' + (error.stack || error.message) + '\n\nPlease check that the Google Drive file IDs are correct and the files are publicly shared.');
        const btn = document.getElementById('refresh-btn');
        if (btn) btn.innerHTML = '<i class="fa-solid fa-arrows-rotate"></i>';
    }
}

/**
 * Parse all sheets into workbookData and initialize sidebar.
 * @param {Object} workbook - XLSX workbook object
 */
function processWorkbook(workbook) {
    workbookData = {};
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        workbookData[sheetName] = XLSX.utils.sheet_to_json(worksheet, { defval: "", cellDates: true });
    });

    dropZone.classList.remove('active');
    dashboardContainer.classList.add('active');

    buildSidebar(workbook.SheetNames, {
        onSelectTab: selectTab,
        onSelectKPI: () => selectKPIView(setCurrentTab),
        onSelectTraining: () => selectTrainingView(setCurrentTab, workbookData),
        onSelectTcvArr: () => {
            currentTab = 'TCV_ARR';
            currentTabTitle.innerText = 'TCV vs ARR';
            document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
            const dataSection = document.querySelector('.data-section');
            if (dataSection) dataSection.classList.add('hidden');
            emptyState.classList.add('hidden');
            dataTable.classList.add('hidden');
            renderTabMetrics([], 'TCV_ARR', null, workbookData, searchInput);
        }
    });

    updateLastUpdatedDate();

    if (workbook.SheetNames.length > 0) {
        const firstTab = workbook.SheetNames.includes('ORDER SHEET') ? 'ORDER SHEET' : workbook.SheetNames[0];
        selectTab(firstTab);
    }
}

/* ═══════════════════════════════════════════════════════════════
   State Setter (passed to sub-modules)
   ═══════════════════════════════════════════════════════════════ */

/**
 * Set the current tab state.
 * @param {string} tab
 */
function setCurrentTab(tab) {
    currentTab = tab;
}

/* ═══════════════════════════════════════════════════════════════
   Tab Selection & Rendering
   ═══════════════════════════════════════════════════════════════ */

/**
 * Select a tab and render its data view.
 * @param {string} tabName
 * @param {string|null} [subTabName=null]
 */
function selectTab(tabName, subTabName = null) {
    currentTab = tabName;
    currentTabTitle.innerText = subTabName ? `${tabName} — ${subTabName}` : tabName;
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    renderTableData("", subTabName);
}

/**
 * Render the data table and delegate metric panels.
 * @param {string} [searchTerm=""]
 * @param {string|null} [filterCountry=null]
 */
function renderTableData(searchTerm = "", filterCountry = null) {
    const data = workbookData[currentTab] || [];
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    const dataSection = document.querySelector('.data-section');

    if (data.length === 0) {
        if (dataSection) dataSection.classList.add('hidden');
        emptyState.classList.remove('hidden');
        dataTable.classList.add('hidden');
        return;
    }

    const headers = Object.keys(data[0]).filter(k => !k.startsWith('__EMPTY'));
    const trHead = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.innerText = h;
        trHead.appendChild(th);
    });
    tableHead.appendChild(trHead);

    const filtered = data.filter(r =>
        isCountryMatch(r, filterCountry) &&
        (!searchTerm || Object.values(r).some(v => String(v).toLowerCase().includes(searchTerm.toLowerCase())))
    );

    if (filtered.length === 0) {
        if (dataSection) dataSection.classList.add('hidden');
        emptyState.classList.remove('hidden');
        dataTable.classList.add('hidden');
    } else {
        if (dataSection) {
            if (currentTab === 'PARTNER') {
                dataSection.classList.add('hidden');
            } else {
                dataSection.classList.remove('hidden');
            }
        }
        emptyState.classList.add('hidden');
        if (currentTab === 'PARTNER') {
            dataTable.classList.add('hidden');
        } else {
            dataTable.classList.remove('hidden');
            filtered.slice(0, 100).forEach(row => {
                const tr = document.createElement('tr');
                headers.forEach(h => {
                    const td = document.createElement('td');
                    td.innerText = row[h] || '';
                    tr.appendChild(td);
                });
                tableBody.appendChild(tr);
            });
        }
    }

    renderTabMetrics(filtered, currentTab, filterCountry, workbookData, searchInput);
}

/* ═══════════════════════════════════════════════════════════════
   Event Listeners
   ═══════════════════════════════════════════════════════════════ */

function initAuth() {
    const overlay = document.getElementById('login-overlay');
    const input = document.getElementById('password-input');
    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');

    if (sessionStorage.getItem('auth') === AUTH.PASSWORD_ENCODED) {
        overlay.classList.add('hidden');
        loadLocalExcel();
        return;
    }

    const tryLogin = () => {
        if (btoa(input.value) === AUTH.PASSWORD_ENCODED) {
            sessionStorage.setItem('auth', AUTH.PASSWORD_ENCODED);
            overlay.classList.add('hidden');
            loadLocalExcel();
        } else {
            error.style.display = 'block';
            input.value = '';
            input.focus();
        }
    };

    btn.addEventListener('click', tryLogin);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') tryLogin(); });
}

window.addEventListener('load', () => {
    initAuth();
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) refreshBtn.addEventListener('click', loadLocalExcel);
});

searchInput.addEventListener('input', (e) => {
    if (currentTab) {
        const activeSubSpan = document.querySelector('.nav-sublist .sub-item.active span');
        const country = activeSubSpan ? activeSubSpan.innerText : null;
        renderTableData(e.target.value, country === 'All' ? null : country);
    }
});

/** Handle filter-country-change events dispatched from views.js */
window.addEventListener('filter-country-change', (e) => {
    renderTableData(e.detail.searchTerm || '', e.detail.country);
});