/**
 * kpi.js — KPI Goals: Admin sets structure/targets, team members enter achievements.
 * Storage: server files (kpi_structure_YEAR.json, kpi_ach_YEAR_NAME.json)
 */
import { parseCurrency, formatCurrency } from './utils.js';
import { getKPIHTML } from './ui.js';

/* ═══════════════════════════════════════════════════════════════
   State
   ═══════════════════════════════════════════════════════════════ */

let kpiStructure = null;    // Admin's data: categories, objectives, targets, sub-item names
let kpiAchievements = null; // Per-user: { user, year, data: {"ci_oi_si": [q1,q2,q3,q4]} }
let currentKPIYear = new Date().getFullYear();
let currentUser = 'admin';
let isAdmin = true;
let availableUsers = [];

/* ═══════════════════════════════════════════════════════════════
   Defaults
   ═══════════════════════════════════════════════════════════════ */

const DEFAULT_STRUCTURE = {
    categories: [
        {
            name: "FINANCIAL", color: "#8b5cf6",
            objectives: [
                { name: "Nett New Revenue", kpis: "", targets: [0,0,0,0], weight: 0,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        },
        {
            name: "CUSTOMER", color: "#f59e0b",
            objectives: [
                { name: "New Strategic Account", kpis: "", targets: [0,0,0,0], weight: 0,
                  subItems: [{name:""},{name:""},{name:""}] },
                { name: "Customer Retention", kpis: "", targets: [0,0,0,0], weight: 0,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        },
        {
            name: "INTERNAL PROCESS", color: "#3b82f6",
            objectives: [
                { name: "Conversion : POC to Deal", kpis: "", targets: [0,0,0,0], weight: 0,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        },
        {
            name: "LEARNING & GROWTH", color: "#22c55e",
            objectives: [
                { name: "Staff Training & Development", kpis: "", targets: [0,0,0,0], weight: 0,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        }
    ]
};

/* ═══════════════════════════════════════════════════════════════
   Data Helpers
   ═══════════════════════════════════════════════════════════════ */

function migrateStructure(s) {
    if (!s || !s.categories) return;
    s.categories.forEach(cat => cat.objectives.forEach(obj => {
        if (!obj.subItems || obj.subItems.length === 0) {
            obj.subItems = [{name:""},{name:""},{name:""}];
        } else {
            // Keep only name, strip achievements (achievements live in user files)
            obj.subItems = obj.subItems.slice(0, 3).map(si =>
                typeof si === 'string' ? { name: si } : { name: si.name || "" }
            );
        }
        delete obj.achievements; // remove legacy field
    }));
}

// Merge structure + current user's achievements into one object for rendering
function getMergedData() {
    if (!kpiStructure) return null;
    const merged = JSON.parse(JSON.stringify(kpiStructure));
    merged.categories.forEach((cat, ci) => {
        cat.objectives.forEach((obj, oi) => {
            if (!obj.subItems || obj.subItems.length < 3) {
                obj.subItems = [{name:""},{name:""},{name:""}];
            }
            obj.subItems.forEach((sub, si) => {
                const key = `${ci}_${oi}_${si}`;
                sub.achievements = kpiAchievements?.data?.[key]
                    ? [...kpiAchievements.data[key]]
                    : [0, 0, 0, 0];
            });
        });
    });
    return merged;
}

/* ═══════════════════════════════════════════════════════════════
   Server API Calls
   ═══════════════════════════════════════════════════════════════ */

async function loadStructure() {
    // Try new structure file first
    try {
        const res = await fetch(`/api/kpi/structure?year=${currentKPIYear}`);
        if (res.ok) { kpiStructure = await res.json(); migrateStructure(kpiStructure); return; }
    } catch (_e) {}
    // Fallback: old unified file
    try {
        const res = await fetch(`/api/kpi/load?year=${currentKPIYear}`);
        if (res.ok) { kpiStructure = await res.json(); migrateStructure(kpiStructure); return; }
    } catch (_e) {}
    // Fallback: localStorage (legacy)
    const stored = localStorage.getItem(`global_dashboard_kpi_${currentKPIYear}`)
        || (currentKPIYear === 2026 ? localStorage.getItem('global_dashboard_kpi') : null);
    if (stored) {
        try { kpiStructure = JSON.parse(stored); migrateStructure(kpiStructure); return; } catch (_e) {}
    }
    kpiStructure = JSON.parse(JSON.stringify(DEFAULT_STRUCTURE));
}

async function loadAchievements(user) {
    try {
        const res = await fetch(`/api/kpi/achievement?year=${currentKPIYear}&user=${encodeURIComponent(user)}`);
        if (res.ok) { kpiAchievements = await res.json(); return; }
    } catch (_e) {}
    kpiAchievements = { user, year: currentKPIYear, data: {} };
}

async function loadAvailableUsers() {
    try {
        const res = await fetch(`/api/kpi/users?year=${currentKPIYear}`);
        if (res.ok) availableUsers = await res.json();
    } catch (_e) { availableUsers = []; }
}

async function saveStructure() {
    const json = JSON.stringify(kpiStructure);
    localStorage.setItem(`global_dashboard_kpi_${currentKPIYear}`, json);
    await fetch(`/api/kpi/structure?year=${currentKPIYear}`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: json
    });
}

async function saveAchievements() {
    if (!kpiAchievements) kpiAchievements = { user: currentUser, year: currentKPIYear, data: {} };
    await fetch(`/api/kpi/achievement?year=${currentKPIYear}&user=${encodeURIComponent(currentUser)}`, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(kpiAchievements)
    });
}

/* ═══════════════════════════════════════════════════════════════
   Render
   ═══════════════════════════════════════════════════════════════ */

async function renderKPIView() {
    if (!kpiStructure) await loadStructure();
    if (!isAdmin && !kpiAchievements) await loadAchievements(currentUser);

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
        metricsGrid.innerHTML = getKPIHTML(getMergedData(), currentKPIYear, isAdmin, currentUser, availableUsers);
        metricsGrid.classList.remove('hidden');
    }
}

/* ═══════════════════════════════════════════════════════════════
   Public API — window handlers called from inline HTML
   ═══════════════════════════════════════════════════════════════ */

window.saveKPIData = async function () {
    if (isAdmin) {
        await saveStructure();
        alert(`KPI structure for ${currentKPIYear} saved!`);
    } else {
        await saveAchievements();
        alert(`${currentUser}'s achievements for ${currentKPIYear} saved!`);
    }
    renderKPIView();
};

window.resetKPIData = async function () {
    if (!isAdmin) { alert('Only Admin can reset the KPI structure.'); return; }
    if (confirm(`Reset all KPI structure for ${currentKPIYear} to default? This cannot be undone.`)) {
        kpiStructure = JSON.parse(JSON.stringify(DEFAULT_STRUCTURE));
        await saveStructure();
        renderKPIView();
    }
};

window.exportKPIData = function () {
    const data = isAdmin ? kpiStructure : kpiAchievements;
    const name = isAdmin
        ? `kpi_structure_${currentKPIYear}`
        : `kpi_ach_${currentKPIYear}_${currentUser}`;
    const anchor = document.createElement('a');
    anchor.href = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(data, null, 2));
    anchor.download = `${name}.json`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
};

window.changeKPIYear = async function (year) {
    currentKPIYear = parseInt(year, 10);
    kpiStructure = null;
    kpiAchievements = null;
    await loadStructure();
    if (!isAdmin) await loadAchievements(currentUser);
    await loadAvailableUsers();
    const titleEl = document.getElementById('current-tab-title');
    if (titleEl) titleEl.innerText = `${currentKPIYear} ANNUAL KPI TARGETS & ACHIEVEMENTS`;
    renderKPIView();
};

window.switchKPIMode = async function (mode) {
    if (mode === 'admin') {
        isAdmin = true;
        currentUser = 'admin';
        kpiAchievements = null;
    } else {
        isAdmin = false;
        currentUser = mode;
        kpiAchievements = null;
        await loadAchievements(currentUser);
        localStorage.setItem('kpi_last_user', currentUser);
    }
    renderKPIView();
};

window.addKPIUser = async function () {
    const name = prompt('Enter team member name:')?.trim();
    if (!name || name.toLowerCase() === 'admin') return;
    isAdmin = false;
    currentUser = name;
    kpiAchievements = { user: name, year: currentKPIYear, data: {} };
    await saveAchievements();
    await loadAvailableUsers();
    localStorage.setItem('kpi_last_user', name);
    renderKPIView();
};

// Admin-only: structure edits
window.updateKPICell = function (el, type, catIdx, objIdx, qIdx) {
    if (!isAdmin) return;
    const val = parseCurrency(el.value);
    kpiStructure.categories[catIdx].objectives[objIdx].targets[qIdx] = val;
    el.value = formatCurrency(val);
};

window.updateKPIText = function (el, field, catIdx, objIdx) {
    if (!isAdmin) return;
    kpiStructure.categories[catIdx].objectives[objIdx][field] = el.innerText || el.value || '';
};

window.updateKPINumber = function (el, field, catIdx, objIdx) {
    if (!isAdmin) return;
    const val = parseFloat(el.value) || 0;
    kpiStructure.categories[catIdx].objectives[objIdx][field] = val;
    el.value = val;
    renderKPIView();
};

window.updateKPICategoryName = function (el, catIdx) {
    if (!isAdmin) return;
    kpiStructure.categories[catIdx].name = el.innerText;
};

window.updateKPIObjectiveName = function (el, catIdx, objIdx) {
    if (!isAdmin) return;
    kpiStructure.categories[catIdx].objectives[objIdx].name = el.innerText;
};

window.updateKPISubItem = function (el, catIdx, objIdx, subIdx) {
    if (!isAdmin) return;
    const sub = kpiStructure.categories[catIdx].objectives[objIdx].subItems[subIdx];
    if (sub) sub.name = el.innerText || el.textContent || '';
};

// User-only: achievement edits
window.updateKPISubItemAchievement = function (el, catIdx, objIdx, subIdx, qIdx) {
    if (isAdmin) return;
    if (!kpiAchievements) kpiAchievements = { user: currentUser, year: currentKPIYear, data: {} };
    const key = `${catIdx}_${objIdx}_${subIdx}`;
    if (!kpiAchievements.data[key]) kpiAchievements.data[key] = [0, 0, 0, 0];
    kpiAchievements.data[key][qIdx] = parseCurrency(el.value);
    el.value = formatCurrency(kpiAchievements.data[key][qIdx]);
    renderKPIView();
};

/* ═══════════════════════════════════════════════════════════════
   Tab Entry Point (called from app.js)
   ═══════════════════════════════════════════════════════════════ */

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

    // Restore last user session
    const lastUser = localStorage.getItem('kpi_last_user');
    if (lastUser && lastUser !== 'admin') {
        isAdmin = false;
        currentUser = lastUser;
    }

    loadAvailableUsers().then(() => renderKPIView());
}
