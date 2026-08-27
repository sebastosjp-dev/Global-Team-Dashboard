/**
 * kpi.js — KPI Goals: Admin sets structure/targets, team members enter achievements.
 * Storage: server files (kpi_structure_YEAR.json, kpi_ach_YEAR_NAME.json)
 */
import { parseCurrency, formatCurrency } from './utils.js';
import { getKPIHTML, getKPIDashboardHTML } from './ui.js';
import { getRevenueTypeQuarterlyMix } from './services.js';

/**
 * Revenue-Type quarterly mix (New / Upsell / Recurring) from the ORDER SHEET.
 * Returns null when the workbook isn't loaded yet.
 */
function computeRevenueMix() {
    const orderRows = window.__workbookData?.['ORDER SHEET'];
    if (!orderRows || !orderRows.length) return null;
    try { return getRevenueTypeQuarterlyMix(orderRows); }
    catch (e) { console.warn('[KPI] revenue mix failed:', e); return null; }
}

/* ═══════════════════════════════════════════════════════════════
   State
   ═══════════════════════════════════════════════════════════════ */

let kpiStructure = null;    // Admin's data: categories, objectives, targets, sub-item names
let kpiAchievements = null; // Per-user: { user, year, data: {"ci_oi_si": [q1,q2,q3,q4]} }
let currentKPIYear = new Date().getFullYear();
let currentUser = 'admin';
let isAdmin = true;
let availableUsers = [];
let kpiViewMode = 'dashboard';   // 'dashboard' (default) | 'edit'

/* ═══════════════════════════════════════════════════════════════
   Defaults
   ═══════════════════════════════════════════════════════════════ */

const DEFAULT_STRUCTURE = {
    categories: [
        {
            name: "FINANCIAL", color: "#8b5cf6",
            objectives: [
                { name: "Nett Base + New Revenue",
                  kpis: "900,000 TVC Korea --> 1.8M TCV IDN\nIDN: 900,000 (Base 200,000 + new 600,000 + upsell 100,000)\nMAL: 200,000\nTHAI: 100,000",
                  targets: [100000, 400000, 500000, 200000], weight: 60,
                  subItems: [{name:""},{name:""},{name:""}] },
                { name: "Up/cross selling",
                  kpis: "US$100,000\nBase recurring = US$200,000 × 50%",
                  targets: [0, 50000, 50000, 0], weight: 10,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        },
        {
            name: "CUSTOMER", color: "#f59e0b",
            objectives: [
                { name: "New Strategic Account",
                  kpis: "Named brands 3ea",
                  targets: [0, 1, 1, 1], weight: 5,
                  subItems: [{name:""},{name:""},{name:""}] },
                { name: "Customer Retention",
                  kpis: "Retention Rate : 100%\nfrom last year (20 customers)",
                  targets: [1, 3, 6, 10], weight: 5,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        },
        {
            name: "INTERNAL PROCESSES", color: "#3b82f6",
            objectives: [
                { name: "Conversion : POC to Deal",
                  kpis: "Conversion Rate : 20% (data needed from BA).",
                  targets: [0, 5, 6, 1], weight: 5,
                  subItems: [{name:""},{name:""},{name:""}] }
            ]
        },
        {
            name: "LEARNING & GROWTH", color: "#22c55e",
            objectives: [
                { name: "Online Training",
                  kpis: "500 Hours (for the whole team, 10hrs a month per person, for 5ppl)",
                  targets: [125, 125, 125, 125], weight: 5,
                  subItems: [{name:""},{name:""},{name:""}] },
                { name: "Fundamental workshop",
                  kpis: "Partner Fundamental Workshop. 4 times.\nPOC generation. 5per session. Total 20",
                  targets: [5, 5, 5, 5], weight: 5,
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
   KPI Sheet Parser — reads BSC structure from a raw worksheet
   ═══════════════════════════════════════════════════════════════ */

const CAT_COLORS = {
    'FINANCIAL': '#8b5cf6',
    'CUSTOMER': '#f59e0b',
    'INTERNAL PROCESS': '#3b82f6',
    'INTERNAL PROCESSES': '#3b82f6',
    'LEARNING & GROWTH': '#22c55e',
    'LEARNING AND GROWTH': '#22c55e'
};

function isCategoryName(s) {
    const u = String(s || '').trim().toUpperCase();
    if (!u) return false;
    return /^(FINANCIAL|CUSTOMER|INTERNAL\s+PROCESS(ES)?|LEARNING\s*(&|AND)\s*GROWTH)$/.test(u);
}

function normalizeCategoryName(s) {
    const u = String(s || '').trim().toUpperCase();
    if (/^INTERNAL\s+PROCESS(ES)?$/.test(u)) return 'INTERNAL PROCESSES';
    if (/^LEARNING\s*(&|AND)\s*GROWTH$/.test(u)) return 'LEARNING & GROWTH';
    return u;
}

function toNum(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).replace(/[,$\s%]/g, ''));
    return isNaN(n) ? 0 : n;
}

function toPct(v) {
    if (typeof v === 'number') return v <= 1.0001 ? Math.round(v * 100) : Math.round(v);
    if (!v) return 0;
    const s = String(v).replace(/[%\s]/g, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : Math.round(n);
}

/**
 * Parse a raw KPI worksheet (SheetJS sheet object) into our BSC structure.
 * Returns null when the sheet doesn't look like a Balanced Scorecard.
 */
function parseKPISheet(rawSheet) {
    if (!rawSheet || typeof XLSX === 'undefined') return null;
    const rows = XLSX.utils.sheet_to_json(rawSheet, { header: 1, defval: '', cellDates: false });
    if (!rows.length) return null;

    // 1) Locate header row containing "STRATEGIC OBJECTIVES" and column layout
    let headerRow = -1, colCat = -1, colObj = -1, colKPI = -1, colWeight = -1;
    for (let r = 0; r < Math.min(rows.length, 30); r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
            const v = String(row[c] || '').trim();
            if (/strategic\s+objective/i.test(v)) { headerRow = r; colObj = c; }
            if (/key\s+performance\s+indicator/i.test(v)) colKPI = c;
            if (/^weight$/i.test(v)) colWeight = c;
        }
        if (headerRow === r && colObj >= 0) break;
    }
    if (headerRow < 0 || colObj < 0 || colKPI < 0) return null;
    colCat = Math.max(0, colObj - 1);

    // 2) Locate Q1 column (the row right after the header that has Q1)
    let colQ1 = -1, qHeaderRow = -1;
    for (let r = headerRow; r <= Math.min(headerRow + 2, rows.length - 1); r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
            if (/^q1$/i.test(String(row[c] || '').trim())) { colQ1 = c; qHeaderRow = r; break; }
        }
        if (colQ1 >= 0) break;
    }
    if (colQ1 < 0) return null;

    // 3) Walk data rows. Each objective spans 2 rows: target row + achievement row
    const categories = [];
    let currentCat = null;
    let currentObj = null;
    let achTotals = null; // [q1..q4] aggregated achievements for current objective

    const startRow = qHeaderRow + 1;
    for (let r = startRow; r < rows.length; r++) {
        const row = rows[r];
        if (!row.some(v => v !== '' && v !== null && v !== undefined)) continue;

        const catCellRaw = String(row[colCat] || '').trim();
        const objCell = String(row[colObj] || '').trim();
        const kpiCell = String(row[colKPI] || '').trim();
        const q = [colQ1, colQ1 + 1, colQ1 + 2, colQ1 + 3].map(c => toNum(row[c]));

        const isAchRow = /^achievement$/i.test(objCell) || /^achievement$/i.test(kpiCell);

        // Category transition
        if (isCategoryName(catCellRaw)) {
            const norm = normalizeCategoryName(catCellRaw);
            if (!currentCat || currentCat.name !== norm) {
                currentCat = {
                    name: norm,
                    color: CAT_COLORS[norm] || '#6366f1',
                    objectives: []
                };
                categories.push(currentCat);
            }
        }

        if (isAchRow) {
            if (currentObj) {
                achTotals = q;
                // Stash achievements onto subItems[0] so getMergedData can find them
                currentObj.subItems = [
                    { name: 'Sheet', achievements: q },
                    { name: '', achievements: [0, 0, 0, 0] },
                    { name: '', achievements: [0, 0, 0, 0] }
                ];
                currentObj.__sheetAch = q.slice();
            }
            continue;
        }

        // Target row — must have a non-empty objective name
        if (objCell && objCell.length > 0) {
            const w = colWeight >= 0 ? toPct(row[colWeight]) : 0;
            currentObj = {
                name: objCell,
                kpis: kpiCell,
                targets: q,
                weight: w,
                subItems: [
                    { name: '', achievements: [0, 0, 0, 0] },
                    { name: '', achievements: [0, 0, 0, 0] },
                    { name: '', achievements: [0, 0, 0, 0] }
                ]
            };
            if (!currentCat) {
                currentCat = { name: 'UNCATEGORIZED', color: '#6366f1', objectives: [] };
                categories.push(currentCat);
            }
            currentCat.objectives.push(currentObj);
        }
    }

    if (!categories.length || categories.every(c => !c.objectives.length)) return null;
    return { categories };
}

/* ═══════════════════════════════════════════════════════════════
   Inline KPI dashboard renderer (called from views.js for the 'KPI' sheet tab)
   ═══════════════════════════════════════════════════════════════ */

/**
 * Render the KPI BSC dashboard directly into a container, using sheet data
 * when available and falling back to the saved structure otherwise.
 *
 * @param {HTMLElement} container - usually #tab-metrics-grid
 * @param {Object[]} sheetRows - workbookData['KPI'] (unused except as a presence signal)
 * @param {Object} [rawSheet] - raw SheetJS worksheet for header:1 reparse
 */
export function renderKPISheetDashboard(container, sheetRows, rawSheet) {
    if (!container) return;

    const parsed = parseKPISheet(rawSheet);
    const data = parsed || kpiStructure || JSON.parse(JSON.stringify(DEFAULT_STRUCTURE));

    const year = currentKPIYear || new Date().getFullYear();
    container.innerHTML = `
        ${parsed ? '' : `
            <div style="margin-bottom:14px; padding:10px 14px; background:#FEF3C7; border:1px solid #FCD34D; border-radius:10px; color:#92400E; font-size:0.82rem; font-weight:600;">
                <i class="fa-solid fa-triangle-exclamation" style="margin-right:8px;"></i>
                Could not parse the KPI sheet — showing default Balanced Scorecard structure.
            </div>
        `}
        ${getKPIDashboardHTML(data, year, true, 'admin', [], { readOnly: true, revenueMix: computeRevenueMix() })}
    `;
}

/**
 * BSC quarterly target baseline — used as a fallback when the KPI tab hasn't
 * been saved yet (or its revenue row is all zero). Applied globally and per-
 * country uniformly; KPI tab values override once any quarter is non-zero.
 */
const BSC_FALLBACK_TARGETS = Object.freeze({
    Q1: 100000, Q2: 400000, Q3: 500000, Q4: 200000,
    objectiveName: 'Nett Base + New Revenue',
    source: 'bsc-default'
});

/**
 * Fetch the KPI structure for `year` and return its quarterly Revenue targets.
 * Used by views (e.g. PIPELINE) to overlay annual revenue targets without
 * forcing the user to open the KPI tab.
 *
 * Resolution order:
 *   1. /api/kpi/structure → /api/kpi/load → localStorage (saved KPI structure)
 *   2. First objective whose name contains "revenue" (case-insensitive),
 *      else categories[0].objectives[0]
 *   3. BSC_FALLBACK_TARGETS when no saved data exists or every quarter is 0
 *
 * @param {number} year
 * @returns {Promise<{Q1:number,Q2:number,Q3:number,Q4:number,objectiveName:string,source?:string}>}
 */
export async function loadKPIQuarterlyTargets(year) {
    let structure = null;
    try {
        const res = await fetch(`/api/kpi/structure?year=${year}`);
        if (res.ok) structure = await res.json();
    } catch (_e) {}
    if (!structure) {
        try {
            const res = await fetch(`/api/kpi/load?year=${year}`);
            if (res.ok) structure = await res.json();
        } catch (_e) {}
    }
    if (!structure) {
        const stored = localStorage.getItem(`global_dashboard_kpi_${year}`)
            || (year === 2026 ? localStorage.getItem('global_dashboard_kpi') : null);
        if (stored) { try { structure = JSON.parse(stored); } catch (_e) {} }
    }

    if (structure && Array.isArray(structure.categories)) {
        let obj = null;
        outer: for (const cat of structure.categories) {
            for (const o of (cat.objectives || [])) {
                if (typeof o.name === 'string' && o.name.toLowerCase().includes('revenue')) {
                    obj = o; break outer;
                }
            }
        }
        if (!obj) obj = structure.categories[0]?.objectives?.[0] || null;
        if (obj && Array.isArray(obj.targets) && obj.targets.length >= 4) {
            const t = obj.targets.map(v => Number(v) || 0);
            // If user has saved any non-zero quarter, honor the KPI tab values.
            if (t[0] + t[1] + t[2] + t[3] > 0) {
                return { Q1: t[0], Q2: t[1], Q3: t[2], Q4: t[3], objectiveName: obj.name || 'Revenue', source: 'kpi' };
            }
        }
    }

    return { ...BSC_FALLBACK_TARGETS };
}

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
        const html = kpiViewMode === 'edit'
            ? wrapEditWithBackButton(getKPIHTML(getMergedData(), currentKPIYear, isAdmin, currentUser, availableUsers))
            : getKPIDashboardHTML(getMergedData(), currentKPIYear, isAdmin, currentUser, availableUsers, { revenueMix: computeRevenueMix() });
        metricsGrid.innerHTML = html;
        metricsGrid.classList.remove('hidden');
    }
}

function wrapEditWithBackButton(innerHtml) {
    return `
        <div style="margin-bottom:14px; display:flex; align-items:center; gap:10px;">
            <button onclick="window.toggleKPIView('dashboard')" style="padding:8px 14px; border-radius:10px; border:1px solid #CBD5E1; background:#FFFFFF; color:#1E293B; font-weight:700; font-size:0.85rem; cursor:pointer;">
                <i class="fa-solid fa-arrow-left" style="margin-right:6px;"></i>Back to Dashboard
            </button>
            <div style="font-size:0.78rem; color:#64748B; font-weight:600;">
                ${isAdmin ? 'Editing structure & quarterly targets' : 'Entering quarterly achievements'}
            </div>
        </div>
        ${innerHtml}
    `;
}

/* ═══════════════════════════════════════════════════════════════
   Targeted DOM Updates — no full re-render, preserves focus
   ═══════════════════════════════════════════════════════════════ */

function computeObjRate(catIdx, objIdx) {
    const obj = kpiStructure?.categories[catIdx]?.objectives[objIdx];
    if (!obj) return 0;
    const totalAch = [0,1,2,3].map(q =>
        [0,1,2].reduce((s, si) => s + (kpiAchievements?.data?.[`${catIdx}_${objIdx}_${si}`]?.[q] || 0), 0)
    );
    const sumT = (obj.targets || []).reduce((a, b) => a + b, 0);
    const sumA = totalAch.reduce((a, b) => a + b, 0);
    if (sumT === 0) return sumA > 0 ? 100 : 0;
    return Math.min(200, Math.round((sumA / sumT) * 100));
}

function refreshKPIFooter() {
    if (!kpiStructure) return;
    let totalWeight = 0, totalWeightedRate = 0;
    kpiStructure.categories.forEach((cat, ci) => {
        cat.objectives.forEach((obj, oi) => {
            const w = obj.weight || 0;
            totalWeight += w;
            totalWeightedRate += computeObjRate(ci, oi) * w / 100;
        });
    });
    const weightGap = Math.abs(totalWeight - 100);
    const labelCell = document.getElementById('kpi-footer-label');
    const weightCell = document.getElementById('kpi-footer-weight');
    const rateCell = document.getElementById('kpi-footer-rate');
    if (labelCell) {
        const warning = weightGap > 0.1
            ? `<span style="color:#FCA5A5; font-size:0.72rem; font-weight:600; margin-left:8px;">(Total: ${Math.round(totalWeight)}% — Must be 100%)</span>`
            : '';
        labelCell.innerHTML = `TOTAL WEIGHT${warning}`;
    }
    if (weightCell) {
        weightCell.style.color = weightGap > 0.1 ? '#FCA5A5' : '#86EFAC';
        weightCell.textContent = Math.round(totalWeight) + '%';
    }
    if (rateCell) {
        const color = totalWeightedRate >= 100 ? '#10B981' : (totalWeightedRate >= 70 ? '#F59E0B' : '#EF4444');
        rateCell.style.color = color;
        rateCell.textContent = Math.round(totalWeightedRate) + '%';
    }
}

function refreshKPIObjectiveRow(catIdx, objIdx) {
    [0,1,2,3].forEach(qi => {
        const total = [0,1,2].reduce((s, si) =>
            s + (kpiAchievements?.data?.[`${catIdx}_${objIdx}_${si}`]?.[qi] || 0), 0);
        const cell = document.getElementById(`kpi-ach-total-${catIdx}-${objIdx}-${qi}`);
        if (cell) cell.value = formatCurrency(total);
    });
    const rate = computeObjRate(catIdx, objIdx);
    const rateCell = document.getElementById(`kpi-rate-${catIdx}-${objIdx}`);
    if (rateCell) {
        rateCell.style.color = rate >= 100 ? '#10B981' : (rate >= 70 ? '#F59E0B' : '#EF4444');
        rateCell.textContent = rate + '%';
    }
    refreshKPIFooter();
}

/* ═══════════════════════════════════════════════════════════════
   Public API — window handlers called from inline HTML
   ═══════════════════════════════════════════════════════════════ */

window.toggleKPIView = function (mode) {
    kpiViewMode = (mode === 'edit') ? 'edit' : 'dashboard';
    renderKPIView();
};

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
    refreshKPIObjectiveRow(catIdx, objIdx);
};

window.updateKPIText = function (el, field, catIdx, objIdx) {
    if (!isAdmin) return;
    kpiStructure.categories[catIdx].objectives[objIdx][field] = el.innerText || el.value || '';
};

window.updateKPINumber = function (el, field, catIdx, objIdx) {
    if (!isAdmin) return;
    const val = Math.min(100, Math.max(0, parseFloat(el.value) || 0));
    kpiStructure.categories[catIdx].objectives[objIdx][field] = val;
    el.value = val;
    refreshKPIFooter();
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
    refreshKPIObjectiveRow(catIdx, objIdx);
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

    loadAvailableUsers().then(() => {
        if (!isAdmin && !availableUsers.includes(currentUser)) {
            isAdmin = true;
            currentUser = 'admin';
            localStorage.removeItem('kpi_last_user');
        }
        renderKPIView();
    });
}
