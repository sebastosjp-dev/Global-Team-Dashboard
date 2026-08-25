/**
 * views.js — Render orchestrators connecting Stats → HTML → Charts.
 * Extracted from app.js for single-responsibility.
 * @module views
 */
import { CONFIG } from './config.js';
import { parseCurrency, formatCurrency, isCountryMatch, findKey, parseQuarterTag } from './utils.js';
import {
    chartRegistry, initOrderSheetCharts, initPipelineCharts,
    initPartnerCharts, initPartnerPerformanceCharts,
    initPocCharts, initEventCharts, initServiceAnalysisCharts,
    initTcvArrChart, initProjectCharts, initDealLostCharts,
    initTaskDashboardCharts
} from './charts.js';
import {
    getOrderSheetStats, getPipelineStats, getPartnerStats,
    getGenericCountryStats, getExpiringContractsStats,
    getPartnerPerformanceStats, getPocStats, getEventStats,
    getCountrySpecificStats, getServiceAnalysisStats,
    getCollectionStats, getCollectionFileLinks,
    getTcvArrStats, getChurnRiskStats,
    getPartnerROIStats, getPipelineCoverageStats,
    getProjectStats, getDealLostStats,
    getQuarterlyForecastStats, getCsmTaskStats,
    getCsmViewStats, getTaskDashboardStats
} from './services.js';
import {
                                                getOrderSheetHTML, getPipelineHTML, getPartnerHTML, getPartnerNetworkDetailsHTML,
    getGenericCountryHTML, getExpiringContractsHTML,
    getPartnerPerformanceHTML, getPocHTML, getEventHTML,
    getCountrySpecificHTML, getServiceAnalysisHTML,
    getKPIHTML, getCollectionHTML,
    getTcvArrHTML, getChurnRiskHTML,
    getPartnerROIHTML, getPipelineCoverageHTML,
    getPipelineChangeLogHTML, getCurrentPipelineListHTML,
    getCollectionChangeLogHTML,
    getProjectHTML, getDealLostHTML,
    getQuarterlyForecastHTML, getCsmTasksByClientHTML,
    getCsmViewHTML, getTaskDashboardHTML
} from './ui.js';
import { loadKPIQuarterlyTargets, renderKPISheetDashboard } from './kpi.js';

/* ═══════════════════════════════════════════════════════════════
   Metrics Router
   ═══════════════════════════════════════════════════════════════ */

/**
 * Top-level metrics router — decides which panels to render.
 * @param {Object[]} data - Filtered rows
 * @param {string} tabName - Current tab name
 * @param {string|null} filterCountry
 * @param {Object} workbookData - Full workbook data map
 * @param {HTMLInputElement} searchInput - Search input element
 */
export function renderTabMetrics(data, tabName, filterCountry, workbookData, searchInput) {
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    let hasMetrics = false;
    const isGlobalTab = tabName && tabName.includes('Global(Contract Date)');
    const isCountryTab = tabName && filterCountry === null &&
        !['ORDER SHEET', 'PIPELINE', 'PARTNER', 'POC', 'EVENT', 'END USER (CSM)', 'COLLECTION', 'PROJECT', 'DEAL LOST', 'TASK', 'KPI'].includes(tabName) &&
        !isGlobalTab;

    if (tabName === 'ORDER SHEET' || isGlobalTab) {
        _renderQuarterlyForecast(workbookData, filterCountry, metricsGrid);
        _renderOrderSheet(data, filterCountry, metricsGrid, tabName, workbookData);
        _renderChurnRisk(data, workbookData, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'PIPELINE' && workbookData['PIPELINE']) {
        _renderPipeline(workbookData, filterCountry, tabName, metricsGrid, searchInput);
        _renderPipelineCoverage(workbookData, filterCountry, metricsGrid);
        _renderPipelineChangeLog(workbookData, filterCountry, metricsGrid, searchInput);
        _renderCurrentPipelineList(workbookData, filterCountry, metricsGrid);
        hasMetrics = true;
    }

    if ((tabName === 'PARTNER' || isCountryTab) && data && data.length > 0) {
        try { _renderPartner(data, filterCountry, tabName, metricsGrid, workbookData, searchInput); } catch(e) { console.error('_renderPartner', e); }
        try { _renderPartnerTopPerformer(data, metricsGrid); } catch(e) { console.error('_renderPartnerTopPerformer', e); }
        try { _renderPartnerROI(workbookData, filterCountry, metricsGrid); } catch(e) { console.error('_renderPartnerROI', e); }
        try { _renderGenericCountry(data, filterCountry, metricsGrid, tabName); } catch(e) { console.error('_renderGenericCountry', e); }
        try { _renderPartnerNetworkDetails(data, filterCountry, metricsGrid, workbookData); } catch(e) { console.error('_renderPartnerNetworkDetails', e); }
        hasMetrics = true;
    }

    if ((String(tabName).trim().toUpperCase() === 'POC' || isCountryTab) && data && data.length > 0) {
        _renderPoc(data, filterCountry, metricsGrid, workbookData);
        hasMetrics = true;
    }

    if (tabName === 'EVENT' && workbookData['EVENT']) {
        _renderEvent(workbookData['EVENT'], filterCountry, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'END USER (CSM)' && workbookData['END USER (CSM)']) {
        _renderServiceAnalysis(workbookData['END USER (CSM)'], filterCountry, tabName, metricsGrid, searchInput, workbookData);
        hasMetrics = true;
    }

    if (tabName === 'COLLECTION' && data && data.length > 0) {
        _renderCollection(data, filterCountry, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'TCV_ARR' && workbookData['ORDER SHEET']) {
        _renderTcvArr(workbookData, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'PROJECT' && workbookData['PROJECT']) {
        _renderProject(workbookData['PROJECT'], filterCountry, metricsGrid, searchInput, workbookData);
        hasMetrics = true;
    }

    if (tabName === 'DEAL LOST' && workbookData['PIPELINE']) {
        _renderDealLost(workbookData['PIPELINE'], metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'TASK' && workbookData['TASK']) {
        _renderTaskDashboard(workbookData['TASK'], filterCountry, metricsGrid, workbookData);
        hasMetrics = true;
    }

    if (tabName === 'KPI') {
        renderKPISheetDashboard(metricsGrid, workbookData['KPI'] || [], window.__rawSheets?.['KPI']);
        hasMetrics = true;
    }

    if (hasMetrics) metricsGrid.classList.remove('hidden');
    else metricsGrid.classList.add('hidden');
}




/* ═══════════════════════════════════════════════════════════════
   Service Analysis Metrics
   ═══════════════════════════════════════════════════════════════ */

function _renderServiceAnalysis(data, filterCountry, tabName, metricsGrid, searchInput, workbookData) {
    if (!data || data.length === 0) return;
    const stats = getServiceAnalysisStats(data, filterCountry);

    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getServiceAnalysisHTML(stats, filterCountry);

    metricsGrid.appendChild(container);

    const taskRows = (workbookData && workbookData['TASK']) || [];
    const taskStats = getCsmTaskStats(taskRows, filterCountry);
    const taskContainer = document.createElement('div');
    taskContainer.style.gridColumn = '1 / -1';
    taskContainer.innerHTML = getCsmTasksByClientHTML(taskStats, filterCountry || 'All');
    metricsGrid.appendChild(taskContainer);

    setTimeout(() => {
        const selector = document.getElementById('csm-filter-country');
        if (selector) {
            selector.addEventListener('change', (e) => {
                const val = e.target.value;
                window.dispatchEvent(new CustomEvent('filter-country-change', {
                    detail: { country: val === 'All' ? null : val, searchTerm: searchInput?.value || '' }
                }));
            });
        }
        if (stats) initServiceAnalysisCharts(stats);
    }, 100);
}

/* ═══════════════════════════════════════════════════════════════
   Private Render Orchestrators
   ═══════════════════════════════════════════════════════════════ */

/** @param {function} renderTableData - closure from app.js */
function _renderOrderSheet(data, filterCountry, metricsGrid, tabName, workbookData) {
    const stats = getOrderSheetStats(data, filterCountry, tabName, workbookData);
    // Stash for the ACCUMULATED KTCV click-to-compare modal in ui.js
    window.__orderSheetStats = stats;
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getOrderSheetHTML(stats, filterCountry);
    metricsGrid.appendChild(container);
    setTimeout(() => initOrderSheetCharts(stats), 120);
}

/**
 * Quarterly Forecast panel — country or Global.
 * Shows Q1–Q4 New (Booked + Forecast) + Renewal targets. Stats are
 * stashed on window so the click-to-expand modal can read full deals.
 */
function _renderQuarterlyForecast(workbookData, filterCountry, metricsGrid) {
    const stats = getQuarterlyForecastStats(workbookData, filterCountry);
    if (!stats) return;
    window.__qForecastStats = stats;
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getQuarterlyForecastHTML(stats);
    metricsGrid.appendChild(container);
}

async function _renderPipeline(workbookData, filterCountry, tabName, metricsGrid, searchInput) {
    const pData = filterCountry
        ? workbookData['PIPELINE'].filter(r => isCountryMatch(r, filterCountry))
        : workbookData['PIPELINE'];
    if (!pData || pData.length === 0) return;

    const oDataRaw = workbookData['ORDER SHEET'] || [];
    const oData = filterCountry ? oDataRaw.filter(r => isCountryMatch(r, filterCountry)) : oDataRaw;

    const stats = getPipelineStats(pData, oData);

    // Reserve grid slot synchronously so sibling renders (coverage, change-log)
    // stay in their original order while we await KPI targets.
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.style.marginTop = '12px';
    container.style.marginBottom = '24px';
    metricsGrid.appendChild(container);

    let kpiTargets = null;
    try { kpiTargets = await loadKPIQuarterlyTargets(new Date().getFullYear()); }
    catch (e) { console.warn('[Pipeline] loadKPIQuarterlyTargets failed:', e); }

    container.innerHTML = getPipelineHTML(stats, filterCountry, tabName, kpiTargets);
    setTimeout(() => {
        const selector = document.getElementById('pipeline-filter-country');
        if (selector) {
            selector.addEventListener('change', (e) => {
                const val = e.target.value;
                // Dispatch custom event for app.js to handle
                window.dispatchEvent(new CustomEvent('filter-country-change', {
                    detail: { country: val === 'All' ? null : val, searchTerm: searchInput?.value || '' }
                }));
            });
        }
        initPipelineCharts(stats, kpiTargets);
    }, 100);
}

function _renderPartner(data, filterCountry, tabName, metricsGrid, workbookData, searchInput) {
    if (!data || data.length === 0) return;
    window.partnerFilters = window.partnerFilters || { year: 'all' };
    const stats = getPartnerStats(data, filterCountry, workbookData, window.partnerFilters.year);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getPartnerHTML(stats, filterCountry, tabName);
    metricsGrid.appendChild(container);
    setTimeout(() => {
        const selector = document.getElementById('partner-filter-country');
        if (selector) {
            selector.addEventListener('change', (e) => {
                const val = e.target.value;
                window.dispatchEvent(new CustomEvent('filter-country-change', {
                    detail: { country: val === 'All' ? null : val, searchTerm: searchInput?.value || '' }
                }));
            });
        }
        const yearSel = document.getElementById('partner-filter-year');
        if (yearSel) {
            yearSel.addEventListener('change', (e) => {
                window.partnerFilters.year = e.target.value;
                window.dispatchEvent(new CustomEvent('filter-country-change', {
                    detail: { country: filterCountry, searchTerm: searchInput?.value || '' }
                }));
            });
        }
    }, 100);
}

function _renderPartnerNetworkDetails(data, filterCountry, metricsGrid, workbookData) {
    if (!data || data.length === 0) return;
    const yearFilter = (window.partnerFilters && window.partnerFilters.year) || 'all';
    const stats = getPartnerStats(data, filterCountry, workbookData, yearFilter);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getPartnerNetworkDetailsHTML(stats, filterCountry);
    metricsGrid.appendChild(container);
    setTimeout(() => { initPartnerCharts(stats, filterCountry); }, 100);
}

function _renderGenericCountry(data, filterCountry, metricsGrid, tabName) {
    if (data.length === 0 || tabName === 'EVENT') return;
    const stats = getGenericCountryStats(data, filterCountry);
    if (!stats) return;
    const wrapper = document.createElement('div');
    wrapper.style.gridColumn = '1 / -1';
    wrapper.innerHTML = getGenericCountryHTML(stats, filterCountry);
    metricsGrid.appendChild(wrapper);
}

function _renderPartnerROI(workbookData, filterCountry, metricsGrid) {
    const pocData = workbookData['POC'] || [];
    console.log('[PartnerROI] pocData rows:', pocData.length);
    if (pocData.length > 0) console.log('[PartnerROI] sample keys:', Object.keys(pocData[0]).join(', '));

    const stats = getPartnerROIStats(pocData, filterCountry);
    console.log('[PartnerROI] stats:', stats);
    if (!stats) {
        console.warn('[PartnerROI] getPartnerROIStats returned null — nothing to render');
        return;
    }
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    try {
        container.innerHTML = getPartnerROIHTML(stats);
    } catch (e) {
        console.error('[PartnerROI] getPartnerROIHTML threw:', e);
        return;
    }
    metricsGrid.appendChild(container);
    console.log('[PartnerROI] rendered OK, partners:', stats.partners.length);
}

function _renderPipelineCoverage(workbookData, filterCountry, metricsGrid) {
    const pData = filterCountry
        ? (workbookData['PIPELINE'] || []).filter(r => isCountryMatch(r, filterCountry))
        : workbookData['PIPELINE'] || [];
    const oData = filterCountry
        ? (workbookData['ORDER SHEET'] || []).filter(r => isCountryMatch(r, filterCountry))
        : workbookData['ORDER SHEET'] || [];
    const stats = getPipelineCoverageStats(pData, oData);
    if (!stats) return;
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.style.marginTop = '12px';
    container.innerHTML = getPipelineCoverageHTML(stats);
    metricsGrid.appendChild(container);
}

/**
 * Extract a normalized deal-level view from PIPELINE rows for diff tracking.
 * Each deal carries a stable key (customer::name) so we can detect added /
 * removed / modified deals across snapshots — even when totals stay identical
 * (e.g. swap of equal-value deals).
 * @returns {Array<{key, name, customer, quarter, amount, weighted}>}
 */
function _extractPipelineDeals(pData) {
    if (!pData || pData.length === 0) return [];
    const keys = Object.keys(pData[0]);
    const nameKey = findKey(keys,
        k => k.toLowerCase().includes('deal name'),
        k => k.toLowerCase().includes('crm deal name'));
    const customerKey = findKey(keys,
        k => k.toLowerCase().includes('customer'),
        k => k.toLowerCase().includes('end user'),
        k => k.toLowerCase().includes('account'));
    const amtKey = findKey(keys, k => (k.toUpperCase().includes('KOR TCV') && k.toUpperCase().includes('USD')) || k === 'Amount') || 'Amount';
    const wAmtKey = findKey(keys, k => (k.toUpperCase().includes('WEIGHTED') && k.toUpperCase().includes('KOR TCV')) || k === 'Weighted Amount') || 'Weighted Amount';
    const qKey = findKey(keys,
        k => k.toLowerCase() === 'quarter',
        k => k.toLowerCase().includes('qtr'),
        k => k.toLowerCase() === 'q');

    const seen = new Map();
    return pData.map((row, i) => {
        const name = nameKey ? String(row[nameKey] || '').trim() : '';
        const customer = customerKey ? String(row[customerKey] || '').trim() : '';
        const amount = Math.round(parseCurrency(row[amtKey]));
        const weighted = Math.round(parseCurrency(row[wAmtKey]));
        // Keep the year suffix for non-current years ("Q1-2027") so a deal
        // pushed to next year surfaces as a quarter change in the weekly diff.
        let quarter = '';
        if (qKey && row[qKey]) {
            const tag = parseQuarterTag(row[qKey]);
            if (tag) {
                quarter = (tag.year !== null && tag.year !== new Date().getFullYear())
                    ? `${tag.q}-${tag.year}`
                    : tag.q;
            }
        }
        // Build a stable-ish key. Disambiguate duplicates by appending #N.
        const baseKey = (customer && name) ? `${customer}::${name}`
            : (name || customer || `row-${i}`);
        const dupCount = seen.get(baseKey) || 0;
        seen.set(baseKey, dupCount + 1);
        const key = dupCount === 0 ? baseKey : `${baseKey}#${dupCount + 1}`;
        return {
            key,
            name: name || '(unnamed)',
            customer,
            quarter,
            amount,
            weighted
        };
    });
}

function _dealsFingerprint(deals) {
    return deals
        .map(d => `${d.key}|${d.amount}|${d.weighted}|${d.quarter}`)
        .sort()
        .join('||');
}

function _diffDeals(beforeDeals, afterDeals) {
    const beforeMap = new Map((beforeDeals || []).map(d => [d.key, d]));
    const afterMap = new Map((afterDeals || []).map(d => [d.key, d]));
    const added = [];
    const removed = [];
    const modified = [];
    afterMap.forEach((a, k) => { if (!beforeMap.has(k)) added.push(a); });
    beforeMap.forEach((b, k) => { if (!afterMap.has(k)) removed.push(b); });
    afterMap.forEach((a, k) => {
        const b = beforeMap.get(k);
        if (!b) return;
        if (b.amount !== a.amount || b.weighted !== a.weighted || b.quarter !== a.quarter) {
            modified.push({ before: b, after: a });
        }
    });
    return { added, removed, modified };
}

/**
 * Auto-snapshot the country's pipeline (totals + deal-level) to localStorage
 * and render the historical change log. Only runs when a country is selected.
 */
function _renderPipelineChangeLog(workbookData, filterCountry, metricsGrid, searchInput) {
    if (!filterCountry) return; // country-specific only
    const pipelineRows = workbookData['PIPELINE'] || [];
    const pData = pipelineRows.filter(r => isCountryMatch(r, filterCountry));
    if (pData.length === 0) return;

    const oData = (workbookData['ORDER SHEET'] || []).filter(r => isCountryMatch(r, filterCountry));
    const stats = getPipelineStats(pData, oData);

    const byQuarter = {};
    (stats.sortedQuarterly || []).forEach(([q, qData]) => {
        const totals = Object.values(qData.countries || {}).reduce((a, c) => ({
            amount: a.amount + (c.amount || 0),
            weighted: a.weighted + (c.weighted || 0),
            count: a.count + (c.count || 0)
        }), { amount: 0, weighted: 0, count: 0 });
        byQuarter[q] = {
            amount: Math.round(totals.amount),
            weighted: Math.round(totals.weighted),
            count: totals.count
        };
    });

    const deals = _extractPipelineDeals(pData);
    const current = {
        count: stats.globalTotalCount || 0,
        amount: Math.round(stats.globalTotalAmount || 0),
        weighted: Math.round(stats.globalTotalWeighted || 0),
        tcv: Math.round(stats.globalTotalTcv || 0),
        byQuarter,
        deals,
        dealsFp: _dealsFingerprint(deals)
    };

    const storageKey = `pipelineChangeLog::${filterCountry}`;
    let history = [];
    try {
        const raw = localStorage.getItem(storageKey);
        if (raw) history = JSON.parse(raw);
        if (!Array.isArray(history)) history = [];
    } catch { history = []; }

    const last = history[history.length - 1];
    const quartersDiffer = (a, b) => {
        if (!a || !b) return true;
        return ['Q1', 'Q2', 'Q3', 'Q4'].some(q => {
            const ax = a[q] || {}; const bx = b[q] || {};
            return (ax.amount || 0) !== (bx.amount || 0)
                || (ax.weighted || 0) !== (bx.weighted || 0)
                || (ax.count || 0) !== (bx.count || 0);
        });
    };
    const changed = !last
        || last.count !== current.count
        || last.amount !== current.amount
        || last.weighted !== current.weighted
        || last.tcv !== current.tcv
        || quartersDiffer(last.byQuarter, current.byQuarter)
        || (last.dealsFp || '') !== current.dealsFp;

    if (changed) {
        history.push({ date: new Date().toISOString(), ...current });
        if (history.length > 100) history = history.slice(-100);
        try { localStorage.setItem(storageKey, JSON.stringify(history)); }
        catch (e) { console.warn('[ChangeLog] localStorage write failed (likely quota):', e); }
    }

    // Pre-compute deal-level diffs between consecutive snapshots so the UI layer
    // doesn't have to know how snapshots are structured.
    const sorted = [...history].sort((a, b) => new Date(a.date) - new Date(b.date));
    const dealDiffs = sorted.map((snap, i) => {
        if (i === 0) return null;
        return _diffDeals(sorted[i - 1].deals || [], snap.deals || []);
    });

    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.style.marginTop = '12px';
    container.innerHTML = getPipelineChangeLogHTML(filterCountry, history, dealDiffs);
    metricsGrid.appendChild(container);

    setTimeout(() => {
        const resetBtn = document.getElementById('pipeline-changelog-reset');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                if (!confirm(`Clear all pipeline change history for ${filterCountry}? This cannot be undone.`)) return;
                try { localStorage.removeItem(storageKey); } catch {}
                window.dispatchEvent(new CustomEvent('filter-country-change', {
                    detail: { country: filterCountry, searchTerm: searchInput?.value || '' }
                }));
            });
        }
    }, 80);
}

/**
 * Render the current deal-level pipeline list at the very bottom of the
 * country pipeline page. This is the live anchor everything in the change log
 * is compared against.
 */
function _renderCurrentPipelineList(workbookData, filterCountry, metricsGrid) {
    if (!filterCountry) return;
    const pipelineRows = workbookData['PIPELINE'] || [];
    const pData = pipelineRows.filter(r => isCountryMatch(r, filterCountry));
    if (pData.length === 0) return;

    const deals = _extractPipelineDeals(pData);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.style.marginTop = '12px';
    container.innerHTML = getCurrentPipelineListHTML(filterCountry, deals);
    metricsGrid.appendChild(container);
}

function _renderChurnRisk(orderData, workbookData, metricsGrid) {
    const csmData = workbookData['END USER (CSM)'] || [];
    const stats = getChurnRiskStats(orderData, csmData);
    if (!stats) return;
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getChurnRiskHTML(stats);
    metricsGrid.appendChild(container);
}

function _renderPartnerTopPerformer(data, metricsGrid) {
    const stats = getPartnerPerformanceStats(data);
    if (!stats) return;
    const div = document.createElement('div');
    div.style.gridColumn = '1 / -1';
    div.innerHTML = getPartnerPerformanceHTML();
    metricsGrid.appendChild(div);
    setTimeout(() => initPartnerPerformanceCharts(stats), 100);
}

function _renderPoc(data, filterCountry, metricsGrid, workbookData) {
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

function _renderProject(data, filterCountry, metricsGrid, searchInput, workbookData) {
    if (!data || data.length === 0) return;
    metricsGrid.innerHTML = '';
    const projectContainer = document.createElement('div');
    projectContainer.id = 'project-dashboard-container';
    projectContainer.style.gridColumn = '1 / -1';
    metricsGrid.appendChild(projectContainer);

    window.projectFilters = window.projectFilters || { country: 'All', poc: 'All', endUser: 'All' };
    // Carry forward older sessions that only stored country.
    if (!window.projectFilters.poc) window.projectFilters.poc = 'All';
    if (!window.projectFilters.endUser) window.projectFilters.endUser = 'All';

    const pocSheet = (workbookData && workbookData['POC']) || [];
    const orderSheet = (workbookData && workbookData['ORDER SHEET']) || [];

    window.renderProjectUI = function () {
        const { stats, uniqueValues } = getProjectStats(data, window.projectFilters, pocSheet, orderSheet);
        const container = document.getElementById('project-dashboard-container');
        if (container) {
            container.innerHTML = getProjectHTML(stats, window.projectFilters, uniqueValues);
            const selC = document.getElementById('project-filter-country');
            if (selC) selC.addEventListener('change', (e) => { window.projectFilters.country = e.target.value; window.renderProjectUI(); });
            const selP = document.getElementById('project-filter-poc');
            if (selP) selP.addEventListener('change', (e) => { window.projectFilters.poc = e.target.value; window.renderProjectUI(); });
            const selE = document.getElementById('project-filter-enduser');
            if (selE) selE.addEventListener('change', (e) => { window.projectFilters.endUser = e.target.value; window.renderProjectUI(); });
            setTimeout(() => initProjectCharts(stats), 50);
        }
    };

    window.renderProjectUI();
}

function _renderDealLost(data, metricsGrid) {
    if (!data) return;
    metricsGrid.innerHTML = '';
    const container = document.createElement('div');
    container.id = 'deallost-dashboard-container';
    container.style.gridColumn = '1 / -1';
    metricsGrid.appendChild(container);

    window.dealLostFilters = window.dealLostFilters || { country: 'All' };

    window.renderDealLostUI = function () {
        const { stats, uniqueValues } = getDealLostStats(data, window.dealLostFilters.country);
        const el = document.getElementById('deallost-dashboard-container');
        if (!el) return;
        el.innerHTML = getDealLostHTML(stats, window.dealLostFilters.country, uniqueValues);
        const sel = document.getElementById('deallost-filter-country');
        if (sel) sel.addEventListener('change', (e) => {
            window.dealLostFilters.country = e.target.value;
            window.renderDealLostUI();
        });
        setTimeout(() => initDealLostCharts(stats), 50);
    };

    window.renderDealLostUI();
}

/**
 * Render the TASK operational dashboard — distributions, throughput,
 * resolution time, backlog aging, country/client/service breakdowns,
 * and a recent activity feed. Country filter handled via filterCountry.
 *
 * Hosts two sub-tabs: End User (operational dashboard) and CSM
 * (a filterable task-log view, originally a top-level sidebar tab).
 *
 * @param {Object[]} taskData
 * @param {string|null} filterCountry
 * @param {HTMLElement} metricsGrid
 * @param {Object} workbookData - Full workbook (CSM sub-view reads TASK here)
 */
function _renderTaskDashboard(taskData, filterCountry, metricsGrid, workbookData) {
    window.taskFilters = window.taskFilters || { view: 'EndUser' };
    if (!['EndUser', 'CSM'].includes(window.taskFilters.view)) {
        window.taskFilters.view = 'EndUser';
    }

    const tabBar = document.createElement('div');
    tabBar.id = 'task-tab-bar';
    tabBar.style.gridColumn = '1 / -1';
    tabBar.style.marginBottom = '6px';

    const container = document.createElement('div');
    container.id = 'task-dashboard-container';
    container.style.gridColumn = '1 / -1';

    metricsGrid.appendChild(tabBar);
    metricsGrid.appendChild(container);

    const renderTabBar = () => {
        const tabs = [
            { key: 'EndUser', label: 'End User' },
            { key: 'CSM', label: 'CSM' }
        ];
        const v = window.taskFilters.view;
        tabBar.innerHTML = `
            <div style="display:flex; gap:4px; border-bottom:2px solid #e5e7eb;">
                ${tabs.map(t => `
                    <button data-task-view="${t.key}" style="
                        padding:10px 22px;
                        background:${v === t.key ? '#FFF' : 'transparent'};
                        border:1px solid ${v === t.key ? '#e5e7eb' : 'transparent'};
                        border-bottom:3px solid ${v === t.key ? '#6366f1' : 'transparent'};
                        border-radius:8px 8px 0 0;
                        color:${v === t.key ? '#4338ca' : '#6b7280'};
                        font-weight:${v === t.key ? '700' : '600'};
                        font-size:0.85rem;
                        cursor:pointer;
                        margin-bottom:-2px;
                    ">${t.label}</button>
                `).join('')}
            </div>
        `;
        tabBar.querySelectorAll('[data-task-view]').forEach(btn => {
            btn.addEventListener('click', () => {
                window.taskFilters.view = btn.dataset.taskView;
                renderAll();
            });
        });
    };

    const renderCsmSubView = () => {
        window.csmViewFilters = window.csmViewFilters || { category: 'All', endUser: 'All', status: 'All' };
        const { stats, uniqueValues } = getCsmViewStats(taskData, window.csmViewFilters);
        container.innerHTML = getCsmViewHTML(stats, uniqueValues);

        const catSel = document.getElementById('csmview-filter-category');
        const euSel = document.getElementById('csmview-filter-enduser');
        const stSel = document.getElementById('csmview-filter-status');
        const reset = document.getElementById('csmview-reset');
        if (catSel) catSel.addEventListener('change', e => { window.csmViewFilters.category = e.target.value; renderCsmSubView(); });
        if (euSel) euSel.addEventListener('change', e => { window.csmViewFilters.endUser = e.target.value; renderCsmSubView(); });
        if (stSel) stSel.addEventListener('change', e => { window.csmViewFilters.status = e.target.value; renderCsmSubView(); });
        if (reset) reset.addEventListener('click', () => {
            window.csmViewFilters = { category: 'All', endUser: 'All', status: 'All' };
            renderCsmSubView();
        });
    };

    const renderAll = () => {
        renderTabBar();
        if (window.taskFilters.view === 'CSM') {
            renderCsmSubView();
            return;
        }
        const stats = getTaskDashboardStats(taskData, filterCountry, window.taskFilters.view);
        container.innerHTML = getTaskDashboardHTML(stats);
        if (stats && stats.totalTasks > 0) setTimeout(() => initTaskDashboardCharts(stats), 80);
    };

    renderAll();
}

function _renderEvent(eventData, filterCountry, metricsGrid) {
    const stats = getEventStats(eventData, filterCountry);
    if (!stats) return;
    const eventHeader = document.createElement('div');
    eventHeader.style.gridColumn = '1 / -1';
    eventHeader.style.marginBottom = '25px';
    eventHeader.innerHTML = getEventHTML(stats);
    metricsGrid.prepend(eventHeader);
    setTimeout(() => initEventCharts(stats), 150);
}

/**
 * Render the COLLECTION dashboard: two sub-tabs (Active Collections /
 * Pending Records), interactive aging + distributor charts with
 * click-to-expand year-of-Due-Date breakdowns, top-outstanding and pending
 * tables with per-contract detail modals, collection trend, payment
 * timeliness, and the change log.
 *
 * @param {Object[]} data - COLLECTION rows (country/search filtered)
 * @param {string|null} filterCountry
 * @param {HTMLElement} metricsGrid
 */
function _renderCollection(data, filterCountry, metricsGrid) {
    const stats = getCollectionStats(data);

    // Stash for the contract detail modal (ui.js window.showCollectionDealDetail)
    // and the KPI drill-down modal (ui.js window.showCollectionKpiDetail).
    window.__collectionDeals = stats.deals;
    window.__collectionStats = stats;
    window.__collectionFileLinks = getCollectionFileLinks(window.__rawSheets ? window.__rawSheets['COLLECTION'] : null);

    if (window.collectionView !== 'active' && window.collectionView !== 'pending') {
        window.collectionView = 'active';
    }

    const tabBar = document.createElement('div');
    tabBar.id = 'collection-tab-bar';
    tabBar.style.gridColumn = '1 / -1';
    tabBar.style.marginBottom = '6px';

    const container = document.createElement('div');
    container.id = 'collection-dashboard-container';
    container.style.gridColumn = '1 / -1';

    metricsGrid.appendChild(tabBar);
    metricsGrid.appendChild(container);

    const destroyCharts = () => {
        ['collectionAgingChart', 'collectionDistChart', 'collectionTrendChart', 'collectionTimelinessChart'].forEach(k => {
            if (window[k]) {
                try { window[k].destroy(); } catch { /* already detached */ }
                window[k] = null;
            }
        });
    };

    /** Year-of-Due-Date drill-down panel shown under a clicked bar. */
    const renderExpand = (hostId, title, byYear) => {
        const host = document.getElementById(hostId);
        if (!host) return;
        const total = byYear.reduce((s, y) => s + y.amount, 0);
        const max = Math.max(...byYear.map(y => y.amount), 1);
        host.style.display = 'block';
        host.innerHTML = `
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 14px;">
                <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:8px;">
                    <div style="font-size:0.68rem; color:#4338ca; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">
                        <i class="fa-solid fa-calendar-days" style="margin-right:5px;"></i>${title} — outstanding by year of Due Date
                    </div>
                    <button data-expand-close style="border:none; background:#e2e8f0; color:#475569; width:22px; height:22px; border-radius:6px; cursor:pointer; font-size:0.7rem; flex-shrink:0;"><i class="fa-solid fa-xmark"></i></button>
                </div>
                ${byYear.length === 0
                    ? '<div style="font-size:0.72rem; color:#94a3b8; font-style:italic;">Nothing outstanding in this bar.</div>'
                    : byYear.map(y => `
                    <div style="display:flex; align-items:center; gap:10px; margin-bottom:5px;">
                        <div style="width:110px; font-size:0.72rem; font-weight:700; color:#334155; white-space:nowrap;">${y.label}</div>
                        <div style="flex:1; background:#e2e8f0; border-radius:5px; height:10px; overflow:hidden;">
                            <div style="width:${Math.max(2, Math.round((y.amount / max) * 100))}%; height:100%; background:#6366f1; border-radius:5px;"></div>
                        </div>
                        <div style="width:100px; text-align:right; font-size:0.72rem; font-weight:800; color:#111827; white-space:nowrap;">$${formatCurrency(y.amount)}</div>
                        <div style="width:70px; text-align:right; font-size:0.62rem; color:#94a3b8; white-space:nowrap;">${y.dealCount} deal${y.dealCount === 1 ? '' : 's'}</div>
                    </div>`).join('')}
                <div style="margin-top:6px; padding-top:6px; border-top:1px dashed #cbd5e1; display:flex; justify-content:space-between; font-size:0.7rem; color:#64748b;">
                    <span>Total</span><span style="font-weight:800; color:#111827;">$${formatCurrency(total)}</span>
                </div>
            </div>`;
        host.querySelector('[data-expand-close]').addEventListener('click', () => {
            host.style.display = 'none';
            host.innerHTML = '';
        });
    };

    /** Installment-level drill-down under a clicked Collection-plan bar. */
    const renderTrendExpand = (month) => {
        const host = document.getElementById('collection-trend-expand');
        if (!host) return;
        const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        const pct = month.expected > 0 ? Math.round((month.collected / month.expected) * 100) : 0;
        const statusPill = (r) => {
            const c = r.status === 'Paid' ? { bg: '#dcfce7', fg: '#166534' }
                : r.status === 'Partial' ? { bg: '#fef3c7', fg: '#92400e' }
                : (r.isOutstanding && r.daysOverdue > 0) ? { bg: '#fee2e2', fg: '#991b1b' }
                : { bg: '#f1f5f9', fg: '#475569' };
            const label = (r.status === 'Unpaid' && r.daysOverdue !== null && r.daysOverdue > 0)
                ? `Overdue ${r.daysOverdue}d` : r.status;
            return `<span style="background:${c.bg}; color:${c.fg}; padding:2px 8px; border-radius:999px; font-size:0.66rem; font-weight:700; white-space:nowrap;">${esc(label)}</span>`;
        };
        // Money still missing first (most overdue at top), then paid rows.
        const sorted = [...month.rows].sort((a, b) => {
            if (a.isOutstanding !== b.isOutstanding) return a.isOutstanding ? -1 : 1;
            return b.balance - a.balance;
        });
        host.style.display = 'block';
        host.innerHTML = `
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 14px;">
                <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:8px;">
                    <div style="font-size:0.68rem; color:#4338ca; font-weight:800; text-transform:uppercase; letter-spacing:0.05em;">
                        <i class="fa-solid fa-calendar-days" style="margin-right:5px;"></i>${esc(month.label)} — installments due this month
                    </div>
                    <button data-expand-close style="border:none; background:#e2e8f0; color:#475569; width:22px; height:22px; border-radius:6px; cursor:pointer; font-size:0.7rem; flex-shrink:0;"><i class="fa-solid fa-xmark"></i></button>
                </div>
                <div style="display:flex; gap:14px; flex-wrap:wrap; margin-bottom:10px; font-size:0.72rem;">
                    <span style="color:#64748b;">Expected <b style="color:#111827;">$${formatCurrency(month.expected)}</b></span>
                    <span style="color:#64748b;">Collected <b style="color:#7c3aed;">$${formatCurrency(month.collected)}</b> (${pct}%)</span>
                    <span style="color:#64748b;">Gap <b style="color:${month.gap > 0.5 ? '#dc2626' : '#111827'};">$${formatCurrency(month.gap)}</b>${month.overdueGap > 0.5 ? ` <span style="color:#dc2626;">· overdue $${formatCurrency(month.overdueGap)}</span>` : ''}</span>
                </div>
                <div style="overflow-x:auto;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.72rem;">
                        <thead>
                            <tr style="text-align:left; border-bottom:1px solid #e2e8f0; color:#64748b;">
                                <th style="padding:6px 8px; font-weight:700;">End User / Deal</th>
                                <th style="padding:6px 8px; font-weight:700;">Distributor</th>
                                <th style="padding:6px 8px; font-weight:700;">Due</th>
                                <th style="padding:6px 8px; font-weight:700;">Status</th>
                                <th style="padding:6px 8px; font-weight:700; text-align:right;">Due amt</th>
                                <th style="padding:6px 8px; font-weight:700; text-align:right;">Paid</th>
                                <th style="padding:6px 8px; font-weight:700; text-align:right;">Balance</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${sorted.map(r => `
                            <tr data-collection-deal="${encodeURIComponent(r.deal)}" title="${esc(r.deal)} — click for contract detail"
                                style="border-bottom:1px solid #f1f5f9; cursor:pointer;"
                                onmouseover="this.style.background='#eef2ff'" onmouseout="this.style.background=''">
                                <td style="padding:6px 8px; font-weight:600; color:#1e293b;">${esc(r.endUser)}<div style="font-size:0.6rem; color:#94a3b8;">${esc(r.deal)}${r.installmentNo ? ` · #${esc(r.installmentNo)}` : ''}</div></td>
                                <td style="padding:6px 8px; color:#475569;">${esc(r.distributor)}</td>
                                <td style="padding:6px 8px; font-family:monospace; white-space:nowrap; color:#64748b;">${esc(r.dueStr)}</td>
                                <td style="padding:6px 8px;">${statusPill(r)}</td>
                                <td style="padding:6px 8px; text-align:right; font-weight:700; white-space:nowrap;">$${formatCurrency(r.amountDue)}</td>
                                <td style="padding:6px 8px; text-align:right; color:#7c3aed; font-weight:700; white-space:nowrap;">${r.amountPaid > 0 ? '$' + formatCurrency(r.amountPaid) : '—'}${r.paidDateStr ? `<div style="font-size:0.58rem; color:#94a3b8; font-weight:400;">${esc(r.paidDateStr)}</div>` : ''}</td>
                                <td style="padding:6px 8px; text-align:right; font-weight:800; white-space:nowrap; color:${r.balance > 0.5 ? '#dc2626' : '#16a34a'};">$${formatCurrency(r.balance)}</td>
                            </tr>`).join('')}
                        </tbody>
                    </table>
                </div>
            </div>`;
        host.querySelector('[data-expand-close]').addEventListener('click', () => {
            host.style.display = 'none';
            host.innerHTML = '';
        });
        host.querySelectorAll('[data-collection-deal]').forEach(el => {
            el.addEventListener('click', () => window.showCollectionDealDetail(el.getAttribute('data-collection-deal')));
        });
        host.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    const barCursor = (evt, els) => {
        const t = evt.native ? evt.native.target : evt.target;
        if (t) t.style.cursor = els.length ? 'pointer' : 'default';
    };
    const moneyTicks = { callback: (v) => '$' + formatCurrency(v) };

    const initActiveCharts = () => {
        const act = stats.active;

        const agingCtx = document.getElementById('collection-aging-chart');
        if (agingCtx) {
            // Stacked by distributor so each bucket shows WHO owes the money.
            // Fixed hue order follows total outstanding (act.distributors is
            // already sorted desc); anything past the palette folds into "Other".
            const AGING_SERIES_COLORS = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300'];
            const AGING_OTHER_COLOR = '#94a3b8';
            const agingBarBase = { stack: 'aging', borderRadius: 3, barPercentage: 0.6, borderColor: '#FFFFFF', borderWidth: 1 };
            const topDistNames = act.distributors.slice(0, AGING_SERIES_COLORS.length).map(d => d.name);
            const topDistSet = new Set(topDistNames);
            const agingDatasets = topDistNames.map((name, i) => ({
                ...agingBarBase,
                label: name,
                data: act.agingBuckets.map(b => (b.byDistributor || {})[name] || 0),
                backgroundColor: AGING_SERIES_COLORS[i]
            }));
            const agingOther = act.agingBuckets.map(b =>
                Object.entries(b.byDistributor || {}).reduce((s, [name, amt]) => s + (topDistSet.has(name) ? 0 : amt), 0));
            if (agingOther.some(v => Math.abs(v) > 0.5)) {
                agingDatasets.push({ ...agingBarBase, label: 'Other', data: agingOther, backgroundColor: AGING_OTHER_COLOR });
            }
            if (agingDatasets.length === 0) {
                agingDatasets.push({ ...agingBarBase, label: 'Outstanding', data: act.agingBuckets.map(b => b.amount), backgroundColor: 'rgba(148,163,184,0.85)' });
            }
            window.collectionAgingChart = new Chart(agingCtx, {
                type: 'bar',
                data: {
                    labels: act.agingBuckets.map(b => b.label),
                    datasets: agingDatasets
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (evt, els) => {
                        if (!els.length) return;
                        const b = act.agingBuckets[els[0].index];
                        renderExpand('collection-aging-expand', b.label, b.byYear);
                    },
                    onHover: barCursor,
                    interaction: { mode: 'index', intersect: false },
                    plugins: {
                        legend: {
                            display: topDistNames.length > 0,
                            position: 'bottom',
                            labels: { boxWidth: 10, boxHeight: 10, font: { size: 10 }, color: '#64748b' }
                        },
                        tooltip: {
                            filter: (item) => Math.abs(item.raw) > 0.005,
                            callbacks: {
                                label: (item) => ` ${item.dataset.label}: $${formatCurrency(item.raw)}`,
                                footer: (items) => {
                                    const b = act.agingBuckets[items[0].dataIndex];
                                    return [
                                        `Total $${formatCurrency(b.amount)} · ${b.dealCount} contract${b.dealCount === 1 ? '' : 's'}`,
                                        'Click for year-by-year breakdown'
                                    ];
                                }
                            }
                        }
                    },
                    scales: {
                        y: { stacked: true, beginAtZero: true, grid: { color: '#F3F4F6' }, ticks: moneyTicks },
                        x: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 } } }
                    }
                }
            });
        }

        const distCtx = document.getElementById('collection-distributor-chart');
        if (distCtx && act.distributors.length > 0) {
            window.collectionDistChart = new Chart(distCtx, {
                type: 'bar',
                data: {
                    labels: act.distributors.map(d => d.name),
                    datasets: [{
                        data: act.distributors.map(d => d.amount),
                        backgroundColor: 'rgba(99,102,241,0.85)',
                        borderRadius: 6,
                        barPercentage: 0.6
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (evt, els) => {
                        if (!els.length) return;
                        const d = act.distributors[els[0].index];
                        renderExpand('collection-dist-expand', d.name, d.byYear);
                    },
                    onHover: barCursor,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: (item) => {
                                    const d = act.distributors[item.dataIndex];
                                    return ` $${formatCurrency(d.amount)} · ${d.dealCount} contract${d.dealCount === 1 ? '' : 's'}`;
                                },
                                footer: () => 'Click for year-by-year breakdown'
                            }
                        }
                    },
                    scales: {
                        x: { beginAtZero: true, grid: { color: '#F3F4F6' }, ticks: moneyTicks },
                        y: { grid: { display: false }, ticks: { font: { size: 10 } } }
                    }
                }
            });
        }

        const trendCtx = document.getElementById('collection-trend-chart');
        if (trendCtx && stats.trend.length > 0) {
            window.collectionTrendChart = new Chart(trendCtx, {
                type: 'bar',
                data: {
                    labels: stats.trend.map(t => t.label),
                    datasets: [
                        {
                            label: 'Collected',
                            data: stats.trend.map(t => t.collected),
                            backgroundColor: 'rgba(139,92,246,0.85)',
                            borderRadius: 3,
                            barPercentage: 0.7
                        },
                        {
                            label: 'Overdue (uncollected)',
                            data: stats.trend.map(t => t.overdueGap),
                            backgroundColor: 'rgba(239,68,68,0.8)',
                            borderRadius: 3,
                            barPercentage: 0.7
                        },
                        {
                            label: 'Not yet due',
                            data: stats.trend.map(t => t.upcomingGap),
                            backgroundColor: 'rgba(203,213,225,0.9)',
                            borderRadius: 3,
                            barPercentage: 0.7
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (evt, els) => {
                        if (!els.length) return;
                        renderTrendExpand(stats.trend[els[0].index]);
                    },
                    onHover: barCursor,
                    interaction: { mode: 'index', intersect: false },
                    plugins: {
                        legend: {
                            display: true,
                            position: 'bottom',
                            labels: { boxWidth: 10, boxHeight: 10, font: { size: 10 }, color: '#64748b' }
                        },
                        tooltip: {
                            callbacks: {
                                label: (item) => ` ${item.dataset.label}: $${formatCurrency(item.raw)}`,
                                footer: (items) => {
                                    const t = stats.trend[items[0].dataIndex];
                                    const pct = t.expected > 0 ? Math.round((t.collected / t.expected) * 100) : 0;
                                    return [
                                        `Expected $${formatCurrency(t.expected)} · collected ${pct}%`,
                                        `Gap $${formatCurrency(t.gap)}`,
                                        'Click for installment detail'
                                    ];
                                }
                            }
                        }
                    },
                    scales: {
                        y: { stacked: true, beginAtZero: true, grid: { color: '#F3F4F6' }, ticks: moneyTicks },
                        x: { stacked: true, grid: { display: false }, ticks: { font: { size: 9 }, maxRotation: 60, minRotation: 40 } }
                    }
                }
            });
        }

        const tlCtx = document.getElementById('collection-timeliness-chart');
        if (tlCtx && stats.timeliness.length > 0) {
            window.collectionTimelinessChart = new Chart(tlCtx, {
                type: 'bar',
                data: {
                    labels: stats.timeliness.map(t => t.label),
                    datasets: [{
                        data: stats.timeliness.map(t => t.pct),
                        backgroundColor: 'rgba(16,185,129,0.85)',
                        borderRadius: 6,
                        barPercentage: 0.5
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (evt, els) => {
                        if (!els.length) return;
                        const t = stats.timeliness[els[0].index];
                        window.showCollectionKpiDetail('timeliness:' + t.year);
                    },
                    onHover: barCursor,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: (item) => {
                                    const t = stats.timeliness[item.dataIndex];
                                    return ` ${t.pct}% on-time · ${t.onTime}/${t.paid} payments`;
                                },
                                footer: () => 'Click for scheduled vs received detail'
                            }
                        }
                    },
                    scales: {
                        y: { beginAtZero: true, max: 100, grid: { color: '#F3F4F6' }, ticks: { callback: (v) => v + '%' } },
                        x: { grid: { display: false } }
                    }
                }
            });
        }
    };

    const bindDealRows = () => {
        container.querySelectorAll('[data-collection-deal]').forEach(el => {
            el.addEventListener('click', () => window.showCollectionDealDetail(el.getAttribute('data-collection-deal')));
        });
        container.querySelectorAll('[data-collection-kpi]').forEach(el => {
            el.addEventListener('click', () => window.showCollectionKpiDetail(el.getAttribute('data-collection-kpi')));
        });
    };

    const renderTabBar = () => {
        const tabs = [
            { key: 'active', label: 'Active Collections', badge: stats.active.contractCount },
            { key: 'pending', label: 'Pending Records', badge: stats.pending.recordCount }
        ];
        const v = window.collectionView;
        tabBar.innerHTML = `
            <div style="display:flex; gap:4px; border-bottom:2px solid #e5e7eb;">
                ${tabs.map(t => `
                    <button data-collection-view="${t.key}" style="
                        padding:10px 22px;
                        background:${v === t.key ? '#FFF' : 'transparent'};
                        border:1px solid ${v === t.key ? '#e5e7eb' : 'transparent'};
                        border-bottom:3px solid ${v === t.key ? '#6366f1' : 'transparent'};
                        border-radius:8px 8px 0 0;
                        color:${v === t.key ? '#4338ca' : '#6b7280'};
                        font-weight:${v === t.key ? '700' : '600'};
                        font-size:0.85rem;
                        cursor:pointer;
                        margin-bottom:-2px;
                    ">${t.label} <span style="background:${v === t.key ? '#eef2ff' : '#f3f4f6'}; color:${v === t.key ? '#4338ca' : '#6b7280'}; font-size:0.68rem; font-weight:800; padding:1px 7px; border-radius:999px; margin-left:4px;">${t.badge}</span></button>
                `).join('')}
            </div>
        `;
        tabBar.querySelectorAll('[data-collection-view]').forEach(btn => {
            btn.addEventListener('click', () => {
                window.collectionView = btn.dataset.collectionView;
                renderTabBar();
                renderView();
            });
        });
    };

    const renderView = () => {
        destroyCharts();
        container.innerHTML = getCollectionHTML(stats, window.collectionView);
        bindDealRows();
        if (window.collectionView === 'active') {
            initActiveCharts();
            _renderCollectionChangeLog(stats, filterCountry, container);
        }
    };

    renderTabBar();
    renderView();
}

/**
 * Snapshot the Collection sheet (totals + row-level) into localStorage and
 * render a change log mirroring the Pipeline pattern. Tracks new rows,
 * removed rows, and modified rows (Total Collected, Outstanding, Payment
 * Status, Next Due Date) with timestamps.
 *
 * @param {Object} stats - Output of getCollectionStats
 * @param {string|null} filterCountry
 * @param {HTMLElement} container - Where to append the change-log card
 */
function _renderCollectionChangeLog(stats, filterCountry, container) {
    const scope = filterCountry || 'ALL';
    const storageKey = `collectionChangeLog::${scope}`;

    const rowsByKey = Object.values(stats.deals).map(d => ({
        key: d.deal,
        distributor: d.distributor,
        endUser: d.endUser,
        collected: Math.round(d.totalPaid),
        outstanding: Math.round(d.totalBalance),
        status: d.aggStatus,
        nextDue: d.nextDueStr || ''
    }));
    const seen = new Map();
    const dedupedRows = rowsByKey.map(r => {
        const dup = seen.get(r.key) || 0;
        seen.set(r.key, dup + 1);
        return dup === 0 ? r : { ...r, key: `${r.key}#${dup + 1}` };
    });

    const rowsFp = dedupedRows
        .map(r => `${r.key}|${r.collected}|${r.outstanding}|${r.status}|${r.nextDue}`)
        .sort()
        .join('||');

    const current = {
        date: new Date().toISOString(),
        totalCollected: Math.round(stats.active.totalPaid),
        totalOutstanding: Math.round(stats.active.totalBalance),
        ktcvNet: Math.round(stats.active.totalDue),
        byStatusAmount: { ...stats.byStatusAmount },
        byStatusCount: { ...stats.byStatusCount },
        rows: dedupedRows,
        rowsFp
    };

    let history = [];
    try {
        const raw = localStorage.getItem(storageKey);
        if (raw) history = JSON.parse(raw);
        if (!Array.isArray(history)) history = [];
    } catch { history = []; }

    const last = history[history.length - 1];
    const changed = !last
        || last.totalCollected !== current.totalCollected
        || last.totalOutstanding !== current.totalOutstanding
        || last.ktcvNet !== current.ktcvNet
        || (last.rowsFp || '') !== current.rowsFp;

    if (changed) {
        history.push(current);
        if (history.length > 100) history = history.slice(-100);
        try { localStorage.setItem(storageKey, JSON.stringify(history)); }
        catch (e) { console.warn('[CollectionChangeLog] localStorage write failed:', e); }
    }

    const sorted = [...history].sort((a, b) => new Date(a.date) - new Date(b.date));
    const rowDiffs = sorted.map((snap, i) => {
        if (i === 0) return null;
        const before = sorted[i - 1].rows || [];
        const after = snap.rows || [];
        const beforeMap = new Map(before.map(r => [r.key, r]));
        const afterMap = new Map(after.map(r => [r.key, r]));
        const added = [];
        const removed = [];
        const modified = [];
        afterMap.forEach((a, k) => { if (!beforeMap.has(k)) added.push(a); });
        beforeMap.forEach((b, k) => { if (!afterMap.has(k)) removed.push(b); });
        afterMap.forEach((a, k) => {
            const b = beforeMap.get(k);
            if (!b) return;
            if (b.collected !== a.collected || b.outstanding !== a.outstanding || b.status !== a.status || b.nextDue !== a.nextDue) {
                modified.push({ before: b, after: a });
            }
        });
        return { added, removed, modified };
    });

    const logHost = document.createElement('div');
    logHost.style.gridColumn = '1 / -1';
    logHost.style.marginTop = '12px';
    logHost.innerHTML = getCollectionChangeLogHTML(scope, history, rowDiffs);
    container.appendChild(logHost);

    setTimeout(() => {
        const resetBtn = document.getElementById('collection-changelog-reset');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                if (!confirm(`Clear all collection change history for ${scope}? This cannot be undone.`)) return;
                try { localStorage.removeItem(storageKey); } catch {}
                logHost.innerHTML = getCollectionChangeLogHTML(scope, [], []);
            });
        }
    }, 80);
}

/* ═══════════════════════════════════════════════════════════════
   TCV vs ARR
   ═══════════════════════════════════════════════════════════════ */

/**
 * Render the TCV vs ARR Revenue Mix dashboard.
 * Uses internal filter state (window.tcvArrFilters) for Country and Contract Yr.
 * @param {Object} workbookData
 * @param {HTMLElement} metricsGrid
 */
function _renderTcvArr(workbookData, metricsGrid) {
    const orderData = workbookData['ORDER SHEET'] || [];
    if (orderData.length === 0) return;

    window.tcvArrFilters = window.tcvArrFilters || { country: 'All', contractYr: 'All' };

    const container = document.createElement('div');
    container.id = 'tcvarr-dashboard-container';
    container.style.gridColumn = '1 / -1';
    metricsGrid.appendChild(container);

    /**
     * Re-render the TCV vs ARR view with current filter state.
     */
    function updateTcvArrUI() {
        const stats = getTcvArrStats(orderData, window.tcvArrFilters);
        const el = document.getElementById('tcvarr-dashboard-container');
        if (!el) return;

        el.innerHTML = getTcvArrHTML(stats, window.tcvArrFilters);

        /* Bind filter change handlers */
        const countrySelect = document.getElementById('tcvarr-filter-country');
        const yearSelect = document.getElementById('tcvarr-filter-year');

        if (countrySelect) {
            countrySelect.addEventListener('change', (e) => {
                window.tcvArrFilters.country = e.target.value;
                updateTcvArrUI();
            });
        }
        if (yearSelect) {
            yearSelect.addEventListener('change', (e) => {
                window.tcvArrFilters.contractYr = e.target.value;
                updateTcvArrUI();
            });
        }

        setTimeout(() => initTcvArrChart(stats), 80);
    }

    updateTcvArrUI();
}
