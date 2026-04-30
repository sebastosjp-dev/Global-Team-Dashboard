/**
 * views.js — Render orchestrators connecting Stats → HTML → Charts.
 * Extracted from app.js for single-responsibility.
 * @module views
 */
import { CONFIG } from './config.js';
import { parseCurrency, formatCurrency, isCountryMatch, findKey } from './utils.js';
import {
    chartRegistry, initOrderSheetCharts, initPipelineCharts,
    initPartnerCharts, initPartnerPerformanceCharts,
    initPocCharts, initEventCharts, initServiceAnalysisCharts,
    initTcvArrChart
} from './charts.js';
import {
    getOrderSheetStats, getPipelineStats, getPartnerStats,
    getGenericCountryStats, getExpiringContractsStats,
    getPartnerPerformanceStats, getPocStats, getEventStats,
    getCountrySpecificStats, getServiceAnalysisStats,
    getCollectionStats, getDetailedCollectionAnalysis,
    getTcvArrStats, getChurnRiskStats,
    getPartnerROIStats, getPipelineCoverageStats
} from './services.js';
import {
                                                getOrderSheetHTML, getPipelineHTML, getPartnerHTML, getPartnerNetworkDetailsHTML,
    getGenericCountryHTML, getExpiringContractsHTML,
    getPartnerPerformanceHTML, getPocHTML, getEventHTML,
    getCountrySpecificHTML, getServiceAnalysisHTML,
    getKPIHTML, getCollectionHTML,
    getTcvArrHTML, getChurnRiskHTML,
    getPartnerROIHTML, getPipelineCoverageHTML,
    getPipelineChangeLogHTML, getCurrentPipelineListHTML
} from './ui.js';

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
        !['ORDER SHEET', 'PIPELINE', 'PARTNER', 'POC', 'EVENT', 'CSM', 'END USER (CSM)', 'COLLECTION'].includes(tabName) &&
        !isGlobalTab;

    if (tabName === 'ORDER SHEET' || isGlobalTab) {
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
        _renderServiceAnalysis(workbookData['END USER (CSM)'], filterCountry, tabName, metricsGrid, searchInput);
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

    if (hasMetrics) metricsGrid.classList.remove('hidden');
    else metricsGrid.classList.add('hidden');
}




/* ═══════════════════════════════════════════════════════════════
   Service Analysis Metrics
   ═══════════════════════════════════════════════════════════════ */

function _renderServiceAnalysis(data, filterCountry, tabName, metricsGrid, searchInput) {
    if (!data || data.length === 0) return;
    const stats = getServiceAnalysisStats(data, filterCountry);

    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getServiceAnalysisHTML(stats, filterCountry);

    metricsGrid.appendChild(container);
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
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getOrderSheetHTML(stats, filterCountry);
    metricsGrid.appendChild(container);
    setTimeout(() => initOrderSheetCharts(stats), 120);
}

function _renderPipeline(workbookData, filterCountry, tabName, metricsGrid, searchInput) {
    const pData = filterCountry
        ? workbookData['PIPELINE'].filter(r => isCountryMatch(r, filterCountry))
        : workbookData['PIPELINE'];
    if (!pData || pData.length === 0) return;

    const oDataRaw = workbookData['ORDER SHEET'] || [];
    const oData = filterCountry ? oDataRaw.filter(r => isCountryMatch(r, filterCountry)) : oDataRaw;

    const stats = getPipelineStats(pData, oData);
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.style.marginTop = '12px';
    container.style.marginBottom = '24px';
    container.innerHTML = getPipelineHTML(stats, filterCountry, tabName);
    metricsGrid.appendChild(container);
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
        initPipelineCharts(stats);
    }, 100);
}

function _renderPartner(data, filterCountry, tabName, metricsGrid, workbookData, searchInput) {
    if (!data || data.length === 0) return;
    const stats = getPartnerStats(data, filterCountry, workbookData);
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
    }, 100);
}

function _renderPartnerNetworkDetails(data, filterCountry, metricsGrid, workbookData) {
    if (!data || data.length === 0) return;
    const stats = getPartnerStats(data, filterCountry, workbookData);
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
        let quarter = '';
        if (qKey && row[qKey]) {
            const qRaw = String(row[qKey]).toUpperCase().trim();
            if (qRaw.includes('Q1')) quarter = 'Q1';
            else if (qRaw.includes('Q2')) quarter = 'Q2';
            else if (qRaw.includes('Q3')) quarter = 'Q3';
            else if (qRaw.includes('Q4')) quarter = 'Q4';
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

function _renderCollection(data, filterCountry, metricsGrid) {
    const stats = getCollectionStats(data);

    // 필터 상태 유지 (세션)
    if (window.collectionShowUnpaid === undefined) window.collectionShowUnpaid = false;

    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    metricsGrid.appendChild(container);

    const updateUI = () => {
        const detailedStats = getDetailedCollectionAnalysis(data);
        const displayStats = {
            ...detailedStats,
            rows: window.collectionShowUnpaid ? detailedStats.rows.filter(r => r.balance > 0) : detailedStats.rows
        };

        container.innerHTML = getCollectionHTML(stats, displayStats, window.collectionShowUnpaid);

        const ctx = document.getElementById('collection-performance-chart');
        if (!ctx) return;

        const years = Object.keys(stats.yearlyStats).sort();
        const targets = years.map(y => stats.yearlyStats[y].target);
        const actuals = years.map(y => stats.yearlyStats[y].actual);

        if (window.collectionChart) window.collectionChart.destroy();
        window.collectionChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: years,
                datasets: [
                    { label: 'Collection Target (ARR)', data: targets, backgroundColor: 'rgba(99, 102, 241, 0.6)', borderRadius: 4, barPercentage: 0.6 },
                    { label: 'Actual Performance (Received)', data: actuals, backgroundColor: 'rgba(16, 185, 129, 0.8)', borderRadius: 4, barPercentage: 0.6 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                onClick: (event, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const year = years[index];
                        const detailData = stats.yearlyDistributorTargets[year] || {};
                        window.renderCollectionYearDetail(year, detailData);
                    }
                },
                plugins: {
                    legend: { position: 'top', labels: { usePointStyle: true, padding: 20 } },
                    title: { display: true, text: 'Click bar to see distributor breakdown', font: { size: 11, weight: 'normal' }, color: '#94a3b8' },
                    tooltip: { callbacks: { label: (ctx) => ` $${formatCurrency(ctx.raw)}` } }
                },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#F3F4F6' }, ticks: { callback: (v) => '$' + formatCurrency(v) } },
                    x: { grid: { display: false } }
                }
            }
        });

        // 필터 토글 이벤트 바인딩
        const toggleBtn = document.getElementById('collection-unpaid-toggle');
        if (toggleBtn) {
            toggleBtn.onclick = () => {
                window.collectionShowUnpaid = !window.collectionShowUnpaid;
                updateUI();
            };
        }
    };

    updateUI();
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
