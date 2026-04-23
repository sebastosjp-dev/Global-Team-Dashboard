/**
 * views.js — Render orchestrators connecting Stats → HTML → Charts.
 * Extracted from app.js for single-responsibility.
 * @module views
 */
import { CONFIG } from './config.js';
import { parseCurrency, formatCurrency, isCountryMatch } from './utils.js';
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
    getTcvArrStats
} from './services.js';
import {
    getOrderSheetHTML, getPipelineHTML, getPartnerHTML,
    getGenericCountryHTML, getExpiringContractsHTML,
    getPartnerPerformanceHTML, getPocHTML, getEventHTML,
    getCountrySpecificHTML, getServiceAnalysisHTML,
    getRenewalHTML, getKPIHTML, getCollectionHTML,
    getTcvArrHTML
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
        _renderExpiringContracts(data, metricsGrid);
        hasMetrics = true;
    }

    if (tabName === 'PIPELINE' && workbookData['PIPELINE']) {
        _renderPipeline(workbookData, filterCountry, tabName, metricsGrid, searchInput);
        hasMetrics = true;
    }

    if ((tabName === 'PARTNER' || isCountryTab) && data && data.length > 0) {
        _renderPartner(data, filterCountry, tabName, metricsGrid, workbookData, searchInput);
        _renderGenericCountry(data, filterCountry, metricsGrid, tabName);
        _renderPartnerTopPerformer(data, metricsGrid);
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
        const renewalFiltered = _getRenewalTableData(workbookData, filterCountry);
        if (renewalFiltered && renewalFiltered.length > 0) {
            const renewalContainer = document.createElement('div');
            renewalContainer.style.gridColumn = '1 / -1';
            renewalContainer.innerHTML = getRenewalHTML(renewalFiltered);
            metricsGrid.appendChild(renewalContainer);
        }

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



/**
 * @param {Object} workbookData
 * @returns {Object[]}
 */
function _getRenewalTableData(workbookData, filterCountry) {
    const data = workbookData['END USER (CSM)'] || [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const sixMonthsLater = new Date(today);
    sixMonthsLater.setMonth(today.getMonth() + 6);

    const getExpDate = (row) => {
        let endDate = row['End License Date'];
        if (!endDate) return null;
        if (endDate instanceof Date) return endDate;
        if (typeof endDate === 'number') return new Date(Math.round((endDate - 25569) * 86400 * 1000));
        return new Date(endDate);
    };

    return data.filter(row => {
        if (!isCountryMatch(row, filterCountry)) return false;
        const d = getExpDate(row);
        return d && d >= today && d <= sixMonthsLater;
    }).map(row => {
        const d = getExpDate(row);
        const diffDays = Math.ceil((d - today) / (1000 * 60 * 60 * 24));
        return {
            ...row,
            'D-Day': diffDays === 0 ? 'D-Day' : (diffDays > 0 ? `D-${diffDays}` : `D+${Math.abs(diffDays)}`),
            diffDays,
            endDateFormatted: d.toISOString().split('T')[0]
        };
    }).sort((a, b) => a.diffDays - b.diffDays || (parseCurrency(b['TCV Amount']) - parseCurrency(a['TCV Amount'])));
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
    container.innerHTML = getOrderSheetHTML(stats);
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
        initPartnerCharts(stats, filterCountry);
    }, 100);
}

function _renderGenericCountry(data, filterCountry, metricsGrid, tabName) {
    if (data.length === 0 || tabName === 'EVENT') return;
    const stats = getGenericCountryStats(data, filterCountry);
    if (!stats) return;
    const div = document.createElement('div');
    div.innerHTML = getGenericCountryHTML(stats, filterCountry);
    metricsGrid.appendChild(div.firstElementChild);
}

function _renderExpiringContracts(data, metricsGrid) {
    const stats = getExpiringContractsStats(data);
    if (!stats) return;
    const container = document.createElement('div');
    container.style.gridColumn = '1 / -1';
    container.innerHTML = getExpiringContractsHTML(stats);
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
