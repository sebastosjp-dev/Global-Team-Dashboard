/**
 * charts.js — Chart registry and all chart initialization functions
 */
import { formatCurrency } from './utils.js';

/** Chart instance registry — ensures proper cleanup on re-render */
export const chartRegistry = {
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

/* ═══ ORDER SHEET Charts ═══ */
export function initOrderSheetCharts(stats) {
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
                    fill: true, tension: 0.4, pointRadius: 4
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

/* ═══ PIPELINE Charts ═══ */
export function initPipelineCharts(stats) {
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
                        borderColor: '#34C759', borderWidth: 1, borderRadius: 4, yAxisID: 'yCount'
                    },
                    {
                        label: 'Pipeline Value (USD)',
                        data: stats.pipelineInfluxData.map(d => d.value),
                        type: 'line',
                        borderColor: '#f59e0b', backgroundColor: 'rgba(245, 158, 11, 0.1)',
                        borderWidth: 3, tension: 0.4, fill: true, pointRadius: 4, yAxisID: 'yValue'
                    }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                scales: {
                    yCount: { type: 'linear', position: 'left', grid: { color: 'rgba(0,0,0,0.06)' }, ticks: { color: '#6B7280', beginAtZero: true, stepSize: 1 }, title: { display: true, text: 'Deal Count', color: '#94a3b8', font: { size: 10 } } },
                    yValue: { type: 'linear', position: 'right', grid: { display: false }, ticks: { color: '#f59e0b', callback: (v) => formatCurrency(v) }, title: { display: true, text: 'Value (USD)', color: '#f59e0b', font: { size: 10 } } },
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
                                    accs.slice(0, 10).forEach(a => lines.push(`· ${a}`));
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

/* ═══ PARTNER Charts ═══ */
export function initPartnerCharts(stats, filterCountry) {
    if (!filterCountry) {
        const ctx = document.getElementById('partner-country-chart');
        if (ctx) {
            chartRegistry.destroyTag('partner-country');
            const chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: stats.sortedCountries,
                    datasets: [{
                        label: 'Partner Count',
                        data: stats.sortedCountries.map(c => stats.partnerGroups[c].length),
                        backgroundColor: 'rgba(0,122,255,0.5)', borderColor: '#007AFF',
                        borderWidth: 2, borderRadius: 10, hoverBackgroundColor: 'rgba(0,122,255,0.85)',
                        barThickness: 'flex', maxBarThickness: 60
                    }]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: { backgroundColor: '#FFFFFF', titleColor: '#111827', bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1, padding: 14, titleFont: { size: 14, weight: 'bold' }, bodyFont: { size: 13 }, cornerRoundness: 8 }
                    },
                    scales: {
                        y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)', drawBorder: false }, ticks: { color: '#6B7280', font: { size: 11 }, stepSize: 1 } },
                        x: { grid: { display: false }, ticks: { color: '#111827', font: { size: 12, weight: '600' } } }
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

/* ═══ PARTNER PERFORMANCE Charts ═══ */
export function initPartnerPerformanceCharts(stats) {
    const ctx = document.getElementById('partner-top-performer-chart');
    if (ctx) {
        chartRegistry.destroyTag('partner-perf');
        chartRegistry.register('partner-perf-summary', new Chart(ctx, {
            type: 'bar',
            data: {
                labels: stats.map(p => `${p.name} (${p.country})`),
                datasets: [{ label: 'TCV (USD)', data: stats.map(p => p.tcv), backgroundColor: 'rgba(0,122,255,0.7)', borderRadius: 4 }]
            },
            options: {
                indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { backgroundColor: '#FFFFFF', titleColor: '#111827', bodyColor: '#374151', borderColor: '#E5E7EB', borderWidth: 1, padding: 12, callbacks: { label: (ctx) => ' US$ ' + formatCurrency(ctx.parsed.x) } }
                },
                scales: {
                    x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => '$' + formatCurrency(v) } },
                    y: { grid: { display: false }, ticks: { color: '#111827', font: { weight: '500' } } }
                }
            }
        }));
    }
}

/* ═══ POC Charts ═══ */
export function initPocCharts(stats) {
    chartRegistry.destroyTag('poc');

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

    const ctxAging = document.getElementById('poc-aging-chart');
    if (ctxAging) {
        chartRegistry.register('poc-aging', new Chart(ctxAging, {
            type: 'doughnut',
            data: { labels: ['100+', '60-100', '<60'], datasets: [{ data: [stats.longTermCount, stats.midTermCount, stats.normalCount], backgroundColor: ['#FF3B30', '#FF9500', '#34C759'] }] },
            options: { responsive: true, maintainAspectRatio: false, cutout: '65%', plugins: { legend: { position: 'bottom' } } }
        }));
    }

    const ctxB = document.getElementById('poc-bottleneck-chart');
    if (ctxB && stats.partnerAvg.length > 0) {
        chartRegistry.register('poc-bottleneck', new Chart(ctxB, {
            type: 'bar',
            data: { labels: stats.partnerAvg.slice(0, 10).map(p => p.partner), datasets: [{ data: stats.partnerAvg.slice(0, 10).map(p => p.avg), backgroundColor: '#007AFF' }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        }));
    }

    const ctxI = document.getElementById('poc-industry-chart');
    if (ctxI && stats.sortedIndustry.length > 0) {
        chartRegistry.register('poc-industry', new Chart(ctxI, {
            type: 'bar',
            data: { labels: stats.sortedIndustry.slice(0, 10).map(i => i.name), datasets: [{ data: stats.sortedIndustry.slice(0, 10).map(i => i.val), backgroundColor: '#a78bfa' }] },
            options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
        }));
    }
}

/* ═══ EVENT Charts ═══ */
export function initEventCharts(stats) {
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

/* ═══ SERVICE ANALYSIS Charts ═══ */
export function initServiceAnalysisCharts(stats) {
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