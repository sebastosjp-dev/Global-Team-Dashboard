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
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#1e293b',
                        titleColor: '#f59e0b',
                        titleFont: { weight: 'bold' },
                        padding: 12,
                        cornerRadius: 8,
                        callbacks: {
                            label: function (context) {
                                let label = context.dataset.label || '';
                                if (label) label += ': ';
                                label += '$' + formatCurrency(context.parsed.y);
                                return label;
                            },
                            afterBody: function (context) {
                                const q = context[0].label;
                                const deals = stats.qDeals[q] || [];
                                if (deals.length === 0) return '';
                                // Sort by TCV descending and take top 10
                                const sorted = [...deals].sort((a, b) => b.tcv - a.tcv).slice(0, 10);
                                const lines = ['', '--- Top Deals ---'];
                                sorted.forEach(d => {
                                    lines.push(`· ${d.name} ($${formatCurrency(d.tcv)})`);
                                });
                                if (deals.length > 10) lines.push(`... (+${deals.length - 10} more)`);
                                return lines;
                            }
                        }
                    }
                },
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

    // Sparklines for ARR and MRR Growth
    const sparklineOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { 
            legend: { display: false }, 
            tooltip: { enabled: true, callbacks: { label: (ctx) => ` US$ ${formatCurrency(ctx.parsed.y)}` } } 
        },
        scales: { 
            x: { 
                display: true, 
                grid: { display: false }, 
                border: { display: false },
                ticks: {
                    color: '#94a3b8',
                    font: { size: 9, weight: '700' },
                    padding: 4
                }
            }, 
            y: { display: false } 
        },
        elements: { 
            point: { radius: 2, hoverRadius: 5 }, 
            line: { borderWidth: 2, tension: 0.4 } 
        },
        layout: {
            padding: { top: 5, bottom: 0, left: 5, right: 5 }
        }
    };

    const arrYears = Object.keys(stats.yearlyArr).sort();
    const arrCtx = document.getElementById('arr-sparkline');
    if (arrCtx && arrYears.length > 0) {
        chartRegistry.register('order-arr-spark', new Chart(arrCtx, {
            type: 'line',
            data: {
                labels: arrYears,
                datasets: [{ 
                    data: arrYears.map(y => stats.yearlyArr[y]),
                    borderColor: '#8b5cf6',
                    backgroundColor: 'rgba(139, 92, 246, 0.1)',
                    fill: true
                }]
            },
            options: sparklineOptions
        }));
    }

    const mrrYears = Object.keys(stats.yearlyMrr).sort();
    const mrrCtx = document.getElementById('mrr-sparkline');
    if (mrrCtx && mrrYears.length > 0) {
        chartRegistry.register('order-mrr-spark', new Chart(mrrCtx, {
            type: 'line',
            data: {
                labels: mrrYears,
                datasets: [{
                    data: mrrYears.map(y => stats.yearlyMrr[y]),
                    borderColor: '#a855f7',
                    backgroundColor: 'rgba(168, 85, 247, 0.1)',
                    fill: true
                }]
            },
            options: sparklineOptions
        }));
    }

    // Country TCV donut chart
    const donutCtx = document.getElementById('country-tcv-donut');
    if (donutCtx && stats.tcvByCountry) {
        const COUNTRY_COLORS = {
            'Indonesia': '#f59e0b', 'Thailand': '#3b82f6', 'Malaysia': '#10b981',
            'USA': '#6366f1', 'Philippines': '#ef4444', 'Singapore': '#ec4899',
            'Vietnam': '#14b8a6', 'Turkey': '#f97316', 'Other': '#94a3b8'
        };
        const FALLBACK = ['#6366f1','#10b981','#f59e0b','#ef4444','#3b82f6','#8b5cf6','#ec4899','#14b8a6'];

        const sorted = Object.entries(stats.tcvByCountry)
            .filter(([, v]) => v > 0)
            .sort((a, b) => b[1] - a[1]);

        const labels = sorted.map(([k]) => k);
        const values = sorted.map(([, v]) => v);
        const total = values.reduce((s, v) => s + v, 0);
        const colors = labels.map((l, i) => COUNTRY_COLORS[l] || FALLBACK[i % FALLBACK.length]);

        chartRegistry.register('order-country-donut', new Chart(donutCtx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{ data: values, backgroundColor: colors, borderWidth: 2, borderColor: '#fff', hoverBorderWidth: 3 }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '65%',
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#1e293b',
                        titleColor: '#f1f5f9',
                        bodyColor: '#94a3b8',
                        padding: 12,
                        cornerRadius: 8,
                        callbacks: {
                            label: (ctx) => {
                                const pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
                                return ` $${formatCurrency(ctx.parsed)}  (${pct}%)`;
                            }
                        }
                    }
                }
            }
        }));

        const legendEl = document.getElementById('country-tcv-legend');
        if (legendEl) {
            legendEl.innerHTML = `
                <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 6px;">
                    ${sorted.map(([country, tcv], i) => {
                        const pct = total > 0 ? ((tcv / total) * 100).toFixed(1) : 0;
                        const color = COUNTRY_COLORS[country] || FALLBACK[i % FALLBACK.length];
                        return `
                            <div style="display:flex; align-items:center; gap:8px; padding:7px 10px; border-radius:8px; background:#f9fafb; border-left:3px solid ${color};">
                                <div style="width:8px; height:8px; border-radius:50%; background:${color}; flex-shrink:0;"></div>
                                <span style="flex:1; font-size:0.75rem; font-weight:600; color:#374151;">${country}</span>
                                <span style="font-size:0.72rem; color:#6b7280;">$${formatCurrency(tcv)}</span>
                                <span style="font-size:0.72rem; font-weight:800; color:${color}; min-width:36px; text-align:right;">${pct}%</span>
                            </div>`;
                    }).join('')}
                </div>`;
        }
    }

    // YoY KTCV growth by country grouped bar chart
    const yoyCtx = document.getElementById('country-yoy-bar');
    if (yoyCtx && stats.tcvByCountryYear) {
        const COUNTRY_COLORS = {
            'Indonesia': '#f59e0b', 'Thailand': '#3b82f6', 'Malaysia': '#10b981',
            'USA': '#6366f1', 'Philippines': '#ef4444', 'Singapore': '#ec4899',
            'Vietnam': '#14b8a6', 'Turkey': '#f97316', 'Other': '#94a3b8'
        };
        const FALLBACK = ['#6366f1','#10b981','#f59e0b','#ef4444','#3b82f6','#8b5cf6','#ec4899','#14b8a6'];

        const thisYear = new Date().getFullYear();
        const lastYear = thisYear - 1;

        // Sort countries by this year's TCV descending
        const countries = Object.keys(stats.tcvByCountryYear).sort((a, b) => {
            const aVal = (stats.tcvByCountryYear[a][thisYear] || 0);
            const bVal = (stats.tcvByCountryYear[b][thisYear] || 0);
            return bVal - aVal;
        });

        const lastYearData = countries.map(c => stats.tcvByCountryYear[c][lastYear] || 0);
        const thisYearData = countries.map(c => stats.tcvByCountryYear[c][thisYear] || 0);

        chartRegistry.register('order-yoy-bar', new Chart(yoyCtx, {
            type: 'bar',
            data: {
                labels: countries,
                datasets: [
                    {
                        label: String(lastYear),
                        data: lastYearData,
                        backgroundColor: 'rgba(148, 163, 184, 0.5)',
                        borderColor: '#94a3b8',
                        borderWidth: 1,
                        borderRadius: 4
                    },
                    {
                        label: String(thisYear),
                        data: thisYearData,
                        backgroundColor: countries.map((c, i) => {
                            const col = COUNTRY_COLORS[c] || FALLBACK[i % FALLBACK.length];
                            return col + 'CC';
                        }),
                        borderColor: countries.map((c, i) => COUNTRY_COLORS[c] || FALLBACK[i % FALLBACK.length]),
                        borderWidth: 1,
                        borderRadius: 4
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: { font: { size: 11 }, boxWidth: 12 }
                    },
                    tooltip: {
                        backgroundColor: '#1e293b',
                        padding: 12,
                        cornerRadius: 8,
                        callbacks: {
                            label: (ctx) => {
                                const val = ctx.parsed.y;
                                const country = countries[ctx.dataIndex];
                                const ly = stats.tcvByCountryYear[country][lastYear] || 0;
                                const ty = stats.tcvByCountryYear[country][thisYear] || 0;
                                let line = ` ${ctx.dataset.label}: $${formatCurrency(val)}`;
                                if (ctx.datasetIndex === 1 && ly > 0) {
                                    const growth = ((ty - ly) / ly * 100).toFixed(1);
                                    const sign = growth >= 0 ? '+' : '';
                                    line += `  (${sign}${growth}% YoY)`;
                                } else if (ctx.datasetIndex === 1 && ly === 0 && ty > 0) {
                                    line += '  (New)';
                                }
                                return line;
                            }
                        }
                    }
                },
                scales: {
                    x: { grid: { display: false }, ticks: { font: { size: 11 } } },
                    y: {
                        grid: { color: 'rgba(0,0,0,0.05)' },
                        ticks: { callback: v => '$' + formatCurrency(v) }
                    }
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
                        label: 'Weighted Pipeline Value',
                        data: stats.pipelineInfluxData.map(d => d.weighted),
                        backgroundColor: 'rgba(52,199,89,0.55)',
                        borderColor: '#34C759', borderWidth: 1, borderRadius: 4, yAxisID: 'yValue'
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
                    yValue: { type: 'linear', position: 'left', grid: { color: 'rgba(0,0,0,0.06)' }, ticks: { color: '#6B7280', callback: (v) => formatCurrency(v) }, title: { display: true, text: 'Value (USD)', color: '#f59e0b', font: { size: 10 } } },
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
            },
            plugins: [{
                id: 'influxLabels',
                afterDatasetsDraw(chart) {
                    const { ctx, data } = chart;
                    ctx.save();
                    ctx.font = 'bold 9px "Inter", sans-serif';
                    ctx.fillStyle = '#ff3b30'; // Red
                    ctx.textAlign = 'center';
                    const metaTotal = chart.getDatasetMeta(1); 
                    metaTotal.data.forEach((point, i) => {
                        const valWeighted = data.datasets[0].data[i];
                        const valTotal = data.datasets[1].data[i];
                        const diff = valWeighted - valTotal;
                        if (valTotal > 0) {
                            const { x, y } = point.tooltipPosition();
                            const formatDiff = (v) => {
                                const absV = Math.abs(v);
                                const sign = v < 0 ? '-' : '';
                                const str = absV >= 1000000 ? (absV/1000000).toFixed(1) + 'M' : (absV/1000).toFixed(0) + 'K';
                                return sign + str;
                            };
                            ctx.fillText(formatDiff(diff), x, y - 10);
                        }
                    });
                    ctx.restore();
                }
            }]
        });
        chartRegistry.register('pipeline-influx', chart);
    }

    // New: Pipeline % by Quarter Pie Chart
    const pieCtx = document.getElementById('pipeline-quarter-pie-chart');
    if (pieCtx) {
        chartRegistry.destroyTag('pipeline-quarter-pie');
        const qLabels = ['Q1', 'Q2', 'Q3', 'Q4'];
        const qValues = qLabels.map(q => {
            const qData = stats.pipelineByQuarter[q];
            if (!qData) return 0;
            // Sum up values from countries in that quarter
            return Object.values(qData.countries).reduce((acc, curr) => acc + curr.amount, 0);
        });

        const totalValue = qValues.reduce((a, b) => a + b, 0);

        const pieChart = new Chart(pieCtx, {
            type: 'pie',
            data: {
                labels: qLabels,
                datasets: [{
                    data: qValues,
                    backgroundColor: [
                        '#10B981', // green-500
                        '#3B82F6', // blue-500
                        '#F59E0B', // amber-500
                        '#EF4444'  // red-500
                    ],
                    borderWidth: 2,
                    borderColor: '#ffffff',
                    hoverOffset: 15
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            usePointStyle: true,
                            boxWidth: 8,
                            padding: 15,
                            font: { size: 11, family: "'Inter', sans-serif", weight: '600' },
                            generateLabels: (chart) => {
                                const data = chart.data;
                                if (data.labels.length && data.datasets.length) {
                                    return data.labels.map((label, i) => {
                                        const value = data.datasets[0].data[i];
                                        const percentage = totalValue > 0 ? ((value / totalValue) * 100).toFixed(1) : 0;
                                        return {
                                            text: `${label}: ${percentage}%`,
                                            fillStyle: data.datasets[0].backgroundColor[i],
                                            hidden: isNaN(data.datasets[0].data[i]) || chart.getDatasetMeta(0).data[i].hidden,
                                            index: i
                                        };
                                    });
                                }
                                return [];
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: '#FFFFFF',
                        titleColor: '#111827',
                        bodyColor: '#374151',
                        borderColor: '#E5E7EB',
                        borderWidth: 1,
                        padding: 12,
                        cornerRadius: 8,
                        titleFont: { size: 13, weight: 'bold' },
                        bodyFont: { size: 12 },
                        callbacks: {
                            label: function (context) {
                                const value = context.raw;
                                const percentage = totalValue > 0 ? ((value / totalValue) * 100).toFixed(1) : 0;
                                return ` US$ ${formatCurrency(value)} (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
        chartRegistry.register('pipeline-quarter-pie', pieChart);
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
            options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { legend: { display: false } } },
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
            options: { responsive: true, maintainAspectRatio: false, cutout: '65%', plugins: { legend: { display: false } } },
            plugins: [{
                id: 'doughnutLabelAging',
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

    const ctxB = document.getElementById('poc-bottleneck-chart');
    if (ctxB && stats.runningList.length > 0) {
        chartRegistry.register('poc-bottleneck', new Chart(ctxB, {
            type: 'bar',
            data: { 
                labels: stats.runningList.slice(0, 10).map(r => r.name), 
                datasets: [{ 
                    label: 'Working Days',
                    data: stats.runningList.slice(0, 10).map(r => r.days), 
                    backgroundColor: '#007AFF',
                    borderRadius: 4
                }] 
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { 
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => ` ${ctx.parsed.y} Working Days`
                        }
                    }
                },
                scales: {
                    x: { ticks: { font: { size: 9 }, maxRotation: 45, minRotation: 45 } },
                    y: { beginAtZero: true, title: { display: true, text: 'Days', font: { size: 10 } } }
                }
            }
        }));
    }

    const ctxI = document.getElementById('poc-industry-chart');
    if (ctxI && stats.sortedIndustry.length > 0) {
        const colors = ['#6366f1', '#0ea5e9', '#10b981', '#f59e0b', '#ef4444', '#a855f7', '#ec4899', '#14b8a6', '#f97316', '#84cc16'];
        chartRegistry.register('poc-industry', new Chart(ctxI, {
            type: 'bar',
            data: { 
                labels: stats.sortedIndustry.slice(0, 10).map(i => i.name), 
                datasets: [{ 
                    data: stats.sortedIndustry.slice(0, 10).map(i => i.val), 
                    backgroundColor: colors,
                    borderRadius: 4
                }] 
            },
            options: { 
                indexAxis: 'y', 
                responsive: true, 
                maintainAspectRatio: false,
                plugins: { 
                    legend: { display: false } ,
                    tooltip: {
                        callbacks: {
                            label: (ctx) => ` $${formatCurrency(ctx.parsed.x)}`
                        }
                    }
                },
                scales: {
                    x: {
                        ticks: { display: false },
                        grid: { display: false },
                        border: { display: false }
                    },
                    y: {
                        grid: { display: false },
                        border: { display: false },
                        ticks: { 
                            font: { size: 10 }, // Slightly reduced font size
                            padding: 6,
                            autoSkip: false,
                            callback: function(val) {
                                let label = this.getLabelForValue(val) || '';
                                if (label.length > 15) {
                                    // Split long labels to prevent truncation
                                    let words = label.split(' ');
                                    let lines = [];
                                    let currentLine = '';
                                    words.forEach(word => {
                                        if (currentLine.length + word.length > 16) {
                                            if (currentLine) lines.push(currentLine);
                                            currentLine = word;
                                        } else {
                                            currentLine += (currentLine.length ? ' ' : '') + word;
                                        }
                                    });
                                    if (currentLine) lines.push(currentLine);
                                    return lines;
                                }
                                return label;
                            }
                        }
                    }
                }
            }
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
                plugins: {
                    tooltip: {
                        callbacks: {
                            title: function(context) {
                                const idx = context[0].dataIndex;
                                const item = stats.comparisonData[idx];
                                return [item.eventName, 'Date: ' + item.name];
                            }
                        }
                    }
                },
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

/* ═══ TCV vs ARR Charts ═══ */

/**
 * Initialize the TCV vs ARR grouped horizontal bar chart.
 * @param {Object} stats - Output from getTcvArrStats
 */
export function initTcvArrChart(stats) {
    if (!stats || !stats.items || stats.items.length === 0) return;
    chartRegistry.destroyTag('tcvarr');

    const displayItems = stats.items.slice(0, 25);
    const labels = displayItems.map(i => _truncateLabel(i.name, 28));
    const tcvData = displayItems.map(i => i.tcv);
    const arrData = displayItems.map(i => i.arr);

    const chartHeight = Math.max(400, displayItems.length * 38);
    const container = document.getElementById('tcvarr-chart-container');
    if (container) container.style.height = chartHeight + 'px';

    const ctx = document.getElementById('tcvarr-bar-chart');
    if (!ctx) return;

    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: 'KOR TCV (USD)',
                    data: tcvData,
                    backgroundColor: 'rgba(30, 64, 175, 0.78)',
                    borderColor: '#1e40af',
                    borderWidth: 1,
                    borderRadius: 4,
                    barPercentage: 0.75,
                    categoryPercentage: 0.7
                },
                {
                    label: 'KOR ARR (USD)',
                    data: arrData,
                    backgroundColor: 'rgba(34, 197, 94, 0.72)',
                    borderColor: '#16a34a',
                    borderWidth: 1,
                    borderRadius: 4,
                    barPercentage: 0.75,
                    categoryPercentage: 0.7
                }
            ]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 20,
                        font: { size: 12, family: "'Inter', sans-serif", weight: '600' }
                    }
                },
                tooltip: {
                    backgroundColor: '#FFFFFF',
                    titleColor: '#111827',
                    bodyColor: '#374151',
                    borderColor: '#E5E7EB',
                    borderWidth: 1,
                    padding: 14,
                    cornerRadius: 10,
                    titleFont: { size: 13, weight: 'bold' },
                    bodyFont: { size: 12 },
                    callbacks: {
                        title(ctx) {
                            const idx = ctx[0].dataIndex;
                            return displayItems[idx].name;
                        },
                        label(ctx) {
                            return ` ${ctx.dataset.label}: $${formatCurrency(ctx.raw)}`;
                        },
                        afterBody(ctx) {
                            const idx = ctx[0].dataIndex;
                            const item = displayItems[idx];
                            const lines = [''];
                            lines.push(`Gap (TCV - ARR): $${formatCurrency(item.gap)}`);
                            lines.push(`Gap %: ${item.gapPct.toFixed(1)}%`);
                            lines.push(item.isPerpetual ? '⚡ Perpetual License (one-time)' : '🔄 Recurring Revenue');
                            return lines;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(0,0,0,0.04)' },
                    ticks: {
                        color: '#6B7280',
                        font: { size: 10 },
                        callback: v => '$' + formatCurrency(v)
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: {
                        color: '#1E293B',
                        font: { size: 11, weight: '500', family: "'Inter', sans-serif" },
                        padding: 8
                    }
                }
            }
        },
        plugins: [{
            id: 'tcvArrDataLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                ctx.save();
                chart.data.datasets.forEach((dataset, dsIdx) => {
                    const meta = chart.getDatasetMeta(dsIdx);
                    meta.data.forEach((bar, idx) => {
                        const value = dataset.data[idx];
                        if (value > 0) {
                            const { x, y } = bar.tooltipPosition();
                            ctx.fillStyle = dsIdx === 0 ? '#1e40af' : '#16a34a';
                            ctx.font = 'bold 9px "Inter", sans-serif';
                            ctx.textAlign = 'left';
                            ctx.textBaseline = 'middle';
                            const label = '$' + _shortCurrency(value);
                            ctx.fillText(label, x + 6, y);
                        }
                    });
                });
                ctx.restore();
            }
        }]
    });

    chartRegistry.register('tcvarr-main', chart);
}

/**
 * Truncate label for chart axis.
 * @param {string} label
 * @param {number} max
 * @returns {string}
 */
function _truncateLabel(label, max) {
    return label.length > max ? label.substring(0, max - 1) + '…' : label;
}

/**
 * Short currency format for data labels (e.g., 1.2M, 450K).
 * @param {number} val
 * @returns {string}
 */
function _shortCurrency(val) {
    if (val >= 1000000) return (val / 1000000).toFixed(1) + 'M';
    if (val >= 1000) return (val / 1000).toFixed(0) + 'K';
    return String(Math.round(val));
}