import { parseExcelDate, parseCurrency, formatCurrency } from './utils.js';

/**
 * training.js — Logic for the Staff Training & Development dashboard.
 * @module training
 */

const TRAINING_PALETTE = [
    '#0ea5e9', // Blue
    '#6366f1', // Indigo
    '#f43f5e', // Rose (Distinct from Indigo)
    '#10b981', // Emerald (Distinct from Blue)
    '#f59e0b', // Amber
    '#a855f7', // Purple
    '#06b6d4', // Cyan
    '#ec4899'  // Pink
];

/**
 * Gets a consistent color based on the index.
 */
function getStaffColor(index) {
    return TRAINING_PALETTE[index % TRAINING_PALETTE.length];
}

/**
 * Find a key in a row using case/whitespace-insensitive matching.
 */
function findRowKey(row, ...candidates) {
    if (!row) return null;
    const keys = Object.keys(row);
    const norm = s => String(s).trim().toLowerCase().replace(/\s+/g, ' ');
    for (const cand of candidates) {
        const target = norm(cand);
        const exact = keys.find(k => norm(k) === target);
        if (exact) return exact;
    }
    for (const cand of candidates) {
        const target = norm(cand);
        const partial = keys.find(k => norm(k).includes(target));
        if (partial) return partial;
    }
    return null;
}

/**
 * Resolve the column keys used by the Staff Training sheet.
 */
function resolveTrainingKeys(trainingData) {
    const sample = trainingData && trainingData.length > 0 ? trainingData[0] : null;
    return {
        name: findRowKey(sample, 'Name', 'Staff', 'Staff Name', '이름'),
        hours: findRowKey(sample, 'Total Training Hours', 'Training Hours', 'Hours'),
        completion: findRowKey(sample, 'Completion Date', 'End Date', 'Date'),
        start: findRowKey(sample, 'Start Date', 'Date'),
        title: findRowKey(sample, 'Training Title', 'Title')
    };
}

/**
 * Pure function to extract staff list.
 */
function getStaffList(workbookData, trainingData, keys) {
    if (trainingData && trainingData.length > 0 && keys.name) {
        const names = [...new Set(
            trainingData.map(r => String(r[keys.name] || '').trim()).filter(Boolean)
        )];
        if (names.length > 0) return names;
    }

    const nameSheetKey = Object.keys(workbookData).find(k =>
        k.trim().toUpperCase() === 'NAME' || k.includes('이름')
    );
    const nameSheet = nameSheetKey ? workbookData[nameSheetKey] : null;
    if (nameSheet && nameSheet.length > 0) {
        const sheetKeys = Object.keys(nameSheet[0]);
        const nameKey = findRowKey(nameSheet[0], 'Name', 'Staff') || sheetKeys[0];
        return [...new Set(nameSheet.map(r => String(r[nameKey] || '').trim()).filter(Boolean))];
    }
    return [];
}

/**
 * Pure function to process training statistics.
 */
export function calculateTrainingStats(trainingData, staffNames, year = new Date().getFullYear(), keys = null) {
    const k = keys || resolveTrainingKeys(trainingData);
    const stats = staffNames.reduce((acc, name) => {
        acc[name] = { name, monthlyHours: Array(12).fill(0), totalHours: 0, avgPerMonth: 0 };
        return acc;
    }, {});

    trainingData.forEach(row => {
        const rowName = String((k.name && row[k.name]) || '').trim();
        if (!stats[rowName]) return;

        const hours = parseCurrency(k.hours ? row[k.hours] : 0);
        const date = parseExcelDate(
            (k.completion && row[k.completion]) || (k.start && row[k.start])
        );

        if (date && !isNaN(date.getTime()) && date.getFullYear() === year) {
            stats[rowName].monthlyHours[date.getMonth()] += hours;
            stats[rowName].totalHours += hours;
        }
    });

    staffNames.forEach(name => {
        stats[name].avgPerMonth = stats[name].totalHours / 12;
    });

    return stats;
}

/**
 * Renders the top metric cards with unique colors.
 */
function renderMetricCards(displayTop, stats, staffNames) {
    return displayTop.map((name) => {
        const s = stats[name] || { totalHours: 0, avgPerMonth: 0 };
        const color = getStaffColor(staffNames.indexOf(name));
        return `
            <div class="stat-card" style="background: white; padding: 24px; border-radius: 20px; border-left: 6px solid ${color}; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                    <span style="font-size: 0.75rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em;">${name}</span>
                    <i class="fa-solid fa-user-graduate" style="color: ${color}; opacity: 0.2;"></i>
                </div>
                <div style="display: flex; align-items: baseline; gap: 8px;">
                    <h1 style="font-size: 2.2rem; font-weight: 800; color: #1e293b;">${s.totalHours} <span style="font-size: 1rem; font-weight: 500; color: #64748b;">hrs</span></h1>
                    <span style="font-size: 0.8rem; color: #94a3b8; font-weight: 500;">Avg: ${s.avgPerMonth.toFixed(1)}/mo</span>
                </div>
            </div>
        `;
    }).join('');
}

/**
 * Populates the Data View table (#data-table) with Staff Training sheet rows.
 * @param {Object[]} trainingData - Raw rows from the 'Staff Training' sheet
 */
function renderTrainingDataView(trainingData) {
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');
    const dataSection = document.querySelector('.data-section');
    const emptyState = document.getElementById('empty-state');
    const dataTable = document.getElementById('data-table');

    if (!tableHead || !tableBody) return;

    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (trainingData.length === 0) {
        if (dataSection) dataSection.classList.add('hidden');
        if (emptyState) emptyState.classList.remove('hidden');
        if (dataTable) dataTable.classList.add('hidden');
        return;
    }

    const headers = Object.keys(trainingData[0]).filter(k => !k.startsWith('__EMPTY'));

    const trHead = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.innerText = h;
        trHead.appendChild(th);
    });
    tableHead.appendChild(trHead);

    trainingData.slice(0, 100).forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(h => {
            const td = document.createElement('td');
            td.innerText = row[h] != null ? row[h] : '';
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    if (dataSection) dataSection.classList.remove('hidden');
    if (emptyState) emptyState.classList.add('hidden');
    if (dataTable) dataTable.classList.remove('hidden');
}

/**
 * Renders the monthly input table.
 */
function renderMonthlyTable(staffNames, stats, months, year) {
    return `
        <div class="stat-card" style="grid-column: 1 / -1; padding: 32px; background: white; border-radius: 24px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05);">
            <div style="margin-bottom: 30px;">
                <h2 style="font-size: 1.25rem; font-weight: 800; color: #1e293b; margin-bottom: 4px;">Monthly Training Input</h2>
                <p style="font-size: 0.85rem; color: #64748b;">Aggregate tracking of training hours per month for ${year}</p>
            </div>
            <div style="overflow-x: auto; margin: -10px;">
                <table style="width: 100%; border-collapse: separate; border-spacing: 0;">
                    <thead>
                        <tr style="background: #f8fafc;">
                            <th style="padding: 16px; text-align: left; font-size: 0.7rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.1em; border-bottom: 2px solid #e2e8f0; border-top-left-radius: 12px; width: 150px;">Staff</th>
                            ${months.map((m, i) => `<th style="padding: 16px; text-align: center; font-size: 0.7rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.1em; border-bottom: 2px solid #e2e8f0; ${i === 11 ? 'border-top-right-radius: 12px;' : ''}">${m}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${staffNames.map((name, idx) => {
                            const s = stats[name];
                            return `
                                <tr style="background: ${idx % 2 === 0 ? 'white' : '#fcfdfe'}; transition: background 0.2s;">
                                    <td style="padding: 16px; font-weight: 700; color: #1e293b; border-bottom: 1px solid #f1f5f9;">
                                        <div style="display: flex; align-items: center; gap: 10px;">
                                            <div style="width: 24px; height: 24px; border-radius: 50%; background: #eef2ff; color: #6366f1; display: flex; align-items: center; justify-content: center; font-size: 0.7rem;">${name.charAt(0)}</div>
                                            ${name}
                                        </div>
                                    </td>
                                    ${s.monthlyHours.map(hrs => `
                                        <td style="padding: 16px; text-align: center; border-bottom: 1px solid #f1f5f9;">
                                            <div style="padding: 6px; border-radius: 8px; font-weight: 500; color: ${hrs > 0 ? '#1e293b' : '#cbd5e1'}; background: ${hrs > 0 ? '#f0f9ff' : 'transparent'}; border: 1px solid ${hrs > 0 ? '#bae6fd' : '#f1f5f9'}; width: 45px; margin: 0 auto; transition: all 0.2s;">
                                                ${hrs}
                                            </div>
                                        </td>
                                    `).join('')}
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

/**
 * Main function to select and render the training view.
 */
export function selectTrainingView(setCurrentTab, workbookData) {
    setCurrentTab('TRAINING_VIEW');
    const titleEl = document.getElementById('current-tab-title');
    if (titleEl) titleEl.innerText = 'Staff Training & Development';

    // UI basic updates
    document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
    document.querySelector('.training-tab')?.classList.add('active');
    
    const metricsGrid = document.getElementById('tab-metrics-grid');
    metricsGrid.innerHTML = '';
    metricsGrid.classList.remove('hidden');

    const trainingSheetKey = Object.keys(workbookData).find(k => {
        const norm = k.trim().toLowerCase();
        return norm === 'staff training' || norm === 'training';
    }) || 'Staff Training';
    const trainingData = workbookData[trainingSheetKey] || [];
    const trainingKeys = resolveTrainingKeys(trainingData);
    const staffNames = getStaffList(workbookData, trainingData, trainingKeys);

    if (staffNames.length === 0) {
        const availableSheets = Object.keys(workbookData);
        const sheetExists = availableSheets.some(k => k.trim().toLowerCase() === 'staff training');
        const cols = trainingData.length > 0 ? Object.keys(trainingData[0]).filter(k => !k.startsWith('__EMPTY')) : [];
        metricsGrid.innerHTML = `
            <div style="grid-column: 1 / -1; padding: 32px; background: white; border-radius: 16px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <h2 style="font-size: 1.25rem; font-weight: 800; color: #1e293b; margin-bottom: 12px;">Staff training data not loaded</h2>
                <p style="color: #475569; margin-bottom: 16px;">
                    ${sheetExists
                        ? `"Staff Training" 시트는 발견됐지만 직원 이름을 인식하지 못했어요. 아래 컬럼 중 이름 컬럼이 있는지 확인해주세요.`
                        : `현재 불러온 워크북에 "Staff Training" 시트가 없습니다. Google Drive 파일에 시트가 추가됐는지, 그리고 우측 상단의 새로고침 버튼을 눌렀는지 확인해주세요.`}
                </p>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
                    <div style="padding: 16px; background: #f8fafc; border-radius: 12px;">
                        <p style="font-size: 0.7rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 8px;">Sheets in workbook</p>
                        <div style="font-size: 0.85rem; color: #334155; line-height: 1.6;">${availableSheets.length ? availableSheets.map(s => `<div>• ${s}</div>`).join('') : '<em>(none)</em>'}</div>
                    </div>
                    <div style="padding: 16px; background: #f8fafc; border-radius: 12px;">
                        <p style="font-size: 0.7rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 8px;">Staff Training columns (rows: ${trainingData.length})</p>
                        <div style="font-size: 0.85rem; color: #334155; line-height: 1.6;">${cols.length ? cols.map(c => `<div>• ${c}</div>`).join('') : '<em>(no columns / no rows)</em>'}</div>
                    </div>
                </div>
            </div>
        `;
        renderTrainingDataView(trainingData);
        return;
    }

    const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const trainingYear = new Date().getFullYear();
    const stats = calculateTrainingStats(trainingData, staffNames, trainingYear, trainingKeys);

    // All-time total across all years
    let allTimeTotal = 0;
    trainingData.forEach(row => {
        const h = parseCurrency(trainingKeys.hours ? row[trainingKeys.hours] : 0);
        if (!isNaN(h)) allTimeTotal += h;
    });

    // Inject total card directly into DOM first
    const totalCard = document.createElement('div');
    totalCard.style.cssText = 'grid-column: 1 / -1; background: linear-gradient(135deg, #1e293b 0%, #334155 100%); padding: 28px 36px; border-radius: 20px; margin-bottom: 4px; box-shadow: 0 10px 30px rgba(15,23,42,0.3);';
    totalCard.innerHTML = `
        <p style="font-size: 0.7rem; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.12em; margin: 0 0 8px 0;">TOTAL ACCUMULATED TRAINING HOURS</p>
        <div style="display: flex; align-items: baseline; gap: 10px;">
            <span style="font-size: 3rem; font-weight: 900; color: #ffffff; line-height: 1;">${allTimeTotal % 1 === 0 ? allTimeTotal : allTimeTotal.toFixed(1)}</span>
            <span style="font-size: 1.2rem; font-weight: 600; color: #94a3b8;">hrs</span>
            <span style="font-size: 0.85rem; color: #64748b; margin-left: 8px;">(${staffNames.length} staff · avg ${(allTimeTotal / (staffNames.length || 1)).toFixed(1)} hrs/person)</span>
        </div>
    `;
    metricsGrid.appendChild(totalCard);

    // Determine Top Staff for cards
    const topStaff = ['Andy', 'Hady', 'Clarissa', 'Kamal'].filter(n => stats[n] || staffNames.includes(n));
    const displayTop = topStaff.length > 0 ? topStaff : staffNames.slice(0, 3);

    // Assembly of HTML
    let html = `<div style="grid-column: 1 / -1; display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px;">`;
    html += renderMetricCards(displayTop, stats, staffNames);
    html += `</div>`;
    html += renderMonthlyTable(staffNames, stats, months, trainingYear);
    
    // Bottom Section: Comparison Chart and Activity
    html += `
        <div style="grid-column: 1 / -1; margin-top: 24px; display: grid; grid-template-columns: 1fr 1fr; gap: 24px;">
            <div class="stat-card" style="background: white; padding: 24px; border-radius: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="font-size: 1rem; font-weight: 800; color: #1e293b; margin-bottom: 20px;">Training Hours Comparison</h3>
                <div style="height: 300px;"><canvas id="staff-hours-comparison"></canvas></div>
            </div>
            <div class="stat-card" style="background: white; padding: 24px; border-radius: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; flex-direction: column;">
                <h3 style="font-size: 1rem; font-weight: 800; color: #1e293b; margin-bottom: 20px;">Recent Training Activity</h3>
                <div style="flex: 1; overflow-y: auto; max-height: 300px;">
                    <table style="width: 100%; border-collapse: collapse;">
                        <thead>
                            <tr style="text-align: left; border-bottom: 2px solid #f1f5f9;">
                                <th style="padding: 10px; font-size: 0.7rem; color: #94a3b8; text-transform: uppercase;">Staff</th>
                                <th style="padding: 10px; font-size: 0.7rem; color: #94a3b8; text-transform: uppercase;">Title</th>
                                <th style="padding: 10px; font-size: 0.7rem; color: #94a3b8; text-align: right; text-transform: uppercase;">Hrs</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${trainingData.slice(0, 10).map(row => {
                                const nm = trainingKeys.name ? row[trainingKeys.name] : '';
                                const ttl = trainingKeys.title ? row[trainingKeys.title] : '';
                                const hrs = trainingKeys.hours ? row[trainingKeys.hours] : '';
                                return `
                                <tr style="border-bottom: 1px solid #f1f5f9;">
                                    <td style="padding: 10px; font-size: 0.8rem; font-weight: 600; color: #334155;">${nm || ''}</td>
                                    <td style="padding: 10px; font-size: 0.8rem; color: #64748b;">${ttl || 'Self Development'}</td>
                                    <td style="padding: 10px; font-size: 0.8rem; font-weight: 700; color: #10b981; text-align: right;">${hrs !== '' && hrs != null ? hrs + 'h' : ''}</td>
                                </tr>
                            `;
                            }).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;

    metricsGrid.innerHTML += html;

    // Initialize Charts
    setTimeout(() => {
        const ctx = document.getElementById('staff-hours-comparison');
        if (!ctx) return;

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: months,
                datasets: staffNames.slice(0, 8).map((name) => ({
                    label: name,
                    data: stats[name].monthlyHours,
                    backgroundColor: getStaffColor(staffNames.indexOf(name)),
                    borderRadius: 4,
                    barPercentage: 0.6
                }))
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: { beginAtZero: true, grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
                    x: { grid: { display: false }, ticks: { font: { size: 10 } } }
                },
                plugins: {
                    legend: { position: 'bottom', labels: { usePointStyle: true, font: { size: 10 } } }
                }
            }
        });
    }, 100);

    // Populate the Data View table with Staff Training sheet data
    renderTrainingDataView(trainingData);
}